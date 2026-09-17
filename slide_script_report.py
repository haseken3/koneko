"""ko-NeKo — スライド↔原稿チェック結果の整形（ステップ括り／Word 出力）。

streamlit に依存しない純ロジック。**画面と Word が同じ `group_findings_by_step` で括る**
（片方だけ直る二重実装を作らない）。種類の文言・色もここが単一の正本で、画面は CSS に、
Word は RGBColor に変換して使う。

Word 側の決め事（先生が印刷して手元で直す前提）:
- 本文12pt・見出し14〜16pt。東アジアフォントを `w:eastAsia` で明示する
  （既定のままだと日本語が欧文フォント側の代替で出る個体がある）。
  ⚠ OOXML には「フォントA→無ければB」を1つの run に列挙する書き方が無い。
  ここで書けるのは第1候補だけで、無い環境での差し替えは閲覧側の代替に委ねる。
  フォント名は欧文表記で書く（「MS ゴシック」のような和名は照合されない環境がある）。
  実測: eastAsia に実在フォントを指定した .docx を LibreOffice で PDF 化すると、
  日本語がそのフォントで出る（指定が効いていることの確認）。未インストールの名前を
  指定した時だけ閲覧側の代替に落ちた。
- 断定しない。Word にも「可能性」「最終判断は先生」を刷る（画面と同じ約束）。
- 確度「低」は末尾の参考章に送る（本編を赤だらけにしない）。
- 所見1件＝1つの囲み（1行1セルの表）。左に種類色の縦帯・地は淡く・ページを跨がない。
  紙に落ちた時に「どこからどこまでが1件か」が目で分かることを、行数の節約より優先する。
- 章は改ページ、所見のあるステップも改ページ。所見0件のステップは流す
  （0件で改ページすると「ありませんでした」1行だけの紙が何枚も出る）。
"""

import re
from datetime import datetime
from io import BytesIO

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# 高齢の読み手が印刷して読む前提で、可読性を優先した順に選ぶ。BIZ UDGothic は
# ユニバーサルデザイン書体で Windows 10 1809 以降に同梱（app.py の CSS も既にこれを積んでいる）。
# 無い環境では閲覧側が代替する＝ここに書けるのは第1候補だけ。
EASTASIA_FONT = "BIZ UDGothic"
LATIN_FONT = "Arial"  # 欧文が明朝系に落ちるのを防ぐ（どの環境にもある sans を明示する）
BODY_PT = 12
SLIDE_HEADING_PT = 13
STEP_HEADING_PT = 14
SECTION_HEADING_PT = 16
TITLE_PT = 18

INK = "#2E2A22"
MUTED = "#5C5346"  # 白地の本文に使える濃さ（コントラスト比 7.5:1）
RULE = "B8AE9E"  # 手書き欄・見出し下の細い罫線（# なしで OOXML にそのまま渡す）
# スライド見出しの帯。意匠full 2026-09-18 指摘(5-b): 旧値EFEBE3は「相違」の
# 淡色地(F2ECE3)と実測でほぼ同色(差3,1,0)だった。帯は淡色地4種のどれとも
# 被らない濃さへ落として区別する（帯＝見出し、淡色地＝所見種類、の役割を分ける）。
BAND_FILL = "DCD5C4"

# 所見ブロックの地色は種類色を白へ寄せて作る。濃くすると本文の黒と競り、
# モノクロ印刷で灰に潰れて字が沈む。薄くすると「まとまり」が見えない。
BLOCK_TINT = 0.12
# スライド見出しの帯（_write_section の band、indent=8pt）と表の左端を揃える値。
# 意匠full 2026-09-18 指摘(5-a): 旧値2ptは実測で3.05mm左にずれていた
# （tblInd はセル内マージンぶんの補正が要らず、帯のindentと同じ値で揃う。
#  実測して直した＝ソース値だけで判断しない）。
BLOCK_INDENT_PT = 8

# 用紙は A4 固定（python-docx の既定は Letter＝日本のプリンタで拡縮される）。
# 左だけ綴じ代ぶん広い。
PAGE_W_MM, PAGE_H_MM = 210, 297
MARGIN_TB_MM = 20
MARGIN_LR_MM = 22
GUTTER_MM = 3
TEXT_W_MM = PAGE_W_MM - MARGIN_LR_MM * 2 - GUTTER_MM

# 種類ごとの文字ラベルと色（画面・Word 共通の正本）。断定しない言い回しで固定する。
KIND_LABEL = {
    "不足": "原稿にあってスライドに見当たらない可能性",
    "相違": "内容が食い違っている可能性",
    "過剰": "スライドにあって原稿に見当たらない可能性",
    "誤字脱字": "誤字・脱字の可能性",
}
KIND_COLOR = {
    "不足": "#A8492C",
    "相違": "#8F5E14",
    "過剰": "#4A5A68",
    "誤字脱字": "#3F5C77",
}

DEFAULT_CONFIDENCES = ("高", "中")
LOW_CONFIDENCE = "低"

ALL_SLIDES_HEADING = "全スライド（ステップ分割ができなかったため、スライド順に並べています）"
OUTSIDE_STEPS_HEADING = "ステップの範囲外のスライド"
NO_FINDINGS_TEXT = "気づいた点はありませんでした"
DISCLAIMER = (
    "この一覧は AI（Claude）が気づいた「可能性」を並べたものです。"
    "誤りを保証するものでも、すべての誤りを見つけるものでもありません。"
    "直すかどうかの最終的な判断は先生が行ってください。"
)


# ─────────────────────────────────────────────
# ステップ括り（画面・Word 共通）
# ─────────────────────────────────────────────
def step_heading(step) -> str:
    """「ステップ 2　16進数の基礎（スライド 5〜8）」の形の見出しを作る。"""
    start, end = step["slide_range"]
    label = (step.get("step_label") or "").strip()
    if re.fullmatch(r"step\s*\d+", label, flags=re.IGNORECASE):
        label = ""  # 連番フォールバックのラベルは「ステップ N」と二重になるので出さない
    span = f"スライド {start}〜{end}" if start != end else f"スライド {start}"
    head = f"ステップ {step['step_num']}" if step.get("step_num") else "ステップ"
    return f"{head}　{label}（{span}）" if label else f"{head}（{span}）"


def count_text(group) -> str:
    """件数行の文言（画面・Word 共通）。ステップ分割前は「このステップ」と言わない。"""
    count = len(group["findings"])
    if not count:
        return NO_FINDINGS_TEXT
    where = "このステップで" if group.get("step_num") else ""
    return f"{where}気づいた点: {count}件"


def _by_slide(findings) -> list:
    """所見をスライド単位に束ねる（枚数順）。"""
    order = []
    buckets = {}
    for f in findings:
        num = f["slide_num"]
        if num not in buckets:
            buckets[num] = {"slide_num": num, "title": f.get("slide_title") or "",
                            "findings": []}
            order.append(num)
        buckets[num]["findings"].append(f)
    return [buckets[n] for n in sorted(order)]


def group_findings_by_step(findings, steps=None) -> list:
    """所見をステップ単位（未分割なら1括り）に束ねる。

    どのステップの範囲にも入らない所見は捨てずに末尾の「範囲外」へ回す
    （境界を手で動かした直後など、steps が全枚を覆わないことがある）。

    Returns:
        list[dict]: heading / step_num / findings / slides（[{slide_num,title,findings}]）
    """
    ordered = sorted(findings or [], key=lambda f: f["slide_num"])
    if not steps:
        return [{"heading": ALL_SLIDES_HEADING, "step_num": None,
                 "findings": ordered, "slides": _by_slide(ordered)}]

    groups = []
    covered = set()
    for step in steps:
        start, end = step["slide_range"]
        picked = [f for f in ordered if start <= f["slide_num"] <= end]
        covered.update(f["slide_num"] for f in picked)
        groups.append({"heading": step_heading(step), "step_num": step.get("step_num"),
                       "findings": picked, "slides": _by_slide(picked)})
    orphans = [f for f in ordered if f["slide_num"] not in covered]
    if orphans:
        groups.append({"heading": OUTSIDE_STEPS_HEADING, "step_num": None,
                       "findings": orphans, "slides": _by_slide(orphans)})
    return groups


def split_by_confidence(findings) -> tuple:
    """(本編に出す所見, 参考へ送る所見) に分ける。"""
    main = [f for f in findings or [] if f.get("confidence") in DEFAULT_CONFIDENCES]
    low = [f for f in findings or [] if f.get("confidence") == LOW_CONFIDENCE]
    return main, low


def quote_label(finding) -> str:
    """原文がどちら側のものかの見出し語。要約になっている引用は正直にそう書く。"""
    if finding.get("where") == "notes":
        return "原稿" if finding.get("quote_verified") else "原稿の要点"
    return "スライド"


def counterpart_label(finding) -> str:
    return "スライド" if finding.get("where") == "notes" else "原稿"


# ─────────────────────────────────────────────
# Word 出力
# ─────────────────────────────────────────────
def _style(run, *, size=BODY_PT, bold=False, color=INK, underline=False):
    run.font.size = Pt(size)
    run.font.name = LATIN_FONT
    run.bold = bold
    run.underline = underline
    run.font.color.rgb = RGBColor.from_string(color.lstrip("#").upper())
    rpr = run._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), EASTASIA_FONT)
    return run


# OOXML の子要素はスキーマの順序どおりに並んでいないと閲覧側に無視される
# （罫線が出ない・網掛けが消える形で現れ、例外は出ない）。順序表を持っておく。
# 今実際に挿し込むのは pBdr/shd（pPr側）・tblW/tblInd/tblBorders（tblPr側）だけ
# だが、将来要素が増えた時に個別対応しないための全順序表（inspector 2026-09-18）。
_PPR_ORDER = (
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
_TBLPR_ORDER = (
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
    "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
    "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
    "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
    "w:tblPrChange",
)
_SIDES = ("top", "left", "bottom", "right")
_PBDR_ORDER = ("top", "left", "bottom", "right", "between", "bar")


def _el(tag, **attrs):
    element = OxmlElement(tag)
    for key, value in attrs.items():
        element.set(qn(f"w:{key}"), str(value))
    return element


def _insert_ordered(parent, element, order):
    tag = "w:" + element.tag.split("}")[-1]
    parent.insert_element_before(element, *order[order.index(tag) + 1:])


def _tbl_child(table, tag):
    """tblPr の子を取る（無ければ順序どおりに差し込む）。tblW は既に在る。"""
    tbl_pr = table._tbl.tblPr
    element = tbl_pr.find(qn(tag))
    if element is None:
        element = OxmlElement(tag)
        _insert_ordered(tbl_pr, element, _TBLPR_ORDER)
    return element


def _tint(color, ratio):
    """色を白へ寄せた淡い地色を作る（`#RRGGBB` → `RRGGBB`）。"""
    raw = color.lstrip("#")
    return "".join(f"{round(int(raw[i:i + 2], 16) * ratio + 255 * (1 - ratio)):02X}"
                   for i in (0, 2, 4))


def _para(container, text="", *, size=BODY_PT, bold=False, color=INK,
          space_before=0, space_after=4, indent=0):
    """段落を1つ足す。container は Document でも表のセルでもよい。"""
    p = container.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.left_indent = Pt(indent)
    if text:
        _style(p.add_run(text), size=size, bold=bold, color=color)
    return p


def _labeled(container, label, value, *, color=MUTED, indent=12):
    p = _para(container, indent=indent)
    _style(p.add_run(f"{label}: "), bold=True, color=color)
    _style(p.add_run(value), color=INK)
    return p


def _para_border(paragraph, **sides):
    """段落に罫線を引く。sides は `bottom=(色, 太さsz, 文字との間隔space)`。

    注意: 罫線の指定がまったく同じ段落が続くと、閲覧側はそれを1つの囲みとして
    まとめ、下罫線を最後の段落にしか描かない（手書き欄を2行にしたいのに1本しか
    出ない）。続けて線を引きたいときは `between` を併せて渡す。
    """
    borders = OxmlElement("w:pBdr")
    for name in _PBDR_ORDER:
        if name in sides:
            color, sz, space = sides[name]
            borders.append(_el(f"w:{name}", val="single", sz=sz, space=space,
                               color=color))
    _insert_ordered(paragraph._p.get_or_add_pPr(), borders, _PPR_ORDER)


def _para_shade(paragraph, fill):
    _insert_ordered(paragraph._p.get_or_add_pPr(),
                    _el("w:shd", val="clear", color="auto", fill=fill), _PPR_ORDER)


def _spacer(doc, *, space_after=10, mark_pt=2):
    """表と表の間に挟む段落。

    隣り合う表は Word が1つの表につなげてしまうので、所見ブロックの間には必ず
    段落が要る。ただし本文と同じ高さだと隙間が開きすぎるため、段落記号を小さく
    落として、間隔は space_after で作る。
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    rpr = OxmlElement("w:rPr")  # w:sz は half-point 単位
    rpr.append(_el("w:sz", val=mark_pt * 2))
    rpr.append(_el("w:szCs", val=mark_pt * 2))
    _insert_ordered(p._p.get_or_add_pPr(), rpr, _PPR_ORDER)
    return p


def _apply_default_font(doc):
    """既定スタイルにも東アジアフォントを入れる（run を貼り忘れた箇所の保険）。"""
    style = doc.styles["Normal"]
    style.font.name = LATIN_FONT
    style.font.size = Pt(BODY_PT)
    style.element.rPr.get_or_add_rFonts().set(qn("w:eastAsia"), EASTASIA_FONT)


def _field_run(paragraph, instr, *, placeholder="1"):
    """PAGE/NUMPAGES のような単純フィールドを段落に足す。

    `w:fldSimple` はキャッシュ済みの表示値(`placeholder`)を子の run に持てる。
    閲覧側がフィールドを再計算しない設定でも、開いた瞬間は placeholder が
    見える（0 や空白でなく「1」を入れておけば、更新されなくても違和感が薄い）。
    """
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), EASTASIA_FONT)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = placeholder
    run.append(t)
    fld.append(run)
    paragraph._p.append(fld)


def _setup_footer(doc):
    """ページ番号を中央に刷る（意匠full 2026-09-18 指摘6-b: 90頁の印刷物に
    ノンブルが1つも無く「45ページの件」が言えない）。"""
    footer = doc.sections[0].footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    _style(p.add_run("- "), size=9, color=MUTED)
    _field_run(p, "PAGE")
    _style(p.add_run(" / "), size=9, color=MUTED)
    _field_run(p, "NUMPAGES")
    _style(p.add_run(" -"), size=9, color=MUTED)


def _setup_page(doc):
    section = doc.sections[0]
    section.page_width = Mm(PAGE_W_MM)
    section.page_height = Mm(PAGE_H_MM)
    section.top_margin = Mm(MARGIN_TB_MM)
    section.bottom_margin = Mm(MARGIN_TB_MM)
    section.left_margin = Mm(MARGIN_LR_MM + GUTTER_MM)
    section.right_margin = Mm(MARGIN_LR_MM)


def _block_table(doc, *, bar_color, fill):
    """所見1件ぶんの囲み（1行1セルの表）を作って返す。

    左だけ種類色の太い縦帯、地は淡く、上下右は罫線なし。行に cantSplit を付けて
    所見がページを跨がないようにする（跨ぐと「原稿／ご提案」が別の紙に散る）。
    """
    width = Mm(TEXT_W_MM).twips - Pt(BLOCK_INDENT_PT).twips
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    tbl_w = _tbl_child(table, "w:tblW")
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(width))
    tbl_ind = _tbl_child(table, "w:tblInd")
    tbl_ind.set(qn("w:type"), "dxa")
    tbl_ind.set(qn("w:w"), str(Pt(BLOCK_INDENT_PT).twips))

    row = table.rows[0]
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))

    # 意匠full 2026-09-18 指摘(5-a)の実測沼: `cell.width=` は tcW だけを書き、
    # `tblGrid/gridCol` は既定値(6インチ)のまま残る python-docx の癖がある
    # （tblW と gridCol が食い違い、閲覧側は gridCol 基準で描いて右端が3.05mm
    # 短くなっていた）。`table.columns[0].width` で gridCol も同じ値に揃える。
    content_width = Mm(TEXT_W_MM) - Pt(BLOCK_INDENT_PT)
    table.columns[0].width = content_width
    cell = table.cell(0, 0)
    cell.width = content_width
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for name in _SIDES:
        borders.append(_el(f"w:{name}", val="single", sz=24, space=0, color=bar_color)
                       if name == "left" else _el(f"w:{name}", val="nil"))
    tc_pr.append(borders)
    tc_pr.append(_el("w:shd", val="clear", color="auto", fill=fill))
    margins = OxmlElement("w:tcMar")
    for name, amount in (("top", 110), ("left", 170), ("bottom", 150), ("right", 130)):
        margins.append(_el(f"w:{name}", w=amount, type="dxa"))
    tc_pr.append(margins)
    return table


def _write_finding(doc, finding):
    kind = finding.get("kind", "")
    color = KIND_COLOR.get(kind, INK)
    table = _block_table(doc, bar_color=color.lstrip("#").upper(),
                         fill=_tint(color, BLOCK_TINT))
    cell = table.cell(0, 0)
    placeholder = cell.paragraphs[0]

    # 意匠full 2026-09-18 指摘(6-a): 同じスライドの2件目以降が改ページ先頭に
    # 来ると、直上のスライド帯が前頁に残り「どのスライドの話か」が分からない
    # 紙になる（実測90頁中23頁）。●行自体にスライド番号を持たせて、帯が
    # 見えない頁でも所見単体で分かるようにする。
    head = _para(cell, space_after=3)
    _style(head.add_run(f"● {KIND_LABEL.get(kind, kind)}"), bold=True, color=color)
    slide_tag = f"スライド{finding['slide_num']}・" if finding.get("slide_num") else ""
    _style(head.add_run(f"　（{slide_tag}AIの確度: {finding.get('confidence', '')}）"), color=MUTED)

    _labeled(cell, quote_label(finding), finding.get("quote", ""), indent=0)
    if finding.get("counterpart"):
        _labeled(cell, counterpart_label(finding), finding["counterpart"], indent=0)
    if finding.get("reason"):
        _labeled(cell, "気づいた理由", finding["reason"], indent=0)
    if finding.get("suggestion"):
        _labeled(cell, "ご提案", finding["suggestion"], indent=0)

    # □ は U+25A1（日本語書体が持つ字）。☐ U+2610 だと別フォントに落ちて字面が揃わない。
    # 書き込み欄の線は run の下線でなく段落の下罫線で引く（空白だけの run に下線を付けても
    # 線が描かれない閲覧環境がある＝実測: LibreOffice で PDF 化したとき線が消えた）。
    # 意匠full 2026-09-18 指摘(4-b): 旧 space=12pt だと罫線間13.5mm(実測)で
    # 手書き欄として広すぎ、ブロックが伸びて1頁2件止まりの主因になっていた。
    # 6pt に詰めて手書きしやすい幅(実測後に確認)へ。
    memo = _para(cell, "□ 対応した　　メモ:", color=MUTED, space_before=6, space_after=0)
    write_line = _para(cell, space_after=0)
    for p in (memo, write_line):
        _para_border(p, bottom=(RULE, 6, 6), between=(RULE, 6, 6))

    # セルが最初から持っている空段落。先頭に余白が1行ぶん残るので外す。
    placeholder._p.getparent().remove(placeholder._p)
    _spacer(doc)


def _write_section(doc, title, findings, steps):
    p = _para(doc, title, size=SECTION_HEADING_PT, bold=True, space_after=10)
    p.paragraph_format.page_break_before = True
    p.paragraph_format.keep_with_next = True
    _para_border(p, bottom=(INK.lstrip("#"), 12, 8))

    # 意匠full 2026-09-18 指摘(4-c): 「0件ステップの次のステップは無条件で改ページ」
    # だと、0件ステップだけが載った紙の残りが丸ごと白紙になる（実測で頁の86%が
    # 白紙の例あり）。改ページは「これより前に所見のあるステップを1つ以上
    # 置いたか」で判定し、0件ステップの直後はそのまま同じ紙に続ける。
    seen_filled = False
    for group in group_findings_by_step(findings, steps):
        has_findings = bool(group["findings"])
        page_break = has_findings and seen_filled
        seen_filled = seen_filled or has_findings
        step = _para(doc, group["heading"], size=STEP_HEADING_PT, bold=True,
                     space_before=0 if page_break else 22, space_after=4)
        step.paragraph_format.page_break_before = page_break
        step.paragraph_format.keep_with_next = True
        _para_border(step, bottom=(RULE, 6, 6))

        count = _para(doc, count_text(group), color=MUTED, space_after=6, indent=4)
        count.paragraph_format.keep_with_next = bool(group["findings"])

        for slide in group["slides"]:
            title_text = f"　{slide['title']}" if slide["title"] else ""
            band = _para(doc, f"スライド {slide['slide_num']}{title_text}",
                         size=SLIDE_HEADING_PT, bold=True, space_before=14,
                         space_after=8, indent=8)
            band.paragraph_format.keep_with_next = True
            _para_shade(band, BAND_FILL)
            _para_border(band, left=(MUTED.lstrip("#"), 12, 6))
            for finding in slide["findings"]:
                _write_finding(doc, finding)


def _cell_text(cell, text, *, bold=False, center=False, size=10):
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _style(p.add_run(text), size=size, bold=bold)
    return p


def _write_step_summary(doc, consistency, typos, steps):
    """表紙に「ステップ別の件数」を1つ。分割前（steps なし）は1行なので出さない。

    ステップ名は `step_heading` から取る（対象スライドも入っている）＝画面・本文と
    同じ正本で、ここで括りのロジックを二重に持たない。
    """
    if not steps:
        return
    by_step = group_findings_by_step(consistency, steps)
    typo_counts = {g["heading"]: len(g["findings"])
                   for g in group_findings_by_step(typos, steps)}
    counts = {g["heading"]: len(g["findings"]) for g in by_step}
    headings = [g["heading"] for g in by_step]
    headings += [h for h in typo_counts if h not in counts]

    _para(doc, "ステップ別の件数", size=SLIDE_HEADING_PT, bold=True,
          space_before=16, space_after=6)
    table = doc.add_table(rows=len(headings) + 1, cols=3)
    table.autofit = False  # 所見ブロックの表と同じ流儀（dxaで明示、Wordの自動列幅に譲らない）
    tbl_w = _tbl_child(table, "w:tblW")
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(Mm(TEXT_W_MM).twips))
    borders = OxmlElement("w:tblBorders")
    for name in (*_SIDES, "insideH", "insideV"):
        borders.append(_el(f"w:{name}", val="single", sz=4, space=0, color=RULE))
    _insert_ordered(table._tbl.tblPr, borders, _TBLPR_ORDER)
    # 意匠full 2026-09-18 指摘(8): 旧配分(107/28/28mm)だと1列目が狭く、実測で
    # 「スライド」が語の内側で改行された。数値2列を22mmへ削って1列目へ回す。
    for index, width in enumerate((TEXT_W_MM - 44, 22, 22)):
        for row in table.rows:
            row.cells[index].width = Mm(width)

    for cell, label in zip(table.rows[0].cells,
                           ("ステップ（対象スライド）", "不一致", "誤字・脱字")):
        _para_shade(_cell_text(cell, label, bold=True,
                               center=label != "ステップ（対象スライド）"), BAND_FILL)
    for row, heading in zip(table.rows[1:], headings):
        _cell_text(row.cells[0], heading)
        _cell_text(row.cells[1], str(counts.get(heading, 0)), center=True)
        _cell_text(row.cells[2], str(typo_counts.get(heading, 0)), center=True)
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def build_docx(result, steps=None, pptx_filename="") -> BytesIO:
    """チェック結果を Word にまとめて BytesIO で返す（st.download_button にそのまま渡せる）。

    Args:
        result: run_checks の返値（consistency / typos / slides / models / errors）
        steps: group_slides_into_steps の返値。None なら1括り
        pptx_filename: 対象資料名（表紙に刷る）
    """
    doc = Document()
    _apply_default_font(doc)
    _setup_page(doc)
    _setup_footer(doc)

    consistency, low_consistency = split_by_confidence(result.get("consistency"))
    typos, low_typos = split_by_confidence(result.get("typos"))
    models = result.get("models") or {}

    _para(doc, "スライドと原稿のチェック結果", size=TITLE_PT, bold=True, space_after=10)
    _labeled(doc, "対象資料", pptx_filename or "（ファイル名なし）", indent=0)
    _labeled(doc, "作成日時", datetime.now().strftime("%Y年%m月%d日 %H:%M"), indent=0)
    _labeled(doc, "スライド枚数", f"{result.get('slides', 0)}枚", indent=0)
    _labeled(doc, "気づいた点",
             f"内容の不一致 {len(consistency)}件 / 誤字・脱字 {len(typos)}件"
             f"（ほかに参考 {len(low_consistency) + len(low_typos)}件）", indent=0)
    if models:
        _labeled(doc, "使用したAIモデル",
                 f"不一致 {models.get('consistency', '')} / "
                 f"誤字脱字 {models.get('typos', '')}", indent=0)

    _para(doc, DISCLAIMER, color=MUTED, space_before=10, space_after=8)
    for message in result.get("errors") or []:
        _para(doc, f"※ {message}", color=MUTED, space_after=2)

    _write_step_summary(doc, consistency, typos, steps)

    _write_section(doc, "内容の不一致の可能性", consistency, steps)
    _write_section(doc, "誤字・脱字の可能性", typos, steps)

    low = low_consistency + low_typos
    if low:
        _write_section(doc, "参考: 確度が低い気づき", low, steps)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def download_filename(pptx_filename: str, *, now=None) -> str:
    stem = re.sub(r"\.pptx$", "", pptx_filename or "", flags=re.IGNORECASE) or "チェック結果"
    stamp = (now or datetime.now()).strftime("%Y%m%d_%H%M")
    return f"{stem}_チェック結果_{stamp}.docx"
