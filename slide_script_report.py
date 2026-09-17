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
"""

import re
from datetime import datetime
from io import BytesIO

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

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


def _para(doc, text="", *, size=BODY_PT, bold=False, color=INK,
          space_before=0, space_after=4, indent=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.left_indent = Pt(indent)
    if text:
        _style(p.add_run(text), size=size, bold=bold, color=color)
    return p


def _labeled(doc, label, value, *, color=MUTED, indent=12):
    p = _para(doc, indent=indent)
    _style(p.add_run(f"{label}: "), bold=True, color=color)
    _style(p.add_run(value), color=INK)
    return p


def _bottom_border(paragraph, color="B8AE9E"):
    """段落の下に手書き用の罫線を1本引く。"""
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for key, value in (("val", "single"), ("sz", "6"), ("space", "6"), ("color", color)):
        bottom.set(qn(f"w:{key}"), value)
    borders.append(bottom)
    paragraph._p.get_or_add_pPr().append(borders)


def _apply_default_font(doc):
    """既定スタイルにも東アジアフォントを入れる（run を貼り忘れた箇所の保険）。"""
    style = doc.styles["Normal"]
    style.font.name = LATIN_FONT
    style.font.size = Pt(BODY_PT)
    style.element.rPr.get_or_add_rFonts().set(qn("w:eastAsia"), EASTASIA_FONT)


def _write_finding(doc, finding):
    kind = finding.get("kind", "")
    color = KIND_COLOR.get(kind, INK)
    head = _para(doc, space_before=8, space_after=2, indent=12)
    _style(head.add_run(f"● {KIND_LABEL.get(kind, kind)}"), bold=True, color=color)
    _style(head.add_run(f"　（AIの確度: {finding.get('confidence', '')}）"), color=MUTED)

    _labeled(doc, quote_label(finding), finding.get("quote", ""))
    if finding.get("counterpart"):
        _labeled(doc, counterpart_label(finding), finding["counterpart"])
    if finding.get("reason"):
        _labeled(doc, "気づいた理由", finding["reason"])
    if finding.get("suggestion"):
        _labeled(doc, "ご提案", finding["suggestion"])

    # □ は U+25A1（日本語書体が持つ字）。☐ U+2610 だと別フォントに落ちて字面が揃わない。
    # 書き込み欄の線は run の下線でなく段落の下罫線で引く（空白だけの run に下線を付けても
    # 線が描かれない閲覧環境がある＝実測: LibreOffice で PDF 化したとき線が消えた）。
    memo = _para(doc, "□ 対応した　　メモ:", color=MUTED, space_after=14, indent=12)
    _bottom_border(memo)


def _write_section(doc, title, findings, steps):
    _para(doc, title, size=SECTION_HEADING_PT, bold=True, space_before=20, space_after=6)
    for group in group_findings_by_step(findings, steps):
        _para(doc, group["heading"], size=STEP_HEADING_PT, bold=True,
              space_before=12, space_after=2)
        _para(doc, count_text(group), color=MUTED, space_after=4, indent=6)
        for slide in group["slides"]:
            title_text = f"　{slide['title']}" if slide["title"] else ""
            _para(doc, f"スライド {slide['slide_num']}{title_text}",
                  size=SLIDE_HEADING_PT, bold=True, space_before=10, space_after=2,
                  indent=6)
            for finding in slide["findings"]:
                _write_finding(doc, finding)


def build_docx(result, steps=None, pptx_filename="") -> BytesIO:
    """チェック結果を Word にまとめて BytesIO で返す（st.download_button にそのまま渡せる）。

    Args:
        result: run_checks の返値（consistency / typos / slides / models / errors）
        steps: group_slides_into_steps の返値。None なら1括り
        pptx_filename: 対象資料名（表紙に刷る）
    """
    doc = Document()
    _apply_default_font(doc)

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

    _para(doc, DISCLAIMER, color=MUTED, space_before=8, space_after=8)
    for message in result.get("errors") or []:
        _para(doc, f"※ {message}", color=MUTED, space_after=2)

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
