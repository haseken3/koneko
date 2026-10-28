"""ko-NeKo の画面を持たない部分（ナレーションカウンター・IDチェッカー）。

Streamlit 版（ui.py・slide_script_check_ui.py）と、NeKo の白磁の Web 版（NeKo/web/koneko_api.py）の
両方がここを呼ぶ（2026-10-28 に ui.py・slide_script_check_ui.py から括り出した・写さない）。
既定・範囲・色の決め方・評価の数え方の正本はここ。画面の文言は各画面が持つ。

このモジュールは flat（slide_script_check_ui の `import koneko_core`）とパッケージ経由
（ui.py・NeKo の `from koneko import koneko_core`）の両方で読まれるので、narration_counter と
同じく自分の居るディレクトリを末尾に足してから flat import する。
"""

import sys
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

_HERE = str(Path(__file__).resolve().parent)
if _HERE not in sys.path:
    sys.path.append(_HERE)

# ─── API キー ───
LOCAL_KEY_PATH = Path.home() / ".config" / "koneko-idcheck" / "anthropic_api_key.txt"


def load_anthropic_key(secret=None):
    """Anthropic API キー。secret（Streamlit Secrets の値）→ ローカルファイル の順。無ければ None。"""
    if secret:
        return secret
    if LOCAL_KEY_PATH.exists():
        return LOCAL_KEY_PATH.read_text().strip()
    return None


# ─── ナレーションカウンター ───
# 読み上げ速度のスライダー（既定・範囲・刻み・説明）
CHARS_PER_MIN = {
    "label": "読み上げ速度（文字/分）",
    "min": 250,
    "max": 400,
    "value": 300,
    "step": 10,
    "help": "日本語ナレーションの目安: ゆっくり250〜速め400文字/分",
}


def heat_color(ratio: float) -> str:
    """文字数の割合（0〜1・いちばん多い枚が1）から、帯と表の棒の色を作る（薄い肌色→橙）。"""
    r = int(245 + (195 - 245) * ratio)
    g = int(230 + (90 - 230) * ratio)
    b = int(211 + (53 - 211) * ratio)
    return f"rgb({r},{g},{b})"


def heat_text_color(ratio: float) -> str:
    """帯の上の番号の色（濃い所は白・薄い所は茶）。"""
    return "#FFFFFF" if ratio > 0.5 else "#7A4A2A"


def max_char_count(slides) -> int:
    """割合の分母（いちばん多い枚の文字数・0 なら 1）。"""
    return max((s["char_count"] for s in slides), default=1) or 1


# ─── スライドと原稿のチェック（不一致・誤字脱字） ───
def check_estimate_minutes(total_slides: int) -> int:
    """実行前に見せる所要時間の目安（分）。1枚3秒・最低1分（104枚で約5分）。"""
    return max(1, round(total_slides * 3 / 60))


def run_slide_script_check(pptx_source, api_key, *, model=None, typo_model=None, progress=None) -> dict:
    """不一致・誤字脱字のチェックを走らせて結果を返す。全ブロックが失敗したら RuntimeError。

    progress(done, total) は点検したブロックの数。
    """
    from slide_script_check import extract_slides, is_total_failure, run_checks

    if hasattr(pptx_source, "seek"):
        pptx_source.seek(0)
    slides = extract_slides(pptx_source)
    overrides = {k: v for k, v in (("model", model), ("typo_model", typo_model)) if v}
    result = run_checks(slides=slides, api_key=api_key, progress=progress, **overrides)
    if is_total_failure(result):
        raise RuntimeError(
            f"全{len(result.get('errors', []))}件のブロックが失敗しました: "
            + "; ".join(result.get("errors", []))
        )
    return result


# ─── IDチェッカー（鈴木克明レイヤーモデル44項目） ───
GRADE_TO_PERSONA = {
    "1年次": "grade1_shin",
    "2年次": "grade2_mai",
    "3年次": "grade3_ken",
    "4年次": "grade4_aya",
}
GRADE_DEFAULT = "2年次"
GRADE_HELP = "評価対象の授業を受講する学年を選択してください。"
SCORE_MARK = {-1: "🔴", 0: "⚪", 1: "🟢"}


def load_idcheck():
    """idcheck/ のモジュールを読む。

    idcheck/ 配下（llm_evaluator 等）は Path(__file__) ベースで checklist.json / personas を
    解決するため、ディレクトリを sys.path に通すだけで import できる。
    """
    idcheck_dir = str(Path(_HERE) / "idcheck")
    if idcheck_dir not in sys.path:
        sys.path.insert(0, idcheck_dir)
    from docx_writer import build_docx
    from llm_evaluator import MODEL_OPUS, evaluate_all, llm_evaluable_only, load_checklist
    from pptx_reader import read_slides
    return SimpleNamespace(build_docx=build_docx, MODEL_OPUS=MODEL_OPUS, evaluate_all=evaluate_all,
                           llm_evaluable_only=llm_evaluable_only, load_checklist=load_checklist,
                           read_slides=read_slides)


def idcheck_items(checklist, item_filter):
    """評価する項目（item_filter が None なら全部）。"""
    items = checklist["items"]
    return [it for it in items if item_filter(it)] if item_filter else items


def idcheck_plan(n_items: int, model: str) -> dict:
    """「実行内容」の数（項目数・モデルの短い名・API コール数・所要時間の目安（分））。"""
    return {"items": n_items, "model": model.split("-")[1], "calls": n_items,
            "min_lo": n_items * 9 // 60, "min_hi": n_items * 12 // 60}


def idcheck_summary(results) -> dict:
    """評価の結果を数える（成功・失敗・スコアごと）。"""
    ok = sum(1 for r in results if r["result"]["ok"])
    scores = {-1: 0, 0: 0, 1: 0}
    for r in results:
        if r["result"]["ok"]:
            s = r["result"]["evaluation"]["score"]
            scores[s] = scores.get(s, 0) + 1
    return {"ok": ok, "total": len(results), "fail": len(results) - ok, "scores": scores}


def idcheck_download_name(source_label, now=None) -> str:
    """Word フィードバックのファイル名（<資料名>_IDフィードバック_<年月日_時分>.docx）。"""
    stem = Path(str(source_label)).stem or "idcheck"
    return f"{stem}_IDフィードバック_{(now or datetime.now()).strftime('%Y%m%d_%H%M')}.docx"
