"""ko-NeKo の画面を持たない部分（授業資料チェッカー＝ナレーションカウンター＋スライドと原稿のチェック）。

IDチェッカーは NeKo の画面から廃止した（ケンタ裁定 2026-10-28）ので、その部分はここから外した。

Streamlit 版（ui.py・slide_script_check_ui.py）と、NeKo の白磁の Web 版（NeKo/web/koneko_api.py）の
両方がここを呼ぶ（2026-10-28 に ui.py・slide_script_check_ui.py から括り出した・写さない）。
既定・範囲・色の決め方の正本はここ。画面の文言は各画面が持つ。

このモジュールは flat（slide_script_check_ui の `import koneko_core`）とパッケージ経由
（ui.py・NeKo の `from koneko import koneko_core`）の両方で読まれるので、narration_counter と
同じく自分の居るディレクトリを末尾に足してから flat import する。
"""

import sys
from pathlib import Path

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
