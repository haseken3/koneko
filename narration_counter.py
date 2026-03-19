"""ko-NeKo — PPTXナレーション文字数カウント＋尺推定ロジック"""

import re
from pptx import Presentation

# ポーズ時間（VoiceSpace デフォルト設定）
PAUSE_COMMA = 0.3   # 「、」1個あたり（秒）
PAUSE_PERIOD = 1.0  # 「。」1個あたり（秒）


def _count_reading_chars(text: str) -> int:
    """読み上げ対象の文字数をカウント（空白・改行を除外）。"""
    return len(re.sub(r'\s', '', text))


def _count_pauses(text: str) -> float:
    """句読点によるポーズ時間を算出（秒）。"""
    commas = text.count('、')
    periods = text.count('。')
    return commas * PAUSE_COMMA + periods * PAUSE_PERIOD


def extract_slide_notes(pptx_file) -> list[dict]:
    """全スライドのノート欄テキスト・文字数を抽出する。

    Args:
        pptx_file: PPTXファイルパスまたはファイルライクオブジェクト

    Returns:
        list[dict]: 各スライドの情報
            slide_num, title, notes, char_count, pause_seconds
    """
    prs = Presentation(pptx_file)
    results = []

    for i, slide in enumerate(prs.slides, start=1):
        title = ""
        if slide.shapes.title:
            title = slide.shapes.title.text.strip()

        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()

        results.append({
            "slide_num": i,
            "title": title,
            "notes": notes,
            "char_count": _count_reading_chars(notes),
            "pause_seconds": _count_pauses(notes),
        })

    return results


def analyze_narration(pptx_file, chars_per_min: int = 320) -> dict:
    """PPTXのナレーション文字数を集計し、推定尺を算出する。

    Args:
        pptx_file: PPTXファイルパスまたはファイルライクオブジェクト
        chars_per_min: 1分あたりの読み上げ文字数（デフォルト320）

    Returns:
        dict: 集計結果
            total_slides: 総スライド数
            slides_with_notes: ナレーション付きスライド数
            total_chars: 合計文字数（空白・改行除外）
            total_pause_seconds: 句読点ポーズ合計（秒）
            estimated_seconds: 推定尺（秒）= 読み上げ時間 + ポーズ
            slides: スライド別データのリスト
    """
    slides = extract_slide_notes(pptx_file)

    total_chars = sum(s["char_count"] for s in slides)
    total_pause = sum(s["pause_seconds"] for s in slides)
    slides_with_notes = sum(1 for s in slides if s["char_count"] > 0)
    reading_seconds = total_chars / chars_per_min * 60 if chars_per_min > 0 else 0
    estimated_seconds = reading_seconds + total_pause

    # 各スライドにも推定時間を追加
    for s in slides:
        reading = s["char_count"] / chars_per_min * 60 if chars_per_min > 0 else 0
        s["estimated_seconds"] = reading + s["pause_seconds"]

    return {
        "total_slides": len(slides),
        "slides_with_notes": slides_with_notes,
        "total_chars": total_chars,
        "total_pause_seconds": total_pause,
        "estimated_seconds": estimated_seconds,
        "slides": slides,
    }


def format_duration(seconds: float) -> str:
    """秒数を「○分○秒」形式にフォーマットする。"""
    minutes = int(seconds) // 60
    secs = int(seconds) % 60
    if minutes > 0:
        return f"{minutes}分{secs:02d}秒"
    return f"{secs}秒"
