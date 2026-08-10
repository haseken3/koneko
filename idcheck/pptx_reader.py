"""PPTX 解析モジュール — 全スライドのタイトル・本文・ノート欄を抽出する。"""

import sys
from pathlib import Path

from pptx import Presentation

# ノート欄の読み取り判断は koneko ルートの pptx_notes に集約する（読み手が2箇所あり、
# 別々に書くと片方だけ直る）。IDチェッカーは idcheck/ を起点に起動するので、
# ルートを末尾に足してから import する（先頭に差すと同名モジュールの解決先を奪う）。
_KONEKO_ROOT = str(Path(__file__).resolve().parent.parent)
if _KONEKO_ROOT not in sys.path:
    sys.path.append(_KONEKO_ROOT)
from pptx_notes import read_notes  # noqa: E402


def read_slides(pptx_file) -> list[dict]:
    """PPTX から全スライドのテキスト情報を抽出する。

    Args:
        pptx_file: PPTX ファイルパスまたはファイルライクオブジェクト

    Returns:
        list[dict]: 各スライドの情報
            slide_num (int): スライド番号（1始まり）
            title (str): スライドタイトル
            body (str): スライド本文（全テキストフレームを結合）
            notes (str): ノート欄テキスト
    """
    prs = Presentation(pptx_file)
    total_slides = len(prs.slides)
    slides = []

    for i, slide in enumerate(prs.slides, start=1):
        title = ""
        body_parts = []

        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            text = sh.text_frame.text.strip()
            if not text:
                continue
            if not title:
                title = text.split("\n")[0][:120]
            body_parts.append(text)

        notes = read_notes(slide, slide_num=i, total_slides=total_slides)

        slides.append({
            "slide_num": i,
            "title": title,
            "body": "\n".join(body_parts),
            "notes": notes,
        })

    return slides


def slides_summary(slides: list[dict]) -> dict:
    """全スライドの軽量サマリ（評価項目で全体参照する用）。

    Returns:
        dict: total_slides, titles (list of "S{num}: {title}")
    """
    return {
        "total_slides": len(slides),
        "titles": [f"S{s['slide_num']}: {s['title']}" for s in slides],
    }
