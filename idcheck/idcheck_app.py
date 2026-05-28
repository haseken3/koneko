"""ko-NeKo IDチェッカー — Streamlit UI

PPTX をアップロード → ペルソナ選択 → 評価開始 → Word DL の最小フロー。
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import anthropic
import streamlit as st

from docx_writer import build_docx
from llm_evaluator import (
    MODEL_OPUS,
    MODEL_SONNET,
    evaluate_all,
    list_personas,
    load_checklist,
    llm_evaluable_only,
)
from pptx_reader import read_slides


# ─────────────────────────────────────────────
# 設定
# ─────────────────────────────────────────────
APP_TITLE = "ko-NeKo IDチェッカー"
LOCAL_KEY_PATH = Path.home() / ".config" / "koneko-idcheck" / "anthropic_api_key.txt"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def load_api_key() -> str | None:
    """Streamlit Secrets（Cloud）→ ローカルファイル の順で API キーを取得。"""
    try:
        if "anthropic_api_key" in st.secrets:
            return st.secrets["anthropic_api_key"]
    except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
        pass
    if LOCAL_KEY_PATH.exists():
        return LOCAL_KEY_PATH.read_text().strip()
    return None


# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.set_page_config(page_title=APP_TITLE, page_icon="📋", layout="wide")
st.title(APP_TITLE)
st.caption(
    "オンデマンド授業 PPTX を、鈴木克明レイヤーモデル44項目で自動評価し、"
    "Word フィードバックを生成します。"
)

GRADE_TO_PERSONA = {
    "1年次": "grade1_shin",
    "2年次": "grade2_mai",
    "3年次": "grade3_ken",
    "4年次": "grade4_aya",
}

# サイドバー
with st.sidebar:
    st.header("評価設定")
    grade_label = st.selectbox(
        "授業の年次",
        options=list(GRADE_TO_PERSONA.keys()),
        index=1,  # 2年次がデフォルト
        help="評価対象の授業を受講する学年を選択してください。",
    )
    persona_key = GRADE_TO_PERSONA[grade_label]

# モデルと評価項目範囲は固定
model = MODEL_OPUS
cl = load_checklist()
item_filter = llm_evaluable_only

# API キー確認
api_key = load_api_key()
if not api_key:
    st.error(
        f"⚠ Anthropic API キーが設定されていません。\n\n"
        f"ローカル: `{LOCAL_KEY_PATH}` にキーを保存（chmod 600 推奨）\n\n"
        f"Streamlit Cloud: Secrets に `anthropic_api_key` を設定"
    )
    st.stop()

# PPTX アップロード
uploaded = st.file_uploader(
    "PPTX ファイルをアップロード",
    type=["pptx"],
    help="評価対象のオンデマンド授業 PPTX を選択してください。",
)

if not uploaded:
    st.info("👆 PPTX ファイルをアップロードしてください。")
    st.stop()

# スライド解析
try:
    slides = read_slides(uploaded)
except Exception as e:
    st.error(f"PPTX 解析に失敗しました: {e}")
    st.stop()

# 解析結果プレビュー
col1, col2, col3 = st.columns(3)
col1.metric("スライド数", len(slides))
col2.metric("ノート付きスライド", sum(1 for s in slides if s["notes"]))
col3.metric(
    "総文字数（本文+ノート）",
    sum(len(s["body"]) + len(s["notes"]) for s in slides),
)

with st.expander("スライド一覧（タイトル）"):
    for s in slides:
        st.markdown(f"- **S{s['slide_num']}**: {s['title'] or '（タイトルなし）'}")

# 評価項目数の事前確認
items = cl["items"]
if item_filter:
    items = [it for it in items if item_filter(it)]

st.divider()
st.subheader("評価実行")
st.markdown(
    f"**実行内容**: {len(items)} 項目 × 1 ペルソナ（{persona_key}）× {model.split('-')[1]} モデル\n\n"
    f"想定 API コール数: {len(items)} 回 / "
    f"想定所要時間: 約 {len(items) * 9 // 60} 〜 {len(items) * 12 // 60} 分"
)

if st.button("🚀 評価を開始", type="primary"):
    progress = st.progress(0.0, text="評価を開始しています…")
    log_area = st.empty()
    log_lines = []

    client = anthropic.Anthropic(api_key=api_key)

    def progress_callback(idx: int, total: int, item: dict, result: dict):
        progress.progress(
            idx / total,
            text=f"評価中 {idx}/{total}: [{item['id']}] {item['title'][:40]}",
        )
        if result["ok"]:
            ev = result["evaluation"]
            mark = {-1: "🔴", 0: "⚪", 1: "🟢"}.get(ev["score"], "❓")
            log_lines.append(
                f"{mark} **[{item['id']}]** {item['title'][:50]} → スコア {ev['score']:+d}"
            )
        else:
            log_lines.append(
                f"⚠ **[{item['id']}]** {item['title'][:50]} → 失敗: {result.get('error', '不明')}"
            )
        # 直近 10 件だけ表示（長くなりすぎ防止）
        log_area.markdown("\n".join(log_lines[-10:]))

    try:
        results = evaluate_all(
            client=client,
            slides=slides,
            persona_key=persona_key,
            model=model,
            progress_callback=progress_callback,
            item_filter=item_filter,
        )
    except ValueError as e:
        st.error(f"⚠ 入力サイズ超過: {e}")
        st.stop()
    except Exception as e:
        st.error(f"⚠ 評価中に予期しないエラー: {type(e).__name__}: {e}")
        st.stop()

    progress.progress(1.0, text="評価完了！")

    # 結果サマリ
    ok_count = sum(1 for r in results if r["result"]["ok"])
    fail_count = len(results) - ok_count
    score_counts = {-1: 0, 0: 0, 1: 0}
    for r in results:
        if r["result"]["ok"]:
            s = r["result"]["evaluation"]["score"]
            score_counts[s] = score_counts.get(s, 0) + 1

    st.success(f"✅ 評価完了: {ok_count}/{len(results)} 件成功")

    smcol1, smcol2, smcol3, smcol4 = st.columns(4)
    smcol1.metric("良好 +1", score_counts[1])
    smcol2.metric("中立 0", score_counts[0])
    smcol3.metric("要改善 -1", score_counts[-1])
    smcol4.metric("失敗", fail_count, delta_color="inverse")

    if fail_count > 0:
        st.warning(
            f"⚠ {fail_count} 件の評価が失敗しました（API エラー等）。"
            "失敗項目も Word に記載されます。"
        )

    # Word 生成
    docx_buf = build_docx(
        results=results,
        pptx_filename=uploaded.name,
        persona_key=persona_key,
        model=model,
        levels_meta=cl["levels"],
    )

    pptx_stem = Path(uploaded.name).stem
    download_name = (
        f"{pptx_stem}_IDフィードバック_"
        f"{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    )

    st.download_button(
        label="📥 Word フィードバックをダウンロード",
        data=docx_buf,
        file_name=download_name,
        mime=DOCX_MIME,
        type="primary",
    )

    # 詳細プレビュー
    with st.expander("評価結果プレビュー（全項目）"):
        for r in results:
            item = r["item"]
            res = r["result"]
            if res["ok"]:
                ev = res["evaluation"]
                mark = {-1: "🔴", 0: "⚪", 1: "🟢"}.get(ev["score"], "❓")
                st.markdown(
                    f"### {mark} [{item['id']}] {item['title']}\n"
                    f"**スコア**: {ev['score']:+d}　**該当スライド**: "
                    f"{', '.join(f'S{n}' for n in ev.get('related_slides', []))}\n\n"
                    f"{ev['rationale']}"
                )
            else:
                st.markdown(
                    f"### ⚠ [{item['id']}] {item['title']}\n"
                    f"評価失敗: {res.get('error', '不明')}"
                )
            st.divider()
