"""ko-NeKo — ナレーション尺カウンター UI"""

import streamlit as st
from narration_counter import (
    analyze_narration,
    format_duration,
    group_slides_into_steps,
    STEP_TARGET_MIN_SEC,
    STEP_TARGET_MAX_SEC,
)
import step_segmenter


def _load_anthropic_key():
    """Streamlit Secrets（Cloud）→ ローカルファイル の順で Anthropic API キーを取得。"""
    from pathlib import Path
    try:
        if "anthropic_api_key" in st.secrets:
            return st.secrets["anthropic_api_key"]
    except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
        pass
    p = Path.home() / ".config" / "koneko-idcheck" / "anthropic_api_key.txt"
    if p.exists():
        return p.read_text().strip()
    return None


def _render_step_table(steps):
    """ステップ別集計テーブルを描画する（app/ui 共通の見た目）。"""
    target_label = f"目標 {STEP_TARGET_MIN_SEC // 60}〜{STEP_TARGET_MAX_SEC // 60}分"
    html = (
        '<table style="width:auto;border-collapse:collapse;font-size:0.9em">'
        '<thead><tr style="background:#E8E0D8;color:#3D3929">'
        '<th style="padding:8px 12px;text-align:left">ステップ</th>'
        '<th style="padding:8px 12px;text-align:left">スライド</th>'
        '<th style="padding:8px 12px;text-align:right">文字数</th>'
        '<th style="padding:8px 12px;text-align:right">推定尺</th>'
        f'<th style="padding:8px 12px;text-align:left">{target_label}</th>'
        '</tr></thead><tbody>'
    )
    for idx, stp in enumerate(steps):
        sec = stp["estimated_seconds"]
        if stp["status"] == "over":
            badge = (f'<span style="color:#C0392B;font-weight:600">⚠ '
                     f'+{format_duration(sec - STEP_TARGET_MAX_SEC)} オーバー</span>')
        elif stp["status"] == "under":
            badge = (f'<span style="color:#B7791F;font-weight:600">⚠ '
                     f'-{format_duration(STEP_TARGET_MIN_SEC - sec)} 不足</span>')
        else:
            badge = '<span style="color:#1E8449;font-weight:600">✓ 目標内</span>'
        rng = stp["slide_range"]
        title_hint = (stp["first_title"] or "")[:20]
        bg = "#FFFFFF" if idx % 2 == 0 else "#FAF7F4"
        html += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:6px 12px"><b style="color:#C35A35">{stp["step_label"]}</b> '
            f'<span style="color:#8A7E6B">{title_hint}</span></td>'
            f'<td style="padding:6px 12px;color:#8A7E6B">'
            f'S{rng[0]}–S{rng[1]}（{stp["slide_count"]}枚）</td>'
            f'<td style="padding:6px 12px;text-align:right">{stp["char_count"]:,}</td>'
            f'<td style="padding:6px 12px;text-align:right">{format_duration(sec)}</td>'
            f'<td style="padding:6px 12px">{badge}</td>'
            f'</tr>'
        )
    html += '</tbody></table>'
    st.markdown(html, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ページ設定
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ko-NeKo：ナレーションカウンター",
    page_icon=":cat:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# カスタムCSS（NeKo共通デザイン）
# ─────────────────────────────────────────────
st.markdown("""
<style>
    [data-testid="stToolbar"], .stDeployButton, .stAppDeployButton,
    #MainMenu, .stActionButton, [data-testid="manage-app-button"] {
        display: none !important; visibility: hidden !important;
        height: 0 !important; overflow: hidden !important;
    }
    [data-testid="stHeader"] {
        background: transparent !important;
        height: 0 !important; min-height: 0 !important;
        max-height: 0 !important; overflow: hidden !important;
    }
    .stMainBlockContainer, .block-container { padding-top: 1rem !important; }
    [data-testid="stSidebar"] { min-width: 280px !important; }

    @import url('https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic:wght@400;500;700&display=swap');
    html, body, .stApp, .stApp *:not([class*="icon"]):not([class*="Icon"]):not([data-testid="stIconMaterial"]):not(.material-symbols-rounded) {
        font-family: 'Zen Maru Gothic', 'Hiragino Maru Gothic ProN', 'BIZ UDGothic', sans-serif !important;
    }
    .stApp { background-color: #FAF7F4; }
    [data-testid="stSidebar"] { background-color: #F0EBE4; }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: #3D3929; }

    .stButton > button[kind="primary"], button[kind="primary"] {
        background-color: #C35A35 !important; border-color: #C35A35 !important; color: white !important;
    }
    .stButton > button[kind="primary"]:hover, button[kind="primary"]:hover {
        background-color: #A8492C !important; border-color: #A8492C !important;
    }
    .stProgress > div > div > div > div { background-color: #C35A35 !important; }

    /* ファイルアップローダーの背景・ボーダーをオレンジ系に */
    [data-testid="stFileUploader"] section {
        border-color: #C35A35 !important;
        background-color: #FFF3EC !important;
    }
    [data-testid="stFileUploader"] button {
        color: #C35A35 !important;
    }

    /* メトリクスの数値を大きく */
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size: 2rem; color: #3D3929;
    }
    /* 最初のメトリクス（推定動画尺）をオレンジに */
    [data-testid="stHorizontalBlock"] > div:first-child [data-testid="stMetricValue"] {
        color: #C35A35 !important; font-weight: 700;
    }
    [data-testid="stMetric"] [data-testid="stMetricLabel"] {
        color: #8A7E6B;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ヘッダー
# ─────────────────────────────────────────────
st.markdown(
    '<div style="margin-bottom:0.5rem">'
    '<span style="font-size:1.8rem;font-weight:700;color:#3D3929">ko-NeKo</span>'
    '<span style="font-size:0.9rem;color:#8A7E6B;margin-left:8px">ナレーションカウンター</span>'
    '</div>',
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# サイドバー：読み上げ速度設定
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("設定")

    chars_per_min = st.slider(
        "読み上げ速度（文字/分）",
        min_value=250,
        max_value=400,
        value=300,
        step=10,
        help="日本語ナレーションの目安: ゆっくり250〜速め400文字/分",
    )


# ─────────────────────────────────────────────
# ファイルアップロード
# ─────────────────────────────────────────────
uploaded = st.file_uploader(
    "PPTXファイルをアップロード",
    type=["pptx"],
    help="パワーポイントをアップロードすると、ノート欄のナレーション文字数と推定尺を表示します",
)

if uploaded:
    # 解析実行
    result = analyze_narration(uploaded, chars_per_min=chars_per_min)
    slides = result["slides"]

    # ─────────────────────────────────────────
    # サマリー（3カラム metric）
    # ─────────────────────────────────────────
    st.markdown('<div style="margin-top:2.5rem"></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("推定動画尺", format_duration(result["estimated_seconds"]))
    with col2:
        st.metric(
            "スライド数",
            f"{result['slides_with_notes']} / {result['total_slides']}",
            help="ナレーション付き / 総スライド数",
        )
    with col3:
        st.metric("合計文字数", f"{result['total_chars']:,}")

    # ─────────────────────────────────────────
    # タイムラインバー
    # ─────────────────────────────────────────
    st.markdown('<div style="margin-top:2.5rem"></div>', unsafe_allow_html=True)

    # 文字数に応じた色（薄ベージュ→オレンジ）
    max_chars = max((s["char_count"] for s in slides), default=1)
    if max_chars == 0:
        max_chars = 1

    timeline_html = '<div style="display:flex;width:100%;height:48px;border-radius:8px;overflow:hidden;gap:1px;background:#E8E0D8">'
    for s in slides:
        if s["char_count"] == 0:
            continue
        width_pct = s["estimated_seconds"] / result["estimated_seconds"] * 100 if result["estimated_seconds"] > 0 else 0
        # 色の補間: #F5E6D3（薄い）→ #C35A35（濃い）
        ratio = s["char_count"] / max_chars
        r = int(245 + (195 - 245) * ratio)
        g = int(230 + (90 - 230) * ratio)
        b = int(211 + (53 - 211) * ratio)
        color = f"rgb({r},{g},{b})"
        # テキスト色: 濃い背景は白、薄い背景は暗め
        text_color = "#FFFFFF" if ratio > 0.5 else "#7A4A2A"
        title_text = s["title"] or f"スライド{s['slide_num']}"
        tooltip = f"S{s['slide_num']}: {title_text}&#10;{s['char_count']}文字 / {format_duration(s['estimated_seconds'])}"
        # スライド番号をバー内に表示
        font_size = "0.7em" if width_pct < 4 else "0.75em"
        timeline_html += (
            f'<div style="width:{width_pct:.2f}%;background:{color};min-width:14px;'
            f'cursor:pointer;transition:opacity 0.2s;'
            f'display:flex;align-items:center;justify-content:center;'
            f'font-size:{font_size};color:{text_color};font-weight:500" '
            f'title="{tooltip}" '
            f'onmouseover="this.style.opacity=0.7" '
            f'onmouseout="this.style.opacity=1">'
            f'{s["slide_num"]}'
            f'</div>'
        )
    timeline_html += '</div>'
    st.markdown(timeline_html, unsafe_allow_html=True)

    # 凡例
    st.markdown(
        '<div style="display:flex;align-items:center;gap:16px;margin-top:4px;font-size:0.8em;color:#8A7E6B">'
        '<span><span style="display:inline-block;width:12px;height:12px;background:#F5E6D3;border-radius:2px;vertical-align:middle"></span> 少ない</span>'
        '<span><span style="display:inline-block;width:12px;height:12px;background:#D9A06A;border-radius:2px;vertical-align:middle"></span> 中間</span>'
        '<span><span style="display:inline-block;width:12px;height:12px;background:#C35A35;border-radius:2px;vertical-align:middle"></span> 多い</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ─────────────────────────────────────────
    # スライド別テーブル
    # ─────────────────────────────────────────
    st.markdown('<div style="margin-top:2.5rem"></div>', unsafe_allow_html=True)

    import streamlit.components.v1 as components

    rows_html = ""
    for s in slides:
        bar_width = s["char_count"] / max_chars * 100 if max_chars > 0 else 0
        ratio = s["char_count"] / max_chars if max_chars > 0 else 0
        r = int(245 + (195 - 245) * ratio)
        g = int(230 + (90 - 230) * ratio)
        b = int(211 + (53 - 211) * ratio)
        bar_color = f"rgb({r},{g},{b})"
        bg = "#FFFFFF" if s["slide_num"] % 2 == 1 else "#FAF7F4"
        rows_html += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:6px 10px;text-align:center;color:#8A7E6B">{s["slide_num"]}</td>'
            f'<td style="padding:6px 10px;text-align:left;font-weight:500">{s["char_count"]:,}</td>'
            f'<td style="padding:6px 10px;text-align:left;color:#8A7E6B">{format_duration(s["estimated_seconds"])}</td>'
            f'<td style="padding:6px 10px">'
            f'<div style="background:#F0EBE4;border-radius:4px;height:16px;overflow:hidden">'
            f'<div style="width:{bar_width:.1f}%;background:{bar_color};height:100%;border-radius:4px"></div>'
            f'</div></td></tr>'
        )

    table_height = 44 + len(slides) * 33  # ヘッダー + 行数 × 行高さ
    components.html(f"""
    <link href="https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic:wght@400;500;700&display=swap" rel="stylesheet">
    <table style="width:auto;border-collapse:collapse;font-size:0.9em;font-family:'Zen Maru Gothic',sans-serif">
    <thead>
        <tr style="background:#E8E0D8;color:#3D3929">
            <th style="padding:8px 10px;text-align:center;border-radius:8px 0 0 0;width:70px">スライドNo.</th>
            <th style="padding:8px 10px;text-align:left;width:60px">文字数</th>
            <th style="padding:8px 10px;text-align:left;width:70px">推定時間</th>
            <th style="padding:8px 10px;text-align:left;border-radius:0 8px 0 0;min-width:400px">文字数の割合</th>
        </tr>
    </thead>
    <tbody>{rows_html}</tbody>
    </table>
    """, height=table_height)

    # ─────────────────────────────────────────
    # ステップ別集計（ノート原稿をAIで意味分割）
    # ─────────────────────────────────────────
    st.markdown('<div style="margin-top:2.5rem"></div>', unsafe_allow_html=True)
    st.markdown("#### 🪜 ステップ別の文字数・尺")
    st.caption(
        "ノート欄の原稿をAI（Claude）が読んで、内容が大きく変わる所で最大4ステップに分けます。"
        "各ステップが目標尺（10〜17分）に収まっているかを確認できます。"
    )

    api_key = _load_anthropic_key()
    file_id = f"{uploaded.name}:{getattr(uploaded, 'size', '')}"

    if not api_key:
        st.info(
            "ステップ分割（AI）には Anthropic API キーが必要です。"
            "管理者の方は Streamlit secrets の `anthropic_api_key`、または "
            "`~/.config/koneko-idcheck/anthropic_api_key.txt` を設定してください。"
        )
    else:
        if st.button("🪜 AIでステップに分割する", key="app_seg_btn", type="primary"):
            with st.spinner("AIがノート原稿を読んでステップを判定中…（10〜30秒ほど）"):
                try:
                    seg = step_segmenter.segment_steps(
                        slides, api_key,
                        lecture_title=uploaded.name.rsplit(".", 1)[0],
                    )
                    st.session_state["app_seg"] = {"file": file_id, **seg}
                except Exception as e:
                    st.error(f"ステップ分割に失敗しました: {e}")

        seg = st.session_state.get("app_seg")
        if seg and seg.get("file") == file_id and seg.get("boundaries"):
            slide_titles = {
                s["slide_num"]: (s["title"] or f"スライド{s['slide_num']}") for s in slides
            }
            all_slide_nums = [s["slide_num"] for s in slides]
            chosen = st.multiselect(
                "ステップの開始スライド（AIの判定。ずれていたら直せます）",
                options=all_slide_nums,
                default=seg["boundaries"],
                format_func=lambda n: f"S{n}: {slide_titles.get(n, '')[:24]}",
                key="app_counter_boundaries",
            )
            boundaries = sorted(chosen) if chosen else seg["boundaries"]
            # 手動で境界を変えたらAIラベルと対応がずれるので連番ラベルに切り替える
            use_labels = seg["labels"] if boundaries == seg["boundaries"] else None
            steps = group_slides_into_steps(slides, boundaries, use_labels)
            _render_step_table(steps)
            if seg.get("rationale"):
                st.caption(f"🤖 AIの判定根拠: {seg['rationale']}")


else:
    st.markdown(
        '<div style="background:#FFF3EC;border-left:4px solid #C35A35;border-radius:8px;padding:1rem 1.2rem;color:#3D3929;line-height:1.8">'
        'パワーポイント（.pptx）をアップロードすると、ノート欄のナレーション文字数と推定動画尺を分析します。<br><br>'
        '<b>使い方</b><br>'
        '1. PPTXファイルをドラッグ＆ドロップ<br>'
        '2. サイドバーで読み上げ速度を調整'
        '</div>',
        unsafe_allow_html=True,
    )
