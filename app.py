"""ko-NeKo — ナレーション尺カウンター UI"""

import csv
import io
import streamlit as st
from narration_counter import analyze_narration, format_duration

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

    /* メトリクスの数値を大きく */
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size: 2rem; color: #3D3929;
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
        value=320,
        step=10,
        help="日本語ナレーションの目安: ゆっくり250〜速め400文字/分",
    )

    st.markdown(
        f"<small style='color:#8A7E6B'>現在の設定: {chars_per_min}文字/分</small>",
        unsafe_allow_html=True,
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
    st.markdown("### サマリー")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            "スライド数",
            f"{result['slides_with_notes']} / {result['total_slides']}",
            help="ナレーション付き / 総スライド数",
        )
    with col2:
        st.metric("合計文字数", f"{result['total_chars']:,}")
    with col3:
        st.metric("推定動画尺", format_duration(result["estimated_seconds"]))

    # ─────────────────────────────────────────
    # タイムラインバー
    # ─────────────────────────────────────────
    st.markdown("### タイムライン")
    st.caption("各スライドの時間配分（ホバーで詳細表示）")

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
    st.markdown("### スライド別詳細")

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
            f'<td style="padding:6px 12px;text-align:center;color:#8A7E6B">{s["slide_num"]}</td>'
            f'<td style="padding:6px 12px;text-align:right;font-weight:500">{s["char_count"]:,}</td>'
            f'<td style="padding:6px 12px;text-align:right;color:#8A7E6B">{format_duration(s["estimated_seconds"])}</td>'
            f'<td style="padding:6px 12px;width:40%">'
            f'<div style="background:#F0EBE4;border-radius:4px;height:16px;overflow:hidden">'
            f'<div style="width:{bar_width:.1f}%;background:{bar_color};height:100%;border-radius:4px"></div>'
            f'</div></td></tr>'
        )

    table_height = 44 + len(slides) * 33  # ヘッダー + 行数 × 行高さ
    components.html(f"""
    <link href="https://fonts.googleapis.com/css2?family=Zen+Maru+Gothic:wght@400;500;700&display=swap" rel="stylesheet">
    <table style="width:100%;border-collapse:collapse;font-size:0.9em;font-family:'Zen Maru Gothic',sans-serif">
    <thead>
        <tr style="background:#E8E0D8;color:#3D3929">
            <th style="padding:8px 12px;text-align:center;border-radius:8px 0 0 0">No.</th>
            <th style="padding:8px 12px;text-align:right">文字数</th>
            <th style="padding:8px 12px;text-align:right">推定時間</th>
            <th style="padding:8px 12px;text-align:left;border-radius:0 8px 0 0">バー</th>
        </tr>
    </thead>
    <tbody>{rows_html}</tbody>
    </table>
    """, height=table_height)

    # ─────────────────────────────────────────
    # CSVダウンロード
    # ─────────────────────────────────────────
    st.markdown("---")

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["スライド番号", "文字数", "推定時間（秒）", "推定時間"])
    for s in slides:
        writer.writerow([
            s["slide_num"],
            s["char_count"],
            round(s["estimated_seconds"], 1),
            format_duration(s["estimated_seconds"]),
        ])
    writer.writerow([])
    writer.writerow(["合計", result["total_chars"], round(result["estimated_seconds"], 1), format_duration(result["estimated_seconds"])])

    csv_data = buf.getvalue().encode("utf-8-sig")  # Excel対応BOM付き
    filename = uploaded.name.replace(".pptx", "") + "_narration.csv"

    st.download_button(
        "CSVダウンロード",
        data=csv_data,
        file_name=filename,
        mime="text/csv",
        type="primary",
    )

else:
    st.info(
        "パワーポイント（.pptx）をアップロードすると、ノート欄のナレーション文字数と推定動画尺を分析します。\n\n"
        "**使い方**\n"
        "1. PPTXファイルをドラッグ＆ドロップ\n"
        "2. サイドバーで読み上げ速度を調整\n"
        "3. 結果をCSVでダウンロード"
    )
