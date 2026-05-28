"""ko-NeKo ナレーションカウンター — 統合UI用 render 関数."""

import streamlit as st
import streamlit.components.v1 as components

from koneko.narration_counter import analyze_narration, format_duration


def render_koneko_counter():
    """ナレーションカウンターの UI を描画する（統合UI埋め込み用）."""

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
    # ファイルアップローダーの見た目をオレンジ系に
    st.markdown("""
    <style>
        [data-testid="stFileUploader"] section {
            border-color: #C35A35 !important;
            background-color: #FFF3EC !important;
        }
        [data-testid="stFileUploader"] button {
            color: #C35A35 !important;
        }
        [data-testid="stMetric"] [data-testid="stMetricValue"] {
            font-size: 2rem; color: #3D3929;
        }
        [data-testid="stHorizontalBlock"] > div:first-child [data-testid="stMetricValue"] {
            color: #C35A35 !important; font-weight: 700;
        }
        [data-testid="stMetric"] [data-testid="stMetricLabel"] {
            color: #8A7E6B;
        }
    </style>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "PPTXファイルをアップロード",
        type=["pptx"],
        help="パワーポイントをアップロードすると、ノート欄のナレーション文字数と推定尺を表示します",
        key="koneko_counter_uploader",
    )

    if uploaded:
        result = analyze_narration(uploaded, chars_per_min=chars_per_min)
        slides = result["slides"]

        # サマリー
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

        # タイムラインバー
        st.markdown('<div style="margin-top:2.5rem"></div>', unsafe_allow_html=True)
        max_chars = max((s["char_count"] for s in slides), default=1)
        if max_chars == 0:
            max_chars = 1

        timeline_html = '<div style="display:flex;width:100%;height:48px;border-radius:8px;overflow:hidden;gap:1px;background:#E8E0D8">'
        for s in slides:
            if s["char_count"] == 0:
                continue
            width_pct = s["estimated_seconds"] / result["estimated_seconds"] * 100 if result["estimated_seconds"] > 0 else 0
            ratio = s["char_count"] / max_chars
            r = int(245 + (195 - 245) * ratio)
            g = int(230 + (90 - 230) * ratio)
            b = int(211 + (53 - 211) * ratio)
            color = f"rgb({r},{g},{b})"
            text_color = "#FFFFFF" if ratio > 0.5 else "#7A4A2A"
            title_text = s["title"] or f"スライド{s['slide_num']}"
            tooltip = f"S{s['slide_num']}: {title_text}&#10;{s['char_count']}文字 / {format_duration(s['estimated_seconds'])}"
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

        # スライド別テーブル
        st.markdown('<div style="margin-top:2.5rem"></div>', unsafe_allow_html=True)
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

        table_height = 44 + len(slides) * 33
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
