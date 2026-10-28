"""ko-NeKo — 統合UI用 render 関数（授業資料チェッカー＝ナレーションカウンター＋スライドと原稿のチェック）.

IDチェッカーは NeKo の画面から廃止した（ケンタ裁定 2026-10-28・使っていない）。idcheck/ は試作の記録として残す。
"""

import os

import streamlit as st
import streamlit.components.v1 as components

from koneko.narration_counter import (
    analyze_narration,
    format_duration,
    group_slides_into_steps,
    STEP_TARGET_MIN_SEC,
    STEP_TARGET_MAX_SEC,
)
from koneko import koneko_core
from koneko.slide_script_check_ui import (
    render_slide_script_check,
    run_step_segmentation,
    make_file_id,
    widget_suffix,
)


def _load_anthropic_key():
    """Streamlit Secrets（Cloud）→ ローカルファイル の順で Anthropic API キーを取得。"""
    secret = None
    try:
        if "anthropic_api_key" in st.secrets:
            secret = st.secrets["anthropic_api_key"]
    except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
        pass
    return koneko_core.load_anthropic_key(secret)


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
# 入力ソース共通ヘルパー（PPTXアップロード or Canva URL）
# ─────────────────────────────────────────────
def _render_source_input(key_prefix: str):
    """PPTXアップロード／Canva URL のタブUIを描画し、入力ソースを返す。

    Canva も最終的に一時PPTXへ化けるので、以降の解析処理は PPTX と完全に共通
    （siam/ui.py の取り込みパターンと同じ思想）。Canva 認証は siam と同じ
    ~/.voicespace-auto/ を共有するため、SiAm でログイン済みならそのまま使える。

    Returns:
        (source, label): source は PPTXファイルパス(str) or ファイルライク。
                         未入力なら (None, None)。
    """
    # Canva 取り込みは siam パッケージに依存する。NeKo 統合UI から起動した時だけ
    # siam が import できる。koneko を単体デプロイした環境（Streamlit Cloud で教員が
    # 使う本番 haseken3/koneko）には siam が無いため、その場合は Canva タブを出さず、
    # 従来どおり PPTX アップロードのみにフォールバックする（import 失敗で画面を壊さない）。
    try:
        from siam.canva_fetcher import fetch_canva_as_pptx, has_valid_token
    except ImportError:
        # siam 非同梱（本番 koneko 単体）でのみ起きる想定。siam 側の実装エラー等は
        # ここで握りつぶさず上に伝播させる（黙ってCanvaタブを消さない）。
        uploaded = st.file_uploader(
            "PPTXファイルをアップロード",
            type=["pptx"],
            key=f"{key_prefix}_uploader",
        )
        return (uploaded, uploaded.name) if uploaded else (None, None)

    path_key = f"{key_prefix}_canva_path"
    label_key = f"{key_prefix}_canva_label"

    tab_pptx, tab_canva = st.tabs(["📊 PPTXアップロード", "🎨 Canva URL"])
    with tab_pptx:
        uploaded = st.file_uploader(
            "PPTXファイルをアップロード",
            type=["pptx"],
            key=f"{key_prefix}_uploader",
        )
    with tab_canva:
        if not has_valid_token():
            st.warning(
                "Canva 連携がまだ設定されていません。先に「SiAm」タブで "
                "Canva ログイン（共有URL取り込み）を一度済ませてください。"
            )
        canva_url = st.text_input(
            "Canva の共有URL",
            placeholder="https://www.canva.com/design/... または https://canva.link/...",
            key=f"{key_prefix}_canva_url",
        ).strip()
        st.caption("Canva デザインのノート欄（プレゼンターノート）を原稿として取り込みます。")
        if st.button("取り込む", key=f"{key_prefix}_canva_fetch",
                     disabled=not canva_url):
            old = st.session_state.get(path_key)
            try:
                with st.spinner("Canva から取り込み中…（原稿の取り出しに最大数分かかります）"):
                    # interactive=False: トークン失効時はブラウザ認可へ入らず例外にする
                    new_path = fetch_canva_as_pptx(canva_url, interactive=False)
                # 取得に成功してから旧ファイルを掃除して差し替える（楽観的更新）。
                # 先に消すと、取り込み失敗時に前回の取り込み結果まで失ってしまう。
                if old and old != new_path and os.path.exists(old):
                    try:
                        os.unlink(old)
                    except OSError:
                        pass
                st.session_state[path_key] = new_path
                st.session_state[label_key] = canva_url
            except Exception as e:
                st.error(f"Canva の取り込みに失敗しました：{e}")
        cached = st.session_state.get(path_key)
        if cached and os.path.exists(cached):
            st.success(f"取り込み済み：{st.session_state.get(label_key, 'Canva デザイン')}")

    # PPTX アップロードが優先（siam と同じ：両方入力時は PPTX を使う）
    if uploaded:
        return uploaded, uploaded.name
    cached = st.session_state.get(path_key)
    if cached and os.path.exists(cached):
        return cached, st.session_state.get(label_key, "Canva デザイン")
    return None, None


def render_koneko_counter():
    """ナレーションカウンターの UI を描画する（統合UI埋め込み用）."""

    # ─────────────────────────────────────────────
    # サイドバー：読み上げ速度設定
    # ─────────────────────────────────────────────
    with st.sidebar:
        st.header("設定")
        _cpm = koneko_core.CHARS_PER_MIN
        chars_per_min = st.slider(
            "読み上げ速度（文字/分）",
            min_value=_cpm["min"],
            max_value=_cpm["max"],
            value=_cpm["value"],
            step=_cpm["step"],
            help=_cpm["help"],
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
        [data-testid="stDownloadButton"] button[kind="primary"] {
            background-color: #C35A35 !important; border-color: #C35A35 !important;
            color: white !important;
        }
        [data-testid="stDownloadButton"] button[kind="primary"] p {
            color: white !important;
        }
    </style>
    """, unsafe_allow_html=True)

    source, source_label = _render_source_input("koneko_counter")

    if source:
        result = analyze_narration(source, chars_per_min=chars_per_min)
        slides = result["slides"]

        # スライドと原稿のチェックは画面の一番上に表示する（2026-09-17・ケンタ指示）。
        # 計算は下（steps確定後）で行い、描画だけこのスロットに差し込む。
        ssc_slot = st.container()

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
        max_chars = koneko_core.max_char_count(slides)

        timeline_html = '<div style="display:flex;width:100%;height:48px;border-radius:8px;overflow:hidden;gap:1px;background:#E8E0D8">'
        for s in slides:
            if s["char_count"] == 0:
                continue
            width_pct = s["estimated_seconds"] / result["estimated_seconds"] * 100 if result["estimated_seconds"] > 0 else 0
            ratio = s["char_count"] / max_chars
            color = koneko_core.heat_color(ratio)
            text_color = koneko_core.heat_text_color(ratio)
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
            bar_color = koneko_core.heat_color(ratio)
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

        # ─────────────────────────────────────────────
        # ステップ別集計（ノート原稿をAIで意味分割）
        # ─────────────────────────────────────────────
        st.markdown('<div style="margin-top:2.5rem"></div>', unsafe_allow_html=True)
        st.markdown("#### 🪜 ステップ別の文字数・尺")
        st.caption(
            "ノート欄の原稿をAI（Claude）が読んで、内容が大きく変わる所で最大4ステップに分けます。"
            "各ステップが目標尺（10〜17分）に収まっているかを確認できます。"
        )

        api_key = _load_anthropic_key()
        file_id = make_file_id(source, source_label)
        steps = None

        if not api_key:
            st.info(
                "ステップ分割（AI）には Anthropic API キーが必要です。"
                "管理者の方は Streamlit secrets の `anthropic_api_key`、または "
                "`~/.config/koneko-idcheck/anthropic_api_key.txt` を設定してください。"
            )
        else:
            seg = run_step_segmentation("koneko", file_id, slides, api_key,
                                        lecture_title=str(source_label).rsplit(".", 1)[0])
            if seg and seg.get("file") == file_id and seg.get("boundaries"):
                slide_titles = {
                    s["slide_num"]: (s["title"] or f"スライド{s['slide_num']}") for s in slides
                }
                all_slide_nums = [s["slide_num"] for s in slides]
                chosen = st.multiselect(
                    "ステップの開始スライド（AIの判定。ずれていたら直せます）",
                    options=all_slide_nums,
                    default=seg["boundaries"],
                    format_func=lambda n: f"スライド{n}: {slide_titles.get(n, '')[:24]}",
                    key=f"koneko_counter_boundaries_{widget_suffix(file_id)}",
                )
                boundaries = sorted(chosen) if chosen else seg["boundaries"]
                # 手動で境界を変えたらAIラベルと対応がずれるので連番ラベルに切り替える
                use_labels = seg["labels"] if boundaries == seg["boundaries"] else None
                steps = group_slides_into_steps(slides, boundaries, use_labels)
                _render_step_table(steps)
                if seg.get("rationale"):
                    st.caption(f"🤖 AIの判定根拠: {seg['rationale']}")

        # 描画は上（ssc_slot）で行う。計算はここまで（steps確定後）で完了している。
        with ssc_slot:
            render_slide_script_check(source, api_key, file_id=file_id, key_prefix="koneko",
                                      steps=steps, source_label=str(source_label),
                                      total_slides=result["total_slides"])

    else:
        st.markdown(
            '<div style="background:#FFF3EC;border-left:4px solid #C35A35;border-radius:8px;padding:1rem 1.2rem;color:#3D3929;line-height:1.8">'
            'パワーポイント（.pptx）または Canva のデザインを取り込むと、'
            'ノート欄のナレーション文字数・推定動画尺・ステップ別の尺を自動で分析します。<br><br>'
            '<b>使い方</b><br>'
            '1. PPTXをドラッグ＆ドロップ、または「Canva URL」タブで共有URLを取り込む<br>'
            '2. サイドバーで読み上げ速度を調整<br>'
            '3. 内容の不一致・誤字脱字のチェックは「チェックを実行する」ボタンで開始'
            '（枚数が多いと数分・費用がかかります）'
            '</div>',
            unsafe_allow_html=True,
        )
