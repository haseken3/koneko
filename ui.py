"""ko-NeKo 道具ツール群 — 統合UI用 render 関数（ナレーションカウンター / IDチェッカー）."""

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
from koneko import step_segmenter


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

    source, source_label = _render_source_input("koneko_counter")

    if source:
        result = analyze_narration(source, chars_per_min=chars_per_min)
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
        file_id = str(source_label)

        if not api_key:
            st.info(
                "ステップ分割（AI）には Anthropic API キーが必要です。"
                "管理者の方は Streamlit secrets の `anthropic_api_key`、または "
                "`~/.config/koneko-idcheck/anthropic_api_key.txt` を設定してください。"
            )
        else:
            if st.button("🪜 AIでステップに分割する", key="koneko_seg_btn", type="primary"):
                with st.spinner("AIがノート原稿を読んでステップを判定中…（10〜30秒ほど）"):
                    try:
                        seg = step_segmenter.segment_steps(
                            slides, api_key,
                            lecture_title=str(source_label).rsplit(".", 1)[0],
                        )
                        st.session_state["koneko_seg"] = {"file": file_id, **seg}
                    except Exception as e:
                        st.error(f"ステップ分割に失敗しました: {e}")

            seg = st.session_state.get("koneko_seg")
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
                    key="koneko_counter_boundaries",
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
            'パワーポイント（.pptx）または Canva のデザインを取り込むと、ノート欄のナレーション文字数と推定動画尺を分析します。<br><br>'
            '<b>使い方</b><br>'
            '1. PPTXをドラッグ＆ドロップ、または「Canva URL」タブで共有URLを取り込む<br>'
            '2. サイドバーで読み上げ速度を調整'
            '</div>',
            unsafe_allow_html=True,
        )


# ─────────────────────────────────────────────
# IDチェッカー（鈴木克明レイヤーモデル44項目の自動評価）
# ─────────────────────────────────────────────
def render_koneko_idcheck():
    """IDチェッカーの UI を描画する（統合UI埋め込み用）.

    idcheck/ 配下のモジュール（llm_evaluator 等）は Path(__file__) ベースで
    checklist.json / personas を解決するため、ディレクトリを sys.path に通すだけで
    import できる。単独アプリ idcheck_app.py の st.stop() は、統合UIでは下部ナビバーまで
    止めてしまうため return に置き換えている。
    """
    import sys
    from pathlib import Path
    from datetime import datetime

    _IDCHECK_DIR = Path(__file__).resolve().parent / "idcheck"
    if str(_IDCHECK_DIR) not in sys.path:
        sys.path.insert(0, str(_IDCHECK_DIR))

    import anthropic
    from docx_writer import build_docx
    from llm_evaluator import (
        MODEL_OPUS,
        evaluate_all,
        load_checklist,
        llm_evaluable_only,
    )
    from pptx_reader import read_slides

    LOCAL_KEY_PATH = Path.home() / ".config" / "koneko-idcheck" / "anthropic_api_key.txt"
    DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    def _load_api_key():
        """Streamlit Secrets（Cloud）→ ローカルファイル の順で API キーを取得。"""
        try:
            if "anthropic_api_key" in st.secrets:
                return st.secrets["anthropic_api_key"]
        except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
            pass
        if LOCAL_KEY_PATH.exists():
            return LOCAL_KEY_PATH.read_text().strip()
        return None

    GRADE_TO_PERSONA = {
        "1年次": "grade1_shin",
        "2年次": "grade2_mai",
        "3年次": "grade3_ken",
        "4年次": "grade4_aya",
    }

    # ファイルアップローダーの見た目をオレンジ系に（counter と揃える）
    st.markdown("""
    <style>
        [data-testid="stFileUploader"] section {
            border-color: #C35A35 !important;
            background-color: #FFF3EC !important;
        }
        [data-testid="stFileUploader"] button { color: #C35A35 !important; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown(
        "<div style='color:#8A7E6B;line-height:1.7;margin-bottom:1rem'>"
        "オンデマンド授業の PPTX / Canva を、鈴木克明レイヤーモデル44項目で自動評価し、"
        "Word フィードバックを生成します。</div>",
        unsafe_allow_html=True,
    )

    # サイドバー：評価設定（年次＝ペルソナ選択）
    with st.sidebar:
        st.header("評価設定")
        grade_label = st.selectbox(
            "授業の年次",
            options=list(GRADE_TO_PERSONA.keys()),
            index=1,  # 2年次がデフォルト
            help="評価対象の授業を受講する学年を選択してください。",
            key="koneko_idcheck_grade",
        )
    persona_key = GRADE_TO_PERSONA[grade_label]

    model = MODEL_OPUS
    cl = load_checklist()
    item_filter = llm_evaluable_only

    api_key = _load_api_key()
    if not api_key:
        st.error(
            f"⚠ Anthropic API キーが設定されていません。\n\n"
            f"ローカル: `{LOCAL_KEY_PATH}` にキーを保存（chmod 600 推奨）"
        )
        return

    # 入力ソース（PPTX or Canva）— counter と共通ヘルパー
    source, source_label = _render_source_input("koneko_idcheck")
    if not source:
        st.markdown(
            '<div style="background:#FFF3EC;border-left:4px solid #C35A35;border-radius:8px;'
            'padding:1rem 1.2rem;color:#3D3929;line-height:1.8;margin-top:1rem">'
            'PPTX をアップロードするか、「Canva URL」タブで共有URLを取り込むと評価を開始できます。'
            '</div>',
            unsafe_allow_html=True,
        )
        return

    # スライド解析
    try:
        slides = read_slides(source)
    except Exception as e:
        st.error(f"PPTX 解析に失敗しました: {e}")
        return

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

    if st.button("🚀 評価を開始", type="primary", key="koneko_idcheck_run"):
        progress = st.progress(0.0, text="評価を開始しています…")
        log_area = st.empty()
        log_lines = []
        client = anthropic.Anthropic(api_key=api_key)

        def progress_callback(idx, total, item, result):
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
            return
        except Exception as e:
            st.error(f"⚠ 評価中に予期しないエラー: {type(e).__name__}: {e}")
            return

        progress.progress(1.0, text="評価完了！")

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

        docx_buf = build_docx(
            results=results,
            pptx_filename=source_label,
            persona_key=persona_key,
            model=model,
            levels_meta=cl["levels"],
        )
        pptx_stem = Path(source_label).stem or "idcheck"
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
