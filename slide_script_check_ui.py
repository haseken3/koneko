"""ko-NeKo — スライド↔原稿チェックの結果画面（app.py / ui.py から呼ぶ共通 render）。

表示の原則（教員にそのまま見える面なので、ここが仕様の本体）:
- 断定しない。見出しも文言も「〜の可能性」「AIが気づいた点」で統一する。
  このチェックは誤りを保証するものでも、全部見つけるものでもない（両方を画面に書く）。
- 既定は確度「高・中」だけを出す。低確度は畳んだ先に置き、教員の画面を赤だらけにしない。
- st.tabs を使わない。ボタン押下の rerun で先頭タブに戻る個体差があるので、
  3機能（文字数／不一致／誤字脱字）は縦に並べたまま見せる。

読み手は高齢の先生が多い前提で、数字で固定した約束（2026-09-17 追加）:
- 本文 1.05rem 以上・行間 1.8。補助文字は #5C5346 以上の濃さ（白地でコントラスト比 4.5:1 を満たす）
- 「S5」のような略記を使わず「スライド 5」と書く
- 種類は色だけで区別しない（色＋文字ラベルの二重表示）
- 所見は「ステップ → スライド」で括る。ステップ分割前は1括り（分割を強制しない）
- 文言・色・括り方の正本は `slide_script_report.py`（Word 出力と同じものを使う）

`ensure_result` と `run_step_segmentation` は app.py / ui.py 共通の「file_idキャッシュ判定＋
ボタン制御」ヘルパー（streamlit依存）。ここに置くのは、ステップ分割の結果画面が
不一致・誤字チェックの括りにも使われ、両者が同じ file_id キャッシュの作法を共有するため
（2026-09-17 inspector指摘: app.py/ui.py で同じ25行がコピペされていたのを解消）。
"""

import hashlib
import html
import os
import sys
from pathlib import Path

import streamlit as st

_HERE = str(Path(__file__).resolve().parent)
if _HERE not in sys.path:
    sys.path.append(_HERE)
from slide_script_check import extract_slides, is_total_failure, run_checks  # noqa: E402
import step_segmenter  # noqa: E402
from slide_script_report import (  # noqa: E402
    DOCX_MIME,
    KIND_COLOR,
    KIND_LABEL,
    NO_FINDINGS_TEXT,
    build_docx,
    count_text,
    counterpart_label,
    download_filename,
    group_findings_by_step,
    quote_label,
    split_by_confidence,
)

INK = "#2E2A22"
MUTED = "#5C5346"
BODY = "font-size:1.05rem;line-height:1.8"


def make_file_id(source, label) -> str:
    """PPTXソースから file_id を作る（app.py / ui.py 共通）。

    app.py は `uploaded.size`、ui.py は Canva 経路等でサイズを持たないローカルパスを
    渡すことがあり、file_id の作り方が両者でズレていた（2026-09-17 verifier full 指摘
    W-3: 同名で内容だけ差し替えた場合の挙動が入口によって違う）。
    """
    size = getattr(source, "size", None)
    if size is None and isinstance(source, str):
        try:
            size = os.path.getsize(source)
        except OSError:
            size = None
    return f"{label}:{size}" if size is not None else str(label)


def widget_suffix(file_id: str) -> str:
    """file_id からwidget key用の短い接尾辞を作る（ステップ境界multiselectの汚染防止）。

    2026-09-17 verifier full 指摘 W-4: multiselect の key がファイルをまたいで固定だと、
    別ファイルの境界（session_state）が新しいファイルに引き継がれる
    （`feedback_streamlit_widget_state`）。key自体をfile_id依存にして解決する。
    """
    return hashlib.md5(file_id.encode()).hexdigest()[:8]


def ensure_result(state, file_id, runner, *, force: bool = False):
    """state に file_id 付きの成功結果が無ければ runner() を呼んで結果を返す（streamlit非依存）。

    自動実行化（2026-09-17・ケンタ指示「取り込んだら自動的に実行されるように」）の核。
    失敗も `{"file": file_id, "failed": "..."}` として返す——成功扱いで焼くと、rerunのたびに
    「結果が無い」と誤認して課金APIを呼び直すループになる（副将戦況更新指摘）。

    Args:
        state: 現在のキャッシュ（session_stateの値相当。None可）
        file_id: 対象ファイルの識別子
        runner: 引数なしで呼ぶと結果dictを返す関数（例外を投げてよい）
        force: True なら file_id が一致していても runner() を呼び直す（再実行ボタン用）

    Returns:
        dict: 新しい state（`{"file": file_id, **result}` または `{"file": file_id, "failed": "..."}`）
              実行不要な場合は state をそのまま返す
    """
    # ⚠ is_same_file は failed かどうかを問わない。failed 状態も「file_id が一致した
    # 直近の結果」であり、force 無しでは呼び直さない（前回失敗を毎rerunで再送すると
    # 課金ループになる・2026-09-17実測でこの判定漏れを1回踏んだ）。
    is_same_file = bool(state) and state.get("file") == file_id
    if is_same_file and not force:
        return state
    try:
        result = runner()
        return {"file": file_id, **result}
    except Exception as e:
        return {"file": file_id, "failed": f"{type(e).__name__}: {e}"}


def run_step_segmentation(key_prefix, file_id, slides, api_key, lecture_title):
    """ステップ分割を自動実行し、seg結果（またはNone）を返す（app.py/ui.py共通）。

    sonnet-5・安価なので取り込み時に自動実行する（2026-09-17・ケンタ指示
    「ステップ分割は自動で良い」）。app.py と ui.py で同一だったゲート判定
    （25行）をここへ集約（2026-09-17 inspector指摘）。

    Args:
        key_prefix: session_stateキーの接頭辞（"app" / "koneko"）
        file_id: 対象ファイルの識別子
        slides: analyze_narration の "slides"
        api_key: Anthropic API キー
        lecture_title: プロンプトに渡す授業タイトル

    Returns:
        dict | None: 成功した seg 結果（`{"file":..., "boundaries":..., "labels":..., "rationale":...}`）。
                     未実行・失敗中は None
    """
    seg_key = f"{key_prefix}_seg"
    cached_seg = st.session_state.get(seg_key)
    seg_fresh = bool(cached_seg) and cached_seg.get("file") == file_id
    seg_failed = seg_fresh and "failed" in cached_seg
    seg_ok = seg_fresh and not seg_failed

    retry_seg = False
    if seg_failed:
        st.error(f"ステップ分割に失敗しました: {cached_seg['failed']}")
        retry_seg = st.button("🪜 もう一度分割する", key=f"{key_prefix}_seg_retry", type="primary")
    elif seg_ok:
        retry_seg = st.button("🔄 もう一度分割する", key=f"{key_prefix}_seg_rerun")

    if retry_seg or not seg_fresh:
        with st.spinner("AIがノート原稿を読んでステップを判定中…（10〜30秒ほど）"):
            def _seg_runner():
                return step_segmenter.segment_steps(slides, api_key, lecture_title=lecture_title)
            cached_seg = ensure_result(cached_seg, file_id, _seg_runner, force=retry_seg)
            st.session_state[seg_key] = cached_seg

    return cached_seg if cached_seg and not cached_seg.get("failed") else None


def _card_html(finding) -> str:
    color = KIND_COLOR.get(finding["kind"], INK)
    esc = html.escape
    title = esc((finding.get("slide_title") or "")[:24])
    rows = [
        f'<div style="margin-bottom:6px">'
        f'<span style="color:{MUTED}">{esc(quote_label(finding))}:</span> '
        f'<span style="background:#FFF3EC;padding:2px 6px;border-radius:4px">'
        f'{esc(finding["quote"])}</span></div>'
    ]
    if finding.get("counterpart"):
        rows.append(
            f'<div style="margin-bottom:6px">'
            f'<span style="color:{MUTED}">{esc(counterpart_label(finding))}:</span> '
            f'<span style="background:#F0EBE4;padding:2px 6px;border-radius:4px">'
            f'{esc(finding["counterpart"])}</span></div>'
        )
    if finding.get("reason"):
        rows.append(f'<div style="margin-bottom:6px">{esc(finding["reason"])}</div>')
    if finding.get("suggestion"):
        rows.append(f'<div style="color:#1A6E3C">ご提案: {esc(finding["suggestion"])}</div>')
    return (
        f'<div style="border-left:5px solid {color};background:#FFFFFF;border-radius:6px;'
        f'padding:12px 16px;margin-bottom:10px">'
        f'<div style="{BODY};color:{MUTED};margin-bottom:8px">'
        f'<b style="color:{INK}">スライド {finding["slide_num"]}</b> {title} ・ '
        f'<span style="color:{color};font-weight:700">'
        f'{esc(KIND_LABEL.get(finding["kind"], finding["kind"]))}</span> ・ '
        f'AIの確度 {esc(finding["confidence"])}</div>'
        f'<div style="{BODY};color:{INK}">{"".join(rows)}</div>'
        f'</div>'
    )


def _group_html(group) -> str:
    """ステップ見出し＋件数行＋所見カードを1つの HTML にまとめる。"""
    parts = [
        f'<div style="background:#E8E0D8;border-radius:6px;padding:10px 16px;'
        f'margin:1.4rem 0 0.5rem;font-size:1.25rem;font-weight:700;color:{INK}">'
        f'{html.escape(group["heading"])}</div>'
    ]
    if group["findings"]:
        parts.append(
            f'<div style="{BODY};color:{MUTED};margin:0 0 10px 4px">'
            f'{html.escape(count_text(group))}</div>'
        )
    else:
        parts.append(
            f'<div style="{BODY};color:#1A6E3C;background:#F2F8F4;border-radius:6px;'
            f'padding:10px 16px;margin-bottom:10px">✓ {html.escape(NO_FINDINGS_TEXT)}</div>'
        )
    parts.extend(_card_html(f) for f in group["findings"])
    return "".join(parts)


def _render_section(findings, steps, *, heading: str):
    st.markdown(f'<div style="font-size:1.2rem;font-weight:700;color:{INK};'
                f'margin-top:1.8rem">{html.escape(heading)}</div>',
                unsafe_allow_html=True)
    main, low = split_by_confidence(findings)
    st.markdown("".join(_group_html(g) for g in group_findings_by_step(main, steps)),
                unsafe_allow_html=True)
    if low:
        with st.expander(f"参考: 確度が低い気づき（{len(low)}件）"):
            st.markdown("".join(_card_html(f) for f in low), unsafe_allow_html=True)


def _run_consistency_check(pptx_source, api_key, *, file_id, key_prefix, model, typo_model,
                           total_slides=None):
    """ゲート判定・ボタン制御・実行を担当する（結果画面の描画はしない）。

    不一致・誤字脱字チェックは opus-5 を使うため課金コストが高い（104枚デッキ実測で
    約$5.8）。ステップ分割と違い自動実行にはせず、ボタン押下でのみ実行する
    （2026-09-17・ケンタ指示「③だけボタン式に戻す」）。金額はケンタが負う費用なので
    先生には見せず、枚数と所要時間の目安だけを示す（2026-09-17 副将裁定 W-5）。

    Returns:
        dict | None: 成功結果（`{"file":..., "consistency":..., "typos":..., ...}`）。
                     未実行・失敗中は None
    """
    state_key = f"{key_prefix}_ssc"
    cached = st.session_state.get(state_key)
    is_fresh = bool(cached) and cached.get("file") == file_id
    has_failed = is_fresh and "failed" in cached
    already_ok = is_fresh and not has_failed

    if has_failed:
        st.error(f"チェックに失敗しました: {cached['failed']}")
        run_clicked = st.button("🔍 もう一度チェックする",
                                key=f"{key_prefix}_ssc_retry", type="primary")
    elif already_ok:
        run_clicked = st.button("🔄 もう一度チェックする", key=f"{key_prefix}_ssc_rerun")
    else:
        run_clicked = st.button("🔍 AIで不一致・誤字脱字をチェックする",
                                key=f"{key_prefix}_ssc_btn", type="primary")
        if total_slides:
            est_min = max(1, round(total_slides * 3 / 60))
            st.markdown(
                f'<div style="font-size:0.9rem;color:{MUTED}">'
                f'{total_slides}枚・目安 約{est_min}分（104枚で約5分）</div>',
                unsafe_allow_html=True,
            )

    if run_clicked:
        bar = st.progress(0.0, text="スライド本文と原稿を読み込み中…")
        warn_slot = st.empty()
        warn_slot.markdown(
            f'<div style="font-size:0.9rem;color:{MUTED}">'
            '⏳ 点検が終わるまで、画面の操作（スライダー等）をしないでください。途中で止まります。</div>',
            unsafe_allow_html=True,
        )

        def _on_progress(done, total):
            bar.progress(done / total if total else 1.0,
                         text=f"AIが点検中… {done}/{total} ブロック")

        def _runner():
            if hasattr(pptx_source, "seek"):
                pptx_source.seek(0)
            slides = extract_slides(pptx_source)
            overrides = {k: v for k, v in (("model", model), ("typo_model", typo_model)) if v}
            result = run_checks(slides=slides, api_key=api_key, progress=_on_progress, **overrides)
            if is_total_failure(result):
                raise RuntimeError(
                    f"全{len(result.get('errors', []))}件のブロックが失敗しました: "
                    + "; ".join(result.get("errors", []))
                )
            return result

        cached = ensure_result(cached, file_id, _runner, force=True)
        st.session_state[state_key] = cached
        bar.empty()
        warn_slot.empty()
        # ボタンは同一run内では再描画されずラベルが押下前のまま残る
        # （高齢の先生が「まだ実行していない」と誤読し再課金する型・2026-09-17 意匠full指摘）。
        # rerunして次のrenderで already_ok 分岐（「もう一度チェックする」）に切り替える。
        st.rerun()

    if not cached or cached.get("file") != file_id or cached.get("failed"):
        return None
    return cached


def _render_check_result(result, steps, *, key_prefix, source_label):
    """結果本体の描画（不一致・誤字脱字セクション）とサイドバーのWordダウンロードを描画する。"""
    for message in result.get("errors") or []:
        st.warning(f"⚠ {message}")

    consistency = result.get("consistency") or []
    typos = result.get("typos") or []
    visible, _ = split_by_confidence(consistency + typos)
    partial_note = " ⚠ 一部のブロックは点検できませんでした。" if result.get("errors") else ""
    st.markdown(
        f'<div style="{BODY};color:{MUTED};margin:0.5rem 0 1rem">'
        f'{result.get("slides", 0)}枚を点検しました。'
        f'確度「高・中」の気づきは {len(visible)}件です。{partial_note}'
        f'{"" if steps else "（ステップ分割ができなかったため、スライド順に並べています）"}</div>',
        unsafe_allow_html=True,
    )

    _render_section(consistency, steps, heading="📝 内容の不一致の可能性")
    _render_section(typos, steps, heading="✍️ 誤字・脱字の可能性")

    st.markdown(
        f'<div style="{BODY};color:{MUTED};margin-top:1rem">'
        '📥 チェック結果のWordファイルは、画面左のサイドバーからダウンロードできます。</div>',
        unsafe_allow_html=True,
    )
    with st.sidebar:
        st.divider()
        st.markdown("#### 📥 ダウンロード")
        st.download_button(
            label="Word でダウンロード",
            data=build_docx(result, steps, source_label),
            file_name=download_filename(source_label),
            mime=DOCX_MIME,
            type="primary",
            key=f"{key_prefix}_ssc_dl",
        )
        st.caption("印刷して、直した所に□のチェックを付けながら作業できます。")


def render_slide_script_check(pptx_source, api_key, *, file_id: str,
                              key_prefix: str = "app", model: str = None,
                              typo_model: str = None, steps=None,
                              source_label: str = "", total_slides: int = None):
    """「不一致の可能性」「誤字脱字の可能性」の2機能を描画する。

    Args:
        pptx_source: PPTX パスまたはファイルライクオブジェクト（アップロード物をそのまま渡してよい）
        api_key: Anthropic API キー（None ならボタンを出さず案内文だけ出す）
        file_id: 結果を紐づけるファイル識別子（別のファイルに差し替わったら結果を捨てる。
                 `make_file_id()` で作ったものを渡すこと）
        key_prefix: ウィジェットキーの接頭辞（app.py と ui.py で衝突させない）
        model / typo_model: 省略時は llm_models の既定（実測で差し替える前提の値）
        steps: group_slides_into_steps の返値。渡すとステップ単位で括る。
               None ならステップ分割前として1括りで出す（分割を強制しない）
        source_label: 資料名（Word の表紙とダウンロード名に使う。省略時はファイル名から推定）
        total_slides: 実行前の目安表示（枚数・所要時間）に使うスライド枚数。省略可
    """
    label = source_label or getattr(pptx_source, "name", "") or Path(str(pptx_source)).name

    st.markdown('<div style="margin-top:2.5rem"></div>', unsafe_allow_html=True)
    st.markdown("#### 🔍 スライドと原稿のチェック（AIが気づいた点）")
    st.caption(
        "スライド本文とノート欄の原稿をAI（Claude）が突き合わせ、"
        "内容の不一致と誤字脱字の「可能性」を並べます。"
        "誤りを保証するものでも、すべての誤りを見つけるものでもありません。"
        "最終的な判断は先生が行ってください。"
    )

    if not api_key:
        st.info(
            "このチェックには Anthropic API キーが必要です。"
            "管理者の方は Streamlit secrets の `anthropic_api_key`、または "
            "`~/.config/koneko-idcheck/anthropic_api_key.txt` を設定してください。"
        )
        return

    result = _run_consistency_check(pptx_source, api_key, file_id=file_id,
                                    key_prefix=key_prefix, model=model, typo_model=typo_model,
                                    total_slides=total_slides)
    if result is None:
        return
    _render_check_result(result, steps, key_prefix=key_prefix, source_label=label)
