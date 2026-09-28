"""
ko-NeKo IDチェッカー Day 1 サブセット観察スパイク

目的: Day 0（N=1）では見えなかった「評価品質」を観察する。
- 5項目 × 4ペルソナ評価（20回、Opus）= ペルソナ別差・rationale長分布・スコア分布
- 5項目 × Sonnet（マイ固定）= 5回（Phase 1 Opus結果と並走比較）
- prompt caching の動作確認（cache_creation / cache_read を usage から取得）

合計25回API、想定コスト約400円。

実行: /usr/bin/python3 run_spike_day1.py
出力: 標準出力 + engawa/logs/idcheck_day1.jsonl（逐次書込、中断OK）

GO/NO-GO判定基準（事前定義 = pre_observation_matrix_as_hypothesis 準拠）:
  GO:
    - silent failure 0件
    - rationale max < 800トークン（段階分割設計の根拠取り）
    - ペルソナ別 rationale 文面が変わる
    - cache hit が機能
  NO-GO:
    - silent failure 1件以上 → リトライ機構見直し
    - rationale max > 1500トークン → max_tokens 設計見直し
    - ペルソナ別差なし → ペルソナ注入再設計
"""

import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import anthropic
from pptx import Presentation

# ─────────────────────────────────────────────
# 設定
# ─────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
IDCHECK_DIR = SCRIPT_DIR.parent
KONEKO_DIR = IDCHECK_DIR.parent
ENGAWA_LOGS = KONEKO_DIR / "engawa" / "logs"

API_KEY_PATH = Path.home() / ".config" / "koneko-idcheck" / "anthropic_api_key.txt"
CHECKLIST_PATH = IDCHECK_DIR / "checklist.json"
PERSONAS_DIR = IDCHECK_DIR / "personas"
SAMPLE_PPTX = Path.home() / "Downloads" / "27-ソフトウェアアーキテクチャⅡ-04 -Rev01.pptx"
OUTPUT_JSONL = ENGAWA_LOGS / "idcheck_day1.jsonl"

MODEL_OPUS = "claude-opus-4-20250514"
MODEL_SONNET = "claude-sonnet-4-5-20250929"
MAX_TOKENS = 4096
MAX_RETRIES = 3

# 5項目 × 4ペルソナ評価対象（項目ID, スライドindex）
# 粒度違うやつを選定: 構造系/内容整合/説明レベル/専門用語/箇条書き
EVAL_ITEMS = [
    ("-1.6", 1),   # Step1冒頭の導入 → S2「この授業の目標と本ステップのテーマ」
    ("0.2", 4),    # 学習目標と内容の齟齬 → S5「サービスとは」
    ("0.5", 2),    # 説明レベル（Day 0比較用） → S3「Ⅰの復習」
    ("0.6", 4),    # 初出専門用語の説明 → S5「サービスとは」
    ("1.3", 6),    # 箇条書き構造化 → S7「責務とは」
]

PERSONAS = ["grade1_shin", "grade2_mai", "grade3_ken", "grade4_aya"]


# ─────────────────────────────────────────────
# データ読み込み
# ─────────────────────────────────────────────
def load_api_key() -> str:
    if not API_KEY_PATH.exists():
        sys.exit(f"❌ APIキーが見つかりません: {API_KEY_PATH}")
    return API_KEY_PATH.read_text().strip()


def load_checklist() -> dict:
    return json.loads(CHECKLIST_PATH.read_text(encoding="utf-8"))


def find_item(checklist: dict, item_id: str) -> dict:
    for it in checklist["items"]:
        if it["id"] == item_id:
            return it
    sys.exit(f"❌ checklist 項目が見つかりません: {item_id}")


def load_persona(persona_key: str) -> str:
    return (PERSONAS_DIR / f"{persona_key}.md").read_text(encoding="utf-8")


def load_slide_data(prs: Presentation, slide_index: int) -> dict:
    slides = list(prs.slides)
    if slide_index >= len(slides):
        sys.exit(f"❌ スライド範囲外: {slide_index} >= {len(slides)}")
    slide = slides[slide_index]
    title = ""
    body_parts = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            text = sh.text_frame.text.strip()
            if text and not title:
                title = text.split("\n")[0][:80]
                body_parts.append(text)
            elif text:
                body_parts.append(text)
    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
    return {
        "slide_num": slide_index + 1,
        "title": title,
        "body": "\n".join(body_parts),
        "notes": notes,
    }


# ─────────────────────────────────────────────
# プロンプト構築
# ─────────────────────────────────────────────
def build_system_prompt(persona_text: str, item: dict) -> str:
    return f"""あなたはオンデマンド授業教材のIDチェッカーです。

レイヤーモデル（鈴木克明, 2006）に基づく評価項目1つについて、
教材スライドの1枚を、特定の学生像を念頭に置いて評価してください。

# 評価軸
- 項目ID: {item['id']}
- レベル: {item['level']}
- 評価観点: {item['title']}
- 補足: {item['note']}

# 念頭に置く学生像
以下のペルソナを念頭に、「この学生にとってこの教材スライドはどうか」を判定してください。

---
{persona_text}
---

# 評価ルール
- スコアは -1（要改善・該当あり）/ 0（中立・判定保留）/ +1（良好・問題なし）の3水準
- 文字量・スライド数は評価対象ではありません（内容の質のみ判定）
- 評価根拠（rationale）は具体的に書く。「〜と感じた」の主観ではなく、スライド内容のどこを根拠としたかを明示
- スライドが評価項目に該当しない場合は score=0, rationale で理由説明
- 評価結果は必ず submit_evaluation ツールで返してください

# 重要
- AI生成かどうかには触れない
- 「長いから減点」「短いから減点」はしない
"""


def build_user_prompt(slide_data: dict) -> str:
    return f"""# 評価対象スライド

スライド番号: {slide_data['slide_num']}
タイトル: {slide_data['title']}

## スライド本文
{slide_data['body']}

## ナレーション原稿（ノート欄）
{slide_data['notes']}

---

上記のスライドを、システムプロンプトで指定された評価項目について評価し、submit_evaluation ツールで結果を返してください。"""


def build_tool_definition() -> dict:
    return {
        "name": "submit_evaluation",
        "description": "IDチェッカー評価結果を構造化して返す",
        "input_schema": {
            "type": "object",
            "properties": {
                "item_id": {"type": "string"},
                "score": {"type": "integer", "enum": [-1, 0, 1]},
                "rationale": {"type": "string"},
                "related_slides": {"type": "array", "items": {"type": "integer"}},
            },
            "required": ["item_id", "score", "rationale", "related_slides"],
        },
    }


# ─────────────────────────────────────────────
# API 呼び出し（cache_control 付き、内側リトライ付き）
# ─────────────────────────────────────────────
def evaluate(
    client: anthropic.Anthropic,
    model: str,
    system: str,
    user: str,
    tool: dict,
) -> dict:
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            t_start = time.time()
            response = client.messages.create(
                model=model,
                max_tokens=MAX_TOKENS,
                system=[
                    {
                        "type": "text",
                        "text": system,
                        "cache_control": {"type": "ephemeral"},
                    }
                ],
                messages=[{"role": "user", "content": user}],
                tools=[tool],
                tool_choice={"type": "tool", "name": tool["name"]},
            )
            t_elapsed = time.time() - t_start

            for block in response.content:
                if block.type == "tool_use" and block.name == tool["name"]:
                    usage = response.usage
                    return {
                        "ok": True,
                        "attempt": attempt,
                        "elapsed_sec": round(t_elapsed, 2),
                        "evaluation": block.input,
                        "usage": {
                            "input_tokens": usage.input_tokens,
                            "output_tokens": usage.output_tokens,
                            "cache_creation_input_tokens": getattr(
                                usage, "cache_creation_input_tokens", 0
                            ),
                            "cache_read_input_tokens": getattr(
                                usage, "cache_read_input_tokens", 0
                            ),
                        },
                        "stop_reason": response.stop_reason,
                    }

            last_error = f"tool_use ブロックが応答に含まれない（stop_reason={response.stop_reason}）"
            print(f"  ⚠️ 試行{attempt}: {last_error}, リトライ", file=sys.stderr)

        except (
            anthropic.APIConnectionError,
            anthropic.APITimeoutError,
            anthropic.RateLimitError,
            anthropic.InternalServerError,
        ) as e:
            last_error = f"{type(e).__name__}: {e}"
            print(f"  ⚠️ 試行{attempt}: {last_error}, リトライ", file=sys.stderr)
            time.sleep(2 * attempt)
        except (
            anthropic.BadRequestError,
            anthropic.AuthenticationError,
            anthropic.PermissionDeniedError,
            anthropic.NotFoundError,
        ) as e:
            # 即時打ち切り（リトライしても回復不能なエラー）
            return {
                "ok": False,
                "attempts": attempt,
                "error": f"{type(e).__name__}: {e}",
                "marker": "🚨 NON_RETRIABLE_ERROR: リトライ無効、即時上位通知",
            }
        except Exception as e:
            # 未知の例外も即時打ち切り（握りつぶし防止）
            return {
                "ok": False,
                "attempts": attempt,
                "error": f"UnknownError({type(e).__name__}): {e}",
                "marker": "🚨 UNKNOWN_ERROR: 想定外の例外、即時上位通知",
            }

    return {
        "ok": False,
        "attempts": MAX_RETRIES,
        "error": last_error,
        "marker": "🚨 SILENT_FAILURE_AVOIDED: 評価未完了として上位に通知すべき",
    }


# ─────────────────────────────────────────────
# JSONL 逐次書込
# ─────────────────────────────────────────────
def append_jsonl(record: dict):
    ENGAWA_LOGS.mkdir(parents=True, exist_ok=True)
    with OUTPUT_JSONL.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


# ─────────────────────────────────────────────
# 実走
# ─────────────────────────────────────────────
def run_one(
    client: anthropic.Anthropic,
    model: str,
    persona_key: str,
    persona_text: str,
    item: dict,
    slide_data: dict,
    tool: dict,
    tag: str,
) -> dict:
    system = build_system_prompt(persona_text, item)
    user = build_user_prompt(slide_data)
    print(
        f"  [{tag}] item={item['id']} persona={persona_key} model={model.split('-')[1]}...",
        end="",
        flush=True,
    )
    result = evaluate(client, model, system, user, tool)

    record = {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "tag": tag,
        "model": model,
        "persona": persona_key,
        "item_id": item["id"],
        "item_title": item["title"][:60],
        "slide_num": slide_data["slide_num"],
        "slide_title": slide_data["title"][:60],
        "result": result,
    }
    append_jsonl(record)

    if result["ok"]:
        ev = result["evaluation"]
        cache_info = ""
        u = result["usage"]
        if u["cache_creation_input_tokens"] > 0:
            cache_info = f" [cache_create={u['cache_creation_input_tokens']}]"
        elif u["cache_read_input_tokens"] > 0:
            cache_info = f" [cache_read={u['cache_read_input_tokens']}]"
        print(
            f" score={ev['score']:+d} rationale={len(ev['rationale'])}字 "
            f"in={u['input_tokens']} out={u['output_tokens']}{cache_info} "
            f"({result['elapsed_sec']}s)"
        )
    else:
        print(f" 🚨 FAILED: {result['error']}")
    return record


def estimate_tokens(text: str) -> int:
    """簡易トークン推定（日本語は1文字≒1.5トークン、英数は1単語≒1.3トークン）"""
    return int(len(text) * 1.2)


def main():
    print("=" * 70)
    print("ko-NeKo IDチェッカー Day 1 サブセット観察スパイク")
    print(f"出力: {OUTPUT_JSONL}")
    print("=" * 70)

    # 初期化
    api_key = load_api_key()
    checklist = load_checklist()
    client = anthropic.Anthropic(api_key=api_key)
    tool = build_tool_definition()
    prs = Presentation(str(SAMPLE_PPTX))

    # ペルソナ・スライドの事前読み込み
    persona_texts = {p: load_persona(p) for p in PERSONAS}
    items = [find_item(checklist, iid) for iid, _ in EVAL_ITEMS]
    slides = {idx: load_slide_data(prs, idx) for _, idx in EVAL_ITEMS}

    # ─────────────────────────────────────────
    # 観察前プローブ: system プロンプトのサイズで cache 1024トークン閾値確認
    # ─────────────────────────────────────────
    print("\n[Probe] system プロンプトサイズ確認（cache 閾値1024トークン）")
    print("-" * 70)
    cache_ok = True
    for i, (item_id, _) in enumerate(EVAL_ITEMS):
        sample_persona_text = persona_texts["grade2_mai"]
        sample_system = build_system_prompt(sample_persona_text, items[i])
        est_tokens = estimate_tokens(sample_system)
        threshold_warn = "" if est_tokens >= 1024 else " ⚠️ 閾値割れリスク"
        print(f"  item={item_id}: system {len(sample_system)}文字 ≒ {est_tokens}tok{threshold_warn}")
        if est_tokens < 1024:
            cache_ok = False
    if not cache_ok:
        print("  ⚠️ いずれかの system が cache 閾値（1024）割れ。cache hit しない可能性")
    else:
        print("  ✅ 全 system が cache 閾値超え見込み")

    # JSONL 開始マーカー（再走時の区切り）
    append_jsonl(
        {
            "ts": datetime.now().isoformat(timespec="seconds"),
            "type": "session_start",
            "model_opus": MODEL_OPUS,
            "model_sonnet": MODEL_SONNET,
            "eval_items": EVAL_ITEMS,
            "personas": PERSONAS,
        }
    )

    all_records = []
    phase1_skipped = False

    # ─────────────────────────────────────────
    # Phase 1: 5項目 × 4ペルソナ（Opus）20回
    # ─────────────────────────────────────────
    print(f"\n[Phase 1] 5項目 × 4ペルソナ × Opus = 20回")
    print("-" * 70)
    for i, (item_id, slide_idx) in enumerate(EVAL_ITEMS):
        item = items[i]
        slide_data = slides[slide_idx]
        for persona_key in PERSONAS:
            rec = run_one(
                client=client,
                model=MODEL_OPUS,
                persona_key=persona_key,
                persona_text=persona_texts[persona_key],
                item=item,
                slide_data=slide_data,
                tool=tool,
                tag="P1_persona",
            )
            all_records.append(rec)

    # ─────────────────────────────────────────
    # Phase 1 → Phase 2 ゲート: fail 2件以上で Phase 2 中断
    # ─────────────────────────────────────────
    p1_fail_count = sum(1 for r in all_records if r["tag"] == "P1_persona" and not r["result"]["ok"])
    if p1_fail_count >= 2:
        print(f"\n🚨 [GATE] Phase 1 で {p1_fail_count} 件失敗、Phase 2 を中断します（構造的失敗の疑い）")
        phase1_skipped = True
    else:
        # ─────────────────────────────────────────
        # Phase 2: 5項目 × Sonnet（マイ固定）5回
        # → 既に Phase 1 で Opus×マイ×5項目 取れてるので Sonnet 5回だけ
        # ─────────────────────────────────────────
        print(f"\n[Phase 2] 5項目 × Sonnet × マイ固定 = 5回（Opus結果と並走比較）")
        print("-" * 70)
        persona_key = "grade2_mai"
        for i, (item_id, slide_idx) in enumerate(EVAL_ITEMS):
            item = items[i]
            slide_data = slides[slide_idx]
            rec = run_one(
                client=client,
                model=MODEL_SONNET,
                persona_key=persona_key,
                persona_text=persona_texts[persona_key],
                item=item,
                slide_data=slide_data,
                tool=tool,
                tag="P2_sonnet",
            )
            all_records.append(rec)

    # ─────────────────────────────────────────
    # 集計サマリ
    # ─────────────────────────────────────────
    print("\n" + "=" * 70)
    print("サマリ")
    print("=" * 70)

    ok_records = [r for r in all_records if r["result"]["ok"]]
    fail_records = [r for r in all_records if not r["result"]["ok"]]

    print(f"\n総数: {len(all_records)} / 成功: {len(ok_records)} / 失敗: {len(fail_records)}")

    # silent failure チェック
    if fail_records:
        print(f"\n🚨 silent failure 検知: {len(fail_records)}件")
        for r in fail_records:
            print(f"  - {r['tag']} {r['model']} {r['persona']} {r['item_id']}: {r['result']['error']}")

    # rationale 長分布
    rat_lens = [len(r["result"]["evaluation"]["rationale"]) for r in ok_records]
    if rat_lens:
        rat_lens_sorted = sorted(rat_lens)
        p50 = rat_lens_sorted[len(rat_lens_sorted) // 2]
        p90 = rat_lens_sorted[int(len(rat_lens_sorted) * 0.9)]
        print(f"\n📏 rationale文字数: min={min(rat_lens)} p50={p50} p90={p90} max={max(rat_lens)}")

    # トークン使用量・コスト試算
    total_in = sum(r["result"]["usage"]["input_tokens"] for r in ok_records)
    total_out = sum(r["result"]["usage"]["output_tokens"] for r in ok_records)
    total_cache_read = sum(
        r["result"]["usage"]["cache_read_input_tokens"] for r in ok_records
    )
    total_cache_create = sum(
        r["result"]["usage"]["cache_creation_input_tokens"] for r in ok_records
    )
    print(
        f"\n💰 トークン: in={total_in} out={total_out} "
        f"cache_create={total_cache_create} cache_read={total_cache_read}"
    )

    # ペルソナ別スコア分布（Phase 1のみ）
    print(f"\n📊 ペルソナ別スコア分布（Phase 1 Opus）:")
    p1 = [r for r in ok_records if r["tag"] == "P1_persona"]
    for persona_key in PERSONAS:
        records = [r for r in p1 if r["persona"] == persona_key]
        scores = [r["result"]["evaluation"]["score"] for r in records]
        print(f"  {persona_key}: scores={scores}")

    # Opus vs Sonnet 並走比較
    print(f"\n🔬 Opus vs Sonnet 並走（マイ × 5項目）:")
    print(f"  {'item_id':<8} {'Opus(score/len)':<20} {'Sonnet(score/len)':<20}")
    mai_opus = {
        r["item_id"]: r for r in ok_records
        if r["tag"] == "P1_persona" and r["persona"] == "grade2_mai"
    }
    mai_sonnet = {r["item_id"]: r for r in ok_records if r["tag"] == "P2_sonnet"}
    for item_id, _ in EVAL_ITEMS:
        o = mai_opus.get(item_id)
        s = mai_sonnet.get(item_id)
        o_str = (
            f"{o['result']['evaluation']['score']:+d} / {len(o['result']['evaluation']['rationale'])}字"
            if o else "—"
        )
        s_str = (
            f"{s['result']['evaluation']['score']:+d} / {len(s['result']['evaluation']['rationale'])}字"
            if s else "—"
        )
        print(f"  {item_id:<8} {o_str:<20} {s_str:<20}")

    # ─────────────────────────────────────────
    # GO/NO-GO 自動判定（事前定義基準）
    # ─────────────────────────────────────────
    print("\n" + "=" * 70)
    print("GO/NO-GO 自動判定（事前定義基準）")
    print("=" * 70)
    judgments = []

    # 1. silent failure 0件
    j1 = len(fail_records) == 0
    judgments.append(("silent_failure == 0", j1, f"fail={len(fail_records)}"))

    # 2. rationale max < 800トークン（≒ 800字 × 1.2 = 960トークン、800字で代理判定）
    rat_max = max(rat_lens) if rat_lens else 0
    j2 = rat_max < 800
    judgments.append(("rationale max < 800字", j2, f"max={rat_max}字"))

    # 3. ペルソナ別 rationale 差あり（マイとシンの完全一致を NG とする最弱判定）
    j3 = True
    p1_mai = {r["item_id"]: r["result"]["evaluation"]["rationale"]
              for r in ok_records if r["tag"] == "P1_persona" and r["persona"] == "grade2_mai"}
    p1_shin = {r["item_id"]: r["result"]["evaluation"]["rationale"]
               for r in ok_records if r["tag"] == "P1_persona" and r["persona"] == "grade1_shin"}
    if p1_mai and p1_shin:
        diff_count = sum(1 for k in p1_mai if k in p1_shin and p1_mai[k] != p1_shin[k])
        total = sum(1 for k in p1_mai if k in p1_shin)
        j3 = diff_count == total and total > 0
        judgments.append(("ペルソナ別 rationale 差あり", j3, f"マイvsシン {diff_count}/{total} 異なる"))
    else:
        j3 = False
        judgments.append(("ペルソナ別 rationale 差あり", False, "判定不能（マイ or シン欠損）"))

    # 4. cache hit が機能（cache_read > 0）
    j4 = total_cache_read > 0
    judgments.append(("cache hit 機能", j4, f"cache_read={total_cache_read}"))

    # 5. Phase 2 完走（中断していない）
    j5 = not phase1_skipped
    judgments.append(("Phase 2 完走（中断なし）", j5, "skipped" if phase1_skipped else "完走"))

    for criterion, ok, detail in judgments:
        mark = "✅" if ok else "❌"
        print(f"  {mark} {criterion}: {detail}")

    all_ok = all(j[1] for j in judgments)
    print()
    if all_ok:
        print("  🎯 GO_AUTO_DETECTED: 全基準クリア。本実装着手OK（最終判断はケンタ）")
    else:
        failed = [c for c, ok, _ in judgments if not ok]
        print(f"  🚫 NO_GO_AUTO_DETECTED: 未達基準 = {failed}")
        print("  → 該当基準について原因分析後、再判定 or 再設計")

    print("\n" + "=" * 70)
    print(f"詳細ログ: {OUTPUT_JSONL}")
    print("=" * 70)
    return 0 if (not fail_records and all_ok) else 1


if __name__ == "__main__":
    sys.exit(main())
