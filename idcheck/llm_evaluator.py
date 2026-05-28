"""LLM 評価モジュール — 44項目チェックリストでスライドを評価する。

spike (run_spike_day1.py) を core 化したもの。

# Day 1 観察（2026-05-28）で見えた構造シグナル — 本実装に未反映（Phase A 完成後に見直し予定）
# - 項目 0.6 マイ偏向: マイだけ +1、他3ペルソナ -1（ペルソナ定義のバイアス疑い）
# - 項目 1.3 ペルソナ間スコア方向バラバラ（評価項目の安定性疑い）
# - Sonnet rationale が Opus の 1.5〜3 倍長（コスト試算前提が崩れる可能性）
# - cache_read=0: 現設計（ペルソナを system 内包）の構造帰結。cache hit は構造修正が必要
"""

import json
import time
from pathlib import Path
from typing import Optional

import anthropic

MODEL_OPUS = "claude-opus-4-20250514"
MODEL_SONNET = "claude-sonnet-4-5-20250929"
MODEL_DEFAULT = MODEL_OPUS
MAX_TOKENS = 4096
MAX_RETRIES = 3

# 入力サイズ事前判定の閾値（Day 1 のサンプル教材で full prompt ≒ 4700 tok / item）
# 全スライド本文+ノートの合計文字数 × 1.2 (≒トークン推定) が以下を超えたら警告/中断
INPUT_TOKEN_BUDGET = 180_000  # Opus 200K の余裕枠

IDCHECK_DIR = Path(__file__).resolve().parent
CHECKLIST_PATH = IDCHECK_DIR / "checklist.json"
PERSONAS_DIR = IDCHECK_DIR / "personas"


def load_checklist() -> dict:
    return json.loads(CHECKLIST_PATH.read_text(encoding="utf-8"))


def load_persona(persona_key: str) -> str:
    return (PERSONAS_DIR / f"{persona_key}.md").read_text(encoding="utf-8")


def list_personas() -> list[str]:
    return sorted(p.stem for p in PERSONAS_DIR.glob("*.md") if p.stem != "README")


def build_system_prompt(persona_text: str, item: dict) -> str:
    return f"""あなたはオンデマンド授業教材のIDチェッカーです。

レイヤーモデル（鈴木克明, 2006）に基づく評価項目1つについて、
教材スライドを、特定の学生像を念頭に置いて評価してください。

# 評価軸
- 項目ID: {item['id']}
- レベル: {item['level']}
- 評価観点: {item['title']}
- 補足: {item.get('note', '')}

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
- rationale ではペルソナの固有名（マイ／シン／ケン／アヤ等）を出さず、「学生」「この学生」「対象学生」のように一般化して書く
"""


def build_user_prompt_all_slides(slides: list[dict]) -> str:
    parts = ["# 評価対象教材（全スライド）\n"]
    for s in slides:
        parts.append(f"\n## スライド {s['slide_num']}: {s['title']}\n")
        if s["body"]:
            parts.append(f"### 本文\n{s['body']}\n")
        if s["notes"]:
            parts.append(f"### ナレーション（ノート欄）\n{s['notes']}\n")
    parts.append(
        "\n---\n\n上記の全スライドを参照し、システムプロンプトで指定された評価項目について"
        "教材全体として評価し、submit_evaluation ツールで結果を返してください。"
        "related_slides には該当スライド番号を入れてください（全体該当なら全番号、特定の数枚なら該当のみ）。"
    )
    return "\n".join(parts)


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


def evaluate_one(
    client: anthropic.Anthropic,
    model: str,
    item: dict,
    slides: list[dict],
    persona_text: str,
    tool: Optional[dict] = None,
) -> dict:
    """1項目を全スライドに対して評価する。

    Returns:
        dict: ok / evaluation / usage / error などを含む
    """
    if tool is None:
        tool = build_tool_definition()

    system = build_system_prompt(persona_text, item)
    user = build_user_prompt_all_slides(slides)

    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            t_start = time.time()
            response = client.messages.create(
                model=model,
                max_tokens=MAX_TOKENS,
                system=[{
                    "type": "text",
                    "text": system,
                    "cache_control": {"type": "ephemeral"},
                }],
                messages=[{"role": "user", "content": user}],
                tools=[tool],
                tool_choice={"type": "tool", "name": tool["name"]},
            )
            elapsed = time.time() - t_start

            for block in response.content:
                if block.type == "tool_use" and block.name == tool["name"]:
                    usage = response.usage
                    # score を int 正規化（LLM が tool schema を無視して str 返す事案対策）
                    evaluation = dict(block.input)
                    raw_score = evaluation.get("score")
                    if isinstance(raw_score, str):
                        try:
                            evaluation["score"] = int(raw_score.strip().lstrip("+"))
                        except (ValueError, AttributeError):
                            evaluation["score"] = 0
                    elif not isinstance(raw_score, int):
                        evaluation["score"] = 0
                    return {
                        "ok": True,
                        "attempt": attempt,
                        "elapsed_sec": round(elapsed, 2),
                        "evaluation": evaluation,
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
                    }

            last_error = f"tool_use ブロックなし（stop_reason={response.stop_reason}）"

        except (
            anthropic.APIConnectionError,
            anthropic.APITimeoutError,
            anthropic.RateLimitError,
            anthropic.InternalServerError,
        ) as e:
            last_error = f"{type(e).__name__}: {e}"
            time.sleep(2 * attempt)
        except (
            anthropic.BadRequestError,
            anthropic.AuthenticationError,
            anthropic.PermissionDeniedError,
            anthropic.NotFoundError,
        ) as e:
            return {
                "ok": False,
                "attempts": attempt,
                "error": f"{type(e).__name__}: {e}",
                "marker": "🚨 NON_RETRIABLE_ERROR",
            }
        except Exception as e:
            return {
                "ok": False,
                "attempts": attempt,
                "error": f"UnknownError({type(e).__name__}): {e}",
                "marker": "🚨 UNKNOWN_ERROR",
            }

    return {
        "ok": False,
        "attempts": MAX_RETRIES,
        "error": last_error,
        "marker": "🚨 SILENT_FAILURE_AVOIDED",
    }


def evaluate_all(
    client: anthropic.Anthropic,
    slides: list[dict],
    persona_key: str,
    model: str = MODEL_DEFAULT,
    progress_callback=None,
    item_filter=None,
) -> list[dict]:
    """全評価項目をループして評価する。

    Args:
        client: anthropic クライアント
        slides: pptx_reader.read_slides() の戻り値
        persona_key: ペルソナのキー（例: 'grade2_mai'）
        model: 使用モデル
        progress_callback: (idx, total, item, result) を受ける callable（Streamlit用）
        item_filter: 評価項目を絞り込む callable (item -> bool)。None なら全項目

    Returns:
        list[dict]: 各項目の評価結果（item情報 + result）
    """
    checklist = load_checklist()
    persona_text = load_persona(persona_key)
    tool = build_tool_definition()

    # 入力サイズ事前判定（44項目連鎖 BadRequestError 回避）
    total_chars = sum(len(s["body"]) + len(s["notes"]) for s in slides)
    est_tokens = int(total_chars * 1.2)
    if est_tokens > INPUT_TOKEN_BUDGET:
        raise ValueError(
            f"教材サイズが大きすぎます: 推定 {est_tokens:,} tok > 予算 {INPUT_TOKEN_BUDGET:,} tok "
            f"(本文+ノート {total_chars:,}字)。スライド分割または評価項目絞り込みを検討してください。"
        )

    items = checklist["items"]
    if item_filter:
        items = [it for it in items if item_filter(it)]

    results = []
    total = len(items)
    for idx, item in enumerate(items, start=1):
        result = evaluate_one(client, model, item, slides, persona_text, tool)
        record = {"item": item, "result": result}
        results.append(record)
        if progress_callback:
            progress_callback(idx, total, item, result)

    return results


def llm_evaluable_only(item: dict) -> bool:
    """teacher_self_check=false かつ llm_evaluable=true の項目だけ通すフィルタ。"""
    return item.get("llm_evaluable", False) and not item.get("teacher_self_check", False)
