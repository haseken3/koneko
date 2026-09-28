"""
ko-NeKo IDチェッカー Day 0 ミニマムスパイク

目的: 1スライド × ペルソナ × 1項目で LLM 評価品質を確認する。
- API: anthropic SDK 0.84.0 / tool_use（function calling）
- 内側リトライ: max 3 回（schema違反時）
- silent failure 検知: リトライ尽きた時は ERROR 明示
- スコープ外: 段階分割 / prompt caching / Streamlit UI（本実装で扱う）

実行: /usr/bin/python3 run_spike.py
出力: 標準出力に評価結果（スコア + rationale + related_slides）
"""

import json
import os
import sys
import time
from pathlib import Path

import anthropic
from pptx import Presentation

# ─────────────────────────────────────────────
# 設定
# ─────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
IDCHECK_DIR = SCRIPT_DIR.parent
KONEKO_DIR = IDCHECK_DIR.parent

API_KEY_PATH = Path.home() / ".config" / "koneko-idcheck" / "anthropic_api_key.txt"
CHECKLIST_PATH = IDCHECK_DIR / "checklist.json"
PERSONA_PATH = IDCHECK_DIR / "personas" / "grade2_mai.md"
SAMPLE_PPTX = Path.home() / "Downloads" / "27-ソフトウェアアーキテクチャⅡ-04 -Rev01.pptx"

MODEL = "claude-opus-4-20250514"  # Phase A決定はピン留め前提、Opus 4系
# 上記モデル名は anthropic SDK 0.84.0 時点で最新の Opus。後でモデル更新時に bump
MAX_TOKENS = 4096
MAX_RETRIES = 3
SCOPE_ITEM_ID = "0.5"  # スパイク対象: 〔0.5〕説明（ナレーション・スライド）は、対象者レベルに合っているか？
SCOPE_SLIDE_INDEX = 2  # S3「Ⅰの復習」（0-indexed）


# ─────────────────────────────────────────────
# データ読み込み
# ─────────────────────────────────────────────
def load_api_key() -> str:
    if not API_KEY_PATH.exists():
        sys.exit(f"❌ APIキーが見つかりません: {API_KEY_PATH}")
    return API_KEY_PATH.read_text().strip()


def load_checklist_item(item_id: str) -> dict:
    data = json.loads(CHECKLIST_PATH.read_text(encoding="utf-8"))
    for item in data["items"]:
        if item["id"] == item_id:
            return item
    sys.exit(f"❌ checklist 項目が見つかりません: {item_id}")


def load_persona() -> str:
    return PERSONA_PATH.read_text(encoding="utf-8")


def load_slide_data(slide_index: int) -> dict:
    if not SAMPLE_PPTX.exists():
        sys.exit(f"❌ サンプルPPTXが見つかりません: {SAMPLE_PPTX}")
    p = Presentation(str(SAMPLE_PPTX))
    slides = list(p.slides)
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
                "item_id": {
                    "type": "string",
                    "description": "評価項目ID（〔-1.1〕等の形式）",
                },
                "score": {
                    "type": "integer",
                    "enum": [-1, 0, 1],
                    "description": "-1=要改善, 0=中立/判定保留, +1=良好",
                },
                "rationale": {
                    "type": "string",
                    "description": "評価根拠（スライドのどこを見てどう判定したか、具体的に）",
                },
                "related_slides": {
                    "type": "array",
                    "items": {"type": "integer"},
                    "description": "評価根拠となったスライド番号のリスト",
                },
            },
            "required": ["item_id", "score", "rationale", "related_slides"],
        },
    }


# ─────────────────────────────────────────────
# API 呼び出し（内側リトライ付き）
# ─────────────────────────────────────────────
def evaluate(client: anthropic.Anthropic, system: str, user: str, tool: dict) -> dict:
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            t_start = time.time()
            response = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=system,
                messages=[{"role": "user", "content": user}],
                tools=[tool],
                tool_choice={"type": "tool", "name": tool["name"]},
            )
            t_elapsed = time.time() - t_start

            # tool_use ブロックを取り出す
            for block in response.content:
                if block.type == "tool_use" and block.name == tool["name"]:
                    return {
                        "ok": True,
                        "attempt": attempt,
                        "elapsed_sec": round(t_elapsed, 2),
                        "evaluation": block.input,
                        "usage": {
                            "input_tokens": response.usage.input_tokens,
                            "output_tokens": response.usage.output_tokens,
                        },
                        "stop_reason": response.stop_reason,
                    }

            last_error = f"tool_use ブロックが応答に含まれない（stop_reason={response.stop_reason}）"
            print(f"  ⚠️ 試行{attempt}: {last_error}, リトライ", file=sys.stderr)

        except anthropic.APIError as e:
            last_error = f"APIError: {e}"
            print(f"  ⚠️ 試行{attempt}: {last_error}, リトライ", file=sys.stderr)
            time.sleep(2 * attempt)
        except Exception as e:
            last_error = f"{type(e).__name__}: {e}"
            print(f"  ⚠️ 試行{attempt}: {last_error}, リトライ", file=sys.stderr)
            time.sleep(2 * attempt)

    # silent failure 検知: リトライ尽きたら明示エラー
    return {
        "ok": False,
        "attempts": MAX_RETRIES,
        "error": last_error,
        "marker": "🚨 SILENT_FAILURE_AVOIDED: 評価未完了として上位に通知すべき",
    }


# ─────────────────────────────────────────────
# メイン
# ─────────────────────────────────────────────
def main():
    print("=" * 70)
    print("ko-NeKo IDチェッカー Day 0 ミニマムスパイク")
    print("=" * 70)

    # 1. 材料読み込み
    print("\n[1/4] 材料読み込み...")
    api_key = load_api_key()
    item = load_checklist_item(SCOPE_ITEM_ID)
    persona_text = load_persona()
    slide_data = load_slide_data(SCOPE_SLIDE_INDEX)

    print(f"  ✅ 項目: 〔{item['id']}〕 {item['title'][:50]}")
    print(f"  ✅ ペルソナ: grade2_mai（春波・マイ・イーリス、{len(persona_text)}文字）")
    print(f"  ✅ スライド S{slide_data['slide_num']}: {slide_data['title']}")
    print(f"     本文 {len(slide_data['body'])}文字 / ノート {len(slide_data['notes'])}文字")

    # 2. プロンプト構築
    print("\n[2/4] プロンプト構築...")
    system = build_system_prompt(persona_text, item)
    user = build_user_prompt(slide_data)
    tool = build_tool_definition()
    print(f"  ✅ system: {len(system)}文字 / user: {len(user)}文字 / tool: {tool['name']}")

    # 3. API 呼び出し
    print(f"\n[3/4] Claude API 呼び出し（モデル: {MODEL}）...")
    client = anthropic.Anthropic(api_key=api_key)
    result = evaluate(client, system, user, tool)

    # 4. 結果表示
    print("\n[4/4] 結果")
    print("-" * 70)
    if result["ok"]:
        ev = result["evaluation"]
        print(f"✅ 評価成功（試行{result['attempt']}回目、{result['elapsed_sec']}秒）")
        print(f"  項目: {ev['item_id']}")
        print(f"  スコア: {ev['score']}  ({'要改善' if ev['score'] == -1 else '中立' if ev['score'] == 0 else '良好'})")
        print(f"  関連スライド: {ev['related_slides']}")
        print(f"  rationale:")
        for line in ev["rationale"].split("\n"):
            print(f"    {line}")
        print(f"\n  トークン: input={result['usage']['input_tokens']} / output={result['usage']['output_tokens']}")
        print(f"  stop_reason: {result['stop_reason']}")
    else:
        print(f"🚨 評価失敗（{result['attempts']}回リトライ後）")
        print(f"  最終エラー: {result['error']}")
        print(f"  {result['marker']}")

    print("=" * 70)
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
