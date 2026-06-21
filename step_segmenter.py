"""ko-NeKo — ナレーション原稿のステップ意味分割（Opus tool-use）。

スライドのノート欄テキスト（ナレーション原稿）を Claude（Opus）に渡し、
「内容・テーマが大きく転換する箇所」で最大4ステップに分割する。
各ステップの開始スライド番号とラベルを返す。

設計（2026-06-22・ケンタの方向転換を反映）:
- スライドのデザイン/レイアウトでの判別は教員横断で約4割しか効かない（実測）。
  ノート欄テキストの文脈シフトで分けるのが唯一の汎用解。
- 1授業=最大4ステップ（大学運用ルール）。maxItems=4 をスキーマで縛り、
  さらにコード側でも truncate（LLMの「4を埋めようとする」バイアス＋外れ値対策）。
- プロンプトは「最大4は上限であって目標ではない／2〜3で十分なことも多い」と
  逆方向のキャンセル圧をかける（レン初稿の核）。
- モデルは claude-opus-4-8（最新Opus）。temperature/top_p 等は渡さない（Opus 4.8で400）。
  構造化出力は tool-use（tool_choice で強制）。
"""

import anthropic

MODEL = "claude-opus-4-8"
MAX_STEPS = 4
MAX_RETRIES = 3
MAX_TOKENS = 2000  # 出力は最大4ステップ＝小さい

# tool-use スキーマ（レン初稿）。maxItems=4 はスキーマ側の縛り（＋コード側でも truncate）。
SEGMENT_TOOL = {
    "name": "segment_lecture",
    "description": (
        "大学オンデマンド授業のナレーション原稿（スライドのノート欄）を読み、"
        "講義の内容が大きく転換する箇所で最大4ステップに分割する。"
        "各ステップの開始スライド番号と見出しラベルを返す。"
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "steps": {
                "type": "array",
                "description": "分割結果のステップ配列。1件以上、最大4件。先頭のstart_slideは必ず最初のスライド番号。",
                "minItems": 1,
                "maxItems": MAX_STEPS,
                "items": {
                    "type": "object",
                    "properties": {
                        "start_slide": {
                            "type": "integer",
                            "description": "このステップが始まるスライド番号（1始まり）。",
                            "minimum": 1,
                        },
                        "label": {
                            "type": "string",
                            "description": (
                                "このステップの内容を表す短い日本語見出し（15字以内目安）。"
                                "例: '導入・学習目標' / '16進数の基礎' / '演習と応用' / 'まとめ'"
                            ),
                        },
                    },
                    "required": ["start_slide", "label"],
                },
            },
            "rationale": {
                "type": "string",
                "description": "分割の根拠を1〜3文で。どのスライドで何が変わったかを具体的に。（品質確認用）",
            },
        },
        "required": ["steps", "rationale"],
    },
}

SYSTEM_PROMPT = """あなたは大学オンデマンド授業の教材構造を分析する専門家です。

この大学では1回の授業を「ステップ」という単位で構成しています：
- 1回の授業は最大4ステップで構成される（大学の運用ルール）
- 1ステップはおおむね10〜17分相当（目安であり絶対条件ではない）
- 各ステップは「学習者が一息ついて区切りを感じられる意味のまとまり」

あなたのタスクは、教員のナレーション原稿を読み、**内容・テーマが大きく転換する箇所**を
特定して、授業全体を最大4つのステップに分割することです。

【分割の判断基準（重要）】
内容の転換とは以下のいずれかを指します：
- 学習する概念・技術・テーマが別のものに切り替わる
- 「導入 → 本論 → 演習 → まとめ」のような授業フェーズが変わる
- 教員が明示的に「次は〜」「ここから〜」「続いて〜」と話題転換を宣言している
  （これは有力な手がかりだが、文字列だけで判断せず内容の転換と合わせて見ること）
- 前のスライド群とは前提知識や問題設定が大きく異なる内容に移行する

【やってはいけないこと】
- 細かいサブトピックの切り替わりでステップを刻まない
- 内容的なまとまりがあれば2〜3ステップで収まることもある（無理に4つにしない）
- ステップ数は「必要な切れ目の数」だけにする。最大4という上限を目標にしない

【ラベルの付け方】
- 15字以内の日本語で、そのステップで何を学ぶかが一言でわかるようにする
- 例: 「導入・学習目標」「16進数の基礎」「データ型と変換」「演習と復習」
- 抽象的な「Step1」「前半」のようなラベルは避ける

【出力形式】
必ず segment_lecture ツールを使って出力してください。
steps の最初の要素の start_slide は、原稿の最初のスライド番号にしてください。
steps の要素数は1以上、4以下です。"""


def build_slides_text(slides) -> str:
    """slides（analyze_narration の "slides"）を番号付きテキストに整形する。

    空のノート欄も省かず渡す（タイトル/区切りスライドは文脈の手がかりになる）。
    """
    parts = []
    for s in slides:
        title = (s.get("title") or "").strip() or "タイトルなし"
        note = (s.get("notes") or "").strip()
        parts.append(f"S{s['slide_num']} [{title}]:")
        parts.append(note if note else "（ノート欄が空のスライドはノート本文なし）")
        parts.append("")  # 空行で区切る
    return "\n".join(parts)


def _sanitize_steps(raw_steps, valid_nums) -> tuple:
    """LLM出力の steps を安全な (boundaries, labels) に正規化する。

    - 整数でない/範囲外/重複の start_slide を除外
    - 昇順ソート
    - 先頭は必ず最初のスライド番号（検出漏れ・ゴミ対策。nene指摘の [1,157,157,157] 等を吸収）
    - 最大 MAX_STEPS に truncate（スキーマに加えコード側でも物理的に縛る）
    """
    valid = set(valid_nums)
    first = min(valid_nums)
    seen = set()
    cleaned = []
    for s in raw_steps or []:
        n = s.get("start_slide") if isinstance(s, dict) else None
        if not isinstance(n, int) or n not in valid or n in seen:
            continue
        seen.add(n)
        label = (s.get("label") or "").strip() if isinstance(s, dict) else ""
        cleaned.append((n, label))
    cleaned.sort(key=lambda x: x[0])

    if not cleaned or cleaned[0][0] != first:
        cleaned.insert(0, (first, ""))  # 先頭を強制（ラベルは group 側で連番フォールバック）

    cleaned = cleaned[:MAX_STEPS]
    boundaries = [n for n, _ in cleaned]
    labels = [lbl for _, lbl in cleaned]
    return boundaries, labels


def segment_steps(slides, api_key, *, model: str = MODEL,
                  lecture_title: str = "") -> dict:
    """slides を Opus で意味分割し、境界とラベルを返す。

    Args:
        slides: analyze_narration の返値 ["slides"]（slide_num/title/notes を含む）
        api_key: Anthropic API キー
        model: 使用モデル（既定 claude-opus-4-8）
        lecture_title: 授業タイトル（任意・プロンプトの文脈に使う）

    Returns:
        dict: {"boundaries": list[int], "labels": list[str], "rationale": str}

    Raises:
        RuntimeError: リトライしても有効な分割が得られなかった場合
        anthropic.APIError: API 呼び出しが恒久的に失敗した場合
    """
    if not slides:
        return {"boundaries": [], "labels": [], "rationale": ""}

    valid_nums = [s["slide_num"] for s in slides]
    slides_text = build_slides_text(slides)
    user_prompt = (
        f"以下はオンデマンド授業"
        f"{('「' + lecture_title + '」') if lecture_title else ''}"
        f"のナレーション原稿です。\nスライド総数: {len(slides)}枚\n\n"
        f"---\n{slides_text}\n---\n\n"
        "上記の原稿全体を読んで、内容が大きく転換する箇所を見つけ、"
        "授業を最大4ステップに分割してください。"
        "segment_lecture ツールで結果を返してください。"
    )

    client = anthropic.Anthropic(api_key=api_key)

    # SDK は 429/5xx を自動リトライ（max_retries=2）。このアプリ層ループはそれでも
    # 落ちた場合に加え、「APIは成功したが中身が不正」（tool_use欠落・空境界）も
    # 拾い直すため意図的に被せている二重リトライ。失敗時は黙らず最後に raise。
    last_err = None
    for attempt in range(MAX_RETRIES):
        try:
            resp = client.messages.create(
                model=model,
                max_tokens=MAX_TOKENS,
                system=SYSTEM_PROMPT,
                tools=[SEGMENT_TOOL],
                tool_choice={"type": "tool", "name": "segment_lecture"},
                messages=[{"role": "user", "content": user_prompt}],
            )
            tool_use = next(
                (b for b in resp.content if b.type == "tool_use"
                 and b.name == "segment_lecture"),
                None,
            )
            if tool_use is None:
                last_err = "segment_lecture ツール呼び出しが返らなかった"
                continue
            data = tool_use.input
            boundaries, labels = _sanitize_steps(data.get("steps"), valid_nums)
            if not boundaries:
                last_err = "有効な境界が得られなかった"
                continue
            return {
                "boundaries": boundaries,
                "labels": labels,
                "rationale": (data.get("rationale") or "").strip(),
            }
        except (anthropic.APIStatusError, anthropic.APIConnectionError) as e:
            last_err = f"{type(e).__name__}: {e}"
            continue
        except (KeyError, TypeError) as e:
            last_err = f"{type(e).__name__}: {e}"
            continue

    raise RuntimeError(f"ステップ分割に失敗しました（{MAX_RETRIES}回試行）: {last_err}")
