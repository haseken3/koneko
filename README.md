# ko-NeKo

オンデマンド授業教材（PowerPoint）の制作支援ツール群。Streamlit Web アプリとして動作。

## 機能

### 1. ナレーションカウンター（`app.py`）

PPTX のノート欄に書かれたナレーション原稿の文字数をカウントし、想定動画尺を推定する。

- スライドごとの文字数・読み上げ時間・句読点ポーズを集計
- 全体の推定尺を分秒形式で表示
- 1分あたりの読み上げ文字数を調整可能

### 2. IDチェッカー（`idcheck/idcheck_app.py`）

PPTX をアップロードし、鈴木克明（2006）レイヤーモデル準拠の44項目チェックリストで自動評価する。Claude API による評価結果を Word フィードバックとしてダウンロード可能。

📘 **教員の方は [`docs/teacher_guide.md`](docs/teacher_guide.md) を参照してください**（説明書 + 仕様書）

- 学年（1〜4年次）に応じた学生像ペルソナで評価視点を切替
- LLM 評価可能な項目に絞った効率評価
- 評価結果は「要改善 → 中立 → 良好 → 失敗」のスコア別グルーピングで Word 出力
- 評価ペルソナの内部実装は教員に開示せず、年次選択 UI のみ表示

## セットアップ

### 依存関係
```
pip install -r requirements.txt
```

`requirements.txt` には Streamlit / python-pptx / python-docx / anthropic がピン留めされています。

### Anthropic API キー（IDチェッカーのみ）

ローカル実行時:
```
~/.config/koneko-idcheck/anthropic_api_key.txt
```
に API キーを保存（推奨 `chmod 600`）。

Streamlit Cloud デプロイ時:
- Streamlit Cloud の Secrets 設定で `anthropic_api_key` を設定

## 起動

### ナレーションカウンター
```
streamlit run app.py
```

### IDチェッカー
```
streamlit run idcheck/idcheck_app.py
```

## ライセンス

Internal use — 開志創造大学情報デザイン学部

## 注意

- IDチェッカーの評価結果は LLM による自動評価であり、教員の最終確認・修正の参考としてご活用ください
- 学年別ペルソナによる評価は「該当学年の代表的な学生像1名」を念頭にした視点であり、その学年全体の総意を代表するものではありません
