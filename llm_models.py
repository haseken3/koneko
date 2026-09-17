"""ko-NeKo — 用途ごとの LLM モデルID集約。

ここに置く理由: どの用途にどのモデルが要るかは当て推量でなく実測で決める方針
（2026-09-17）。呼び出し側はモデルIDを直書きせずここを import し、実際の呼び出し関数は
いずれも model を引数で受け取る＝ベンチはこの既定値を上書きして回せる。

下の値は「確定」ではなく「既定」——2026-09-17 実測（n=1・7枚/40枚デッキ）で決めた値。
①は3モデルとも境界一致でsonnet-5に降格。②③はsonnet-5/haiku-4.5だと検出漏れ
（②: 手直し前デッキで不足所見が opus=4枚 vs sonnet=2枚・ゴール条件4枚未達／
③: S4「保証→保障」の誤変換を拾えたのは opus のみ、haiku・sonnetは0件で取りこぼし）
が出たため opus-5 に昇格。実測: `koneko/engawa/experiments/model_bench_20260917.py`・
`typo_model_bench_20260917.py`。降格発火点はケンタのテストデッキで①③を sonnet 並走
1回（`typo_model_bench_20260917.py` の DECKS 差し替え）——これ以外の追加ベンチは立てない。
"""

MODEL_SEGMENT = "claude-sonnet-5"          # ①ナレーション原稿のステップ意味分割
MODEL_MISMATCH = "claude-opus-5"           # ②スライド本文↔原稿の不一致検出
MODEL_TYPO = "claude-opus-5"               # ③誤字脱字検出
