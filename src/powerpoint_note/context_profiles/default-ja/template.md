---
version: 1.0.0
language: Japanese
labels:
  local: 文字と構造の抽出結果です。図・配置・色の意味はAIで解釈していません。
  ai: 画像と抽出情報に基づくAIの説明です。原文と推論を区別して参照してください。
  personal: ここから下に手書きのメモを追記できます。
  evidence: 抽出根拠
  slide: スライド
  column: 列
  merged_cells: 結合セルの情報は抽出根拠のJSONに記録しています。
  chart: グラフ
  smartart: SmartArt
  object: 図形
  endpoints: 接続先・スタイル（意味は未解釈）
  text: テキスト
  location_outside: スライド外
  location_partial: 一部がスライド外
  location_unknown: 位置不明
  hidden_yes: はい
  hidden_no: いいえ
  unknown_value: 不明
---
### スライドのcontext（AI生成）

{{summary}}

{{#sections}}
{{sections}}
{{/sections}}

{{#uncertainties}}
#### 不確かな点

{{uncertainties}}
{{/uncertainties}}
