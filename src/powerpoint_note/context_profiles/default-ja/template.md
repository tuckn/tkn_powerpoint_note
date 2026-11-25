---
version: 2.0.0
language: Japanese
labels:
  personal: "ここから下に手書きのメモを追記できます。"
  slide: "スライド"
  column: "列"
  merged_cells: "結合セルを含む表です。セルの値は、結合の左上のセルに出力しています。"
  chart: "グラフ"
  smartart: "SmartArt"
  picture: "画像"
  location_outside: "スライド外"
  location_partial: "一部がスライド外"
  location_unknown: "位置不明"
  stale: "この説明の生成後に、スライドの内容が変わっています。`--context` を付けて実行すると更新します。"
  skipped: "このスライドの説明は生成していません"
---
#### スライドの説明（AI生成）

{{#stale}}
> [!WARNING]
> {{stale}}
{{/stale}}

{{summary}}

{{#sections}}
{{sections}}
{{/sections}}

{{#uncertainties}}
##### 不確かな点

{{uncertainties}}
{{/uncertainties}}

{{#image}}
##### 出典画像

{{image}}
{{/image}}
