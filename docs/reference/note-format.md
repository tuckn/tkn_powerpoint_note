# ノートの形式

`export` で書き出す Markdown と、`pull` で作成する代理ノートの Frontmatter・本文の構成を定義します。
ノートを読み取るスクリプトや Obsidian の Bases を作るときに参照してください。
使い方は [README](../../README.md) を参照してください。

## 書き出しと代理ノートの違い

| 観点 | 書き出し（`export`） | 代理ノート（`pull`） |
| --- | --- | --- |
| `type` | `PowerPointExport` | `PowerPoint` |
| `schemaVersion` | `"1.0"` | `"5.0.0"` |
| 保存先 | 既定は各 PowerPoint の隣。単体のファイルは `--output` で変更可能 | `sources.<id>.notes.dir` |
| 管理マーカー | なし | ツールが管理する本文を `powerpoint-note` マーカーで囲む |
| 更新 | `--force` で全体を置き換え | 管理マーカー内と、ツールが管理する Frontmatter だけを更新 |
| 手書きの部分 | 保持しません | 管理マーカーの外と、ツールが知らない Frontmatter 項目を保持 |

## 代理ノートの Frontmatter

| 項目 | 内容 | 更新 |
| --- | --- | --- |
| `type` | 常に `PowerPoint`（引用符なし） | ノート更新時 |
| `schemaVersion` | 代理ノートの形式の版。引用符付きの文字列 `"5.0.0"` | ノート更新時 |
| `title` | PowerPoint の文書プロパティのタイトル。空の場合は、拡張子を除いたファイル名 | `pull` のたびに PowerPoint の値で更新 |
| `subject`、`author`、`keywords`、`categories`、`comments` | PowerPoint の文書プロパティ（件名、作成者、キーワード、分類、コメント） | `pull` のたびに PowerPoint の値で更新 |
| `description` | ノート全体の概要 | 利用者 |
| `cover` | カード表示用の画像へのリンク。ツールは作成しません | 利用者 |
| `sourceId` | `<source の ID>:<入力フォルダからの相対パス>` | `pull` |
| `sourceRoot` | source の ID | `pull` |
| `sourceFileName` | 入力フォルダからの相対パス（区切りは `/`） | `pull` |
| `sourceFullPath` | 元の PowerPoint の絶対パス | `pull` |
| `sourceCreated` / `sourceModified` | PowerPoint の作成日時・更新日時 | `pull` |
| `contextStatus` | AI の説明の状態（`not-generated`・`current`・`partial`・`stale`） | `pull` |
| `tags` | ノートのタグ。空の場合は `[]` | 利用者 |
| `created` / `updated` | ノートの作成日時・更新日時。`updated` は、ノートの内容が変わったときだけ更新 | ノート更新時 |
| `noteId` | ノートの識別子（UUID）。作成後は変わりません | 作成時のみ |

AI の説明がある場合は、次の項目も書きます。

| 項目 | 内容 |
| --- | --- |
| `contextProfiles` | ノートにある説明と要約の生成に使った文章の profile 名の一覧（例: `[default-ja]`） |
| `contextAnalyzedSlides` | 説明があるスライド |
| `contextOmittedSlides` | 説明をまだ生成していないスライド |
| `contextStaleSlides` | 説明の生成後に内容が変わったスライド |
| `contextSourceFingerprint` | 資料の要約の生成時点の、PowerPoint の内容のハッシュ値。要約がない場合は現在の内容のハッシュ値 |
| `contextGeneratedAt` | 資料の要約の生成日時。要約がない場合は、最も新しい説明の生成日時 |

いずれも、直近の実行で選んだスライドだけでなく、ノートにある説明全体を表します。
スライドは、`3-6`、`1-2,7-400` のように、1から数えた位置の範囲で書きます。
該当するスライドがない場合は `""` です。
内容のハッシュ値は、文書プロパティと ZIP の格納方法を除いて計算するため、プロパティだけの変更では変わりません。

先頭は `type → schemaVersion → title → description → cover`、末尾は `tags → created → updated → noteId` に固定します。
中央は、文書情報、元ファイルの識別・所在、元ファイルの日時、AI 生成情報の順で、コメント行（`# --- ... ---`）と空行で区切ります。
利用者が追加した項目は、値を変えずに、末尾の共通項目の前に置きます。
階層構造は追加しません。
終端の `---` と本文の間には空行を入れます。

日時は `"2026-06-21T05:44:56+09:00"` のように、ダブルクォート付き・日本時間・秒単位に統一します。
時刻のない日付だけの値は維持します。
タイムゾーンのない日時と不正な日時はエラーにし、推測で補完しません。
`type` 以外の先頭4項目・日時・ハッシュ値は、値によらずダブルクォートで囲みます。
`sourceFullPath` はシングルクォートで囲みます。
その他の文字列は引用符なしで書き、`yes`・`1.0`・`2026-10-11` のように、YAML で真偽値・数値・日付として読まれる値と、`[[...]]` のリンクだけをダブルクォートで囲みます。
空文字列は `""`、空配列は `[]` です。

`keywords` と `categories` は、文書プロパティの値を `;` で区切った一覧です。
`sources.<id>.notes.frontmatter_term_format` が `obsidian-link` なら `[[...]]` 形式のリンク、`plain` なら通常の文字列です。

## 書き出しの Frontmatter

| 項目 | 内容 |
| --- | --- |
| `type` / `schemaVersion` | `PowerPointExport` / `"1.0"` |
| `title`、`subject`、`author`、`keywords`、`categories`、`comments` | PowerPoint の文書プロパティ。`keywords` と `categories` は通常の文字列の一覧 |
| `sourceFileName` | 拡張子付きのファイル名 |
| `sourceFullPath` | 元の PowerPoint の絶対パス |
| `sourceCreated` / `sourceModified` | PowerPoint の作成日時・更新日時 |
| `sourceSnapshotSha256` | 読み取ったファイルの SHA-256 ハッシュ値 |
| `generatedAt` | 書き出した日時 |
| `generationMethod` | AI の説明を含むなら `context`、含まないなら `text` |
| `contextStatus` と `context` で始まる項目 | 代理ノートと同じ |

`sourceId`、`sourceRoot`、`noteId` は持ちません。

## 本文の構成

`--context` で説明を生成した代理ノートは、次の構成になります。
書き出しの場合も同じ構成で、管理マーカーだけを除きます。

```markdown
# <タイトル>

<!-- powerpoint-note:begin -->
## Presentation Map

## 資料の要約（AI生成）
### 要点
### 不確かな点

## スライド

### スライド <番号>: <タイトル>
セクション: <セクション名>

#### スライドの説明（AI生成）
##### <内容に応じた見出し>
##### 不確かな点
##### 出典画像

#### Extracted Text
#### 発表者ノート
#### コメント
#### スライド外の要素
<!-- powerpoint-note:end -->

ここから下に手書きのメモを追記できます。
```

見出しの文言と言語は、文章の profile のテンプレートが定義します。
上記は既定の `default-ja` の場合です。

- スライドは、資料内の並び順に配置します。非表示のスライドは、`extraction.include_hidden` が `true` のときだけ含めます。
- 「資料の要約」と「スライドの説明」は、AI の説明があるときだけ出力します。「要点」には、根拠にしたスライドの番号を添えます。
- 「スライドの説明」の下の見出しは、スライドの内容に応じて AI が構成します。
- 説明の生成後に内容が変わったスライドは、説明の先頭に注意書き（`> [!WARNING]`）を付けます。
- 「出典画像」には、AI に送ったスライド画像へのリンクを置きます。画像は、ノートと同じフォルダの `img/<PowerPointのファイル名>/slides/<キー>/slide.png` にあり、相対パスで参照します。
- 「発表者ノート」「コメント」「スライド外の要素」は、該当する内容がなければ省略します。「スライド外の要素」は `extraction.off_slide: append` のときだけ出力します。
- `--context` で対象にしたのに説明を作らなかったスライドには、理由を引用（`>`）で示します。

### Extracted Text

「Extracted Text」は、AI を使わずに、スライドの上にある要素を上から下、左から右の順に並べます。
タイトルは見出しにあるため、繰り返しません。

```markdown
- 段落の文字
  - 字下げした段落
- 複数行の段落は、
  2行目以降を同じ項目の続きにします

| 列 1 | 列 2 |
| --- | --- |
| 表のセル | 表のセル |

- グラフ（barChart）: 売上

|  | 2026 |
| --- | --- |
| Q1 | 10 |

- SmartArt:
  - 図の中の文字
- 画像: 人が書いた代替テキスト
```

- 文字は原文のまま出力します。Markdown の構造を壊す箇所だけを書き換えます（行頭の `#`・`>`・`-`・番号、`<` で始まるタグ、表の中の `|` と改行）。
- 表の見出しは `列 1`、`列 2` とします。1行目が見出しかデータかを推測しないためです。結合セルがある表には、その旨を示します。
- グラフは、種類とタイトルに続けて、分類を行、系列を列とする表で、保存されている値を出力します。
- 画像は、人が書いた代替テキストがあるときだけ出力します。PowerPoint が自動で付けた説明と、ファイル名だけの代替テキストは出力しません。
- コネクタと、文字も代替テキストもない図形は出力しません。
- スライド番号のフィールドは `‹#›` と表示します。
- 一部がスライドの外にある要素には、「（一部がスライド外）」を添えます。

## Presentation Map

Presentation Map はタイトルの直下に置き、すべてのスライドの概要を最初に確認できる表です。
非表示のスライドも並べます。

| 列 | 意味 |
| --- | --- |
| `Slide` | 資料内の位置（1から） |
| `ID` | PowerPoint が持つスライドの ID。並べ替えても変わりません。 |
| `Title` / `Section` | タイトルとセクション名。ない場合は `—` |
| `State` | `visible` または `hidden` |
| `Text shapes` | 文字のある図形の数 |
| `Pictures` / `Tables` / `Charts` | 画像・表・グラフの数 |
| `Notes` | 発表者ノートがあれば `yes` |
| `Comments` | コメントと返信の数 |

## ノートの隣に保存するファイル

代理ノートと、`--context` を付けた書き出しは、ノートと同じフォルダの `img/<PowerPointのファイル名>/` に次のファイルを保存します。

| ファイル | 内容 |
| --- | --- |
| `state.json` | ノートの生成部分のハッシュ値、スライドごとの説明の記録、資料の要約の記録 |
| `slides/<キー>/slide.png` | AI に送ったスライド画像 |
| `slides/<キー>/request.json` | AI に送ったプロンプト・根拠データ・出力スキーマ |
| `slides/<キー>/response.json` | 検証した AI の出力 |
| `slides/<キー>/usage.json` | 呼び出しの記録（接続先、モデル、所要時間、成否） |
| `deck/<キー>/` | 資料の要約の `request.json`・`response.json`・`usage.json` |

`<キー>` は、再利用の判定に使う値の先頭16桁です。
ノートが参照しなくなったフォルダは、成功した実行の最後に削除します。
出力が検証に通らなかった場合は、返された内容を `rejected.json` に残します。
