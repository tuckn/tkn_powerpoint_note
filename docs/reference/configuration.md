# 設定リファレンス

`tkn-powerpoint-note` の設定ファイルの場所、読み込む順序、すべての設定キーと既定値をまとめます。
初めて設定ファイルを作る手順は、[README の「設定ファイルを作る」](../../README.md#41-設定ファイルを作る)を参照してください。

## 設定ファイルの場所と優先順位

設定は次の順に読み込み、後に読み込んだ値を優先します。

1. 組み込みの既定値。
2. ユーザー共通の設定 `~/.tkn/powerpoint_note/config.yaml`。`config init` で作成します。
3. コマンドを実行したフォルダの `.tkn/config.yaml`。
4. `--config` で指定したファイル。
5. 個別の CLI オプション（`--profile`、`--max-slides`、`--include-hidden` など）。

`generation` や `genai_profiles` のような項目の集まりは、ファイルをまたいで項目単位で上書きします。
一方、`sources` は、後のファイルに書くと全体を置き換えます。
同じ source の設定を複数のファイルに分けて書くことはできません。
`include` / `ignore` / `prompt_profile_dirs` のような一覧も、後の値で全体を置き換えます。

設定ファイル内の相対パスは、コマンドを実行したフォルダを基準に解決します。
`genai_profiles` の `cli.executable` だけは、その値を書いた設定ファイルのフォルダが基準です。

読み込まれた設定ファイルと、有効な値は次のコマンドで確認できます。
1行に1つの `key=value` を表示し、入れ子の項目は `.`、配列は `[0]` などで表します。

```shell
tkn-powerpoint-note config list
tkn-powerpoint-note config list --source library
tkn-powerpoint-note config list --generator my-codex-en --json
```

| 出力項目 | 内容 |
| --- | --- |
| `config` | 解決済みの設定です。 |
| `loaded_sources` | 読み込んだ各ファイルの絶対パスと `schema_version` を、優先順位の低い順に並べます。 |
| `winning_sources` | 各設定値の決定元です。ファイルの絶対パス、`built-in`、`CLI` のいずれかです。 |
| `effective_schema_version` | 内部で使用する設定の形式の版です。 |
| `selected_generator` / `selected_prompt_profile` / `selected_genai_profile` | 選ばれた generator と、その generator が選ぶ文章の profile・接続先です。`--source` を付けると、その source の設定で選びます。 |
| `prompt_profile_source` | 文章の profile を読み込んだ場所です。 |

`config list` は設定ファイルを作成・更新しません。

## コマンドが使う設定

| コマンド | 使う設定 |
| --- | --- |
| `export` | `extraction` と、`--context` を付けた場合の `generation`・`genai_profiles`。`sources` は使いません。 |
| `pull` | `sources`、`extraction`。`--context` を付けた場合は `generation`・`genai_profiles` も使います。 |
| `status` | `sources`、`extraction` |
| `presentation list-slides` | `extraction` |

## source の設定

source は、入力する PowerPoint 側（`presentations`）と代理ノート側（`notes`）を組にした取り込み対象です。
`sources.<id>` の `<id>` が、`--source` で指定する名前になります。

| 設定キー（`sources.<id>.` 以下） | 既定値 | 意味 |
| --- | --- | --- |
| `presentations.dir` | なし（必須） | 入力する PowerPoint のフォルダです。 |
| `presentations.recursive` | `false` | `true` にすると、サブフォルダの PowerPoint も対象にします。 |
| `presentations.include` | `['*.pptx', '*.pptm']` | 対象にするファイルのパターンです。 |
| `presentations.ignore` | `[]` | 対象から外すファイルのパターンです。 |
| `notes.dir` | なし（必須） | 代理ノートを保存するフォルダです。 |
| `notes.frontmatter_term_format` | `obsidian-link` | `keywords` と `categories` の書式です。`obsidian-link` は `[[...]]` のリンク、`plain` は通常の文字列です。 |
| `generation.generator` | なし | この source で使う generator です。 |

`include` と `ignore` には、入力フォルダからの相対パスのパターンを、区切り文字 `/` で書きます。
大文字と小文字は区別しません。
`/` を含まないパターン（`*.pptx` など）は、どのフォルダのファイル名にも当てはめます。
`ignore` に `archive/**` を指定すると、`archive` フォルダ以下を除外します。
PowerPoint が編集中に作る一時ファイル（`~$` で始まるファイル）は、常に対象外です。

`recursive: true` の場合、入力フォルダのフォルダ構成を、代理ノートの保存先にも再現します。
入力フォルダの `2026/deck.pptx` は、`notes.dir` の `2026/deck.pptx.md` になります。

```yaml
sources:
  library:
    presentations:
      dir: 'C:\path\to\presentations'
      recursive: true
      include:
        - "*.pptx"
      ignore:
        - "archive/**"
    notes:
      dir: 'C:\path\to\notes'
      frontmatter_term_format: plain
    generation:
      generator: my-codex-def
```

Windows のパスをシングルクォートで囲むと、バックスラッシュを二重にせずに書けます。

## 抽出の設定

| 設定キー（`extraction.` 以下） | 既定値 | 意味 | CLI オプション |
| --- | --- | --- | --- |
| `include_hidden` | `false` | 非表示のスライドをノートに含めます。 | `--include-hidden` / `--no-include-hidden` |
| `off_slide` | `exclude` | 完全にスライドの外にある要素の扱いです。`append` で、スライドごとの別の欄に含めます。 | `--off-slide` |
| `include_notes` | `true` | 発表者ノートを抽出します。 | `--notes` / `--no-notes` |
| `include_comments` | `true` | コメントを抽出します。 | `--comments` / `--no-comments` |
| `max_file_mb` | `100` | 読み取るファイルの大きさの上限です。単位は MiB です。 | なし |
| `max_uncompressed_mb` | `1024` | ファイルを展開した大きさの上限です。 | なし |

一部だけがスライドの外にはみ出す要素は、文字をすべて残し、「一部がスライド外」と示します。
抽出の設定を変えると、AI に送る根拠データが変わるため、該当するスライドの説明は次の `--context` で生成し直します。

## AI 生成の設定

トップレベルの `generation` は、`export --context` と `pull --context` で使います。
AI の接続先・モデル・認証は、同じファイルの `genai_profiles` で管理します。
GenAI Bridge の共有設定 `~/.tkn/genai_bridge/config.yaml` は読み込みません。

```yaml
genai_profiles:
  codex-default:
    provider: codex
    model: null
  local-vision:
    provider: ollama
    model: example-vision-model
    local_only: true
```

組み込みの接続先は `codex-default`、`claude-default`、`copilot-default`、`antigravity-default` です。
各接続先の項目は、固定している版の GenAI Bridge の `Profile` に従い、未知の項目や不正な値を拒否します。
接続先は名前・項目単位で重ね合わせます。
秘密の値は直接書かず、認証済みの CLI や、環境変数の名前を指定します。
未定義の接続先は、`--context` の実行前にエラーとなり、ほかの接続先へ切り替えません。

### 名前付き generator

generator は、AI の接続先（`genai_profiles`）、文章の profile、接続先の値の上書きを組にして名前を付けた設定です。

```yaml
generation:
  default_generator: my-codex-def
  generators:
    my-codex-def:
      genai_profile: codex-default
      prompt_profile: default-ja
      overrides:
        timeout_seconds: 600
    my-codex-en:
      genai_profile: codex-default
      prompt_profile: default-en
```

| 設定キー（`generation.generators.<id>.` 以下） | 省略時の値 | 意味 |
| --- | --- | --- |
| `genai_profile` | `codex-default` | `genai_profiles` にある接続先の名前です。 |
| `prompt_profile` | `default-ja` | 文章の profile です。 |
| `overrides` | `{}` | 接続先の値を、この generator だけで上書きします。`model`・`reasoning_effort`・`timeout_seconds`・`max_output_tokens`・`local_only` を指定できます。 |

使う設定は次の規則で決まります。

| 対象 | 決まり方 |
| --- | --- |
| generator | `--generator` → `sources.<id>.generation.generator` → `generation.default_generator` の順で、最初に指定されたものを使います。どれもなければ、`generation` 直下の `genai_profile`・`prompt_profile`・`overrides` を使います。 |
| 文章の profile | `--profile` を最優先します。指定がなければ、選んだ generator の `prompt_profile` を使います。 |

`export` は source の設定を使わないため、generator は `--generator` または `generation.default_generator` で決まります。
`--generator`、`--profile`、`--slides`、`--section`、`--max-slides` は、`--context` と一緒に指定します。

### 上限

| 設定キー（`generation.limits.` 以下） | 既定値 | 意味 |
| --- | --- | --- |
| `max_slides` | `10` | 1回の実行で AI に送るスライドの数の上限です。資料ごとに数え、再利用するスライドは数えません。超える資料は、AI を呼ぶ前にエラーにします。`--max-slides` で、その実行だけ変えられます。 |
| `max_input_chars` | `120000` | AI に送る根拠データ（JSON）の文字数の上限です。スライドごとに適用します。 |
| `image_width` | `2400` | AI に送るスライド画像の幅（ピクセル）です。600～8000 の整数を指定します。 |
| `max_render_seconds` | `600` | 1つの資料のスライドを PowerPoint で画像にする処理の時間の上限です。単位は秒です。 |

数値は正の整数で指定します。
上限を超えた場合でも、値を途中で切り捨てません。

- `max_input_chars` を超えたスライドは、AI を呼ばずにそのスライドだけを飛ばし、残りのスライドの生成を続けます。
- 資料の要約に渡すスライドの説明の合計が `max_input_chars` を超えた場合は、要約だけを作らず、結果の `warnings` に示します。
- `max_slides`・`max_input_chars`・`max_render_seconds` を変えても、生成済みの説明は再利用します。
- `image_width` を変えると、AI に渡す画像が変わるため、次の `--context` で生成し直します。

### 文章の profile

| 設定キー（`generation.` 以下） | 既定値 | 意味 |
| --- | --- | --- |
| `prompt_profile` | `default-ja` | generator を使わない場合の文章の profile です。英語は `default-en` です。 |
| `prompt_profile_dirs` | `[]` | 独自の文章 profile を置いた親フォルダの一覧です。記載順に探します。 |

`prompt_profile_dirs` に `C:\path\to\profiles` を指定して `prompt_profile: technical-notes` を選ぶと、`C:\path\to\profiles\technical-notes\` を読み込みます。
同名のフォルダがなければ、同梱の `default-ja` / `default-en` から探します。
文章の profile は、見出しと表示文言も持つため、`--context` を付けない実行でも読み込み、検証します。
ファイルの構成と作り方は「[独自の文章 profile を作る](../guides/slide-context.md#独自の文章-profile-を作る)」を参照してください。

## 設定ファイルの版

`schema_version` は設定ファイルの形式の版で、現在は `"4.0.0"` です。
省略した場合も `"4.0.0"` として扱います。
ほかの版と、このページにないキーはエラーになります。
以前の版の設定ファイルは読み替えないため、`config init` で作り直してください。
