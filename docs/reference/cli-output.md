# CLI 出力契約

コマンドの結果の形式、項目、対象ごとの状態、終了コードを定義します。
スクリプトから `tkn-powerpoint-note` を呼ぶときに参照してください。

## 形式

- 標準出力には、結果を1つだけ出力します。既定は1行ごとの `key=value`、`--json` を付けると1行の JSON です。
- 進捗・警告・診断は標準エラーに出力します。`--quiet` / `--verbose` / `--no-color` が変えるのは、標準エラーのログだけです。
- 形式は、端末・パイプ・リダイレクトのどれに出力しても変わりません。処理の内容と終了コードも、形式によって変わりません。
- `--json`、`--config`、`--quiet`、`--verbose`、`--no-color` は、コマンドの前後のどちらにも指定できます。`--` の後ろは位置引数として扱います。
- `--help` と `--version` は、通常の案内を出力します。

`key=value` の表示では、入れ子を `.`、配列を `[0]` のような番号でつなぎます。
空の配列・項目は `[]` / `{}`、真偽値は `true` / `false`、未設定値は `null` です。
文字列は引用符で囲まず、Windows のパスの `\` は重ねません。
改行・タブなどの制御文字だけをエスケープします。

## 結果の項目

| コマンド | 主な項目 |
| --- | --- |
| `pull` / `status` | `status`、`command`、`dryRun`、`files`、`changed`、`errors`、`statusCounts`、`aiCalls`、`plannedAiCalls`、`results`。`pull` は `sources` も出力します。 |
| `export` | `status`、`command`、`dryRun`、`files`、`exported`、`planned`、`errors`、`statusCounts`、`aiCalls`、`plannedAiCalls`、`results` |
| `presentation list-slides` | `status`、`command`、`presentationPath`、`totalSlides`、`hiddenSlides`、`sections`、`slides`、`warnings` |
| `config init` | `status`（`created` / `unchanged`）、`command`、`dryRun`、`configPath` |
| `config list` | `status`、`command`、`config`、`loaded_sources`、`winning_sources` など（[設定リファレンス](configuration.md)） |
| エラー | `status`、`command`、`message` と、分かる範囲の `configPath`、`inputPath`、`outputPath`、`presentationPath`、`sourceId`、`targetPath` |

全体の `status` は、失敗した対象が1つでもあれば `error`、なければ `success` です。
`files` は処理した PowerPoint の数、`changed` は作成・更新したノートの数、`errors` は失敗した対象の数です。
`aiCalls` は実際に行った AI の呼び出し回数、`plannedAiCalls` は `--dry-run` で見積もった回数です。

### 対象ごとの結果（`results`）

| 項目 | 内容 |
| --- | --- |
| `status` | 下の表の状態 |
| `sourceRoot` / `sourcePath` / `presentationPath` / `notePath` | `pull`・`status` で出力します。source の ID、入力フォルダからの相対パス、PowerPoint とノートの絶対パス |
| `inputPath` / `outputPath` | `export` で出力します。PowerPoint と Markdown の絶対パス |
| `message` | 失敗・保護・要確認の理由 |
| `slides` | ノートに含めたスライドの数 |
| `contextStatus` / `staleSlides` | ノートの AI の説明の状態と、内容が変わったスライド |
| `aiCalls` / `plannedAiCalls` | この対象の呼び出し回数と、見積もった回数 |
| `generatedSlides` / `reusedSlides` | AI で生成したスライドの数と、保存済みの結果を再利用したスライドの数 |
| `skippedSlides` | 説明を作らなかったスライドと理由（`<番号>: <理由>` の一覧） |
| `warnings` | 資料の要約を作らなかった理由など |
| `backupPath` | `--force` で置き換える前のノートを控えた場所 |

| `status` | 意味 | 終了コードへの影響 |
| --- | --- | --- |
| `created` / `updated` | 代理ノートを作成・更新しました。 | なし |
| `unchanged` | ノートは最新で、何も書き込んでいません。 | なし |
| `would-create` / `would-update` | `--dry-run`・`status` で、作成・更新する予定です。 | なし |
| `exported` / `would-export` | `export` で書き出しました・書き出す予定です。 | なし |
| `missing-source` | 代理ノートに対応する PowerPoint が見つかりません。ノートはそのままです。 | なし |
| `protected` | ノートの生成部分が手で直されているため、更新しませんでした。 | 1 |
| `error` / `export-error` | 読み取り・生成・書き込みに失敗しました。 | 1 |

## 終了コード

| 終了コード | 全体の `status` | 意味 |
| --- | --- | --- |
| `0` | `success` など | 成功しました。 |
| `1` | `error` | 一部またはすべての対象で失敗しました。予期しないエラーも含みます。 |
| `2` | `argument-error` | コマンドの引数に誤りがあります。 |
| `3` | `config-error` | 設定ファイル、source の指定、入力ファイルの指定に誤りがあります。 |
| `130` | `interrupted` | 中断しました。完了したノートと、保存済みの AI の結果は残ります。 |

## 呼び出しの例

人が読む場合は既定の表示を使い、スクリプトでは `--json` の標準出力と終了コードで判定します。

```shell
tkn-powerpoint-note pull --source library --dry-run
tkn-powerpoint-note --json pull --source library --dry-run
```

表示形式の指定は、処理の内容、`--dry-run` の範囲、保存するノートと `img/` のファイルに影響しません。
