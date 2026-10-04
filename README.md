# tkn-powerpoint-note: Tkn PowerPoint Note — PowerPoint を Markdown にする

PowerPoint の内容を、人と生成 AI が理解・参照できる Markdown ノートにします。
通常は文字と構造をローカルで抽出し、`--context` を付けるとスライド画像から図・矢印・配置・色の意味を文章化します。

たとえば、構成図に描かれたシステム同士の関係とその意味を説明し、元のスライド番号へ戻れるノートを作れます。
抽出した原文、発表者ノート、コメント、AI の説明は別の見出しにまとめます。
ファイル単位の変換ツールであり、`sources` の登録、`pull/push` 同期、PowerPoint への書き戻しは実装していません。

## インストールする

Python 3.11 以降と uv を用意します。
以下のフォルダを実際のリポジトリへ置き換えて実行してください。

```shell
cd "C:\path\to\tkn_powerpoint_note"
uv tool install .
tkn-powerpoint-note --help
tkn-powerpoint-note --version
```

バージョン番号が表示されれば、インストールを確認できます。
通常の文字抽出には、PowerPoint 本体や AI の契約・設定は不要です。
GenAI Bridge と、Windows 用の画像化に必要な Python 依存関係は、`uv tool install .` でまとめてインストールされます。
視覚的な context を生成する場合は、Windows とデスクトップ版 Microsoft PowerPoint を用意します。

画像を扱える [GenAI Bridge](https://github.com/tuckn/tkn_genai_bridge) の接続設定と認証も必要です。
接続先・モデルは `~/.tkn/genai_bridge/config.yaml` に設定します。
この CLI の `generation.generators.<id>.bridge_profile` はその接続設定を選びます。
`prompt_profile` は、言語・プロンプト・出力構造をまとめたプロファイルを選びます。

## 最初のノートを作る

PowerPoint を保存してから、スライド番号とセクション名を確認し、ノートを生成・検証します。

```shell
tkn-powerpoint-note inspect "C:\path\to\deck.pptx"
tkn-powerpoint-note export "C:\path\to\deck.pptx" --output "C:\path\to\deck.pptx.md"
tkn-powerpoint-note verify "C:\path\to\deck.pptx.md"
```

`export` は Markdown と、同じフォルダに根拠情報を保存するフォルダを作成します。
`--output` を省略すると、`~/.tkn/powerpoint_note/data/` の下へ、元ファイルのパスごとに分けて保存します。
原本の PowerPoint は常に読み取り専用です。

`--dry-run` を付けると、入力・選択範囲・設定・既存ノートの保護条件を確認し、変更予定だけを表示します。
ファイル作成、PowerPoint の起動、認証、通信、AI 呼び出しは行いません。
標準出力には結果の JSON（`config list` は項目ごとの表示）、標準エラーには進捗を出力します。
成功時は終了コード 0、失敗時は非 0 になります。

## 対象スライドを選び、context を生成する

```shell
tkn-powerpoint-note export "C:\path\to\deck.pptx" --slides "2,5-8" --output "C:\path\to\selection.md"
tkn-powerpoint-note export "C:\path\to\deck.pptx" --section "Architecture" --slides "1-20"
tkn-powerpoint-note export "C:\path\to\deck.pptx" --slides "2-3" --context --dry-run
tkn-powerpoint-note export "C:\path\to\deck.pptx" --slides "2-3" --context
```

スライド番号は、ファイル内の並び順を 1 から数えた位置です。
範囲の両端を含み、複数のセクションは `--section` を繰り返して指定します。
セクション名は完全一致で、ページ範囲を同時指定すると両方に該当するスライドを選びます。

非表示スライドは既定で除外し、`--include-hidden` で含められます。
完全にスライド外にある要素も既定で除外し、`--off-slide append` で別枠の補足として含められます。
一部だけはみ出す要素は全文を残し、そのことを明示します。

> [!IMPORTANT]
> `--context` は、選択スライドの画像・抽出文字・有効な発表者ノートやコメントを設定済み AI 接続先へ送信し、費用や利用枠を消費する場合があります。
> 未生成の場合、スライドごとに 1 回、選択範囲の統合説明に 1 回呼び出します。
> 既定の上限は 10 スライドです。上限超過時は開始前に停止し、対象を勝手に切り詰めません。

画像化するのも選択したスライドだけです。
スライド外の補足は抽出情報として渡し、画像には含めません。
対応範囲と制約は [処理・出力仕様](docs/reference/behavior.md) を参照してください。

## 更新と手書き部分の保護

原本を保存した後、同じ `export` を再実行すると更新できます。
入力と生成条件が同一なら、既存の根拠を検証して再利用し、`unchanged` を返します。
生成部分は今回の選択範囲で置き換わり、以前の別範囲を蓄積する方式ではありません。
`--context` を付けずに再生成すると、以前のノートに AI の説明があっても、抽出のみの内容になります。
範囲や生成方式ごとに残す場合は、出力先を分けてください。

生成マーカーの外側の文章と、CLI が管理しない Frontmatter 項目は保持します。
生成部分への手編集と、`reviewStatus: reviewed` のノートは保護します。

> [!WARNING]
> `--force` は旧ノートをバックアップしてから、保護された生成部分を置き換えます。
> マーカー外の手書き部分は保持します。無関係な Markdown や、別の PowerPoint のノートは置き換えられません。

`--refresh` は同じ入力でも AI を再実行するため、再度費用・利用枠を消費する場合があります。
手編集の保護を解除するオプションではありません。
生成に失敗した場合は以前のノートを残し、失敗した実行のフォルダへ取得済みの根拠・診断情報を保存します。
未完了の実行結果は自動再利用しません。

## 既定の設定を変える

通常の抽出には設定ファイルは不要です。
必要な場合だけ、次の操作で設定を作成・確認します。

```shell
tkn-powerpoint-note config init
tkn-powerpoint-note config list
```

`config list` は `config.generation.max_slides=10` のように1項目ずつ改行して表示します。
Windows パスは `\` を重ねず表示します。JSON が必要な場合は `config list --json` を使えます。

保存先は `~/.tkn/powerpoint_note/config.yaml` です。
編集済みの設定ファイルは上書きしません。
優先順位は、組み込み既定値 → ユーザー設定 → `./.tkn/config.yaml` → `--config FILE` → 個別 CLI オプションです。
相対パスは実行時の作業フォルダを基準にします。

新しい設定ファイルには `schema_version: "2.1.0"` を指定します。従来の `2.0.x` と共通の `generation.bridge_profile`／`prompt_profile` も使えます。
ノートの言語はプロファイルで選びます。組み込みは `default-ja`（既定・日本語）と `default-en`（英語）の2種類です。

```yaml
schema_version: "2.1.0"

generation:
  default_generator: my-codex-def
  generators:
    my-codex-def:
      bridge_profile: codex-default
      prompt_profile: default-ja
      overrides: {}
    my-codex-en:
      bridge_profile: codex-default
      prompt_profile: default-en
```

通常は `default_generator` を使い、一度だけ切り替える場合は `--generator` を指定します。

```shell
tkn-powerpoint-note export "C:\path\to\deck.pptx" --generator my-codex-en --context
tkn-powerpoint-note config list --generator my-codex-en
```

個別の `--bridge-profile`／`--prompt-profile` は選択したgeneratorより優先します。
`overrides` にはモデル・推論強度・タイムアウトなどを指定できます。
スライド数や画像幅の設定は、引き続き `generation.max_slides`／`image_width` に置きます。
通常の抽出でもプロファイルの表示文言とテンプレートを使います。原文の翻訳は行いません。

旧 `1.0.x` 設定を使っている場合は `schema_version` を `"2.1.0"` に変更し、
`generation.language: ja`／`en` を `generation.prompt_profile: default-ja`／`default-en` に置き換えてください。
旧 `--language` も `--prompt-profile` に置き換わります。設定ファイルの自動書き換えは行いません。
不正な版、未知のキー、型の誤りは処理前にエラーにします。
既定値と変更の影響は [設定リファレンス](docs/reference/configuration.md) にまとめています。

## コマンドを選ぶ

| 目的 | コマンド |
| --- | --- |
| 選択スライド・セクション・図形数を確認する | `inspect FILE` |
| 1 ファイルからノートを作成・更新する | `export FILE` |
| 生成本文・根拠・原本のハッシュを検証する | `verify NOTE` |
| 編集済み設定を保護して初期設定を作る | `config init` |
| 有効な設定と、その値を決めた設定元を確認する | `config list` |

0.3.0 ではコマンド名を `build` から `export` に変更しました。旧版の `build FILE` は `export FILE` に置き換えてください。

詳細は `COMMAND --help` で確認できます。
`--quiet` は進捗を省略し、`--verbose` は診断情報を追加します。同時指定はできません。
利用者に関わる変更は [CHANGELOG](CHANGELOG.md) に記録します。

## 開発と検証

```shell
cd "C:\path\to\tkn_powerpoint_note"
uv sync --locked
uv run pytest
uv run ruff check .
uv run mypy src
uv build
```

テストには小さな架空の OOXML と AI・Office の代替処理を使い、個人の資料を外部送信しません。
プロンプト・JSONスキーマ・Markdownテンプレート・表示文言は
[`context_profiles/default-ja`](src/powerpoint_note/context_profiles/default-ja/) と
[`context_profiles/default-en`](src/powerpoint_note/context_profiles/default-en/) にまとめています。
言語もテンプレートのメタデータから取得し、Python内では切り替えません。
完全な外部プロファイルの配置方法とファイル構成は [設定リファレンス](docs/reference/configuration.md#content-profiles) を参照してください。
Office による画像化は Windows 専用です。
文字抽出は移植可能な構成ですが、実動作の確認環境は Windows です。
ソースや同梱リソースを更新した後は、`uv tool install . --reinstall` で再インストールしてください。
