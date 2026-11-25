# Changelog

## Unreleased

0.7.0 からの作り直しです。
設定・ノート・コマンド・保存先のすべてが変わり、以前の形式との互換はありません（非互換変更）。
以前のノートと設定ファイルは読み込めないため、「移行」のとおりに作り直してください。

### 追加

- フォルダの PowerPoint を代理ノートとして管理する `pull` と、変更予定だけを表示する `status` を追加。設定ファイルの `sources.<id>` に、入力フォルダ（`presentations.dir`）とノートの保存先（`notes.dir`）を書く。`pull` は `--source <id>` または `--all-sources` で実行し、ファイル名を付けると1つに限定できる。ノートと PowerPoint は削除しない。元の PowerPoint が見つからないノートは `missing-source` として報告する。
- AI の説明を、スライド単位で再利用する。再利用の判定は、そのスライドと、スライドが使うレイアウト・マスター・テーマ・画像・ノートの内容、AI に送る根拠データ、プロンプト、出力スキーマ、接続先のモデルで行う。スライドの位置とファイル名は判定に含めないため、スライドの挿入・並べ替え・保存し直しでは、変わったスライドだけを生成し直す。
- 途中で失敗した後に同じコマンドを実行すると、済んだスライドを再利用して、残りから続ける。
- `--context` を付けない `pull` は、作成済みの説明を残す。生成後に内容が変わったスライドは `contextStatus: stale` と `contextStaleSlides` で示し、説明の上に注意書きを付ける。
- `--slides`・`--section` で一部のスライドだけに説明を付けた場合も、以前に作った説明を残して積み重ねる。選ばなかったスライドの説明は AI に送らないため、資料の要約は、要約に使う説明がすべて今回の対象に入っているときだけ生成する。それ以外では以前の要約を残し、古いことを注意書きで示す。
- 1つの資料の中に内容がまったく同じスライドが複数ある場合は、1回の呼び出しの結果をそれぞれに使う。
- `export --force` は、`export` が書き出した Markdown だけを置き換える。代理ノートやほかの Markdown は置き換えない。
- `--all-sources` で、1つの source の入力フォルダが見つからなくても、ほかの source の処理を続ける。
- 根拠データが `generation.limits.max_input_chars` を超えるスライドは、全体を止めずにそのスライドだけを飛ばす（結果の `skippedSlides`）。資料全体の要約が上限を超える場合は、要約だけを飛ばして `warnings` に示す。
- `.pptm` の文字抽出に対応。マクロは実行しない。`--context` の画像化は、これまでどおりマクロを含むファイルを拒否する。
- ノートの先頭に、全スライドのタイトル・セクション・表示状態・要素の数を並べた「Presentation Map」を追加。
- `--no-color` を追加。

### 変更

- AI に送る根拠データを短くした。空白のない JSON にし、色は図形ごとに重複を除き、位置はスライドに対する千分率の整数にし、図形名と空の項目を省いた。400枚の資料で、合計が約1,709万字から約209万字になり、上限（12万字）を超えるスライドが23枚から0枚になった。
- 抽出した文字を、位置の順（上から下、左から右）に並べる。図形の内部名と ID の見出し、コネクタ、説明のない画像を本文から外した。PowerPoint が自動で付けた代替テキストと、ファイル名だけの代替テキストは出力しない。
- 抽出した文字のエスケープをやめ、原文のまま出力する（`R&D`・`snake_case` など）。Markdown の構造を壊す箇所（行頭の `#`、`<tag`、表の `|`）だけを書き換える。
- グラフを、種類・タイトルと、系列ごとの表で出力する。
- コメントの返信を、返信した人と日時を付けて別の項目にする。
- スライド番号のフィールドを `‹#›` として扱う。PowerPoint がスライドの挿入で書き換える番号で、内容が変わったとみなさないため。
- コマンドと出力を `tkn-excel-note` に合わせた。
  - `export` の出力先の既定を、元ファイルの隣の `<ファイル名>.md` に変更。フォルダを指定すると、中のすべての PowerPoint を書き出す。既存の Markdown は `--force` がなければ置き換えない。
  - `inspect` を `presentation list-slides --presentation <パス>` に変更。
  - `--prompt-profile` を `--profile` に変更。`--genai-profile` を廃止（接続先は generator で選ぶ）。
  - `--slides`・`--section` は、`--context` で説明を付けるスライドを選ぶオプションになった。ノートには、常に表示中のすべてのスライドの抽出結果を書く。
  - 結果のキーを `camelCase` に統一し、`status`・`command`・`statusCounts`・`results` の形にそろえた。
  - 終了コードを、0＝成功、1＝処理の失敗（一部の失敗を含む）、2＝引数の誤り、3＝設定・対象の指定の誤り、130＝中断に変更。
- 設定の形式を `"4.0.0"` に変更。`sources` を追加し、`selection` を廃止。非表示スライドとスライド外の要素の扱いは `extraction.include_hidden`・`extraction.off_slide` に移動。`generation.profile_dirs` を `generation.prompt_profile_dirs` に改名し、上限を `generation.limits`（`max_slides`・`max_input_chars`・`image_width`・`max_render_seconds`）にまとめた。
- `generation.limits.max_slides` は、1回の実行で AI に送るスライドの数（1資料あたり）の上限になった。再利用するスライドは数えない。
- 画像化の待ち時間の既定を 180 秒から 600 秒に変更（`generation.limits.max_render_seconds`）。
- 代理ノートの Frontmatter を `tkn-excel-note` の並びに合わせ、`schemaVersion` を `"5.0.0"` にした。`sourceId`・`sourceRoot`・`contextStaleSlides` を追加し、検証用の項目（`sourceFile`・`sourceSha256`・`selectedSlides`・`context`・`generator`・`generation*`・`generatedBy*`・`promptProfile*`・`reviewStatus`・`evidencePath`・`buildKey`・`generatedSha256`）を廃止。`keywords`・`categories` は、既定で `[[...]]` のリンクにする（`notes.frontmatter_term_format`）。
- `export` の Markdown は `type: PowerPointExport`・`schemaVersion: "1.0"` とし、管理マーカーを付けない。
- AI の結果と画像の保存先を、ノートと同じフォルダの `img/<PowerPointのファイル名>/` に変更。ノートが参照しなくなった結果は、成功した実行の最後に削除する。
- 文章の profile の版を 2.0 に変更。スライドの出力から `slide_number` を外し、テンプレートの項目と表示文言を入れ替えた。独自の profile は作り直しが必要。

### 修正

- 標準の「タイトルとコンテンツ」レイアウトで、本文のプレースホルダーの位置を取得できず「位置不明」の警告が出ていたのを修正。位置をスライド → レイアウト → マスターの順に、PowerPoint と同じ対応で探す。
- AI の出力がスキーマに合わない場合に、原因が分からないまま終了コード 1 になっていたのを修正。合わない項目をメッセージに示し、返された出力を `rejected.json`、理由を `usage.json` に保存する。
- 画像化に失敗したとき、PowerPoint 側のエラーの内容をメッセージと `--verbose` のログに出す。時間切れのとき、一時フォルダの削除の失敗で元のエラーが隠れないようにした。
- グラフと SmartArt が、専用の欄と図形の一覧に二重に出力されていたのを修正。

### 廃止

- `verify` コマンドと、ノートごとの根拠フォルダ（`<ノート>.assets/<ハッシュ>/`、`evidence.json`・`body.md`・`manifest.json`）。手直しの検出は、ノートの隣の `img/<PowerPointのファイル名>/state.json` で行う。
- `reviewStatus: reviewed`・`accepted` によるノートの保護。
- 以前の形式の読み替え（ノートの schema 1.0.0〜4.0.0、設定の 1.0.x〜3.0.x、旧オプション `--bridge-profile`）。

### 移行

1. 設定ファイル `~/.tkn/powerpoint_note/config.yaml` を別名で退避し、`tkn-powerpoint-note config init` で作り直して、必要な値を書き写す。
2. 以前の `export` で作ったノート（`~/.tkn/powerpoint_note/data/` の下や、`--output` で指定した場所）は、新しい版では更新できない。手書きの部分を退避してから、`pull` または `export` で作り直す。以前の `.assets` フォルダの AI の結果は再利用できない。

### そのほか

- リポジトリのテキストファイルの改行コードを LF に統一し、`.gitattributes` を追加。
- CHANGELOG を日本語にし、版ごとの見出しと日付を付けた。0.7.0 の後に `Unreleased` として記録していた変更（Frontmatter の平坦化、`genai_profiles`、`type: PowerPoint` など）は、公開した版がなく、この作り直しに含まれるため、上の内容にまとめた。

## 0.7.0 - 2026-10-29

- すべてのコマンドの結果を、既定で1行ごとの `key=value` で表示。`--json` で JSON に切り替える。`--json` はコマンドの前後のどちらにも指定でき、エラーと中断にも適用する。結果のパスはエラーメッセージと別の項目にした。

## 0.6.0 - 2026-10-07

- Frontmatter を Vault と Excel の代理ノートの並びに合わせた（先頭と末尾の共通項目、コメントで区切ったグループ、入れ子のない項目）。ノートの schema は 2.0.0。
- `date` を `created` に置き換え、ノートの ID と利用者の項目を保持。日時を日本時間・秒単位・引用符付きにそろえ、あいまいな日時は生成の前にエラーにする。
- Frontmatter の後に空行を入れ、新しいノートに H1 を付ける。
- `reviewStatus: accepted` のノートも `reviewed` と同じく保護する。

## 0.5.0 - 2026-10-04

- `config show` を `config list` に改名し、既定で1行ごとの `key=value` を表示。`--json` で JSON に切り替える。
- GenAI Bridge と Windows 用の pywin32 を標準の依存関係に含め、`uv tool install .` だけで `--context` の準備ができるようにした。

## 0.4.0 - 2026-10-01

- `generation.default_generator`、名前付きの `generation.generators`、`--generator` を追加。
- 個別の CLI オプションを、選んだ generator の後に適用。重ねた設定を項目単位で合わせ、有効な値の設定元を表示する。
- GenAI Bridge の値の上書き（`overrides`）を追加し、選んだ generator をノート・根拠・使用量の記録に残す。

## 0.3.0 - 2026-10-01

- コマンド `build FILE` を `export FILE` に改名。

## 0.2.0 - 2026-10-01

- プロンプト・出力スキーマ・Markdown テンプレート・表示文言を、文章の profile（`default-ja`・`default-en`）にまとめた。
- `generation.language`・`--language` を `generation.prompt_profile`・`--prompt-profile` に置き換えた。
- `generation.profile_dirs` で、外部の profile を読み込めるようにした。
- AI に送る出力スキーマから、接続先が対応しない一意性の制約を外し、検証は手元で行う。

## 0.1.0 - 2026-10-01

最初の実装です。

- 1つの PowerPoint から、原本を変更せずに Markdown ノートを作る。
- スライドの文字、表、グラフの保存値、SmartArt の文字、発表者ノート、コメント、文書プロパティ、図形の位置を抽出する。
- スライドを、1から数えた位置・範囲と、セクション名で選ぶ。
- 非表示スライドと、完全にスライド外にある要素を既定で除外する。
- GenAI Bridge を使い、スライド画像からの説明と、出典のスライド番号を付けた資料の要約を生成する（`--context`）。
- 手書きの部分と、ツールが知らない Frontmatter の項目を保持する。
- 重ねた設定、`config init`、`--dry-run`、JSON の結果を提供する。
