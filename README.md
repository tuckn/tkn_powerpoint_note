# Tkn PowerPoint Note

Create a source-backed Markdown note from PowerPoint for people and AI.
Ordinary builds extract text and metadata locally. Add `--context` to explain diagrams, arrows, placement and meaningful colors from slide images.

[日本語](README_ja.md)

A context note can explain which systems a diagram connects and why, citing the original slide numbers.
Source text, speaker notes, comments and AI interpretation remain separately labeled.
This CLI converts individual files; there is no source catalog, pull/push synchronization or PowerPoint write-back.

## Install

Use Python 3.11+ and uv. Replace this example with your repository folder:

```shell
cd "C:\path\to\tkn_powerpoint_note"
uv tool install .
tkn-powerpoint-note --help
tkn-powerpoint-note --version
```

A version number confirms installation. Ordinary extraction requires no Office or AI account.
For context generation, use Windows with desktop Microsoft PowerPoint:

```shell
cd "C:\path\to\tkn_powerpoint_note"
uv tool install ".[context]" --reinstall
```

Configure and authenticate an image-capable [GenAI Bridge](https://github.com/tuckn/tkn_genai_bridge) profile.
Bridge connections/models belong to `~/.tkn/genai_bridge/config.yaml`.
This app's `generation.bridge_profile` chooses the connection; `generation.language` chooses the note language.

## Create your first note

Save your presentation first. Inspect available slide numbers and section names, then build and verify:

```shell
tkn-powerpoint-note inspect "C:\path\to\deck.pptx"
tkn-powerpoint-note build "C:\path\to\deck.pptx" --output "C:\path\to\deck.pptx.md"
tkn-powerpoint-note verify "C:\path\to\deck.pptx.md"
```

`build` writes the note and a sibling evidence folder.
Without `--output`, notes are stored under `~/.tkn/powerpoint_note/data/`, separated by source path.
The original presentation is always read-only.

`--dry-run` validates inputs, selection, configuration and existing-note protection without creating files,
opening PowerPoint, authenticating, contacting a provider or calling AI.
One JSON result goes to stdout; progress goes to stderr. Success exits 0; errors exit nonzero.

## Select slides and generate context

```shell
tkn-powerpoint-note build "C:\path\to\deck.pptx" --slides "2,5-8" --output "C:\path\to\selection.md"
tkn-powerpoint-note build "C:\path\to\deck.pptx" --section "Architecture" --slides "1-20"
tkn-powerpoint-note build "C:\path\to\deck.pptx" --slides "2-3" --context --dry-run
tkn-powerpoint-note build "C:\path\to\deck.pptx" --slides "2-3" --context
```

Slide numbers are 1-based positions in presentation order. Ranges are inclusive.
Repeat `--section` for multiple exact names. Section and page selections intersect.
Hidden slides are excluded unless `--include-hidden` is specified.
Fully off-slide objects are excluded by default; `--off-slide append` adds them as separate supplements.
Partially clipped objects retain full text and carry a warning.

> [!IMPORTANT]
> `--context` sends selected images and extracted evidence, including enabled notes/comments, to the configured provider and may incur charges.
> It makes one call per selected slide plus one summary call.
> The default limit is 10 selected slides. Exceeding it fails before generation without truncation.

Only selected slides are rendered. Off-slide supplements are supplied as text evidence, not images.
See [processing and output specifications](docs/reference/behavior.md) for limits.

## Update and protect notes

Repeat the same command after saving source changes.
Identical builds reuse verified results and return `unchanged`.
The generated section reflects the current selection, replacing earlier selections.
A build without `--context` produces extraction-only content, even for an earlier context note.
Use separate output paths to retain different selections or modes.

Text outside generated markers and unknown Frontmatter properties are preserved.
Edits inside the generated section and `reviewStatus: reviewed` notes are protected.

> [!WARNING]
> `--force` replaces protected generated content after backing up the previous note.
> It preserves handwriting outside the markers. It cannot replace unrelated Markdown or another source's note.

`--refresh` requests fresh AI results and can incur the same calls again; it does not bypass edit protection.
Failed generation keeps the previous note and retains diagnostics in a failed run folder.
Incomplete runs are not reused.

## Configure defaults

Ordinary extraction needs no config. To customize defaults:

```shell
tkn-powerpoint-note config init
tkn-powerpoint-note config show
```

Settings are written to `~/.tkn/powerpoint_note/config.yaml`; edited files are preserved.
Precedence: defaults → user config → `./.tkn/config.yaml` → `--config FILE` → CLI options.
Relative command paths use the current working directory.
Each file needs `schema_version: "1.0.0"`; invalid versions, unknown keys and wrong types fail before processing.
See [configuration reference](docs/reference/configuration.md) for defaults and effects.

## Commands

| Purpose | Command |
| --- | --- |
| List selected slides, sections and object counts | `inspect FILE` |
| Create/update one note | `build FILE` |
| Verify content, evidence and source hashes | `verify NOTE` |
| Create user settings while preserving edits | `config init` |
| Show settings and winning sources | `config show` |

Use `COMMAND --help` for options. `--quiet` hides progress; `--verbose` adds diagnostics; they cannot be combined.
[CHANGELOG](CHANGELOG.md) records user-visible changes.

## Develop and verify

```shell
cd "C:\path\to\tkn_powerpoint_note"
uv sync --locked --all-extras
uv run pytest
uv run ruff check .
uv run mypy src
uv build
```

Tests use synthetic OOXML and mocked AI/Office adapters and never upload private presentations.
Runtime resources live in `src/powerpoint_note/resources/`.
Office rendering is Windows-only. Local extraction is designed to be portable; Windows is the validation environment.
After updating code or resources, run `uv tool install . --reinstall`, or use `".[context]"` for visual context.
