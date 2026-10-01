# Configuration

The installed package contains [config.example.yaml](../../src/powerpoint_note/resources/config.example.yaml).
Run `tkn-powerpoint-note config init` to create a user copy; run `config show` to inspect values and winning sources.
Configuration is optional for ordinary extraction.

Files merge in this order: built-in → `~/.tkn/powerpoint_note/config.yaml` → current directory's `.tkn/config.yaml` → `--config` → CLI options.
Each file is validated before merging. Omitted keys inherit earlier values.
Nested mappings merge by property; lists replace earlier lists.
CLI options may precede or follow the subcommand.

Every file needs `schema_version: "1.0.0"`.
This version accepts 1.0.x, including newer patch versions. Other major/minor versions and missing versions fail.
There are no legacy migrations. Reading configuration never rewrites it.
Relative input/output/config paths resolve from the current working directory; `~` expands to the user's home.
Credentials are owned by GenAI Bridge, not this configuration.

| Key | Default | Meaning |
| --- | --- | --- |
| `selection.slides` | `null` | All eligible slides; otherwise a quoted list such as `"1,3-5"`. |
| `selection.sections` | `[]` | All sections; otherwise exact names. Multiple names form a union, intersected with slides. |
| `selection.include_hidden` | `false` | Include slides marked hidden. |
| `selection.off_slide` | `exclude` | `append` retains fully outside objects as supplemental evidence. |
| `extraction.include_notes` | `true` | Include speaker notes; CLI: `--no-notes` to disable. |
| `extraction.include_comments` | `true` | Include comments; CLI: `--no-comments` to disable. |
| `extraction.max_file_mb` | `100` | Maximum compressed source size, in MiB. |
| `extraction.max_uncompressed_mb` | `1024` | Maximum total declared uncompressed ZIP size, in MiB. |
| `generation.bridge_profile` | `codex-default` | Named connection in GenAI Bridge. CLI: `--bridge-profile`. |
| `generation.language` | `ja` | `ja` or `en`; affects note labels and AI writing language. |
| `generation.max_slides` | `10` | Maximum selected slides per context build; CLI: `--max-slides`. A selection above it fails before AI. |
| `generation.max_input_chars` | `120000` | Maximum evidence JSON characters per slide or summary request; excludes image bytes and prompt instructions. Exceeding it fails; no text truncation. |
| `generation.image_width` | `2400` | PNG width in pixels, 600–8000. Height preserves slide aspect ratio. |
| `generation.render_timeout_seconds` | `180` | Timeout for the entire selected-slides rendering process. AI timeout comes from Bridge. |

`--context` is always an explicit invocation option; a config file cannot turn it on.
There are no `sources`, sync rules, pull or push settings.
Disabling notes/comments also excludes them from AI evidence.
Unknown keys, blank strings, nonpositive limits and wrong types are errors.

For example, this is a complete valid configuration file that changes language and the context selection limit, inheriting other defaults:

```yaml
schema_version: "1.0.0"
generation:
  language: en
  max_slides: 5
```
