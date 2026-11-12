# Configuration

The installed package contains [config.example.yaml](../../src/powerpoint_note/resources/config.example.yaml).
Run `tkn-powerpoint-note config init` to create a user copy; run `config list` to inspect values and winning sources.
`config list` prints one `key=value` entry per line. Nested keys use dots, list items use `[0]`, and Windows paths display with a single backslash so values can be copied. Empty lists/maps appear as `[]`/`{}`. Use `config list --json` for structured JSON output.
Configuration is optional for ordinary extraction.

Files merge in this order: built-in → `~/.tkn/powerpoint_note/config.yaml` → current directory's `.tkn/config.yaml` → `--config` → CLI options.
Each file is validated before merging. Omitted keys inherit earlier values.
Nested mappings merge by property; lists replace earlier lists.
CLI options may precede or follow the subcommand.

Files use `schema_version: "3.0.0"`.
This version accepts 3.0.x, including newer patch versions. Other major/minor versions and missing versions fail.
Old 1.0.x, 2.0.x and 2.1.x settings are rejected with migration instructions. Reading configuration never rewrites it.
To migrate 2.x, set `schema_version: "3.0.0"`, rename `bridge_profile` to `genai_profile` (shared and named generators),
define the connections you used in GenAI Bridge under `genai_profiles` with the same names, and replace `--bridge-profile` with `--genai-profile`.
To migrate 1.0.x, also replace `generation.language: ja`/`en` with
`generation.prompt_profile: default-ja`/`default-en`, and replace `--language` with `--prompt-profile`.
Relative input/output/config paths resolve from the current working directory; `~` expands to the user's home.
Do not store tokens or passwords in this configuration; CLI providers reuse their existing login, and Azure reads an API key from the environment variable named by `azure.api_key_env`.

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
| `generation.default_generator` | `null` | Named preset used unless CLI `--generator` selects another. The initialized example selects `my-codex-def`. |
| `generation.generators` | `{}` | Named presets, merged by ID and field across configuration layers. |
| `generation.overrides` | `{}` | Shared scalar connection overrides, inherited by named presets. |
| `generation.genai_profile` | `codex-default` | Connection name in `genai_profiles`. CLI: `--genai-profile`. |
| `generation.prompt_profile` | `default-ja` | Complete content bundle. Built-ins: `default-ja` and `default-en`. CLI: `--prompt-profile`. |
| `generation.profile_dirs` | `[]` | Optional parent directories containing complete content bundles; searched in list order before packaged profiles. |
| `generation.max_slides` | `10` | Maximum selected slides per context export; CLI: `--max-slides`. A selection above it fails before AI. |
| `generation.max_input_chars` | `120000` | Maximum evidence JSON characters per slide or summary request; excludes image bytes and prompt instructions. Exceeding it fails; no text truncation. |
| `generation.image_width` | `2400` | PNG width in pixels, 600–8000. Height preserves slide aspect ratio. |
| `generation.render_timeout_seconds` | `180` | Timeout for the entire selected-slides rendering process. AI timeout comes from the selected connection's `timeout_seconds`. |
| `genai_profiles.<name>` | four built-in connections | AI connections and models; see [AI connections](#ai-connections). |

`--context` is always an explicit invocation option; a config file cannot turn it on.
There are no `sources`, sync rules, pull or push settings.
Disabling notes/comments also excludes them from AI evidence.
Unknown keys, blank strings, nonpositive limits and wrong types are errors.

For example, this is a complete valid configuration file that changes language and the context selection limit, inheriting other defaults:

```yaml
schema_version: "3.0.0"
generation:
  prompt_profile: default-en
  max_slides: 5
```

## AI connections

AI connections live in the top-level `genai_profiles` of this configuration.
The application validates the selected connection and passes it directly to the GenAI Bridge `Runtime`.
It never reads `~/.tkn/genai_bridge/config.yaml` or other Bridge configuration files, so installing this CLI is enough.
The current directory's `.tkn/config.yaml` is read as this application's configuration, as before.

The built-in connections use each CLI's default model and existing login:

```yaml
genai_profiles:
  codex-default:
    provider: codex
    model: null
  claude-default:
    provider: claude-code
    model: null
  copilot-default:
    provider: github-copilot
    model: null
  antigravity-default:
    provider: antigravity
    model: null
```

Connections merge by name and field across configuration layers, so a user file can set only `genai_profiles.codex-default.model`.
Add new names for other providers such as Ollama or Azure OpenAI.
Fields follow the pinned GenAI Bridge `Profile`: `provider`, `model`, `reasoning_effort`, `timeout_seconds`, `max_output_tokens`, `local_only`, `cli`, `ollama`, `azure` and `pricing`.
See the [GenAI Bridge configuration reference](https://github.com/tuckn/tkn_genai_bridge/blob/27ac8e48c8038ee86e029f5b845cd41f6da121af/docs/reference/configuration.md) for their meaning.
Names use letters/digits with optional `.`, `_` and `-`, starting with a letter or digit.
Each file's fields are validated before merging. Required fields of a nested entry, such as a `pricing.<model>` entry, may be split across files and are checked after merging; when any file defines `genai_profiles`, every merged connection is also validated as a whole.
Unknown fields, wrong types and invalid provider combinations are errors.
A relative `cli.executable` path resolves from the folder of the file that sets it; a bare command such as `claude` is found on `PATH`.

`generation.genai_profile` selects a connection. A name missing from `genai_profiles` fails before rendering or AI calls when `--context` is used; it never falls back to a Bridge setting or another connection.

## Named generators

A generator groups `genai_profile`, `prompt_profile` and optional `overrides`.
It is a reusable preset, not another AI provider. Connections are defined in `genai_profiles`.

```yaml
schema_version: "3.0.0"
generation:
  default_generator: my-codex-def
  generators:
    my-codex-def:
      genai_profile: codex-default
      prompt_profile: default-ja
      overrides: {}
    my-codex-en:
      genai_profile: codex-default
      prompt_profile: default-en
      overrides:
        timeout_seconds: 450
  max_slides: 10
```

```shell
tkn-powerpoint-note export "C:\path\to\deck.pptx" --generator my-codex-en --context
tkn-powerpoint-note config list --generator my-codex-en
```

Selection is CLI `--generator` → `generation.default_generator`. `null` uses the shared settings.
Effective fields resolve in this order: individual CLI `--genai-profile` / `--prompt-profile` → selected generator → shared generation settings → built-in defaults.
Omitted preset fields inherit shared settings. Definitions with the same ID deep-merge by field, including `overrides`, across config files. Each fragment is validated before merging, even for unused presets.
IDs use letters/digits with optional `.`, `_` and `-`, starting with a letter or digit.
Unknown selections fail before output or AI calls. Only the selected content profile's resources are loaded.

`overrides` accepts `model` and `reasoning_effort` (nonempty strings or `null`), `timeout_seconds` (positive finite number up to 86400), `max_output_tokens` (positive integer or `null`) and `local_only` (boolean).
Provider-specific combinations are checked against the selected connection when `--context` is used; for example, output-token limits require an API provider.
These scalar settings can be validated during ordinary extraction without loading GenAI Bridge or calling AI. Context dependencies are included in the standard installation.
Rendering/selection limits and `profile_dirs` remain shared `generation` settings.

Generator selection also controls the template language during ordinary extraction; AI remains opt-in with `--context`.
Flat `generation.genai_profile` / `prompt_profile` settings remain supported and do not require named presets. Configuration reads never migrate or overwrite user files.
`config list` reports the definitions, selected ID, effective values and per-field winning sources.
The selected ID appears as `generationGenerator` in note Frontmatter and `generation_generator` in evidence, results and AI usage.
Context cache keys use resolved settings and the selected ID. Editing an unused preset does not invalidate a note; changing selected settings can regenerate AI explanations. Reviewed/edited notes keep their normal protection.

## Content profiles

The two built-in bundles are [default-ja](../../src/powerpoint_note/context_profiles/default-ja/)
and [default-en](../../src/powerpoint_note/context_profiles/default-en/).
Each bundle contains these application-owned resources:

| File | Purpose |
| --- | --- |
| `prompt.md` | Instructions for interpreting one slide. |
| `output.schema.json` | Canonical structured slide response schema. |
| `template.md` | Slide context Markdown; YAML Frontmatter supplies bundle `version`, `language` and `labels`. |
| `deck-prompt.md` | Instructions for synthesizing the selected slide analyses. |
| `deck-output.schema.json` | Canonical presentation response schema, including citations. |
| `deck-template.md` | Presentation context Markdown. |
| `note-template.md` | Generated note body: coverage, mode, evidence link, overview and slides. |
| `source-template.md` | Source slide structure: title, metadata, image/context, extracted content, notes, comments and supplements. |

`template.md` owns the writing language, such as `Japanese` or `English`.
Both prompts substitute its `{{language}}` placeholder. All other templates use the bundle's own text.
Ordinary builds use the same bundle for labels and structure; extracted source text is never translated.
CLI diagnostics, JSON keys and machine evidence retain their stable technical names.

For an additional curated bundle, copy an entire built-in directory under an explicitly configured parent directory:

```yaml
schema_version: "3.0.0"
generation:
  prompt_profile: team-ja
  profile_dirs:
    - "C:/path/to/profiles"
```

This loads `C:/path/to/profiles/team-ja/`. Relative directories resolve from the execution working directory.
Names use lowercase letters, digits, dots, underscores and hyphens, starting with a letter or digit.
The first matching directory supplies the **whole bundle**. Missing files never fall back individually to packaged files.
Unknown names and malformed selected bundles fail before Office setup, AI calls or note writes, including dry-run.
`config list` reports the resolved name, language, version, resource hashes and package/custom source.
Unselected profiles are neither loaded nor validated.

Keep prompt, schema and template changes consistent. Individual resource path overrides are not supported.
Bundle metadata uses version `1.0.x`. Templates use `{{field}}` once per required renderer field;
optional sections use `{{#field}} ... {{field}} ... {{/field}}`, without nesting.
Keep the placeholders provided by the built-in templates when changing headings or order.
Source text containing template tokens is inserted literally, without a second evaluation.
Structured schema fields and types must remain compatible with the renderer.
Schemas must be self-contained: references, including remote references, are rejected.
Reserved `powerpoint-note:` management markers belong to the application and cannot occur in profile resources.

The selected bundle's name, version and all eight resource hashes participate in the build key.
Changing any selected resource can rebuild the note. Context exports reuse verified analysis when only Markdown templates or display labels change; prompt/schema, profile name/language or AI setting changes require new calls.
The note Frontmatter, evidence manifest and AI usage records retain bundle provenance.
Editing an unused profile does not invalidate existing output. Reviewed/edited notes retain their normal protections.
After changing packaged resources, reinstall a non-editable tool installation; custom profile directories are read directly.
