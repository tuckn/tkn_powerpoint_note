# Changelog

## Unreleased

- Rename `config show` to `config list` and print copyable `key=value` lines by default, with unescaped Windows paths; add `--json` for structured output. Listing logs at INFO level.
- Include GenAI Bridge and Windows-only pywin32 in standard dependencies, so `uv tool install .` also prepares `--context`. The separate `context` extra is removed; existing installations should use `uv tool install . --reinstall`. Desktop PowerPoint and AI connection settings are still required for visual context.

## 0.6.0

- Align PowerPoint Frontmatter with the Vault/Excel proxy-note layout: fixed common head/tail fields, comment-separated groups and flat Obsidian properties (note schema 2.0.0).
- Replace legacy `date` with `created`, preserve note identity and user properties, normalize known timestamps to quoted JST seconds, and reject ambiguous dates before generation.
- Add a blank line after Frontmatter and an H1 for new notes; retain legacy verification and compare flattened profile provenance against evidence.
- Protect `reviewStatus: accepted` alongside `reviewed`. Existing notes change only on explicit export; schema changes invalidate the build cache.

0.4.0:

- Add generation.default_generator, named generation.generators and --generator for export/config show.
- Resolve individual CLI options after the selected generator, deep-merge layered definitions and report effective setting sources.
- Add validated scalar GenAI Bridge overrides and record the selected generator in notes, evidence and usage.
- Keep flat generation settings and config schema 2.0.x readable; new config examples use 2.1.0.
- Exclude unselected generator definitions from output cache keys; keep edited/reviewed note protection.

0.3.0:

- Rename the public `build FILE` command to `export FILE`; use `export` in existing scripts. The old command is no longer accepted.
- Keep export options, note updates, evidence reuse, dry-run and content protection unchanged.

0.2.0:

- Move prompts, structured response schemas, Markdown templates and display labels into complete `default-ja` / `default-en` content profiles.
- Replace `generation.language` and `--language` with `generation.prompt_profile` and `--prompt-profile`; the profile owns the language.
- Add optional `generation.profile_dirs` to load complete curated external bundles.
- Change the config schema to `2.0.0`; old `1.0.x` configs fail with explicit migration instructions and are never rewritten automatically.
- Validate profile resources before processing, record provenance in outputs and invalidate caches when selected resources change.
- Keep source extraction, opt-in AI calls and protection of handwritten/reviewed notes.

## 0.1.0

Initial implementation:

- Build a source-backed Markdown note from one read-only PowerPoint file.
- Extract slide text, tables, cached chart data, SmartArt text, notes, comments, document properties and geometry.
- Select slides by 1-based page positions/ranges and exact section names.
- Exclude hidden slides and fully off-slide objects by default, with explicit inclusion options.
- Add optional visual context and a cited presentation summary through GenAI Bridge.
- Bound context processing to the selected slides, with a default 10-slide guard.
- Preserve handwritten note sections and unknown Frontmatter; protect reviewed/edited generated content.
- Verify hashes and provenance, reuse unchanged results, and retain evidence when generation fails.
- Provide layered settings, config init/show, dry-run, JSON results, and Japanese/English documentation.
- No source catalog, pull/push synchronization or PowerPoint write-back.

- Send provider-compatible output schemas while retaining citation-uniqueness validation locally.
