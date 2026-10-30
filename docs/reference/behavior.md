# Processing and output contract

## Inputs and selection

Only ordinary Transitional OOXML `.pptx` is supported.
Legacy `.ppt`, macro-enabled `.pptm`, encrypted files and Strict OOXML are not supported; save a normal `.pptx` first.
The input is read-only. AI rendering opens a temporary snapshot, never the original.

Slide positions come from the presentation's relationship list, not slide XML filenames.
`--slides "2,5-8"` selects 1-based positions in file order; ranges are inclusive and duplicates collapse.
The note also records the stable OOXML slide ID and the presentation's starting page-number offset.
A displayed page-number textbox may contain arbitrary text; it is retained as text and is not the selector.
`--section` matches exact names. Duplicate section names select all matching slides.
Hidden slides require `--include-hidden`, even when their page number was explicitly selected.
Invalid ranges, unknown section names and empty selections are errors.

## Extraction

The local reader uses guarded XML parsing and follows internal package relationships without fetching external links.
It extracts document properties, titles, text paragraphs and indentation, tables, cached chart series, SmartArt text,
speaker notes, comments, geometry, group membership, alternative text, colors and connector endpoints.

Placeholder geometry can inherit from layouts/masters. Group scaling, translation, rotation and flips are applied.
Bounding boxes use slide coordinates in EMUs; 914400 EMU = 1 inch.
Fully off-slide objects are excluded by default. `--off-slide append` places them in a separate supplement.
Partially intersecting objects retain all their text and carry a warning; text is not geometrically clipped.
Unknown geometry is retained and flagged rather than silently discarded.
Object geometry cannot prove the author's intended reading order, lead sentence or semantic relationships.

Speaker notes omit slide-image, slide-number, date, header and footer placeholders.
Legacy comments retain author, timestamp and text. Modern comment elements are read on a best-effort basis;
modern threading/status variants are not guaranteed to reproduce the complete review UI.
Tables retain row/cell text, with merged-cell metadata in JSON. Generic column labels avoid assuming the first row is a header.
Chart values are cached source values, which may differ from a linked workbook's current data.
Paragraph text is written in package order with recorded indentation.
Document properties are separate from slide text; application vector properties containing all slide titles are omitted.

## Visual context

The renderer exports only selected slides as PNG through desktop PowerPoint on Windows.
It uses a read-only temporary copy, disables macros during opening and closes only its own presentation.
PowerPoint can share an existing application process: the adapter never calls Application.Quit.
The child process has a timeout. A hung Office dialog may require manually closing a leftover snapshot window.
Linked resources and active content are rejected before rendering to avoid implicit resource retrieval or execution.
Hyperlinks are recorded but never followed.

Each selected slide receives one image and that slide's filtered structure, notes and comments.
A second stage synthesizes only those slide analyses and cites their source slide numbers.
The app does not send unselected slide images/text to the AI, nor the original presentation package.
Fully off-slide supplements have no image; uncertain visual relationships there must remain uncertain.
All provider calls are explicit `--context` operations. A selection of N slides requires N + 1 calls when uncached.
There is no automatic retry. The default maximum selection is 10; `--max-slides` explicitly changes it.

Prompts, JSON schemas, Markdown templates and labels form a complete selected content profile.
Built-in profiles are `default-ja` and `default-en`; their metadata owns the writing language.
The full selected bundle is validated before ordinary or context builds. See [content profiles](configuration.md#content-profiles).
The app validates returned JSON, source slide numbers, summary citations and reserved management markers.
The provider receives a schema without the unsupported `uniqueItems` constraint; the full packaged schema still validates citation uniqueness locally.
These checks do not establish that every generated statement is accurate.
Small text in raster images, videos, animations, OLE contents, font substitution and master/layout-only artwork
are limitations. Master artwork appears in Office-rendered images but is not fully extracted by the XML reader.
No OCR is performed in ordinary exports. No LibreOffice renderer is currently implemented.

## Notes and evidence

A note has YAML Frontmatter, a single generated block between `powerpoint-note:begin/end` comments,
and user-editable text outside that block.
Unknown scalar/flat-list Frontmatter fields and text outside the markers survive updates.
Source metadata and application-owned provenance fields refresh on successful exports.
Edits inside the generated block and `reviewStatus: reviewed` / `accepted` are protected.
`--force` backs up the old note and replaces protected generated content; it resets review status to unreviewed.
Unrelated files, unsupported note schemas and a different source identity are never overwritten.

The sibling `<note>.assets/<build-key>/` folder contains:

- `evidence.json`: filtered selected-slide evidence, deck metadata and source hash.
- `body.md`: the exact generated Markdown block.
- `manifest.json`: schema version, generator version, selected generation preset, selected content profile provenance, source/build hashes and file checksums.
- Context exports add slide PNGs, combined context JSON, per-call prompts/schemas, validated responses and usage records.

Keep the note and its assets together. Moving both together preserves relative links.
Source paths are absolute in Frontmatter; use `verify --source FILE` to verify after moving the source.
Private evidence belongs outside version control. This repository ignores `.local/`, which is for retained development evidence.
Source packages are not retained in successful bundles.

Build keys include source bytes, selected evidence/settings, generator version, note schema, selected generation preset, selected content profile identity/resource hashes and connection settings. Unselected generator definitions are excluded.
Any source byte change invalidates the Markdown bundle. A change only to document properties or ZIP bookkeeping can reuse visual context while refreshing the source metadata and `sourceSha256`.
Context also has a separate `context_build_key` in the manifest: it includes the source content fingerprint, extracted slide evidence/selection, profile name/language, prompts, response schemas, generation settings and connection plan. It excludes document metadata, file size/raw hash, note schema, tool version, Markdown templates, profile search directories and execution limits. This lets an export rebuild note properties or display templates using verified context JSON/images without rendering or calling AI again. The original `contextGeneratedAt` is retained.
`contextSourceFingerprint` hashes sorted OOXML package entry names and decompressed bytes, excluding directory entries and `docProps/core.xml`, `docProps/app.xml`, `docProps/custom.xml`. ZIP order, compression and entry timestamps do not affect it. Slide parts, images, charts, notes, comments, themes, layouts and all other package entries participate, including unselected slides. A fingerprint alone never authorizes reuse: AI inputs/settings and bundle checksums must also match. `sourceSha256` independently hashes the entire original file for exact verification. Changes to OOXML bytes (including formatting within XML) outside the excluded properties conservatively require fresh analysis.
Reuse is local to the existing note and its assets; there is no global cache across different output paths.
There is no per-slide incremental synchronization or accumulation across selections.
Identical exports verify existing artifacts and return unchanged.
`--refresh` creates a distinct context bundle with fresh calls.
Failed runs retain available evidence under `failed-<id>/`; they do not replace the previous note or resume automatically.
A crash may leave a `pending-<id>/` folder or note lock; confirm no export is running before removing a stale lock.
Completed unreferenced bundles can remain if publication was interrupted; they are not automatically deleted.

Bundles are validated before publication. Markdown publication uses a same-directory temporary file and atomic replacement.
Source and note content are rechecked immediately before publication.
An application lock prevents competing exports to the same note.
`verify` checks generated-body hash, bundle file hashes, note/bundle provenance and the current source hash.
It reports semantic accuracy as not assessed.

### Frontmatter

PowerPoint proxy notes use `type: powerpoint` and `schemaVersion: "2.1.0"`.
Properties are scalars or flat lists; grouping uses YAML comments and blank lines, never nested mappings.

| Order | Properties |
| --- | --- |
| Common note properties | `type`, `schemaVersion`, `title`, `description`, `cover` |
| Presentation metadata | `subject`, `author`, `keywords`, `categories`, `comments` |
| Source identity and location | `sourceFileName`, `sourceFullPath`, `sourceFile`, `sourceSha256` |
| Source timestamps | `sourceCreated`, `sourceModified` |
| Slide selection | `selectedSlides`, `context` |
| Context generation | `contextStatus`, `contextProfiles`, `contextAnalyzedSlides`, `contextOmittedSlides`, `contextSourceFingerprint`, `contextGeneratedAt` |
| Generation | `generator`, `generatorVersion`, `generationGenerator`, `promptProfile`, `promptProfileVersion`, `promptProfileLanguage`, `promptProfileSha256`, `promptProfileResources`, `reviewStatus` |
| Evidence and verification | `evidencePath`, `buildKey`, `generatedSha256` |
| Additional properties | User-defined scalar/flat-list properties, in their existing order |
| Common file management | `tags`, `created`, `updated`, `noteId` |

The title uses the document title, falling back to the source filename without its extension.
`subject`, `author`, `comments` come from OOXML subject, creator and description.
Keywords and categories split on semicolons, preserving commas within a term.
All original core/application/custom document properties remain in `evidence.json`; only the fields above are projected into Frontmatter.
`promptProfile` is the profile name. `promptProfileResources` is a flat list of `filename: SHA-256` strings.
The full structured profile provenance remains in `manifest.json`, and `verify` compares all flat profile fields against it.
`sourceFileName` and `sourceFullPath` follow the Excel note naming; `sourceFile` remains an equal full-path compatibility alias. There is no source catalog, so `sourceId`/`sourceRoot` are not fabricated.
`contextProfiles` lists the profile used by the context currently in the note (for example `[default-ja]`); selections/profiles replace prior content rather than accumulating it. `contextStatus` is `current` when all presentation slides were analyzed, `partial` when some were omitted (including hidden slides), or `not-generated` for extraction-only output. Analyzed/omitted lists contain 1-based slide positions. Extraction-only output clears the context lists, fingerprint and generated timestamp. `verify` checks these properties against the completed bundle.

New timestamps use JST (`+09:00`) with seconds and double quotes, for example `"2026-06-21T05:44:56+09:00"`.
Known Frontmatter timestamps normalize to that form on export; fractional seconds are dropped.
Date-only values remain `"YYYY-MM-DD"`; missing dates stay `""`. Invalid dates and datetimes without a timezone fail before generation; no timezone is guessed.
Custom date strings (including list elements) are double-quoted without changing their values.
Versions, empty strings, Obsidian links and strings that resemble YAML numbers/booleans/null are double-quoted.
Windows paths are single-quoted. Empty lists use `[]`.

New notes receive `created`, `updated`, a UUID `noteId`, empty `description`/`cover`, and `tags: []`.
Regeneration preserves `created`, `noteId`, description, cover and tags, and refreshes `updated`.
Schemas 1.0.0 and 2.0.0 remain readable by `verify`; the next export writes schema 2.1.0, expanding legacy `sourceMetadata` and nested profile provenance into flat properties.
An old `date` becomes `created`, preserving its represented date/time (subject to timestamp normalization).
Conflicting `date`/`created` values, including subsecond differences, fail before writes or AI calls.
If neither creation field is available, `created` stays empty; the export time is not substituted.
A missing note ID is assigned once. Unknown nested user properties are rejected without changing the note.
Use `export --dry-run` to check the update; this does not rewrite existing notes.
Legacy context bundles without `context_build_key` are reused when their original build key can be reproduced using the recorded tool version, a supported older note schema and the current inputs/settings. If that match cannot be established, a context export plans fresh AI calls; check `--dry-run` first. Existing reviewed/accepted protection still applies to migrations.

There is a blank line after the closing Frontmatter fence. New notes (and legacy notes with no text before the generated block)
receive an H1 before the generated block. Existing text outside the markers, including an edited H1, is preserved.

## Dry-run, output and errors

`export --dry-run` reads the input/config/current note, resolves selection and checks protection.
For context, it additionally checks Bridge settings/executable and evidence limits. Office availability and linked resources are checked when fresh rendering is required; verified context reuse does not require PowerPoint.
It does not open Office, authenticate, call a provider, make network requests, or write data/cache/report/state files.
It reports planned counts and `planned_ai_calls`; actual `ai_calls` remains 0.
Summary size is checked after slide analyses exist during a real export; that future size cannot be validated in advance.

`inspect`, `verify` and `config list` are read-only.
`config init` writes only the selected config path and preserves existing edits.
Normal exports write by default. All commands print one `key=value` entry per line on stdout by default; `--json` prints one JSON object instead. Nested fields use dots and list indexes use `[0]`. Paths are complete and unescaped in text output; control characters are escaped. JSON keys, types, and underlying results retain their existing meanings. Leveled progress goes to stderr; listing a configuration logs `[INFO] Showing resolved configuration`. Quiet and verbose affect only stderr logs, and output format never depends on TTY or redirection. No details are omitted, so `--full-output` is unnecessary.
For a failed command, `status=failed`, `error`, and available `source_path`, `note_path`, `output_path`, or `config_path` fields are separate. Exit codes: 0 success, 2 input/configuration/protection or expected I/O failure, 1 unexpected processing failure, 130 interruption. Help and version retain their usual output.
The same display rules apply to `inspect`, `export`, `verify`, `config init`, and `config list`. `--json` works before or after commands, including both levels of `config init` and `config list`. CLI formatting does not change the Python API, saved notes, evidence, manifests, or dry-run behavior.
`--help` and `--version` work without loading configuration.

## References and implementation basis

The design follows the companion Excel CLI's separation of exact evidence, optional visual interpretation,
packaged prompts, GenAI Bridge integration and protected Markdown sections.
PowerPoint-specific selection and extraction are implemented independently; Excel synchronization is not included.

The renderer follows Microsoft's [Presentations.Open](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open)
and [Slide.Export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export) interfaces.
OOXML relationship traversal follows the [PresentationML structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document).

Structured-output transport constraints follow the official [OpenAI schema guidance](https://developers.openai.com/api/docs/guides/prompt-generation).
