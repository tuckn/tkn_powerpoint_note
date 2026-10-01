# Changelog

## Unreleased

Initial 0.1.0 implementation:

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
