# AGENTS.md

## Project guardrails

- Never modify, move or delete a presentation. Rendering opens a temporary copy only.
- Never delete a proxy note. `pull` reports a note without its presentation as
  `missing-source` and leaves it in place.
- Preserve unknown Frontmatter fields and text outside the managed Markdown markers.
  Replace a hand-edited managed block only with `--force`, after backing the note up.
- Never truncate what is sent to AI or what is written to a note. A slide or summary
  over a limit is skipped and reported.
- AI calls happen only with `--context`. `--dry-run` and `status` make no AI calls,
  start no PowerPoint and write no files.
- Keep progress on stderr and one result on stdout (`key=value` lines, or JSON with `--json`).
- Do not put private paths, presentation contents or real user data in tracked files.
  Tests use small made-up presentations and stand-ins for AI and PowerPoint.

## Conventions

- This CLI follows `tkn_excel_note` for command names, configuration keys, result keys
  (`camelCase`), exit codes and the proxy-note Frontmatter. Check that repository before
  adding a command, option or result field that it already has.
- A change to what is sent to AI (`evidence.py`, prompts, schemas) regenerates existing
  explanations. Bump `EVIDENCE_CONTRACT` deliberately and say so in the CHANGELOG.
- Text files are LF (`.gitattributes`). User-facing documentation is written in Japanese.
