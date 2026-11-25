"""Find the presentations of a configured source and the notes that stand for them."""

from __future__ import annotations

from fnmatch import fnmatchcase
from pathlib import Path, PurePosixPath
from typing import Any

from .extract import SUFFIXES
from .frontmatter import PROXY_TYPE
from .markdown import read_frontmatter
from .models import ConfigError, NoteError


def _matches(relative: str, patterns: list[str]) -> bool:
    """Patterns use `/`; one without `/` also matches the file name in any folder."""
    name = PurePosixPath(relative).name
    lowered, lowered_name = relative.lower(), name.lower()
    for pattern in patterns:
        pattern = pattern.lower()
        if fnmatchcase(lowered, pattern) or (
            "/" not in pattern and fnmatchcase(lowered_name, pattern)
        ):
            return True
        if pattern.startswith("**/") and fnmatchcase(lowered, pattern[3:]):
            return True
    return False


def discover(spec: dict[str, Any]) -> list[str]:
    """Relative paths of the presentations in a source, in name order."""
    settings = spec["presentations"]
    root = Path(settings["dir"])
    if not root.is_dir():
        raise ConfigError(f"presentations.dir does not exist: {root}")
    paths = root.rglob("*") if settings["recursive"] else root.glob("*")
    found = []
    for path in paths:
        relative = path.relative_to(root).as_posix()
        if (
            path.is_file()
            and path.suffix.lower() in SUFFIXES
            and not path.name.startswith("~$")  # Office lock files
            and _matches(relative, settings["include"])
            and not _matches(relative, settings["ignore"])
        ):
            found.append(relative)
    return sorted(found, key=str.lower)


def resolve_file(spec: dict[str, Any], value: Path, known: list[str]) -> str:
    """One presentation named on the command line, relative to or inside presentations.dir."""
    root = Path(spec["presentations"]["dir"])
    path = value.expanduser()
    path = path.resolve() if path.is_absolute() else (root / path).resolve()
    try:
        relative = path.relative_to(root.resolve()).as_posix()
    except ValueError as exc:
        raise ConfigError(f"Presentation is outside presentations.dir: {path}") from exc
    match = next((item for item in known if item.lower() == relative.lower()), None)
    if match is None:
        raise ConfigError(
            f"Presentation is not in this source (check recursive, include and ignore): {relative}"
        )
    return match


def note_path(spec: dict[str, Any], relative: str) -> Path:
    return Path(spec["notes"]["dir"]) / (relative + ".md")


def missing_sources(spec: dict[str, Any], source_id: str, known: list[str]) -> list[dict[str, Any]]:
    """Proxy notes of this source whose presentation is no longer found. Never deleted here."""
    root = Path(spec["notes"]["dir"])
    present = {item.lower() for item in known}
    rows: list[dict[str, Any]] = []
    if not root.is_dir():
        return rows
    for note in sorted(root.rglob("*.md")):
        if "img" in note.relative_to(root).parts[:-1]:
            continue
        try:
            metadata, _ = read_frontmatter(note.read_text("utf-8-sig"))
        except (NoteError, OSError, UnicodeError):
            continue
        name = metadata.get("sourceFileName")
        if (
            metadata.get("type") == PROXY_TYPE
            and metadata.get("sourceRoot") == source_id
            and isinstance(name, str)
            and name.lower() not in present
        ):
            rows.append(
                {
                    "status": "missing-source",
                    "sourceRoot": source_id,
                    "sourcePath": name,
                    "notePath": str(note),
                    "message": "The presentation was not found; the note was left as it is.",
                }
            )
    return rows
