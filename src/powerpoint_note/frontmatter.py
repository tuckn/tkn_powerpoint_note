"""Flat, ordered Obsidian properties for PowerPoint proxy notes and exports."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from typing import Any

import yaml

from .models import NoteError

PROXY_TYPE = "PowerPoint"
PROXY_SCHEMA = "5.0.0"
EXPORT_TYPE = "PowerPointExport"
EXPORT_SCHEMA = "1.0"
JST = timezone(timedelta(hours=9))
HEAD = ("type", "schemaVersion", "title", "description", "cover")
TAIL = ("tags", "created", "updated", "noteId")
DOCUMENT = ("subject", "author", "keywords", "categories", "comments")
CONTEXT = (
    "contextStatus",
    "contextProfiles",
    "contextAnalyzedSlides",
    "contextOmittedSlides",
    "contextStaleSlides",
    "contextSourceFingerprint",
    "contextGeneratedAt",
)
GROUPS = (
    ("Presentation metadata", DOCUMENT),
    (
        "Source identity and location",
        ("sourceId", "sourceRoot", "sourceFileName", "sourceFullPath"),
    ),
    ("Source timestamps", ("sourceCreated", "sourceModified")),
    ("Context generation", CONTEXT),
)
EXPORT_ORDER = (
    "type",
    "schemaVersion",
    "title",
    *DOCUMENT,
    "sourceFileName",
    "sourceFullPath",
    "sourceCreated",
    "sourceModified",
    "sourceSnapshotSha256",
    "generatedAt",
    "generationMethod",
    *CONTEXT,
)
KNOWN = {*HEAD, *TAIL, *(key for _, keys in GROUPS for key in keys)}
# Always double-quoted, whatever the value looks like.
QUOTED = {
    "schemaVersion",
    "title",
    "description",
    "cover",
    "sourceCreated",
    "sourceModified",
    "created",
    "updated",
    "generatedAt",
    "contextGeneratedAt",
    "contextSourceFingerprint",
    "sourceSnapshotSha256",
}


class NoteLoader(yaml.SafeLoader):
    """Preserve date strings and reject ambiguous keys before updating a note."""

    def construct_mapping(self, node: yaml.MappingNode, deep: bool = False) -> dict[Any, Any]:
        self.flatten_mapping(node)
        result: dict[Any, Any] = {}
        for key_node, value_node in node.value:
            key = self.construct_object(key_node, deep=deep)
            if not isinstance(key, str) or key in result:
                raise NoteError("Frontmatter requires unique string property names")
            result[key] = self.construct_object(value_node, deep=deep)
        return result


NoteLoader.yaml_implicit_resolvers = {
    key: [(tag, pattern) for tag, pattern in values if tag != "tag:yaml.org,2002:timestamp"]
    for key, values in NoteLoader.yaml_implicit_resolvers.items()
}


class Quoted(str):
    """A string written in double quotes regardless of its content."""


class NoteDumper(yaml.SafeDumper):
    """Quote strings without changing their values or YAML types."""


def _string(dumper: NoteDumper, value: str) -> yaml.ScalarNode:
    style = None
    if re.match(r"^(?:[A-Za-z]:\\|\\\\)", value) or (
        "\\" in value and _needs_quotes(value) and "'" not in value
    ):
        # Backslashes stay literal in single quotes.
        style = "'"
    elif _needs_quotes(value):
        style = '"'
    return dumper.represent_scalar("tag:yaml.org,2002:str", value, style=style)


def _needs_quotes(value: str) -> bool:
    return (
        not value
        or value.startswith("[[")
        or bool(re.match(r"^\d{4}-\d{2}-\d{2}(?:$|[T ])", value))
        or bool(re.fullmatch(r"\d+\.\d+(?:\.\d+)?(?:[-+].*)?", value))
        or bool(re.fullmatch(r"[yYnN]|\d+[eE]\d+|0o[0-7]+", value))
        or any(
            pattern.match(value)
            for _, pattern in yaml.SafeLoader.yaml_implicit_resolvers.get(value[0], [])
        )
    )


def _quoted(dumper: NoteDumper, value: Quoted) -> yaml.ScalarNode:
    return dumper.represent_scalar("tag:yaml.org,2002:str", str(value), style='"')


NoteDumper.add_representer(str, _string)
NoteDumper.add_representer(Quoted, _quoted)


def now_iso() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def timestamp(value: Any, field: str) -> str:
    """Japan time to the second; date-only values stay dates; no timezone is ever guessed."""
    if value in (None, ""):
        return ""
    if not isinstance(value, str):
        raise NoteError(f"{field} must be an ISO 8601 string")
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as exc:
        raise NoteError(f"{field} contains an invalid timestamp") from exc
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return value
    if parsed.utcoffset() is None:
        raise NoteError(f"{field} has no timezone; refusing to assume one")
    return parsed.astimezone(JST).isoformat(timespec="seconds")


def _terms(value: str, term_format: str) -> list[str]:
    terms = [s.strip() for s in value.split(";") if s.strip()]
    return [f"[[{term}]]" for term in terms] if term_format == "obsidian-link" else terms


def source_properties(
    metadata: dict[str, Any], fallback_title: str, term_format: str = "plain"
) -> dict[str, Any]:
    """Document properties of the presentation, as written in both note types."""
    core = metadata.get("core.xml", {})
    return {
        "title": core.get("title", "") or fallback_title,
        "subject": core.get("subject", ""),
        "author": core.get("creator", ""),
        "keywords": _terms(core.get("keywords", ""), term_format),
        "categories": _terms(core.get("category", ""), term_format),
        "comments": core.get("description", ""),
        "sourceCreated": timestamp(core.get("created"), "sourceCreated"),
        "sourceModified": timestamp(core.get("modified"), "sourceModified"),
    }


def _block(values: dict[str, Any], keys: Any) -> str:
    selected = {
        key: Quoted(values[key])
        if key in QUOTED and isinstance(values[key], str)
        else ("" if values[key] is None else values[key])
        for key in keys
        if key in values
    }
    if not selected:
        return ""
    return yaml.dump(
        selected, Dumper=NoteDumper, allow_unicode=True, sort_keys=False, width=1000
    ).rstrip()


def dump(values: dict[str, Any]) -> str:
    """Proxy-note properties: fixed head and tail, comment-separated groups, extras kept."""
    for key in (*HEAD, *TAIL):
        if key not in values:
            raise NoteError(f"Frontmatter lacks {key}")
    parts = [_block(values, HEAD)]
    for label, keys in GROUPS:
        if any(key in values for key in keys):
            parts.append(f"# --- {label} ---\n" + _block(values, keys))
    extras = [key for key in values if key not in KNOWN]
    if extras:
        parts.append("# --- Additional properties ---\n" + _block(values, extras))
    parts.append(_block(values, TAIL))
    return "\n\n".join(parts) + "\n"


def dump_export(values: dict[str, Any]) -> str:
    return _block(values, EXPORT_ORDER) + "\n"
