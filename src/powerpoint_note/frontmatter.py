"""Flat, ordered Obsidian properties for PowerPoint proxy notes."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from typing import Any

import yaml

from .models import NoteError

SCHEMA_VERSION = "2.0.0"
JST = timezone(timedelta(hours=9))
HEAD = ("type", "schemaVersion", "title", "description", "cover")
TAIL = ("tags", "created", "updated", "noteId")
GROUPS = (
    ("Presentation metadata", ("subject", "author", "keywords", "categories", "comments")),
    ("Source identity and location", ("sourceFile", "sourceSha256")),
    ("Source timestamps", ("sourceCreated", "sourceModified")),
    ("Slide selection", ("selectedSlides", "context")),
    (
        "Generation",
        (
            "generator",
            "generatorVersion",
            "generationGenerator",
            "promptProfile",
            "promptProfileVersion",
            "promptProfileLanguage",
            "promptProfileSha256",
            "promptProfileResources",
            "reviewStatus",
        ),
    ),
    ("Evidence and verification", ("evidencePath", "buildKey", "generatedSha256")),
)
KNOWN = {*HEAD, *TAIL, *(key for _, keys in GROUPS for key in keys)}
TIMESTAMPS = ("created", "updated", "sourceCreated", "sourceModified")


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


class NoteDumper(yaml.SafeDumper):
    """Quote strings without changing their values or YAML types."""


def _string(dumper: NoteDumper, value: str) -> yaml.ScalarNode:
    style = None
    if re.match(r"^(?:[A-Za-z]:\\|\\\\)", value):
        style = "'"
    elif (
        not value
        or value.startswith("[[")
        or re.match(r"^\d{4}-\d{2}-\d{2}(?:$|[T ])", value)
        or re.fullmatch(r"\d+\.\d+\.\d+(?:[-+].*)?", value)
        or any(
            pattern.match(value)
            for _, pattern in yaml.SafeLoader.yaml_implicit_resolvers.get(value[0], [])
        )
    ):
        style = '"'
    return dumper.represent_scalar("tag:yaml.org,2002:str", value, style=style)


NoteDumper.add_representer(str, _string)


def now_iso() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def timestamp(value: Any, field: str) -> str:
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


def profile_properties(provenance: dict[str, Any]) -> dict[str, Any]:
    return {
        "promptProfile": provenance["name"],
        "promptProfileVersion": provenance["version"],
        "promptProfileLanguage": provenance["language"],
        "promptProfileSha256": provenance["sha256"],
        "promptProfileResources": [
            f"{name}: {sha}" for name, sha in sorted(provenance["resources"].items())
        ],
    }


def source_properties(metadata: dict[str, Any]) -> dict[str, Any]:
    core = metadata.get("core.xml", {})
    return {
        "title": core.get("title", ""),
        "subject": core.get("subject", ""),
        "author": core.get("creator", ""),
        "keywords": [s.strip() for s in core.get("keywords", "").split(";") if s.strip()],
        "categories": [s.strip() for s in core.get("category", "").split(";") if s.strip()],
        "comments": core.get("description", ""),
        "sourceCreated": timestamp(core.get("created"), "sourceCreated"),
        "sourceModified": timestamp(core.get("modified"), "sourceModified"),
    }


def normalize(values: dict[str, Any]) -> dict[str, Any]:
    result = dict(values)
    old, new = result.get("date"), result.get("created")
    if old not in (None, "") and new not in (None, "") and old != new:
        timestamp(old, "date")
        timestamp(new, "created")
        if datetime.fromisoformat(old) != datetime.fromisoformat(new):
            raise NoteError("date and created disagree; resolve the conflict before exporting")
    result["created"] = new if new not in (None, "") else old or ""
    result.pop("date", None)
    source = result.pop("sourceMetadata", None)
    if isinstance(source, dict):
        for key, value in source_properties(source).items():
            result.setdefault(key, value)
    if isinstance(result.get("promptProfile"), dict):
        result.update(profile_properties(result["promptProfile"]))
    result.update(type="powerpoint", schemaVersion=SCHEMA_VERSION)
    for key in (*HEAD, *TAIL):
        result.setdefault(key, [] if key == "tags" else "")
    for key in TIMESTAMPS:
        if key in result:
            result[key] = timestamp(result[key], key)
    for key, value in result.items():
        if value is None:
            result[key] = ""
        elif isinstance(value, list):
            if any(not isinstance(item, (str, int, float, bool)) for item in value):
                raise NoteError(f"{key} must be a flat list of scalars; existing note preserved")
        elif not isinstance(value, (str, int, float, bool)):
            raise NoteError(f"{key} must be a scalar or flat list; existing note preserved")
    return result


def dump(values: dict[str, Any]) -> str:
    values = normalize(values)

    def block(keys: Any) -> str:
        return yaml.dump(
            {key: values[key] for key in keys if key in values},
            Dumper=NoteDumper,
            allow_unicode=True,
            sort_keys=False,
            width=1000,
        ).rstrip()

    parts = [block(HEAD)]
    for label, keys in GROUPS:
        if any(key in values for key in keys):
            parts.append(f"# --- {label} ---\n" + block(keys))
    extras = [key for key in values if key not in KNOWN]
    if extras:
        parts.append("# --- Additional properties ---\n" + block(extras))
    parts.append(block(TAIL))
    return "\n\n".join(parts) + "\n"
