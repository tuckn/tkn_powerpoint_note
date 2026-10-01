"""Layered, strict configuration; no source discovery or synchronization."""

from __future__ import annotations

import copy
import re
from importlib.resources import files
from pathlib import Path
from typing import Any

import yaml

from .io import atomic_write
from .models import NoteError

SCHEMA_VERSION = "1.0.0"


def example() -> str:
    return files("powerpoint_note").joinpath("resources/config.example.yaml").read_text("utf-8")


def user_config() -> Path:
    return Path.home() / ".tkn" / "powerpoint_note" / "config.yaml"


def _validate(raw: Any, label: str, defaults: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise NoteError(f"{label}: expected a YAML mapping")
    version = raw.get("schema_version")
    if not isinstance(version, str) or not re.fullmatch(
        r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)", version
    ):
        raise NoteError(f"{label}: schema_version must be a quoted SemVer string (1.0.0)")
    if tuple(map(int, version.split(".")))[:2] != (1, 0):
        raise NoteError(f"{label}: unsupported schema {version}; supported: 1.0.x")
    for group, values in raw.items():
        if group == "schema_version":
            continue
        if group not in defaults:
            raise NoteError(f"{label}: unknown key {group}")
        if not isinstance(values, dict):
            raise NoteError(f"{label}: {group} must be a mapping")
        for key, value in values.items():
            name = f"{group}.{key}"
            if key not in defaults[group]:
                raise NoteError(f"{label}: unknown key {name}")
            expected = defaults[group][key]
            valid = (
                (value is None or isinstance(value, str))
                if name == "selection.slides"
                else type(value) is type(expected)
            )
            if not valid:
                raise NoteError(f"{label}: invalid type for {name}")
            if isinstance(value, list) and not all(isinstance(x, str) and x.strip() for x in value):
                raise NoteError(f"{label}: {name} requires nonempty strings")
            if isinstance(value, str) and not value.strip():
                raise NoteError(f"{label}: {name} must not be empty")
            if type(value) is int and value <= 0:
                raise NoteError(f"{label}: {name} must be positive")
            if name == "selection.off_slide" and value not in ("exclude", "append"):
                raise NoteError(f"{label}: off_slide must be exclude or append")
            if name == "generation.language" and value not in ("ja", "en"):
                raise NoteError(f"{label}: language must be ja or en")
            if name == "generation.image_width" and not 600 <= value <= 8000:
                raise NoteError(f"{label}: image_width must be 600..8000 pixels")
    return raw


def resolve(
    explicit: Path | None = None,
    overrides: dict[str, dict[str, Any]] | None = None,
    *,
    home: Path | None = None,
    cwd: Path | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    defaults: dict[str, Any] = yaml.safe_load(example())
    effective = copy.deepcopy(defaults)
    origins = {
        f"{group}.{key}": "built-in"
        for group, values in defaults.items()
        if isinstance(values, dict)
        for key in values
    }
    loaded = []
    global_path = (home / ".tkn/powerpoint_note/config.yaml") if home else user_config()
    candidates = [global_path, (cwd or Path.cwd()) / ".tkn/config.yaml"]
    if explicit is not None:
        explicit = explicit.expanduser().resolve()
        if not explicit.is_file():
            raise NoteError(f"Explicit config does not exist: {explicit}")
        candidates.append(explicit)
    for path in candidates:
        if not path.exists():
            continue
        try:
            raw = yaml.safe_load(path.read_text("utf-8-sig"))
            data = _validate(raw, str(path), defaults)
        except (OSError, yaml.YAMLError) as exc:
            raise NoteError(f"Cannot read config: {path} ({type(exc).__name__})") from exc
        loaded.append(
            {
                "path": str(path.resolve()),
                "schema_version": data["schema_version"],
                "migrated": False,
            }
        )
        for group, values in data.items():
            if group == "schema_version":
                continue
            effective[group].update(values)
            origins.update({f"{group}.{key}": str(path.resolve()) for key in values})
    if overrides:
        _validate({"schema_version": SCHEMA_VERSION, **overrides}, "CLI options", defaults)
        for group, values in overrides.items():
            effective[group].update(values)
            origins.update({f"{group}.{key}": "CLI" for key in values})
    return effective, {
        "sources": loaded,
        "winning_sources": origins,
        "effective_schema_version": SCHEMA_VERSION,
    }


def initialize(path: Path, dry_run: bool) -> dict[str, Any]:
    content = example()
    if path.exists():
        if path.read_text("utf-8-sig") != content:
            raise NoteError(f"Config already exists and was edited; preserved: {path}")
        status = "unchanged"
    else:
        status = "created"
        if not dry_run:
            atomic_write(path, content, overwrite=False)
    return {"status": status, "dry_run": dry_run, "path": str(path.resolve())}
