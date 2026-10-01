"""Layered, strict configuration and reusable named generation settings."""

from __future__ import annotations

import copy
import math
import re
from importlib.resources import files
from pathlib import Path
from typing import Any

import yaml

from .context_profiles import load_profile
from .io import atomic_write
from .models import NoteError

SCHEMA_VERSION = "2.1.0"
GENERATOR_DEFAULTS: dict[str, Any] = {
    "bridge_profile": "codex-default",
    "prompt_profile": "default-ja",
    "overrides": {},
}


def example() -> str:
    return files("powerpoint_note").joinpath("resources/config.example.yaml").read_text("utf-8")


def user_config() -> Path:
    return Path.home() / ".tkn" / "powerpoint_note" / "config.yaml"


def _name(value: Any, label: str) -> None:
    if not isinstance(value, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", value):
        raise NoteError(f"{label}: expected a nonempty name using letters, digits, '.', '_' or '-'")


def _validate_overrides(values: Any, label: str) -> None:
    if not isinstance(values, dict):
        raise NoteError(f"{label}: expected a mapping")
    for key, value in values.items():
        valid = False
        if key in {"model", "reasoning_effort"}:
            valid = value is None or (
                isinstance(value, str) and bool(value.strip()) and "\x00" not in value
            )
        elif key == "timeout_seconds":
            valid = type(value) in (int, float) and 0 < value <= 86400 and math.isfinite(value)
        elif key == "max_output_tokens":
            valid = value is None or (type(value) is int and value > 0)
        elif key == "local_only":
            valid = type(value) is bool
        else:
            raise NoteError(
                f"{label}: unknown key {key}; configure connection details in GenAI Bridge"
            )
        if not valid:
            raise NoteError(f"{label}: invalid value for {key}")


def _validate_value(value: Any, expected: Any, name: str, label: str) -> None:
    valid = (
        (value is None or isinstance(value, str))
        if name == "selection.slides"
        else type(value) is type(expected)
    )
    if not valid:
        raise NoteError(f"{label}: invalid type for {name}")
    if isinstance(value, list) and not all(isinstance(x, str) and x.strip() for x in value):
        raise NoteError(f"{label}: {name} requires nonempty strings")
    if isinstance(value, str) and (not value.strip() or "\x00" in value):
        raise NoteError(f"{label}: {name} must not be empty or contain NUL")
    if type(value) is int and value <= 0:
        raise NoteError(f"{label}: {name} must be positive")
    if name == "selection.off_slide" and value not in ("exclude", "append"):
        raise NoteError(f"{label}: off_slide must be exclude or append")
    if name == "generation.image_width" and not 600 <= value <= 8000:
        raise NoteError(f"{label}: image_width must be 600..8000 pixels")


def _validate(raw: Any, label: str, defaults: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise NoteError(f"{label}: expected a YAML mapping")
    version = raw.get("schema_version")
    if not isinstance(version, str) or not re.fullmatch(
        r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)", version
    ):
        raise NoteError(
            f"{label}: schema_version must be a quoted SemVer string ({SCHEMA_VERSION})"
        )
    if version.startswith("1.0."):
        raise NoteError(
            f"{label}: config 1.0.x was replaced by 2.x; set schema_version: '{SCHEMA_VERSION}' "
            "and replace generation.language ja/en with generation.prompt_profile default-ja/default-en"
        )
    if tuple(map(int, version.split(".")))[:2] not in {(2, 0), (2, 1)}:
        raise NoteError(f"{label}: unsupported schema {version}; supported: 2.0.x and 2.1.x")
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
            if name == "generation.default_generator":
                if value is not None:
                    _name(value, f"{label}: {name}")
            elif name == "generation.generators":
                if not isinstance(value, dict):
                    raise NoteError(f"{label}: {name} must be a mapping")
                for identifier, spec in value.items():
                    location = f"{label}: {name}.{identifier}"
                    _name(identifier, f"{label}: {name} key")
                    if not isinstance(spec, dict):
                        raise NoteError(f"{location}: expected a mapping")
                    for field, setting in spec.items():
                        if field not in GENERATOR_DEFAULTS:
                            raise NoteError(f"{location}: unknown key {field}")
                        if field == "overrides":
                            _validate_overrides(setting, f"{location}.overrides")
                        else:
                            _validate_value(setting, GENERATOR_DEFAULTS[field], field, location)
            elif name == "generation.overrides":
                _validate_overrides(value, f"{label}: {name}")
            else:
                _validate_value(value, defaults[group][key], name, label)
    return raw


def _merge(target: dict[str, Any], overlay: dict[str, Any]) -> None:
    for key, value in overlay.items():
        if isinstance(value, dict) and isinstance(target.get(key), dict):
            _merge(target[key], value)
        else:
            target[key] = copy.deepcopy(value)


def _paths(values: dict[str, Any], prefix: str = "") -> list[str]:
    paths = []
    for key, value in values.items():
        name = f"{prefix}.{key}" if prefix else key
        paths.append(name)
        if isinstance(value, dict):
            paths.extend(_paths(value, name))
    return paths


def resolve(
    explicit: Path | None = None,
    overrides: dict[str, dict[str, Any]] | None = None,
    *,
    generator: str | None = None,
    home: Path | None = None,
    cwd: Path | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    defaults: dict[str, Any] = yaml.safe_load(example())
    # The example activates named presets; an existing flat config must retain its defaults.
    defaults["generation"].update(copy.deepcopy(GENERATOR_DEFAULTS))
    defaults["generation"].update(default_generator=None, generators={})
    effective = copy.deepcopy(defaults)
    origins = dict.fromkeys(_paths(defaults), "built-in")
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
        source = str(path.resolve())
        loaded.append({"path": source, "schema_version": data["schema_version"], "migrated": False})
        fragment = {key: value for key, value in data.items() if key != "schema_version"}
        _merge(effective, fragment)
        origins.update(dict.fromkeys(_paths(fragment), source))
    generation = effective["generation"]
    default_generator = generation["default_generator"]
    if default_generator is not None and default_generator not in generation["generators"]:
        raise NoteError(f"Unknown generation.default_generator: {default_generator!r}")
    if generator is not None:
        _name(generator, "CLI --generator")
    selected = generator if generator is not None else default_generator
    if selected is not None:
        if selected not in generation["generators"]:
            raise NoteError(f"Unknown generator: {selected!r}")
        spec = generation["generators"][selected]
        _merge(generation, spec)
        prefix = f"generation.generators.{selected}"
        for field_path in _paths(spec):
            origins[f"generation.{field_path}"] = origins[f"{prefix}.{field_path}"]
    if overrides:
        _validate({"schema_version": SCHEMA_VERSION, **overrides}, "CLI options", defaults)
        _merge(effective, overrides)
        origins.update(dict.fromkeys(_paths(overrides), "CLI"))
    generation["generator_id"] = selected
    origins["generation.generator_id"] = (
        "CLI" if generator is not None else origins["generation.default_generator"]
    )
    generation["profile_dirs"] = [
        str(((cwd or Path.cwd()) / Path(root).expanduser()).resolve())
        for root in generation["profile_dirs"]
    ]
    profile = load_profile(generation)
    return effective, {
        "selected_generator": selected,
        "prompt_profile": {**profile.provenance(), "source": profile.source},
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
