"""Layered, strict configuration: sources, extraction, connections and generation settings."""

from __future__ import annotations

import copy
import importlib
import math
import os
import re
from importlib.resources import files
from pathlib import Path
from typing import Any

import yaml

from .context_profiles import load_profile
from .io import atomic_write
from .models import ConfigError

APPLICATION_ID = "powerpoint_note"
SCHEMA_VERSION = "4.0.0"
GENERATOR_DEFAULTS: dict[str, Any] = {
    "genai_profile": "codex-default",
    "prompt_profile": "default-ja",
    "overrides": {},
}
SOURCE_DEFAULTS: dict[str, Any] = {
    "presentations": {
        "dir": "",
        "recursive": False,
        "include": ["*.pptx", "*.pptm"],
        "ignore": [],
    },
    "notes": {"dir": "", "frontmatter_term_format": "obsidian-link"},
    "generation": {"generator": None},
}
TERM_FORMATS = ("obsidian-link", "plain")


def example() -> str:
    return files("powerpoint_note").joinpath("resources/config.example.yaml").read_text("utf-8")


def user_root() -> Path:
    return Path.home() / ".tkn" / APPLICATION_ID


def user_config() -> Path:
    return user_root() / "config.yaml"


def state_root() -> Path:
    return user_root() / "state"


def default_genai_profiles() -> dict[str, Any]:
    """Built-in connections, as written by config init."""
    profiles: dict[str, Any] = yaml.safe_load(example())["genai_profiles"]
    return profiles


def bridge() -> Any:
    """Import GenAI Bridge only when connection settings or AI calls need it."""
    try:
        return importlib.import_module("tkn_genai_bridge")
    except ImportError as exc:
        raise ConfigError(
            "GenAI Bridge could not be imported; reinstall with uv tool install . --reinstall"
        ) from exc


def _name(value: Any, label: str) -> None:
    if not isinstance(value, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", value):
        raise ConfigError(
            f"{label}: expected a nonempty name using letters, digits, '.', '_' or '-'"
        )


def _validate_genai_profiles(values: Any, label: str) -> None:
    """Check each supplied connection field before merging; partial layers are allowed."""
    if not isinstance(values, dict):
        raise ConfigError(f"{label}: genai_profiles must be a mapping")
    from pydantic import TypeAdapter, ValidationError

    fields = bridge().Profile.model_fields
    for identifier, spec in values.items():
        _name(identifier, f"{label}: genai_profiles key")
        location = f"{label}: genai_profiles.{identifier}"
        if not isinstance(spec, dict):
            raise ConfigError(f"{location}: expected a mapping")
        for key, value in spec.items():
            if key not in fields:
                raise ConfigError(f"{location}: unknown key {key}")
            try:
                TypeAdapter(fields[key].rebuild_annotation()).validate_python(value, strict=True)
            except ValidationError as exc:
                # A layer may set only part of a nested entry (such as one pricing field);
                # required fields are checked after merging. Messages can include values,
                # so report the field only.
                errors = exc.errors(include_input=False, include_url=False)
                if any(error["type"] != "missing" for error in errors):
                    raise ConfigError(f"{location}: invalid value for {key}") from None
            if (
                key in {"model", "reasoning_effort"}
                and isinstance(value, str)
                and (not value.strip() or "\x00" in value)
            ):
                raise ConfigError(f"{location}: {key} must not be empty or contain NUL")


def _resolve_genai_paths(data: dict[str, Any], base: Path) -> None:
    """Resolve executable paths from the file that wrote them; bare names use PATH."""
    for spec in data.get("genai_profiles", {}).values():
        cli = spec.get("cli")
        executable = cli.get("executable") if isinstance(cli, dict) else None
        if isinstance(executable, str) and (
            "/" in executable or "\\" in executable or executable.startswith("~")
        ):
            cli["executable"] = str((base / Path(executable).expanduser()).resolve())


def connection(
    genai_profiles: dict[str, Any], name: str, overrides: dict[str, Any] | None = None
) -> Any:
    """Validated Bridge Profile from this application's genai_profiles only."""
    if name not in genai_profiles:
        raise ConfigError(f"GenAI profile is not defined in genai_profiles: {name}")
    from pydantic import ValidationError

    try:
        return bridge().Profile.model_validate({**genai_profiles[name], **(overrides or {})})
    except ValidationError as exc:
        reasons = "; ".join(
            error["msg"] for error in exc.errors(include_input=False, include_url=False)
        )
        raise ConfigError(
            f"genai_profiles.{name}: invalid connection settings: {reasons}"
        ) from None


def _validate_overrides(values: Any, label: str) -> None:
    if not isinstance(values, dict):
        raise ConfigError(f"{label}: expected a mapping")
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
            raise ConfigError(
                f"{label}: unknown key {key}; define other connection settings in genai_profiles"
            )
        if not valid:
            raise ConfigError(f"{label}: invalid value for {key}")


def _validate_value(value: Any, expected: Any, name: str, label: str) -> None:
    if type(value) is not type(expected):
        raise ConfigError(f"{label}: invalid type for {name}")
    if isinstance(value, list) and not all(isinstance(x, str) and x.strip() for x in value):
        raise ConfigError(f"{label}: {name} requires nonempty strings")
    if isinstance(value, str) and (not value.strip() or "\x00" in value):
        raise ConfigError(f"{label}: {name} must not be empty or contain NUL")
    if type(value) is int and value <= 0:
        raise ConfigError(f"{label}: {name} must be positive")
    if name == "extraction.off_slide" and value not in ("exclude", "append"):
        raise ConfigError(f"{label}: off_slide must be exclude or append")
    if name == "generation.limits.image_width" and not 600 <= value <= 8000:
        raise ConfigError(f"{label}: image_width must be 600..8000 pixels")


def _validate_mapping(
    values: Any, expected: dict[str, Any], name: str, label: str
) -> dict[str, Any]:
    if not isinstance(values, dict):
        raise ConfigError(f"{label}: {name} must be a mapping")
    for key in values:
        if key not in expected:
            raise ConfigError(f"{label}: unknown key {name}.{key}")
    return values


def _validate_sources(values: Any, label: str) -> None:
    if not isinstance(values, dict):
        raise ConfigError(f"{label}: sources must be a mapping of source IDs")
    for identifier, spec in values.items():
        _name(identifier, f"{label}: sources key")
        where = f"sources.{identifier}"
        _validate_mapping(spec, SOURCE_DEFAULTS, where, label)
        for group in ("presentations", "notes"):
            section = _validate_mapping(
                spec.get(group), SOURCE_DEFAULTS[group], f"{where}.{group}", label
            )
            if "dir" not in section:
                raise ConfigError(f"{label}: {where}.{group}.dir is required")
            for key, value in section.items():
                _validate_value(value, SOURCE_DEFAULTS[group][key], f"{where}.{group}.{key}", label)
        term_format = spec["notes"].get("frontmatter_term_format", TERM_FORMATS[0])
        if term_format not in TERM_FORMATS:
            raise ConfigError(
                f"{label}: {where}.notes.frontmatter_term_format must be obsidian-link or plain"
            )
        generation = _validate_mapping(
            spec.get("generation", {}), SOURCE_DEFAULTS["generation"], f"{where}.generation", label
        )
        if generation.get("generator") is not None:
            _name(generation["generator"], f"{label}: {where}.generation.generator")


def _validate_generation(values: dict[str, Any], label: str, defaults: dict[str, Any]) -> None:
    for key, value in values.items():
        name = f"generation.{key}"
        if key not in defaults:
            raise ConfigError(f"{label}: unknown key {name}")
        if key == "default_generator":
            if value is not None:
                _name(value, f"{label}: {name}")
        elif key == "generators":
            if not isinstance(value, dict):
                raise ConfigError(f"{label}: {name} must be a mapping")
            for identifier, spec in value.items():
                location = f"{label}: {name}.{identifier}"
                _name(identifier, f"{label}: {name} key")
                if not isinstance(spec, dict):
                    raise ConfigError(f"{location}: expected a mapping")
                for field, setting in spec.items():
                    if field not in GENERATOR_DEFAULTS:
                        raise ConfigError(f"{location}: unknown key {field}")
                    if field == "overrides":
                        _validate_overrides(setting, f"{location}.overrides")
                    else:
                        _validate_value(setting, GENERATOR_DEFAULTS[field], field, location)
        elif key == "overrides":
            _validate_overrides(value, f"{label}: {name}")
        elif key == "limits":
            for field, setting in _validate_mapping(value, defaults[key], name, label).items():
                _validate_value(setting, defaults[key][field], f"{name}.{field}", label)
        else:
            _validate_value(value, defaults[key], name, label)


def _validate(raw: Any, label: str, defaults: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise ConfigError(f"{label}: expected a YAML mapping")
    version = raw.get("schema_version", SCHEMA_VERSION)
    if version != SCHEMA_VERSION:
        raise ConfigError(
            f"{label}: unsupported schema_version {version!r}; this version reads only "
            f'"{SCHEMA_VERSION}". Rewrite the file from `config init` (see CHANGELOG).'
        )
    for group, values in raw.items():
        if group == "schema_version":
            continue
        if group == "genai_profiles":
            _validate_genai_profiles(values, label)
        elif group == "sources":
            _validate_sources(values, label)
        elif group == "generation":
            _validate_generation(
                _validate_mapping(values, defaults[group], group, label), label, defaults[group]
            )
        elif group == "extraction":
            for key, value in _validate_mapping(values, defaults[group], group, label).items():
                _validate_value(value, defaults[group][key], f"{group}.{key}", label)
        else:
            raise ConfigError(f"{label}: unknown key {group}")
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


def _complete_sources(sources: dict[str, Any], cwd: Path) -> dict[str, Any]:
    """Fill defaults and resolve directories from the working directory."""
    completed: dict[str, Any] = {}
    for identifier, spec in sources.items():
        item = copy.deepcopy(SOURCE_DEFAULTS)
        _merge(item, spec)
        for group in ("presentations", "notes"):
            # Absolute, but not resolved: a folder reached through a link keeps the given path.
            item[group]["dir"] = os.path.abspath(cwd / Path(item[group]["dir"]).expanduser())
        completed[identifier] = item
    return completed


def resolve(
    explicit: Path | None = None,
    overrides: dict[str, dict[str, Any]] | None = None,
    *,
    generator: str | None = None,
    source: str | None = None,
    home: Path | None = None,
    cwd: Path | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Merge built-in, user, working-directory and explicit settings, then CLI options."""
    cwd = cwd or Path.cwd()
    defaults: dict[str, Any] = yaml.safe_load(example())
    # The example shows a sample source and named presets; neither is a built-in default.
    defaults["sources"] = {}
    defaults["generation"].update(copy.deepcopy(GENERATOR_DEFAULTS))
    defaults["generation"].update(default_generator=None, generators={})
    effective = copy.deepcopy(defaults)
    origins = dict.fromkeys(_paths(defaults), "built-in")
    loaded = []
    global_path = (home / ".tkn" / APPLICATION_ID / "config.yaml") if home else user_config()
    candidates = [global_path, cwd / ".tkn/config.yaml"]
    if explicit is not None:
        explicit = explicit.expanduser().resolve()
        if not explicit.is_file():
            raise ConfigError(f"Explicit config does not exist: {explicit}")
        candidates.append(explicit)
    custom_connections = False
    for path in candidates:
        if not path.exists():
            continue
        try:
            raw = yaml.safe_load(path.read_text("utf-8-sig"))
        except (OSError, yaml.YAMLError) as exc:
            raise ConfigError(f"Cannot read config: {path} ({type(exc).__name__})") from exc
        data = _validate(raw, str(path), defaults)
        if "genai_profiles" in data:
            custom_connections = True
            _resolve_genai_paths(data, path.parent)
        location = str(path.resolve())
        loaded.append({"path": location, "schema_version": data.get("schema_version")})
        fragment = {key: value for key, value in data.items() if key != "schema_version"}
        if "sources" in fragment:
            # A later file replaces the whole list of sources; one source is never split.
            effective["sources"] = {}
            origins = {k: v for k, v in origins.items() if not k.startswith("sources.")}
        _merge(effective, fragment)
        origins.update(dict.fromkeys(_paths(fragment), location))
    effective["sources"] = _complete_sources(effective["sources"], cwd)
    generation = effective["generation"]
    default_generator = generation["default_generator"]
    if default_generator is not None and default_generator not in generation["generators"]:
        raise ConfigError(f"Unknown generation.default_generator: {default_generator!r}")
    if source is not None and source not in effective["sources"]:
        known = ", ".join(sorted(effective["sources"])) or "none"
        raise ConfigError(f"Unknown source: {source!r}; configured sources: {known}")
    source_generator = (
        effective["sources"][source]["generation"]["generator"] if source is not None else None
    )
    selected, origin = default_generator, origins["generation.default_generator"]
    if source_generator is not None:
        selected = source_generator
        origin = origins.get(f"sources.{source}.generation.generator", origin)
    if generator is not None:
        _name(generator, "CLI --generator")
        selected, origin = generator, "CLI"
    if selected is not None:
        if selected not in generation["generators"]:
            raise ConfigError(f"Unknown generator: {selected!r}")
        spec = generation["generators"][selected]
        _merge(generation, spec)
        prefix = f"generation.generators.{selected}"
        for field_path in _paths(spec):
            origins[f"generation.{field_path}"] = origins[f"{prefix}.{field_path}"]
    if overrides:
        _validate({"schema_version": SCHEMA_VERSION, **overrides}, "CLI options", defaults)
        _merge(effective, overrides)
        origins.update(dict.fromkeys(_paths(overrides), "CLI"))
    if custom_connections:
        # Built-ins are known valid; configured connections also need cross-field checks.
        for name in effective["genai_profiles"]:
            connection(effective["genai_profiles"], name)
    generation["generator_id"] = selected
    origins["generation.generator_id"] = origin
    generation["prompt_profile_dirs"] = [
        str((cwd / Path(root).expanduser()).resolve()) for root in generation["prompt_profile_dirs"]
    ]
    profile = load_profile(generation)
    return effective, {
        "selected_generator": selected,
        "selected_prompt_profile": profile.name,
        "selected_genai_profile": generation["genai_profile"],
        "prompt_profile_source": profile.source,
        "loaded_sources": loaded,
        "winning_sources": origins,
        "effective_schema_version": SCHEMA_VERSION,
    }


def initialize(path: Path, dry_run: bool) -> dict[str, Any]:
    content = example()
    if path.exists():
        if path.read_text("utf-8-sig") != content:
            raise ConfigError(f"Config already exists and was edited; preserved: {path}")
        status = "unchanged"
    else:
        status = "created"
        if not dry_run:
            atomic_write(path, content, overwrite=False)
    return {
        "status": status,
        "command": "config init",
        "dryRun": dry_run,
        "configPath": str(path.resolve()),
    }
