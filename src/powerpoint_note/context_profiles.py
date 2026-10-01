"""Load and validate complete prompt/schema/template bundles from package or profile roots."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from importlib.resources import files
from pathlib import Path
from typing import Any

import yaml
from jsonschema import Draft202012Validator, SchemaError

from .io import digest, fingerprint
from .models import NoteError

RESOURCES = (
    "prompt.md",
    "output.schema.json",
    "template.md",
    "deck-prompt.md",
    "deck-output.schema.json",
    "deck-template.md",
    "note-template.md",
    "source-template.md",
)
FIELDS = {
    "slide": {"summary", "sections", "uncertainties"},
    "deck": {"summary", "key_points", "uncertainties"},
    "note": {"coverage", "mode", "evidence", "overview", "slides"},
    "source": {
        "number",
        "title",
        "slide_id",
        "page_number",
        "hidden",
        "section",
        "image",
        "context",
        "content",
        "notes",
        "comments",
        "outside",
        "warnings",
    },
}
LABELS = {
    "local",
    "ai",
    "personal",
    "evidence",
    "slide",
    "column",
    "merged_cells",
    "chart",
    "smartart",
    "object",
    "endpoints",
    "text",
    "location_outside",
    "location_partial",
    "location_unknown",
    "hidden_yes",
    "hidden_no",
    "unknown_value",
}
TOKEN = re.compile(r"{{\s*([#/]?)([a-z][a-z0-9_]*)\s*}}")
# Renderer compatibility is code; semantic constraints and descriptions live in the schemas.
CONTRACTS: dict[str, Any] = {
    "slide": {
        "slide_number": "integer",
        "summary": "string",
        "sections": [{"heading": "string", "body": "string"}],
        "uncertainties": ["string"],
    },
    "deck": {
        "summary": "string",
        "key_points": [{"text": "string", "slides": ["integer"]}],
        "uncertainties": ["string"],
    },
}


def _validate_template(template: str, fields: set[str], label: str) -> None:
    counts = dict.fromkeys(fields, 0)
    stack: list[str] = []
    for match in TOKEN.finditer(template):
        kind, field = match.groups()
        if field not in fields:
            raise NoteError(f"{label}: unknown template field {field}")
        if kind == "#":
            if stack:
                raise NoteError(f"{label}: nested conditional sections are unsupported")
            stack.append(field)
        elif kind == "/":
            if not stack or stack.pop() != field:
                raise NoteError(f"{label}: unmatched template conditional {field}")
        else:
            if stack and stack[-1] != field:
                raise NoteError(f"{label}: a conditional can contain only its own field")
            counts[field] += 1
    remainder = TOKEN.sub("", template)
    if stack or "{{" in remainder or "}}" in remainder or any(n != 1 for n in counts.values()):
        raise NoteError(f"{label}: each field must occur once with balanced conditional sections")


def render_template(template: str, values: dict[str, str]) -> str:
    """Substitute only original template tokens, never tokens inside source/model text."""
    output: list[str] = []
    active = True
    position = 0
    for match in TOKEN.finditer(template):
        if active:
            output.append(template[position : match.start()])
        kind, field = match.groups()
        if kind == "#":
            active = bool(values[field].strip())
        elif kind == "/":
            active = True
        elif active:
            output.append(values[field])
        position = match.end()
    if active:
        output.append(template[position:])
    return "".join(output).strip() + "\n"


def _check_contract(schema: dict[str, Any], contract: Any, label: str) -> None:
    if not isinstance(schema, dict):
        raise NoteError(f"{label}: schema nodes must be objects with explicit types")
    if isinstance(contract, dict):
        properties = schema.get("properties", {})
        if (
            schema.get("type") != "object"
            or set(properties) != set(contract)
            or set(schema.get("required", [])) != set(contract)
            or schema.get("additionalProperties") is not False
        ):
            raise NoteError(f"{label}: schema fields do not match the renderer contract")
        for name, value in contract.items():
            _check_contract(properties[name], value, f"{label}.{name}")
    elif isinstance(contract, list):
        if schema.get("type") != "array" or not isinstance(schema.get("items"), dict):
            raise NoteError(f"{label}: expected array items")
        _check_contract(schema["items"], contract[0], label + "[]")
    elif schema.get("type") != contract:
        raise NoteError(f"{label}: expected {contract}")


def _check_references(node: Any, label: str) -> None:
    if isinstance(node, dict):
        if any(key in node for key in ("$ref", "$dynamicRef", "$recursiveRef")):
            raise NoteError(
                f"{label}: schema references are unsupported; use a self-contained schema"
            )
        for child in node.values():
            _check_references(child, label)
    elif isinstance(node, list):
        for child in node:
            _check_references(child, label)


@dataclass(frozen=True)
class ContextProfile:
    name: str
    version: str
    language: str
    source: str
    labels: dict[str, str]
    prompts: dict[str, str]
    schemas: dict[str, dict[str, Any]]
    templates: dict[str, str]
    resource_hashes: dict[str, str]

    def provenance(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "version": self.version,
            "language": self.language,
            "sha256": fingerprint(self.resource_hashes),
            "resources": self.resource_hashes,
        }

    def render(self, stage: str, values: dict[str, str]) -> str:
        return render_template(self.templates[stage], values)


def load_profile(settings: dict[str, Any]) -> ContextProfile:
    name = settings["prompt_profile"]
    if not isinstance(name, str) or not re.fullmatch(r"[a-z0-9][a-z0-9._-]*", name):
        raise NoteError(f"Invalid generation.prompt_profile: {name!r}")
    packaged = files("powerpoint_note").joinpath("context_profiles", name)
    custom = next(
        (
            Path(root).expanduser().resolve() / name
            for root in settings["profile_dirs"]
            if (Path(root).expanduser().resolve() / name).exists()
        ),
        None,
    )
    root = custom if custom is not None else packaged
    source = (
        str(custom) if custom is not None else "package:powerpoint_note/context_profiles/" + name
    )
    if not root.is_dir():
        raise NoteError(
            f"Unknown or invalid prompt profile {name!r}; built-ins: default-ja, default-en; "
            "check generation.prompt_profile and generation.profile_dirs"
        )
    try:
        resources = {
            name: root.joinpath(name).read_text(encoding="utf-8").replace("\r\n", "\n")
            for name in RESOURCES
        }
    except (OSError, UnicodeError) as exc:
        raise NoteError(f"Incomplete or unreadable prompt profile {name!r}: {source}") from exc
    if any("powerpoint-note:" in value for value in resources.values()):
        raise NoteError(f"Profile {name!r} contains reserved note management markers")
    match = re.match(r"\A---\n(.*?)\n---\n(.*)\Z", resources["template.md"], re.DOTALL)
    if not match:
        raise NoteError(f"Profile {name!r}: template.md needs YAML metadata")
    try:
        metadata = yaml.safe_load(match[1])
        if (
            not isinstance(metadata, dict)
            or not isinstance(metadata.get("version"), str)
            or not re.fullmatch(r"1\.0\.(0|[1-9]\d*)", metadata["version"])
        ):
            raise NoteError(f"Profile {name!r}: unsupported template version; supported: 1.0.x")
        language, labels = metadata.get("language"), metadata.get("labels")
        if not isinstance(language, str) or not language.strip():
            raise NoteError(f"Profile {name!r}: language must be nonempty template metadata")
        if (
            not isinstance(labels, dict)
            or not labels.keys() >= LABELS
            or not all(isinstance(v, str) and v.strip() for v in labels.values())
        ):
            raise NoteError(f"Profile {name!r}: template labels are missing or invalid")
        schemas: dict[str, dict[str, Any]] = {}
        for stage, filename in (
            ("slide", "output.schema.json"),
            ("deck", "deck-output.schema.json"),
        ):
            schema = json.loads(resources[filename])
            if not isinstance(schema, dict):
                raise NoteError(f"{name}/{filename}: schema must be an object")
            _check_references(schema, filename)
            Draft202012Validator.check_schema(schema)
            _check_contract(schema, CONTRACTS[stage], f"{name}/{filename}")
            schemas[stage] = schema
    except (yaml.YAMLError, ValueError, SchemaError) as exc:
        raise NoteError(
            f"Profile {name!r}: invalid metadata or JSON schema ({type(exc).__name__})"
        ) from exc
    templates = {
        "slide": match[2],
        "deck": resources["deck-template.md"],
        "note": resources["note-template.md"],
        "source": resources["source-template.md"],
    }
    for stage, template in templates.items():
        _validate_template(template, FIELDS[stage], f"{name}/{stage}")
    prompts = {}
    for stage, filename in (("slide", "prompt.md"), ("deck", "deck-prompt.md")):
        prompt = resources[filename]
        if prompt.count("{{language}}") != 1:
            raise NoteError(f"{name}/{filename}: expected one language placeholder")
        remainder = prompt.replace("{{language}}", "")
        if "{{" in remainder or "}}" in remainder:
            raise NoteError(f"{name}/{filename}: unknown prompt placeholder")
        prompts[stage] = prompt.replace("{{language}}", language)
    return ContextProfile(
        name,
        metadata["version"],
        language,
        source,
        labels,
        prompts,
        schemas,
        templates,
        {name: digest(value.encode("utf-8")) for name, value in resources.items()},
    )
