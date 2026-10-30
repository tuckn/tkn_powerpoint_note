"""Optional GenAI Bridge adapter with bounded calls and persisted usage."""

from __future__ import annotations

import importlib
import json
import logging
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import jsonschema

from .context_profiles import ContextProfile, load_profile
from .io import atomic_write, digest, fingerprint, json_text
from .models import NoteError


def provider_schema(schema: dict[str, Any]) -> dict[str, Any]:
    """Keep canonical validation while omitting provider-unsupported uniqueness constraints."""
    projected: dict[str, Any] = json.loads(json.dumps(schema))

    def visit(node: Any) -> None:
        if isinstance(node, dict):
            if node.get("type") == "array":
                node.pop("uniqueItems", None)
            for value in node.values():
                visit(value)
        elif isinstance(node, list):
            for value in node:
                visit(value)

    visit(projected)
    return projected


def generation_properties(folder: Path, files: dict[str, Any]) -> dict[str, Any]:
    """Publish actual saved call provenance, including mixed or unknown values."""
    values: dict[str, list[str]] = {
        key: []
        for key in (
            "generator",
            "generationProvider",
            "generationModel",
            "generationResponseModel",
            "generationReasoningEffort",
        )
    }

    def known(value: Any) -> str:
        return value if isinstance(value, str) else ""

    if "context.json" in files:
        for name in sorted(files):
            if not name.endswith(".usage.json"):
                continue
            usage = json.loads((folder / name).read_text("utf-8"))
            if usage.get("status") != "success":
                raise NoteError("Completed context contains an unsuccessful usage record")
            record = usage.get("record", {})
            provider = known(record.get("provider") or usage.get("provider"))
            requested = known(record.get("requested_model", usage.get("model")))
            response = known(record.get("response_model"))
            effort = known(record.get("reasoning_effort"))
            model = response or requested
            row = {
                "generator": f"{provider} ({model})" if provider and model else provider or model,
                "generationProvider": provider,
                "generationModel": requested,
                "generationResponseModel": response,
                "generationReasoningEffort": effort,
            }
            for key, value in row.items():
                if value not in values[key]:
                    values[key].append(value)
    return {
        key: ", ".join(items)
        if key == "generator"
        else items[0]
        if len(items) == 1
        else items
        if items
        else ""
        for key, items in values.items()
    }


class Generator:
    def __init__(self, settings: dict[str, Any], *, content_profile: ContextProfile | None = None):
        self.settings = settings
        self.content_profile = content_profile or load_profile(settings)
        try:
            self.bridge = importlib.import_module("tkn_genai_bridge")
        except ImportError as exc:
            raise NoteError(
                "GenAI Bridge could not be imported; reinstall with uv tool install . --reinstall"
            ) from exc
        try:
            self.profile = self.bridge.load_profile(
                settings["bridge_profile"], overrides=settings.get("overrides", {})
            )
            with self.bridge.Runtime(self.profile) as runtime:
                plan = runtime.plan(
                    self.bridge.GenerationRequest(
                        prompt=self.content_profile.prompts["slide"],
                        output_schema=provider_schema(self.content_profile.schemas["slide"]),
                        schema_name="powerpoint_slide",
                    ),
                    check_executable=True,
                )
            self.plan = {
                key: plan.model_dump(mode="json")[key]
                for key in (
                    "provider",
                    "model",
                    "profile_name",
                    "bridge_version",
                    "generation_settings_sha256",
                    "local_only",
                    "timeout_seconds",
                )
            }
            self.plan["reasoning_effort"] = plan.model_dump(mode="json").get("reasoning_effort")
        except self.bridge.GenAIError as exc:
            raise NoteError(
                f"GenAI Bridge configuration failed ({exc.code}); check the bridge profile"
            ) from exc

    def generate(
        self, stage: str, evidence: dict[str, Any], image: Path | None, folder: Path, name: str
    ) -> dict[str, Any]:
        payload = json_text(evidence)
        if len(payload) > self.settings["max_input_chars"]:
            raise NoteError(
                "AI evidence exceeds generation.max_input_chars; select fewer slides or raise the limit"
            )
        schema = self.content_profile.schemas[stage]
        prompt = self.content_profile.prompts[stage]
        prompt += "\nSOURCE EVIDENCE (JSON):\n" + payload
        request = self.bridge.GenerationRequest(
            prompt=prompt,
            output_schema=provider_schema(schema),
            schema_name="powerpoint_" + stage,
            images=[self.bridge.ImageInput.from_file(image)] if image else [],
        )
        journal = {
            "status": "started",
            "started_at": datetime.now(UTC).isoformat(),
            "prompt_sha256": digest(prompt.encode()),
            "schema_sha256": fingerprint(schema),
            **self.plan,
            "prompt_profile": self.content_profile.provenance(),
            "generation_generator": self.settings.get("generator_id"),
            "provider_schema_sha256": fingerprint(request.output_schema),
        }
        atomic_write(
            folder / (name + ".request.json"),
            json_text(
                {"prompt": prompt, "schema": schema, "provider_schema": request.output_schema}
            ),
        )
        journal_path = folder / (name + ".usage.json")
        atomic_write(journal_path, json_text(journal))
        started = time.monotonic()
        try:
            with self.bridge.Runtime(self.profile) as runtime:
                result = runtime.generate(request)
            data: dict[str, Any] = result.data
            jsonschema.Draft202012Validator(schema).validate(data)
            if stage == "slide" and data["slide_number"] != evidence["slide"]["number"]:
                raise NoteError("AI output cited a different slide number")
            if stage == "deck":
                allowed = {item["slide_number"] for item in evidence["slides"]}
                if any(set(item["slides"]) - allowed for item in data["key_points"]):
                    raise NoteError("AI summary cited a slide outside selected coverage")
            if "powerpoint-note:" in json.dumps(data):
                raise NoteError("AI returned reserved Markdown management markers")
            journal["status"] = "success"
            journal["record"] = result.record.model_dump(mode="json")
            atomic_write(folder / (name + ".response.json"), json_text(data))
            return data
        except self.bridge.GenAIError as exc:
            journal["status"] = "failed"
            journal["error_code"] = exc.code
            journal["error_message"] = str(exc)
            if exc.record is not None:
                journal["record"] = exc.record.model_dump(mode="json")
            raise NoteError(
                f"AI generation failed ({exc.code}); no automatic retry; inspect usage JSON"
            ) from exc
        except BaseException:
            journal["status"] = "failed"
            raise
        finally:
            journal["duration_seconds"] = round(time.monotonic() - started, 3)
            journal["finished_at"] = datetime.now(UTC).isoformat()
            atomic_write(journal_path, json_text(journal))
            logging.getLogger("powerpoint_note").info(
                "AI %s: %s (%.1fs)", name, journal["status"], journal["duration_seconds"]
            )
