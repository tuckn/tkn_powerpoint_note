"""GenAI Bridge adapter: one bounded call per slide or summary, with a saved journal."""

from __future__ import annotations

import json
import logging
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import jsonschema

from .config import bridge, connection
from .context_profiles import ContextProfile
from .evidence import generation_key
from .io import atomic_write, digest, json_text
from .models import NoteError

LOGGER = logging.getLogger("powerpoint_note")
PLAN_FIELDS = (
    "provider",
    "model",
    "profile_name",
    "bridge_version",
    "generation_settings_sha256",
    "local_only",
    "timeout_seconds",
    "reasoning_effort",
)


class ContextError(NoteError):
    """An AI call failed or returned output this application cannot accept."""


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


def generator_label(usage: dict[str, Any]) -> str:
    """`provider (model)` from a saved call record; unknown parts are left out, not guessed."""
    record = usage.get("record") or {}
    provider = record.get("provider") or usage.get("provider") or ""
    model = record.get("response_model") or record.get("requested_model") or usage.get("model")
    if not isinstance(provider, str):
        provider = ""
    if not isinstance(model, str):
        model = ""
    return f"{provider} ({model})" if provider and model else provider or model


class Generator:
    def __init__(
        self,
        settings: dict[str, Any],
        genai_profiles: dict[str, Any],
        content_profile: ContextProfile,
    ):
        self.settings = settings
        self.content_profile = content_profile
        self.bridge = bridge()
        # Connections come only from this application's configuration, never Bridge files.
        self.profile_name = settings["genai_profile"]
        self.profile = connection(genai_profiles, self.profile_name, settings.get("overrides", {}))
        try:
            with self.bridge.Runtime(self.profile, profile_name=self.profile_name) as runtime:
                plan = runtime.plan(
                    self.bridge.GenerationRequest(
                        prompt=self.content_profile.prompts["slide"],
                        output_schema=provider_schema(self.content_profile.schemas["slide"]),
                        schema_name="powerpoint_slide",
                    ),
                    check_executable=True,
                )
        except self.bridge.GenAIError as exc:
            raise ContextError(
                f"GenAI Bridge configuration failed ({exc.code}); "
                f"check genai_profiles.{self.profile_name}"
            ) from exc
        dumped = plan.model_dump(mode="json")
        self.plan = {key: dumped.get(key) for key in PLAN_FIELDS}

    def key(self, stage: str, content: str) -> str:
        """Reuse key: content, prompt, schema, profile and the model that answers."""
        return generation_key(
            content,
            stage,
            {
                "name": self.content_profile.name,
                "language": self.content_profile.language,
                "prompt": self.content_profile.prompts[stage],
                "schema": self.content_profile.schemas[stage],
            },
            self.plan,
        )

    def generate(
        self,
        stage: str,
        payload: str,
        image: Path | None,
        folder: Path,
        key: str,
        *,
        allowed_slides: set[int] | None = None,
    ) -> dict[str, Any]:
        """Call the provider once and save request, journal and the validated response."""
        schema = self.content_profile.schemas[stage]
        prompt = self.content_profile.prompts[stage] + "\nSOURCE EVIDENCE (JSON):\n" + payload
        request = self.bridge.GenerationRequest(
            prompt=prompt,
            output_schema=provider_schema(schema),
            schema_name="powerpoint_" + stage,
            images=[self.bridge.ImageInput.from_file(image)] if image else [],
        )
        journal: dict[str, Any] = {
            "status": "started",
            "key": key,
            "stage": stage,
            "started_at": datetime.now(UTC).isoformat(),
            "input_chars": len(payload),
            "prompt_sha256": digest(prompt.encode()),
            **self.plan,
            "prompt_profile": self.content_profile.name,
            "generation_generator": self.settings.get("generator_id"),
        }
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "response.json").unlink(missing_ok=True)
        atomic_write(
            folder / "request.json",
            json_text(
                {"prompt": prompt, "schema": schema, "provider_schema": request.output_schema}
            ),
        )
        journal_path = folder / "usage.json"
        atomic_write(journal_path, json_text(journal))
        started = time.monotonic()
        data: Any = None
        try:
            with self.bridge.Runtime(self.profile, profile_name=self.profile_name) as runtime:
                result = runtime.generate(request)
            data = result.data
            journal["record"] = result.record.model_dump(mode="json")
            try:
                jsonschema.Draft202012Validator(schema).validate(data)
            except jsonschema.ValidationError as exc:
                where = "/".join(str(part) for part in exc.absolute_path) or "(root)"
                raise ContextError(f"AI output does not match the schema at {where}") from exc
            if (
                stage == "deck"
                and allowed_slides is not None
                and any(set(item["slides"]) - allowed_slides for item in data["key_points"])
            ):
                raise ContextError("AI summary cited a slide outside the supplied slides")
            if "powerpoint-note:" in json.dumps(data):
                raise ContextError("AI returned reserved Markdown management markers")
            journal["status"] = "success"
            atomic_write(folder / "response.json", json_text(data))
            validated: dict[str, Any] = data
            return validated
        except self.bridge.GenAIError as exc:
            journal["status"] = "failed"
            journal["error_code"] = exc.code
            journal["error_message"] = str(exc)
            if exc.record is not None:
                journal["record"] = exc.record.model_dump(mode="json")
            raise ContextError(
                f"AI generation failed ({exc.code}); no automatic retry; see {journal_path}"
            ) from exc
        except ContextError as exc:
            # Keep what the provider returned so the rejection can be diagnosed.
            journal["status"] = "rejected"
            journal["error_message"] = str(exc)
            atomic_write(folder / "rejected.json", json_text(data))
            raise ContextError(f"{exc}; the rejected output is in {folder}") from exc
        except BaseException as exc:
            journal["status"] = "failed"
            journal["error_message"] = type(exc).__name__
            raise
        finally:
            journal["duration_seconds"] = round(time.monotonic() - started, 3)
            journal["finished_at"] = datetime.now(UTC).isoformat()
            atomic_write(journal_path, json_text(journal))
            LOGGER.info(
                "AI %s %s: %s (%.1fs)",
                stage,
                folder.name,
                journal["status"],
                journal["duration_seconds"],
            )
