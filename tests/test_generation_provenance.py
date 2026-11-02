from __future__ import annotations

import json

import pytest
from test_pipeline import fake_context as fake_context

from powerpoint_note import pipeline
from powerpoint_note.context import generation_properties
from powerpoint_note.io import file_digest
from powerpoint_note.markdown import split_note


@pytest.mark.parametrize(
    "requested, response, effort, expected",
    [
        ("requested-model", "actual-model", "high", "codex (actual-model)"),
        ("requested-model", None, None, "codex (requested-model)"),
        (None, None, None, "codex"),
    ],
)
def test_saved_record_identity_and_effort(tmp_path, requested, response, effort, expected):
    usage = tmp_path / "deck.usage.json"
    usage.write_text(
        json.dumps(
            {
                "status": "success",
                "provider": "codex",
                "model": "plan-model",
                "reasoning_effort": "must-not-infer-from-plan",
                "record": {
                    "provider": "codex",
                    "requested_model": requested,
                    "response_model": response,
                    "reasoning_effort": effort,
                },
            }
        ),
        encoding="utf-8",
    )
    result = generation_properties(tmp_path, {"context.json": "hash", usage.name: "hash"})
    assert result["generator"] == expected
    assert result["generationModel"] == (requested or "")
    assert result["generationResponseModel"] == (response or "")
    assert result["generationReasoningEffort"] == (effort or "")


def test_mixed_call_models_and_unknown_efforts_are_visible(tmp_path):
    files = {"context.json": "hash"}
    for n, (model, effort) in enumerate([("model-a", "high"), ("model-b", None)]):
        name = f"slide-{n:04d}.usage.json"
        (tmp_path / name).write_text(
            json.dumps(
                {
                    "status": "success",
                    "record": {
                        "provider": "codex",
                        "requested_model": None,
                        "response_model": model,
                        "reasoning_effort": effort,
                    },
                }
            ),
            encoding="utf-8",
        )
        files[name] = "hash"
    result = generation_properties(tmp_path, files)
    assert result["generator"] == "codex (model-a), codex (model-b)"
    assert result["generationResponseModel"] == ["model-a", "model-b"]
    assert result["generationReasoningEffort"] == ["high", ""]


@pytest.mark.parametrize(
    "field", ["generator", "generationResponseModel", "generationReasoningEffort"]
)
def test_generation_fields_match_saved_calls(deck_path, config, tmp_path, fake_context, field):
    import yaml

    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config, context=True)
    text = output.read_text("utf-8")
    fm = split_note(text)[0]
    assert fm["generator"] == "test (actual-fake)"
    assert fm["generationModel"] == "fake"
    assert fm["generationResponseModel"] == "actual-fake"
    assert fm["generationReasoningEffort"] == "high"
    assert fm["generatedBy"] == "tkn-powerpoint-note"
    fm[field] = "tampered"
    output.write_text(
        "---\n" + yaml.safe_dump(fm) + "---\n" + text.split("\n---\n", 1)[1], encoding="utf-8"
    )
    with pytest.raises(pipeline.NoteError, match="model/effort provenance"):
        pipeline.verify(output)


def test_schema_21_upgrade_and_bridge_update_reuse_analysis(
    deck_path, config, tmp_path, fake_context, monkeypatch
):
    original_fingerprint = pipeline.fingerprint
    original_init = fake_context.__init__

    def old_init(self, settings, **kwargs):
        original_init(self, settings, **kwargs)
        self.plan["bridge_version"] = "0.10.0"

    def legacy_key(value):
        if isinstance(value, dict) and "note_schema_version" in value:
            value = {**value, "note_schema_version": "2.1.0"}
        return original_fingerprint(value)

    monkeypatch.setattr(fake_context, "__init__", old_init)
    monkeypatch.setattr(pipeline, "fingerprint", legacy_key)
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config, context=True)
    text = output.read_text("utf-8")
    fm = split_note(text)[0]
    folder = output.parent / fm["evidencePath"]
    fm.update(schemaVersion="2.1.0", generator="tkn-powerpoint-note")
    fm["generatorVersion"] = fm.pop("generatedByVersion")
    fm.pop("generatedBy")
    fm["selectedSlides"] = fm["contextAnalyzedSlides"] = [1, 3]
    fm["contextOmittedSlides"] = [2]
    for key in (
        "generationProvider",
        "generationModel",
        "generationResponseModel",
        "generationReasoningEffort",
    ):
        fm.pop(key)
    # Simulate old journals where Bridge did not return model or Effort.
    manifest_path = folder / "manifest.json"
    manifest = json.loads(manifest_path.read_text("utf-8"))
    for name in manifest["files"]:
        if name.endswith(".usage.json"):
            path = folder / name
            data = json.loads(path.read_text("utf-8"))
            data["record"].pop("reasoning_effort")
            data["record"]["requested_model"] = data["record"]["response_model"] = None
            path.write_text(json.dumps(data), encoding="utf-8")
            manifest["files"][name] = file_digest(path)
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
    import yaml

    output.write_text(
        "---\n" + yaml.safe_dump(fm) + "---\n" + text.split("\n---\n", 1)[1], encoding="utf-8"
    )
    assert pipeline.verify(output)["status"] == "verified"
    before = output.read_bytes()
    calls = len(fake_context.calls)

    def new_init(self, settings, **kwargs):
        original_init(self, settings, **kwargs)
        self.plan.update(bridge_version="0.11.0", reasoning_effort="high")

    monkeypatch.setattr(pipeline, "fingerprint", original_fingerprint)
    monkeypatch.setattr(fake_context, "__init__", new_init)
    monkeypatch.setattr(pipeline, "preflight", lambda: pytest.fail("No rendering required"))
    plan = pipeline.build(deck_path, output, config, context=True, dry_run=True)
    assert plan["planned_ai_calls"] == 0 and output.read_bytes() == before
    result = pipeline.build(deck_path, output, config, context=True)
    assert result["ai_calls"] == 0 and len(fake_context.calls) == calls
    new = split_note(output.read_text("utf-8"))[0]
    assert new["schemaVersion"] == "4.0.0" and new["type"] == "PowerPoint"
    assert new["selectedSlides"] == "1,3"
    assert new["generator"] == "test"
    assert new["generationReasoningEffort"] == ""
    assert new["contextGeneratedAt"] == fm["contextGeneratedAt"]
    assert new["noteId"] == fm["noteId"] and new["created"] == fm["created"]
    assert pipeline.verify(output)["selected_slides"] == [1, 3]
    pipeline.build(deck_path, output, config)
    local = split_note(output.read_text("utf-8"))[0]
    assert local["generator"] == local["generationReasoningEffort"] == ""
