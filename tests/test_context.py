from __future__ import annotations

import json
from types import SimpleNamespace

import pytest
from conftest import R, write_deck
from jsonschema import ValidationError

from powerpoint_note.context import Generator
from powerpoint_note.models import NoteError
from powerpoint_note.render import check_renderable


@pytest.fixture
def bridge(monkeypatch, config):
    import tkn_genai_bridge as bridge

    config["genai_profiles"]["codex-default"]["model"] = "test-model"
    actual_plan = bridge.Runtime.plan
    monkeypatch.setattr(
        bridge.Runtime,
        "plan",
        lambda self, request, **kwargs: actual_plan(self, request, check_executable=False),
    )
    return bridge


def test_actual_bridge_request_plan_contract(config, bridge):
    generator = Generator(config["generation"], genai_profiles=config["genai_profiles"])
    assert generator.plan["provider"] == "codex"
    assert generator.plan["model"] == "test-model"
    assert generator.plan["profile_name"] == "codex-default"
    assert generator.plan["generation_settings_sha256"]
    assert generator.plan["reasoning_effort"] is None


def test_success_validation_journal_and_no_image_for_summary(config, bridge, monkeypatch, tmp_path):
    calls = []

    def generate(self, request):
        calls.append(request)
        assert not request.images
        return SimpleNamespace(
            data={
                "summary": "Supported",
                "key_points": [{"text": "Point", "slides": [3]}],
                "uncertainties": [],
            },
            record=SimpleNamespace(
                model_dump=lambda **kwargs: {
                    "usage": {"input_tokens": 10, "output_tokens": 5},
                    "response_model": "observed-model",
                    "reasoning_effort": "high",
                }
            ),
        )

    monkeypatch.setattr(bridge.Runtime, "generate", generate)
    generator = Generator(config["generation"], genai_profiles=config["genai_profiles"])
    result = generator.generate("deck", {"slides": [{"slide_number": 3}]}, None, tmp_path, "deck")
    assert result["summary"] == "Supported"
    assert len(calls) == 1
    assert json.loads((tmp_path / "deck.usage.json").read_text("utf-8"))["status"] == "success"
    record = json.loads((tmp_path / "deck.usage.json").read_text("utf-8"))["record"]
    assert record["response_model"] == "observed-model"
    assert record["reasoning_effort"] == "high"


@pytest.mark.parametrize(
    "result",
    [
        {"slide_number": 99, "summary": "Wrong citation", "sections": [], "uncertainties": []},
        {
            "slide_number": 3,
            "summary": "<!-- powerpoint-note:begin -->",
            "sections": [],
            "uncertainties": [],
        },
        {"slide_number": 3, "summary": "", "sections": [], "uncertainties": []},
    ],
)
def test_invalid_ai_output_never_accepted(config, bridge, monkeypatch, tmp_path, result):
    monkeypatch.setattr(
        bridge.Runtime,
        "generate",
        lambda self, req: SimpleNamespace(
            data=result, record=SimpleNamespace(model_dump=lambda **kwargs: {})
        ),
    )
    generator = Generator(config["generation"], genai_profiles=config["genai_profiles"])
    with pytest.raises((NoteError, ValidationError)):
        generator.generate("slide", {"slide": {"number": 3}}, None, tmp_path, "invalid")
    assert json.loads((tmp_path / "invalid.usage.json").read_text("utf-8"))["status"] == "failed"
    assert not (tmp_path / "invalid.response.json").exists()


def test_summary_cannot_cite_unselected_slides(config, bridge, monkeypatch, tmp_path):
    monkeypatch.setattr(
        bridge.Runtime,
        "generate",
        lambda self, req: SimpleNamespace(
            data={
                "summary": "Claim",
                "key_points": [{"text": "Point", "slides": [99]}],
                "uncertainties": [],
            },
            record=SimpleNamespace(model_dump=lambda **kwargs: {}),
        ),
    )
    with pytest.raises(NoteError):
        Generator(config["generation"], genai_profiles=config["genai_profiles"]).generate(
            "deck", {"slides": [{"slide_number": 3}]}, None, tmp_path, "invalid"
        )


def test_evidence_limit_prevents_provider_call(config, bridge, monkeypatch, tmp_path):
    def forbidden(*args, **kwargs):
        pytest.fail("Provider must not run")

    monkeypatch.setattr(bridge.Runtime, "generate", forbidden)
    config["generation"]["max_input_chars"] = 1
    with pytest.raises(NoteError):
        Generator(config["generation"], genai_profiles=config["genai_profiles"]).generate(
            "deck", {"slides": []}, None, tmp_path, "large"
        )
    assert list(tmp_path.iterdir()) == []


def test_linked_resources_blocked_for_rendering(tmp_path):
    path = write_deck(
        tmp_path / "external.pptx",
        {
            "ppt/slides/_rels/slide2.xml.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="x" Type="{R}/image" Target="https://invalid.example/image.png" TargetMode="External"/></Relationships>'
        },
    )
    with pytest.raises(NoteError):
        check_renderable(path.read_bytes())


def test_provider_schema_preserves_local_uniqueness_contract():
    from powerpoint_note.context import provider_schema
    from powerpoint_note.context_profiles import load_profile

    schema = load_profile({"prompt_profile": "default-ja", "profile_dirs": []}).schemas["deck"]
    projected = provider_schema(schema)
    original_citations = schema["properties"]["key_points"]["items"]["properties"]["slides"]
    projected_citations = projected["properties"]["key_points"]["items"]["properties"]["slides"]
    assert original_citations["uniqueItems"] is True
    assert "uniqueItems" not in projected_citations
    assert projected_citations["minItems"] == 1


def test_duplicate_citations_rejected_locally_after_transport_projection(
    config, bridge, monkeypatch, tmp_path
):
    def generate(self, request):
        assert "uniqueItems" not in json.dumps(request.output_schema)
        return SimpleNamespace(
            data={
                "summary": "Claim",
                "key_points": [{"text": "Point", "slides": [3, 3]}],
                "uncertainties": [],
            },
            record=SimpleNamespace(model_dump=lambda **kwargs: {}),
        )

    monkeypatch.setattr(bridge.Runtime, "generate", generate)
    with pytest.raises(ValidationError):
        Generator(config["generation"], genai_profiles=config["genai_profiles"]).generate(
            "deck", {"slides": [{"slide_number": 3}]}, None, tmp_path, "duplicates"
        )
    assert json.loads((tmp_path / "duplicates.usage.json").read_text("utf-8"))["status"] == "failed"


@pytest.mark.parametrize("name,language", [("default-ja", "Japanese"), ("default-en", "English")])
def test_selected_profile_reaches_provider_and_usage(
    config, bridge, monkeypatch, tmp_path, name, language
):
    config["generation"]["prompt_profile"] = name

    def generate(self, request):
        assert language in request.prompt
        assert "{{language}}" not in request.prompt
        return SimpleNamespace(
            data={"summary": "Supported", "key_points": [], "uncertainties": []},
            record=SimpleNamespace(model_dump=lambda **kwargs: {}),
        )

    monkeypatch.setattr(bridge.Runtime, "generate", generate)
    generator = Generator(config["generation"], genai_profiles=config["genai_profiles"])
    generator.generate("deck", {"slides": [{"slide_number": 3}]}, None, tmp_path, "deck")
    usage = json.loads((tmp_path / "deck.usage.json").read_text("utf-8"))
    assert usage["prompt_profile"] == generator.content_profile.provenance()
    assert usage["prompt_profile"]["name"] == name


def test_generator_overrides_reach_bridge_plan_and_usage(config, bridge, monkeypatch, tmp_path):
    monkeypatch.setattr(
        bridge.Runtime,
        "generate",
        lambda self, request: SimpleNamespace(
            data={"summary": "Supported", "key_points": [], "uncertainties": []},
            record=SimpleNamespace(model_dump=lambda **kwargs: {}),
        ),
    )
    config["genai_profiles"]["named-connection"] = {"provider": "codex", "model": "base"}
    config["generation"].update(
        generator_id="named",
        genai_profile="named-connection",
        overrides={"model": "overridden-model", "timeout_seconds": 450},
    )
    generator = Generator(config["generation"], genai_profiles=config["genai_profiles"])
    assert generator.plan["profile_name"] == "named-connection"
    assert config["genai_profiles"]["named-connection"]["model"] == "base"
    assert generator.plan["model"] == "overridden-model"
    assert generator.plan["timeout_seconds"] == 450
    generator.generate("deck", {"slides": []}, None, tmp_path, "named")
    usage = json.loads((tmp_path / "named.usage.json").read_text("utf-8"))
    assert usage["generation_generator"] == "named"
    assert usage["model"] == "overridden-model"


def test_unknown_genai_profile_never_falls_back(config, bridge):
    config["generation"]["genai_profile"] = "shared-only"
    with pytest.raises(NoteError, match="not defined in genai_profiles: shared-only"):
        Generator(config["generation"], genai_profiles=config["genai_profiles"])
    # Built-in names are defaults of the app config, not an implicit fallback.
    config["generation"]["genai_profile"] = "codex-default"
    with pytest.raises(NoteError, match="codex-default"):
        Generator(config["generation"], genai_profiles={"custom": {"provider": "codex"}})


def test_invalid_connection_combination_fails_before_ai(config, bridge):
    config["generation"]["overrides"] = {"local_only": True}
    with pytest.raises(NoteError, match="genai_profiles.codex-default"):
        Generator(config["generation"], genai_profiles=config["genai_profiles"])


def test_context_key_keeps_the_2x_connection_field_name(config):
    from powerpoint_note import pipeline

    evidence = {"source_fingerprint": "fingerprint", "slides": []}
    profile = SimpleNamespace(name="default-ja", language="Japanese", prompts={}, schemas={})
    plan = {"provider": "codex", "generation_settings_sha256": "hash"}
    legacy = {
        ("bridge_profile" if key == "genai_profile" else key): value
        for key, value in config["generation"].items()
    }
    assert "genai_profile" in config["generation"] and "bridge_profile" in legacy
    assert pipeline._context_key(evidence, config, profile, plan) == pipeline._context_key(
        evidence, {"generation": legacy}, profile, plan
    )
