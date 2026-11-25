from __future__ import annotations

import json
import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest
from conftest import R, write_deck

from powerpoint_note import render as render_module
from powerpoint_note.context import ContextError, Generator, generator_label, provider_schema
from powerpoint_note.context_profiles import load_profile
from powerpoint_note.render import RenderError, check_renderable, render

SLIDE = {"summary": "Supported", "sections": [], "uncertainties": []}
DECK = {
    "summary": "Supported",
    "key_points": [{"text": "Point", "slides": [3]}],
    "uncertainties": [],
}


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


def make(config):
    return Generator(
        config["generation"], config["genai_profiles"], load_profile(config["generation"])
    )


def answer(monkeypatch, bridge, data, record=None):
    requests = []

    def generate(self, request):
        requests.append(request)
        return SimpleNamespace(
            data=data, record=SimpleNamespace(model_dump=lambda **kwargs: record or {})
        )

    monkeypatch.setattr(bridge.Runtime, "generate", generate)
    return requests


def test_plan_comes_from_this_applications_connection(config, bridge):
    generator = make(config)
    assert generator.plan["provider"] == "codex"
    assert generator.plan["model"] == "test-model"
    assert generator.plan["profile_name"] == "codex-default"
    assert generator.plan["reasoning_effort"] is None


def test_success_saves_request_response_and_journal(config, bridge, monkeypatch, tmp_path):
    record = {"provider": "codex", "response_model": "observed-model"}
    requests = answer(monkeypatch, bridge, DECK, record)
    generator = make(config)
    key = generator.key("deck", "content")
    data = generator.generate(
        "deck", '{"slides":[]}', None, tmp_path / "deck", key, allowed_slides={3}
    )
    assert data == DECK and not requests[0].images
    assert requests[0].prompt.endswith('SOURCE EVIDENCE (JSON):\n{"slides":[]}')
    usage = json.loads((tmp_path / "deck" / "usage.json").read_text("utf-8"))
    assert usage["status"] == "success" and usage["key"] == key and usage["input_chars"] == 13
    assert usage["prompt_profile"] == "default-ja" and usage["record"] == record
    assert generator_label(usage) == "codex (observed-model)"
    assert json.loads((tmp_path / "deck" / "response.json").read_text("utf-8")) == DECK
    assert "uniqueItems" not in json.dumps(requests[0].output_schema)


@pytest.mark.parametrize(
    ("stage", "data", "message"),
    [
        ("slide", {"summary": "", "sections": [], "uncertainties": []}, "schema at summary"),
        ("slide", {**SLIDE, "extra": 1}, "schema at"),
        ("slide", {**SLIDE, "summary": "<!-- powerpoint-note:end -->"}, "reserved"),
        ("deck", {**DECK, "key_points": [{"text": "P", "slides": [9]}]}, "outside the supplied"),
        ("deck", {**DECK, "key_points": [{"text": "P", "slides": [3, 3]}]}, "schema at key_points"),
    ],
)
def test_unacceptable_output_is_rejected_with_a_reason_and_kept(
    config, bridge, monkeypatch, tmp_path, stage, data, message
):
    answer(monkeypatch, bridge, data)
    generator = make(config)
    with pytest.raises(ContextError, match=message):
        generator.generate(
            stage, "{}", None, tmp_path, generator.key(stage, "c"), allowed_slides={3}
        )
    usage = json.loads((tmp_path / "usage.json").read_text("utf-8"))
    assert usage["status"] == "rejected" and message.split()[0] in usage["error_message"]
    assert json.loads((tmp_path / "rejected.json").read_text("utf-8")) == data
    assert not (tmp_path / "response.json").exists()


def test_provider_failure_is_reported_without_retry(config, bridge, monkeypatch, tmp_path):
    calls = []

    def generate(self, request):
        calls.append(request)
        raise bridge.GenAIError("provider timed out", code="timeout")

    monkeypatch.setattr(bridge.Runtime, "generate", generate)
    generator = make(config)
    with pytest.raises(ContextError, match="timeout"):
        generator.generate("slide", "{}", None, tmp_path, generator.key("slide", "c"))
    usage = json.loads((tmp_path / "usage.json").read_text("utf-8"))
    assert usage["status"] == "failed" and usage["error_code"] == "timeout" and len(calls) == 1


def test_reuse_key_follows_content_prompt_and_model_but_not_timeouts(config, bridge):
    generator = make(config)
    base = generator.key("slide", "content")
    assert generator.key("slide", "content") == base
    assert generator.key("slide", "other") != base and generator.key("deck", "content") != base
    config["generation"]["overrides"] = {"timeout_seconds": 999}
    assert make(config).key("slide", "content") == base
    config["generation"]["overrides"] = {"model": "another-model"}
    assert make(config).key("slide", "content") != base
    config["generation"]["prompt_profile"] = "default-en"
    config["generation"]["overrides"] = {}
    assert make(config).key("slide", "content") != base


def test_provider_schema_drops_only_unique_items():
    schema = {"type": "array", "uniqueItems": True, "items": {"type": "array", "uniqueItems": True}}
    assert provider_schema(schema) == {"type": "array", "items": {"type": "array"}}
    assert schema["uniqueItems"] is True


def test_linked_resources_and_macros_are_not_rendered(tmp_path):
    linked = write_deck(
        tmp_path / "linked.pptx",
        {
            "ppt/slides/_rels/slide2.xml.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="x" Type="{R}/image" Target="file:///C:/x.png" TargetMode="External"/></Relationships>'
        },
    )
    with pytest.raises(RenderError, match="linked resources"):
        check_renderable(linked.read_bytes())
    macro = write_deck(tmp_path / "macro.pptx", {"ppt/vbaProject.bin": "x"})
    with pytest.raises(RenderError, match="Active content"):
        check_renderable(macro.read_bytes())
    hyperlink = write_deck(
        tmp_path / "hyperlink.pptx",
        {
            "ppt/slides/_rels/slide2.xml.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="x" Type="{R}/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>'
        },
    )
    check_renderable(hyperlink.read_bytes())


def test_render_failure_reports_the_cause_and_timeout_names_the_setting(
    deck_path, tmp_path, monkeypatch
):
    monkeypatch.setattr(render_module, "preflight", lambda: None)
    monkeypatch.setattr(
        render_module.subprocess,
        "run",
        lambda *args, **kwargs: SimpleNamespace(
            returncode=1, stderr="Traceback...\ncom_error: PowerPoint could not open the file\n"
        ),
    )
    with pytest.raises(RenderError, match="com_error: PowerPoint could not open the file"):
        render(deck_path.read_bytes(), {1: tmp_path / "slide.png"}, 1200, 5)

    def hang(*args, **kwargs):
        raise subprocess.TimeoutExpired("render", 5)

    monkeypatch.setattr(render_module.subprocess, "run", hang)
    with pytest.raises(RenderError, match="max_render_seconds=5"):
        render(deck_path.read_bytes(), {1: tmp_path / "slide.png"}, 1200, 5)

    def draw(command, **kwargs):
        request = json.loads(Path(command[-1]).read_text("utf-8"))
        assert request["width"] == 1200 and list(request["targets"]) == ["1"]
        (tmp_path / "slide.png").write_bytes(b"\x89PNG\r\n\x1a\n")
        return SimpleNamespace(returncode=0, stderr="")

    monkeypatch.setattr(render_module.subprocess, "run", draw)
    render(deck_path.read_bytes(), {1: tmp_path / "slide.png"}, 1200, 5)
