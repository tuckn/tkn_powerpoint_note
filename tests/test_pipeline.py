from __future__ import annotations

from pathlib import Path

import pytest

from powerpoint_note import pipeline
from powerpoint_note.io import file_digest
from powerpoint_note.markdown import split_note
from powerpoint_note.models import NoteError


def test_build_verify_repeat_and_read_only(deck_path, config, tmp_path):
    output = tmp_path / "note.md"
    original = file_digest(deck_path)
    result = pipeline.build(deck_path, output, config)
    assert result["status"] == "created"
    text = output.read_text("utf-8")
    assert "A &amp; B" in text
    assert "Speaker evidence" in text and "Review comment" in text
    assert "Private off-canvas" not in text and "Hidden" not in text
    assert pipeline.verify(output)["status"] == "verified"
    files = {str(p): p.stat().st_mtime_ns for p in tmp_path.rglob("*") if p.is_file()}
    assert pipeline.build(deck_path, output, config)["status"] == "unchanged"
    assert files == {str(p): p.stat().st_mtime_ns for p in tmp_path.rglob("*") if p.is_file()}
    assert file_digest(deck_path) == original


def test_dry_run_creates_nothing(deck_path, config, tmp_path):
    output = tmp_path / "not-created" / "note.md"
    before = set(tmp_path.rglob("*"))
    result = pipeline.build(deck_path, output, config, dry_run=True)
    assert result["dry_run"] and result["ai_calls"] == 0
    assert set(tmp_path.rglob("*")) == before


def test_outside_append_and_handwriting_preserved(deck_path, config, tmp_path):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    text = output.read_text("utf-8").replace("schemaVersion:", "custom: keep\nschemaVersion:")
    text += "\nMy handwritten note.\n"
    output.write_text(text, encoding="utf-8")
    config["selection"]["off_slide"] = "append"
    assert pipeline.build(deck_path, output, config)["status"] == "updated"
    updated = output.read_text("utf-8")
    assert "Private off-canvas" in updated
    assert updated.endswith("\nMy handwritten note.\n")
    assert split_note(updated)[0]["custom"] == "keep"


def test_generated_edits_protected_and_force_backed_up(deck_path, config, tmp_path):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    edited = output.read_text("utf-8").replace("Visible body", "Hand edited")
    output.write_text(edited, encoding="utf-8")
    with pytest.raises(NoteError):
        pipeline.build(deck_path, output, config, dry_run=True)
    with pytest.raises(NoteError):
        pipeline.build(deck_path, output, config)
    assert output.read_text("utf-8") == edited
    result = pipeline.build(deck_path, output, config, force=True)
    assert Path(result["backup"]).read_text("utf-8") == edited
    assert "Visible body" in output.read_text("utf-8")


def test_reviewed_note_and_foreign_note(deck_path, config, tmp_path):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    output.write_text(
        output.read_text("utf-8").replace("reviewStatus: unreviewed", "reviewStatus: reviewed"),
        encoding="utf-8",
    )
    assert pipeline.build(deck_path, output, config)["status"] == "unchanged"
    config["selection"]["slides"] = "1"
    with pytest.raises(NoteError):
        pipeline.build(deck_path, output, config)
    other = tmp_path / "foreign.md"
    other.write_text("# Mine")
    with pytest.raises(NoteError):
        pipeline.build(deck_path, other, config, force=True)


def test_source_and_evidence_tamper_detected(deck_path, config, tmp_path):
    output = tmp_path / "note.md"
    result = pipeline.build(deck_path, output, config)
    evidence = Path(result["evidence"]) / "evidence.json"
    evidence.write_text("{}")
    with pytest.raises(NoteError):
        pipeline.verify(output)
    with pytest.raises(NoteError):
        pipeline.build(deck_path, output, config)


class FakeGenerator:
    calls = []

    def __init__(self, settings):
        self.plan = {"provider": "test", "model": "fake", "generation_settings_sha256": "test"}

    def generate(self, stage, evidence, image, folder, name):
        self.calls.append((stage, evidence))
        if stage == "slide":
            return {
                "slide_number": evidence["slide"]["number"],
                "summary": "Explained context",
                "sections": [],
                "uncertainties": [],
            }
        return {
            "summary": "Selected presentation context",
            "key_points": [
                {
                    "text": "A supported claim",
                    "slides": [x["slide_number"] for x in evidence["slides"]],
                }
            ],
            "uncertainties": [],
        }


def fake_render(source, expected_hash, numbers, folder, width, timeout):
    paths = []
    for n in numbers:
        path = folder / f"slide-{n:04d}.png"
        path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"fake")
        paths.append(path)
    return paths


@pytest.fixture
def fake_context(monkeypatch):
    FakeGenerator.calls = []
    monkeypatch.setattr(pipeline, "Generator", FakeGenerator)
    monkeypatch.setattr(pipeline, "preflight", lambda: None)
    monkeypatch.setattr(pipeline, "render", fake_render)
    return FakeGenerator


def test_context_selected_only_summary_cache_and_refresh(deck_path, config, tmp_path, fake_context):
    config["selection"]["slides"] = "3"
    output = tmp_path / "context.md"
    result = pipeline.build(deck_path, output, config, context=True)
    assert result["ai_calls"] == 2 and len(fake_context.calls) == 2
    assert fake_context.calls[0][1]["slide"]["number"] == 3
    assert "Selected presentation context" in output.read_text("utf-8")
    assert "Speaker evidence" not in output.read_text("utf-8")
    assert pipeline.build(deck_path, output, config, context=True)["status"] == "unchanged"
    assert len(fake_context.calls) == 2
    pipeline.build(deck_path, output, config, context=True, refresh=True)
    assert len(fake_context.calls) == 4
    assert pipeline.build(deck_path, output, config, context=True)["status"] == "unchanged"


def test_context_dry_run_and_limit_never_generate(deck_path, config, tmp_path, fake_context):
    output = tmp_path / "no" / "note.md"
    assert (
        pipeline.build(deck_path, output, config, context=True, dry_run=True)["planned_ai_calls"]
        == 3
    )
    assert not output.parent.exists() and fake_context.calls == []
    config["generation"]["max_slides"] = 1
    with pytest.raises(NoteError):
        pipeline.build(deck_path, output, config, context=True)
    assert fake_context.calls == []


def test_generation_failure_keeps_previous_note(
    deck_path, config, tmp_path, fake_context, monkeypatch
):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    old = output.read_bytes()

    def fail(*args, **kwargs):
        raise NoteError("Provider failed")

    monkeypatch.setattr(FakeGenerator, "generate", fail)
    with pytest.raises(NoteError):
        pipeline.build(deck_path, output, config, context=True)
    assert output.read_bytes() == old
    assert list(output.with_name(output.name + ".assets").glob("failed-*/evidence.json"))
    assert not output.with_name(output.name + ".lock").exists()


def test_change_during_generation_preserves_note(
    deck_path, config, tmp_path, fake_context, monkeypatch
):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    old = output.read_bytes()
    original = fake_render

    def mutate(*args, **kwargs):
        paths = original(*args, **kwargs)
        with deck_path.open("ab") as stream:
            stream.write(b"changed")
        return paths

    monkeypatch.setattr(pipeline, "render", mutate)
    with pytest.raises(NoteError):
        pipeline.build(deck_path, output, config, context=True)
    assert output.read_bytes() == old
