from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

import pytest
import yaml
from conftest import write_deck

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


@pytest.mark.parametrize("review_status", ["reviewed", "accepted"])
def test_reviewed_note_and_foreign_note(deck_path, config, tmp_path, review_status):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    output.write_text(
        output.read_text("utf-8").replace(
            "reviewStatus: unreviewed", f"reviewStatus: {review_status}"
        ),
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

    def __init__(self, settings, *, content_profile=None):
        self.plan = {"provider": "test", "model": "fake", "generation_settings_sha256": "test"}

    def generate(self, stage, evidence, image, folder, name):
        self.calls.append((stage, evidence))
        (folder / (name + ".usage.json")).write_text(
            json.dumps(
                {
                    **self.plan,
                    "status": "success",
                    "record": {
                        "provider": "test",
                        "requested_model": "fake",
                        "response_model": "actual-fake",
                        "reasoning_effort": "high",
                    },
                }
            ),
            encoding="utf-8",
        )
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


def test_context_frontmatter_coverage_profiles_and_clear(deck_path, config, tmp_path, fake_context):
    output = tmp_path / "context.md"
    pipeline.build(deck_path, output, config, context=True)
    metadata = split_note(output.read_text("utf-8"))[0]
    assert metadata["contextStatus"] == "partial"
    assert metadata["contextProfiles"] == ["default-ja"]
    assert metadata["contextAnalyzedSlides"] == "1,3"
    assert metadata["contextOmittedSlides"] == "2"
    evidence = json.loads(
        (output.parent / metadata["evidencePath"] / "evidence.json").read_text("utf-8")
    )
    assert metadata["contextSourceFingerprint"] == evidence["source_fingerprint"]
    assert len(metadata["contextSourceFingerprint"]) == 64
    assert metadata["contextGeneratedAt"].endswith("+09:00")
    config["selection"]["include_hidden"] = True
    config["generation"]["prompt_profile"] = "default-en"
    pipeline.build(deck_path, output, config, context=True)
    metadata = split_note(output.read_text("utf-8"))[0]
    assert metadata["contextStatus"] == "current"
    assert metadata["contextProfiles"] == ["default-en"]
    assert metadata["contextAnalyzedSlides"] == "1-3"
    assert metadata["contextOmittedSlides"] == ""
    pipeline.build(deck_path, output, config)
    metadata = split_note(output.read_text("utf-8"))[0]
    assert metadata["contextStatus"] == "not-generated"
    assert metadata["contextProfiles"] == []
    assert metadata["contextAnalyzedSlides"] == ""
    assert metadata["contextSourceFingerprint"] == metadata["contextGeneratedAt"] == ""


@pytest.mark.parametrize("change", ["version", "template"])
def test_context_reused_when_only_note_presentation_changes(
    deck_path, config, tmp_path, fake_context, monkeypatch, change
):
    from test_profiles import copy_profile

    profile_dir = copy_profile(tmp_path / "profiles")
    config["generation"]["profile_dirs"] = [str(profile_dir.parent)]
    output = tmp_path / "context.md"
    pipeline.build(deck_path, output, config, context=True)
    original = split_note(output.read_text("utf-8"))[0]
    output.write_text(output.read_text("utf-8") + "\nHandwritten.\n", encoding="utf-8")
    calls = len(fake_context.calls)
    if change == "version":
        monkeypatch.setattr(pipeline, "__version__", "99.0.0")
    else:
        path = profile_dir / "note-template.md"
        path.write_text(path.read_text("utf-8") + "\nNew display layout.\n", encoding="utf-8")
    monkeypatch.setattr(pipeline, "preflight", lambda: pytest.fail("No rendering required"))
    before = {p: p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}
    plan = pipeline.build(deck_path, output, config, context=True, dry_run=True)
    assert plan["planned_ai_calls"] == 0 and plan["reused_context"]
    assert before == {p: p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}
    result = pipeline.build(deck_path, output, config, context=True)
    assert result["ai_calls"] == 0 and result["reused_context"]
    assert len(fake_context.calls) == calls
    text = output.read_text("utf-8")
    metadata = split_note(text)[0]
    assert text.endswith("\nHandwritten.\n")
    assert metadata["buildKey"] != original["buildKey"]
    for key in ("contextGeneratedAt", "contextSourceFingerprint", "created", "noteId"):
        assert metadata[key] == original[key]
    if change == "template":
        assert "New display layout." in text
    assert pipeline.verify(output)["status"] == "verified"
    assert pipeline.build(deck_path, output, config, context=True)["status"] == "unchanged"


@pytest.mark.parametrize("legacy_schema", ["1.0.0", "2.0.0"])
def test_legacy_context_migration_avoids_ai(
    deck_path, config, tmp_path, fake_context, monkeypatch, legacy_schema
):
    schema = pipeline.SCHEMA_VERSION
    original_fingerprint = pipeline.fingerprint

    def legacy_fingerprint(value):
        if isinstance(value, dict) and "note_schema_version" in value:
            value = {
                **value,
                "evidence": {
                    k: v for k, v in value["evidence"].items() if k != "source_fingerprint"
                },
            }
            if legacy_schema == "1.0.0":
                value.pop("note_schema_version")
            else:
                value["note_schema_version"] = "2.0.0"
        return original_fingerprint(value)

    monkeypatch.setattr(pipeline, "fingerprint", legacy_fingerprint)
    monkeypatch.setattr(pipeline, "__version__", "0.6.0")
    output = tmp_path / "context.md"
    pipeline.build(deck_path, output, config, context=True)
    text = output.read_text("utf-8")
    metadata = split_note(text)[0]
    folder = output.parent / metadata["evidencePath"]
    manifest_path = folder / "manifest.json"
    manifest = json.loads(manifest_path.read_text("utf-8"))
    manifest.pop("context_build_key")
    manifest.pop("context_generated_at")
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
    metadata["schemaVersion"] = legacy_schema
    metadata["generator"] = "tkn-powerpoint-note"
    if legacy_schema == "1.0.0":
        for key in list(metadata):
            if key.startswith("promptProfile"):
                metadata.pop(key)
        metadata["promptProfile"] = manifest["prompt_profile"]
    for key in list(metadata):
        if key.startswith("context") and key != "context":
            metadata.pop(key)
    metadata.pop("sourceFileName")
    metadata.pop("sourceFullPath")
    metadata["sourceMetadata"] = json.loads((folder / "evidence.json").read_text("utf-8"))[
        "metadata"
    ]
    output.write_text(
        "---\n" + yaml.safe_dump(metadata) + "---\n" + text.split("\n---\n", 1)[1],
        encoding="utf-8",
    )
    calls = len(fake_context.calls)
    monkeypatch.setattr(pipeline, "fingerprint", original_fingerprint)
    monkeypatch.setattr(pipeline, "SCHEMA_VERSION", schema)
    monkeypatch.setattr(pipeline, "__version__", "99.0.0")
    monkeypatch.setattr(pipeline, "preflight", lambda: pytest.fail("Migration must not render"))
    result = pipeline.build(deck_path, output, config, context=True)
    assert result["ai_calls"] == 0 and len(fake_context.calls) == calls
    migrated = split_note(output.read_text("utf-8"))[0]
    assert migrated["schemaVersion"] == schema
    assert migrated["contextProfiles"] == ["default-ja"]
    assert migrated["contextGeneratedAt"] == pipeline.timestamp(manifest["created_at"], "date")
    assert "sourceMetadata" not in migrated
    assert pipeline.verify(output)["status"] == "verified"


def test_schema_3_note_moves_to_proper_noun_type_without_ai(
    deck_path, config, tmp_path, fake_context, monkeypatch
):
    original_fingerprint = pipeline.fingerprint

    def schema_3_key(value):
        if isinstance(value, dict) and "note_schema_version" in value:
            value = {**value, "note_schema_version": "3.0.0"}
        return original_fingerprint(value)

    monkeypatch.setattr(pipeline, "fingerprint", schema_3_key)
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config, context=True)
    text = output.read_text("utf-8")
    old = split_note(text)[0]
    old.update(type="powerpoint", schemaVersion="3.0.0")
    output.write_text(
        "---\n" + yaml.safe_dump(old) + "---\n" + text.split("\n---\n", 1)[1], encoding="utf-8"
    )
    assert pipeline.verify(output)["status"] == "verified"
    calls = len(fake_context.calls)
    monkeypatch.setattr(pipeline, "fingerprint", original_fingerprint)
    monkeypatch.setattr(pipeline, "preflight", lambda: pytest.fail("Migration must not render"))
    result = pipeline.build(deck_path, output, config, context=True)
    assert result["ai_calls"] == 0 and len(fake_context.calls) == calls
    migrated_text = output.read_text("utf-8")
    migrated = split_note(migrated_text)[0]
    assert migrated["type"] == "PowerPoint" and migrated["schemaVersion"] == "4.0.0"
    assert migrated_text.startswith("---\ntype: PowerPoint\n")
    assert migrated["noteId"] == old["noteId"] and migrated["created"] == old["created"]
    assert pipeline.verify(output)["status"] == "verified"


@pytest.mark.parametrize("change", ["properties", "zip-container"])
def test_document_properties_and_zip_changes_reuse_context(
    deck_path, config, tmp_path, fake_context, monkeypatch, change
):
    output = tmp_path / "context.md"
    pipeline.build(deck_path, output, config, context=True)
    before = split_note(output.read_text("utf-8"))[0]
    with ZipFile(deck_path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    if change == "properties":
        parts["docProps/core.xml"] = parts["docProps/core.xml"].replace(
            b"Example Deck", b"New presentation title"
        )
    with ZipFile(
        deck_path, "w", ZIP_STORED if change == "zip-container" else ZIP_DEFLATED
    ) as archive:
        for name, content in reversed(list(parts.items())):
            archive.writestr(name, content)
    assert file_digest(deck_path) != before["sourceSha256"]
    calls = len(fake_context.calls)
    monkeypatch.setattr(pipeline, "preflight", lambda: pytest.fail("Content is unchanged"))
    result = pipeline.build(deck_path, output, config, context=True)
    assert result["ai_calls"] == 0 and result["reused_context"]
    after = split_note(output.read_text("utf-8"))[0]
    assert after["sourceSha256"] != before["sourceSha256"]
    assert after["contextSourceFingerprint"] == before["contextSourceFingerprint"]
    assert after["contextGeneratedAt"] == before["contextGeneratedAt"]
    assert len(fake_context.calls) == calls
    if change == "properties":
        assert after["title"] == "New presentation title"
    assert pipeline.verify(output)["status"] == "verified"


def test_damaged_context_cannot_be_reused(deck_path, config, tmp_path, fake_context, monkeypatch):
    output = tmp_path / "context.md"
    pipeline.build(deck_path, output, config, context=True)
    metadata = split_note(output.read_text("utf-8"))[0]
    context_file = output.parent / metadata["evidencePath"] / "context.json"
    context_file.write_text("{}", encoding="utf-8")
    original_note = output.read_bytes()
    calls = len(fake_context.calls)
    monkeypatch.setattr(pipeline, "__version__", "99.0.0")
    with pytest.raises(NoteError, match="missing or modified"):
        pipeline.build(deck_path, output, config, context=True)
    assert output.read_bytes() == original_note
    assert len(fake_context.calls) == calls


@pytest.mark.parametrize("change", ["source", "selection", "profile", "prompt"])
def test_changed_context_inputs_require_new_ai(deck_path, config, tmp_path, fake_context, change):
    from test_profiles import copy_profile

    profile_dir = copy_profile(tmp_path / "profiles")
    config["generation"]["profile_dirs"] = [str(profile_dir.parent)]
    output = tmp_path / "context.md"
    pipeline.build(deck_path, output, config, context=True)
    calls = len(fake_context.calls)
    if change == "source":
        write_deck(deck_path, {"ppt/media/new.png": "new image bytes"})
    elif change == "selection":
        config["selection"]["slides"] = "3"
    elif change == "profile":
        config["generation"]["prompt_profile"] = "default-en"
    else:
        path = profile_dir / "prompt.md"
        path.write_text(path.read_text("utf-8") + "\nNew instruction.\n", encoding="utf-8")
    result = pipeline.build(deck_path, output, config, context=True)
    assert result["ai_calls"] > 0 and not result["reused_context"]
    assert len(fake_context.calls) > calls


@pytest.mark.parametrize(
    "field", ["contextProfiles", "contextSourceFingerprint", "contextGeneratedAt"]
)
def test_context_provenance_tampering_is_detected(deck_path, config, tmp_path, fake_context, field):
    output = tmp_path / "context.md"
    pipeline.build(deck_path, output, config, context=True)
    text = output.read_text("utf-8")
    metadata = split_note(text)[0]
    metadata[field] = ["changed"] if field == "contextProfiles" else "changed"
    output.write_text(
        "---\n" + yaml.safe_dump(metadata) + "---\n" + text.split("\n---\n", 1)[1],
        encoding="utf-8",
    )
    with pytest.raises(NoteError, match="context provenance"):
        pipeline.verify(output)


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
