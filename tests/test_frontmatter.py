from __future__ import annotations

import json
import re
from uuid import UUID

import pytest
import yaml
from conftest import write_deck

from powerpoint_note import pipeline
from powerpoint_note.frontmatter import NoteLoader, dump, normalize, timestamp
from powerpoint_note.markdown import split_note
from powerpoint_note.models import NoteError


def test_export_order_flat_values_dates_and_heading(deck_path, config, tmp_path):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    text = output.read_text("utf-8")
    metadata = split_note(text)[0]
    keys = list(metadata)
    assert keys[:5] == ["type", "schemaVersion", "title", "description", "cover"]
    assert keys[-4:] == ["tags", "created", "updated", "noteId"]
    assert metadata["type"] == "powerpoint" and metadata["schemaVersion"] == "2.0.0"
    assert metadata["title"] == "Example Deck"
    assert metadata["tags"] == []
    assert UUID(metadata["noteId"])
    for value in metadata.values():
        assert not isinstance(value, dict)
        if isinstance(value, list):
            assert all(not isinstance(item, (dict, list)) for item in value)
    assert "\n---\n\n# Example Deck\n" in text
    assert 'schemaVersion: "2.0.0"' in text
    for field in ("created", "updated"):
        assert re.search(
            rf'{field}: "\d{{4}}-\d{{2}}-\d{{2}}T\d{{2}}:\d{{2}}:\d{{2}}\+09:00"', text
        )
    config["selection"]["slides"] = "1"
    pipeline.build(deck_path, output, config)
    updated = split_note(output.read_text("utf-8"))[0]
    assert updated["created"] == metadata["created"]
    assert updated["noteId"] == metadata["noteId"]


@pytest.mark.parametrize(
    "date", ["2026-06-21", "2026-06-20T20:44:56Z", "2026-06-21T05:44:56+09:00"]
)
def test_legacy_export_migrates_date_and_provenance(deck_path, config, tmp_path, date):
    output = tmp_path / "note.md"
    result = pipeline.build(deck_path, output, config)
    text = output.read_text("utf-8")
    metadata, _, _, _ = split_note(text)
    bundle = output.parent / metadata["evidencePath"]
    manifest = json.loads((bundle / "manifest.json").read_text("utf-8"))
    evidence = json.loads((bundle / "evidence.json").read_text("utf-8"))
    legacy = {
        key: value
        for key, value in metadata.items()
        if not key.startswith("promptProfile") and key not in {"created", "type"}
    }
    legacy.update(
        schemaVersion="1.0.0",
        date=date,
        promptProfile=manifest["prompt_profile"],
        sourceMetadata=evidence["metadata"],
        custom="kept",
        cover="[[Cover]]",
        tags=["test"],
    )
    rest = text.split("\n---\n", 1)[1].lstrip("\n") + "\nHandwritten.\n"
    output.write_text(
        "---\n" + yaml.safe_dump(legacy, sort_keys=False) + "---\n" + rest, encoding="utf-8"
    )
    assert pipeline.verify(output)["status"] == "verified"
    # A changed selection triggers a real export of this older managed format.
    config["selection"]["slides"] = "1"
    original = output.read_bytes()
    assert pipeline.build(deck_path, output, config, dry_run=True)["status"] == "updated"
    assert output.read_bytes() == original
    pipeline.build(deck_path, output, config)
    migrated = output.read_text("utf-8")
    new = split_note(migrated)[0]
    assert "date" not in new and "sourceMetadata" not in new
    assert new["created"] == timestamp(date, "date")
    assert new["noteId"] == metadata["noteId"]
    assert new["custom"] == "kept" and new["cover"] == "[[Cover]]" and new["tags"] == ["test"]
    assert migrated.endswith("\nHandwritten.\n")
    assert "\n---\n\n# Example Deck\n" in migrated
    assert pipeline.verify(output)["status"] == "verified"
    assert result["ai_calls"] == 0


@pytest.mark.parametrize(
    "extra, message",
    [
        ({"date": "2000-01-01"}, "date and created disagree"),
        ({"custom": {"nested": "value"}}, "scalar or flat list"),
        ({"custom": [{"nested": "value"}]}, "flat list"),
    ],
)
def test_invalid_existing_metadata_fails_before_side_effects(
    deck_path, config, tmp_path, monkeypatch, extra, message
):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    text = output.read_text("utf-8")
    metadata = split_note(text)[0]
    metadata.update(extra)
    output.write_text(
        "---\n" + yaml.safe_dump(metadata) + "---\n" + text.split("\n---\n", 1)[1], encoding="utf-8"
    )
    before = {p: p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}
    monkeypatch.setattr(pipeline, "preflight", lambda: pytest.fail("Office preflight must not run"))
    for dry_run in (True, False):
        with pytest.raises(NoteError, match=message):
            pipeline.build(deck_path, output, config, context=True, dry_run=dry_run)
    assert before == {p: p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}


def test_unknown_creation_date_is_not_invented():
    assert normalize({"updated": "2026-06-21T05:44:56+09:00"})["created"] == ""


def test_source_metadata_projection_keeps_full_evidence(config, tmp_path):
    source = write_deck(
        tmp_path / "metadata.pptx",
        {
            "docProps/core.xml": """<cp:coreProperties
          xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
          xmlns:dc="http://purl.org/dc/elements/1.1/"
          xmlns:dcterms="http://purl.org/dc/terms/">
          <dc:subject>Subject</dc:subject><dc:creator>Example Author</dc:creator>
          <dc:description>Source comments</dc:description>
          <cp:keywords>one, two; three</cp:keywords><cp:category>Category</cp:category>
          <dcterms:created>2026-06-20T20:44:56Z</dcterms:created>
          <dcterms:modified>2026-06-21T00:00:00Z</dcterms:modified>
          <cp:revision>42</cp:revision></cp:coreProperties>""",
        },
    )
    output = tmp_path / "note.md"
    result = pipeline.build(source, output, config)
    metadata = split_note(output.read_text("utf-8"))[0]
    assert metadata["title"] == "metadata"
    assert metadata["subject"] == "Subject" and metadata["author"] == "Example Author"
    assert metadata["comments"] == "Source comments"
    assert metadata["keywords"] == ["one, two", "three"]
    assert metadata["categories"] == ["Category"]
    assert metadata["sourceCreated"] == "2026-06-21T05:44:56+09:00"
    assert metadata["sourceModified"] == "2026-06-21T09:00:00+09:00"
    evidence = json.loads(
        (output.parent / metadata["evidencePath"] / "evidence.json").read_text("utf-8")
    )
    assert evidence["metadata"]["core.xml"]["revision"] == "42"
    assert result["verified"] == "verified"


def test_timestamp_normalization_and_yaml_round_trip():
    values = {
        "sourceCreated": "2026-06-20T20:44:56.123Z",
        "created": "2026-06-21",
        "sourceFile": r"C:\path\to\deck.pptx",
        "cover": "[[Cover]]",
        "customDates": ["2026-01-01"],
        "customCode": "00123",
    }
    rendered = dump(values)
    assert 'sourceCreated: "2026-06-21T05:44:56+09:00"' in rendered
    assert 'created: "2026-06-21"' in rendered
    assert "sourceFile: 'C:\\path\\to\\deck.pptx'" in rendered
    assert 'cover: "[[Cover]]"' in rendered and '- "2026-01-01"' in rendered
    assert 'customCode: "00123"' in rendered
    assert yaml.safe_load(rendered) == normalize(values)
    assert yaml.load("created: 2026-06-21", Loader=NoteLoader)["created"] == "2026-06-21"
    with pytest.raises(NoteError, match="unique string"):
        yaml.load("created: a\ncreated: b", Loader=NoteLoader)


@pytest.mark.parametrize("value", ["2026-06-21T05:44:56", "not-a-date", "2026-02-30"])
def test_invalid_or_ambiguous_timestamps_are_rejected(value):
    with pytest.raises(NoteError):
        timestamp(value, "created")


def test_conflicting_subseconds_are_not_silently_discarded():
    with pytest.raises(NoteError, match="disagree"):
        normalize({"date": "2026-06-21T05:44:56.1+09:00", "created": "2026-06-21T05:44:56.2+09:00"})


@pytest.mark.parametrize(
    "field",
    [
        "promptProfile",
        "promptProfileVersion",
        "promptProfileLanguage",
        "promptProfileSha256",
        "promptProfileResources",
    ],
)
def test_flat_provenance_tampering_is_detected(deck_path, config, tmp_path, field):
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    text = output.read_text("utf-8")
    metadata = split_note(text)[0]
    metadata[field] = ["tampered"] if field == "promptProfileResources" else "tampered"
    output.write_text(
        "---\n" + yaml.safe_dump(metadata) + "---\n" + text.split("\n---\n", 1)[1], encoding="utf-8"
    )
    with pytest.raises(NoteError, match="prompt profile provenance"):
        pipeline.verify(output)
