from __future__ import annotations

import json
from pathlib import Path

import pytest
from conftest import make_deck, write_deck

from powerpoint_note import pipeline
from powerpoint_note.io import file_digest
from powerpoint_note.markdown import BEGIN, read_frontmatter, split_note
from powerpoint_note.pipeline import Options, export, pull, summarize
from powerpoint_note.store import assets_dir, load_state

SLIDES = [(256, "One", "first body"), (257, "Two", "second body"), (258, "Three", "third body")]
BIG = {"generation": {"limits": {"max_slides": 100}}}


def run(source, **options):
    overrides = options.pop("overrides", None)
    file = options.pop("file", None)
    return pull("demo", source["load"](overrides), Options(**options), file)


def one(source, **options):
    rows = run(source, **options)
    assert len(rows) == 1, rows
    return rows[0]


def note_of(source, name="deck.pptx"):
    return source["notes"] / (name + ".md")


def tree(root: Path) -> dict[str, int]:
    return {str(p): p.stat().st_mtime_ns for p in root.rglob("*") if p.is_file()}


def test_pull_creates_then_leaves_an_unchanged_note_alone(source):
    deck = write_deck(source["decks"] / "deck.pptx")
    original = file_digest(deck)
    row = one(source)
    assert row["status"] == "created" and row["contextStatus"] == "not-generated"
    assert row["sourcePath"] == "deck.pptx" and row["sourceRoot"] == "demo"
    text = note_of(source).read_text("utf-8")
    metadata, _, body, _ = split_note(text)
    assert metadata["type"] == "PowerPoint" and metadata["schemaVersion"] == "5.0.0"
    assert metadata["sourceId"] == "demo:deck.pptx" and metadata["title"] == "Example Deck"
    assert "## Presentation Map" in body and "| 2 | 100 | Hidden |" in body
    assert "Visible body" in body and "Speaker evidence" in body and "Review comment" in body
    assert "Private off-canvas" not in body and "### スライド 2" not in body
    before = tree(source["notes"])
    assert one(source)["status"] == "unchanged"
    assert tree(source["notes"]) == before
    assert file_digest(deck) == original


def test_dry_run_and_status_write_nothing(source):
    write_deck(source["decks"] / "deck.pptx")
    assert one(source, dry_run=True)["status"] == "would-create"
    assert not source["notes"].exists()
    one(source)
    before = tree(source["notes"])
    assert one(source, dry_run=True)["status"] == "unchanged"
    assert tree(source["notes"]) == before


def test_text_outside_markers_and_unknown_properties_survive(source):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    one(source)
    note = note_of(source)
    text = note.read_text("utf-8").replace("tags: []", "custom: keep\ntags:\n- mine")
    note.write_text(text + "\nMy handwritten note.\n", encoding="utf-8")
    make_deck(source["decks"] / "deck.pptx", [*SLIDES, (259, "Four", "new body")])
    assert one(source)["status"] == "updated"
    updated = note.read_text("utf-8")
    metadata = split_note(updated)[0]
    assert updated.endswith("\nMy handwritten note.\n")
    assert metadata["custom"] == "keep" and metadata["tags"] == ["mine"]
    assert "new body" in updated


def test_hand_edited_generated_part_is_protected_until_forced(source, tmp_path, monkeypatch):
    monkeypatch.setattr("powerpoint_note.store.state_root", lambda: tmp_path / "state")
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    one(source)
    note = note_of(source)
    edited = note.read_text("utf-8").replace("first body", "Hand edited")
    note.write_text(edited, encoding="utf-8")
    row = one(source)
    assert row["status"] == "protected" and note.read_text("utf-8") == edited
    assert summarize("pull", [row], Options())["errors"] == 1
    forced = one(source, force=True)
    assert forced["status"] == "updated"
    assert Path(forced["backupPath"]).read_text("utf-8") == edited
    assert "first body" in note.read_text("utf-8")


def test_a_note_without_state_is_adopted_when_it_matches(source):
    """Notes copied to another PC lose `img/`; an untouched note must not look edited."""
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    one(source)
    (assets_dir(note_of(source)) / "state.json").unlink()
    assert one(source)["status"] == "unchanged"
    assert load_state(assets_dir(note_of(source)))["generatedSha256"]


def test_context_is_generated_once_per_slide_and_reused(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    planned = one(source, context=True, dry_run=True)
    assert planned["status"] == "would-create" and planned["plannedAiCalls"] == 4
    assert not source["notes"].exists() and not ai.slide_calls
    row = one(source, context=True)
    assert (row["aiCalls"], row["generatedSlides"], row["contextStatus"]) == (4, 3, "current")
    assert sorted(ai.rendered) == [1, 2, 3]
    text = note_of(source).read_text("utf-8")
    assert text.count("Explained") == 3 and "Deck summary" in text
    assert "Point（スライド 1）" in text and "img/deck.pptx/slides/" in text
    metadata = split_note(text)[0]
    assert metadata["contextAnalyzedSlides"] == "1-3" and metadata["contextProfiles"] == [
        "default-ja"
    ]
    again = one(source, context=True)
    assert (again["status"], again["aiCalls"], again["reusedSlides"]) == ("unchanged", 0, 3)
    assert len(ai.slide_calls) == 3 and ai.deck_calls == 1


def test_evidence_sent_to_ai_is_compact_and_has_no_position(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES, numbered=True)
    one(source, context=True)
    payload = ai.slide_calls[0].split("SOURCE EVIDENCE (JSON):\n", 1)[1]
    assert payload == json.dumps(
        json.loads(payload), ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    data = json.loads(payload)
    assert "number" not in data and "page_number" not in data
    assert data["shapes"][0] == {
        "box": [10, 13, 210, 80],
        "id": "1",
        "kind": "sp",
        "role": "title",
        "text": ["One"],
    }


def test_inserting_a_slide_only_explains_the_new_slide(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES, numbered=True)
    one(source, context=True)
    make_deck(
        source["decks"] / "deck.pptx",
        [SLIDES[0], (300, "New", "inserted"), *SLIDES[1:]],
        numbered=True,
    )
    row = one(source, context=True)
    assert (row["generatedSlides"], row["reusedSlides"], row["aiCalls"]) == (1, 3, 2)
    assert ai.rendered[3:] == [2]
    assert split_note(note_of(source).read_text("utf-8"))[0]["contextAnalyzedSlides"] == "1-4"


def test_plain_pull_keeps_explanations_and_marks_changed_slides_stale(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    one(source, context=True)
    make_deck(source["decks"] / "deck.pptx", [SLIDES[0], (257, "Two", "rewritten"), SLIDES[2]])
    row = one(source)
    assert (row["status"], row["contextStatus"], row["staleSlides"]) == ("updated", "stale", "2")
    text = note_of(source).read_text("utf-8")
    assert text.count("Explained") == 3 and "rewritten" in text and "[!WARNING]" in text
    assert split_note(text)[0]["contextStaleSlides"] == "2"
    refreshed = one(source, context=True)
    assert (refreshed["generatedSlides"], refreshed["contextStatus"]) == (1, "current")
    assert "[!WARNING]" not in note_of(source).read_text("utf-8")


def test_selected_slides_accumulate_explanations(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    first = one(source, context=True, slides="1")
    assert (first["generatedSlides"], first["contextStatus"]) == (1, "partial")
    second = one(source, context=True, slides="3")
    assert (second["generatedSlides"], second["contextStatus"]) == (1, "partial")
    metadata = split_note(note_of(source).read_text("utf-8"))[0]
    assert metadata["contextAnalyzedSlides"] == "1,3" and metadata["contextOmittedSlides"] == "2"


def test_a_failed_run_keeps_the_note_and_resumes_from_finished_slides(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    one(source)
    before = note_of(source).read_text("utf-8")
    ai.fail_on = "third body"
    with pytest.raises(RuntimeError):
        run(source, context=True)
    assert note_of(source).read_text("utf-8") == before and len(ai.slide_calls) == 2
    assert not list(source["notes"].glob(".*.lock"))
    ai.fail_on = None
    row = one(source, context=True)
    assert (row["generatedSlides"], row["reusedSlides"], row["aiCalls"]) == (1, 2, 2)


def test_rejected_ai_output_is_reported_and_saved(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES[:1])
    ai.invalid = True
    row = one(source, context=True)
    assert row["status"] == "error" and "does not match the schema" in row["message"]
    rejected = list(source["notes"].rglob("rejected.json"))
    assert len(rejected) == 1 and json.loads(rejected[0].read_text("utf-8")) == {"summary": ""}
    usage = json.loads((rejected[0].parent / "usage.json").read_text("utf-8"))
    assert usage["status"] == "rejected" and usage["error_message"]
    assert not note_of(source).exists()


def test_a_slide_over_the_input_limit_is_skipped_and_the_rest_continue(source, ai):
    slides = [SLIDES[0], (257, "Two", "x" * 3000), SLIDES[2]]
    make_deck(source["decks"] / "deck.pptx", slides)
    row = one(source, context=True, overrides={"generation": {"limits": {"max_input_chars": 2000}}})
    assert (row["status"], row["generatedSlides"], row["contextStatus"]) == (
        "created",
        2,
        "partial",
    )
    assert row["skippedSlides"][0].startswith("2: evidence needs")
    text = note_of(source).read_text("utf-8")
    assert "このスライドの説明は生成していません" in text and text.count("Explained") == 2


def test_more_new_slides_than_max_slides_stops_before_any_call(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    row = one(source, context=True, overrides={"generation": {"limits": {"max_slides": 2}}})
    assert row["status"] == "error" and "max_slides=2" in row["message"]
    assert not ai.slide_calls and not ai.rendered
    assert (
        one(
            source,
            context=True,
            slides="1-2",
            overrides={"generation": {"limits": {"max_slides": 2}}},
        )["generatedSlides"]
        == 2
    )
    # Slides already explained do not count against the limit.
    assert (
        one(source, context=True, overrides={"generation": {"limits": {"max_slides": 2}}})[
            "generatedSlides"
        ]
        == 1
    )


def test_unreferenced_results_are_removed_after_a_successful_run(source, ai):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    one(source, context=True)
    folder = assets_dir(note_of(source)) / "slides"
    assert len(list(folder.iterdir())) == 3
    make_deck(source["decks"] / "deck.pptx", SLIDES[:2])
    one(source, context=True)
    assert len(list(folder.iterdir())) == 2
    assert len(list((folder.parent / "deck").iterdir())) == 1


def test_one_file_many_files_and_missing_presentations(source):
    make_deck(source["decks"] / "a.pptx", SLIDES)
    (source["decks"] / "sub").mkdir()
    make_deck(source["decks"] / "sub" / "b.pptx", SLIDES)
    (source["decks"] / "~$a.pptx").write_text("lock")
    (source["decks"] / "broken.pptx").write_text("not a zip")
    rows = run(source)
    assert {r["sourcePath"]: r["status"] for r in rows} == {
        "a.pptx": "created",
        "broken.pptx": "error",
        "sub/b.pptx": "created",
    }
    assert (source["notes"] / "sub" / "b.pptx.md").is_file()
    summary = summarize("pull", rows, Options())
    assert (summary["status"], summary["errors"], summary["changed"]) == ("error", 1, 2)
    (source["decks"] / "a.pptx").unlink()
    (source["decks"] / "broken.pptx").unlink()
    rows = run(source)
    assert {r["sourcePath"]: r["status"] for r in rows} == {
        "sub/b.pptx": "unchanged",
        "a.pptx": "missing-source",
    }
    assert (source["notes"] / "a.pptx.md").is_file()
    assert [r["status"] for r in run(source, file=Path("sub/b.pptx"))] == ["unchanged"]


def test_export_writes_beside_the_presentation_and_never_overwrites_silently(source, ai, tmp_path):
    deck = make_deck(tmp_path / "talk.pptx", SLIDES)
    config = source["load"]()
    result = export(deck, None, config, Options())
    output = tmp_path / "talk.pptx.md"
    assert (result["status"], result["exported"]) == ("success", 1)
    assert result["results"][0]["outputPath"] == str(output)
    metadata, rest = read_frontmatter(output.read_text("utf-8"))
    assert metadata["type"] == "PowerPointExport" and metadata["generationMethod"] == "text"
    assert metadata["sourceSnapshotSha256"] == file_digest(deck)
    assert BEGIN not in rest and rest.startswith("\n# Example Deck\n")
    assert not (tmp_path / "img").exists()
    again = export(deck, None, config, Options())
    assert again["status"] == "error" and again["results"][0]["status"] == "export-error"
    assert export(deck, None, config, Options(dry_run=True, force=True))["planned"] == 1
    forced = export(deck, None, config, Options(force=True, context=True))
    assert forced["aiCalls"] == 4 and (tmp_path / "img" / "talk.pptx" / "state.json").is_file()
    assert export(deck, None, config, Options(force=True, context=True))["aiCalls"] == 0


def test_export_a_folder_and_an_explicit_output(source, tmp_path):
    folder = tmp_path / "many"
    (folder / "sub").mkdir(parents=True)
    make_deck(folder / "a.pptx", SLIDES)
    make_deck(folder / "sub" / "b.pptx", SLIDES)
    config = source["load"]()
    assert export(folder, None, config, Options())["exported"] == 2
    assert (folder / "sub" / "b.pptx.md").is_file()
    with pytest.raises(pipeline.ConfigError):
        export(folder, tmp_path / "x.md", config, Options())
    target = tmp_path / "out" / "named.md"
    assert export(folder / "a.pptx", target, config, Options())["exported"] == 1
    assert target.is_file()
    with pytest.raises(pipeline.ConfigError):
        export(tmp_path / "absent.pptx", None, config, Options())
