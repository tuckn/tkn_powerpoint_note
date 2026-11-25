from __future__ import annotations

from pathlib import Path

import pytest
import yaml
from conftest import make_deck

from powerpoint_note.context_profiles import load_profile
from powerpoint_note.evidence import content_key, payload_text, slide_payload
from powerpoint_note.extract import read_deck
from powerpoint_note.frontmatter import NoteLoader, dump, dump_export, source_properties, timestamp
from powerpoint_note.markdown import (
    BEGIN,
    END,
    compose,
    extracted_text,
    item,
    meaningful_alt,
    plain,
    split_note,
)
from powerpoint_note.models import ConfigError, Deck, NoteError, Shape, Slide
from powerpoint_note.sources import discover, note_path, resolve_file

WORDS = load_profile({"prompt_profile": "default-ja", "prompt_profile_dirs": []}).labels
DECK = Deck("sha", 1, 1000, 750, {}, [])


def text_shape(sid, text, top, left=0, **extra):
    paragraphs = [{"text": t, "level": 0, "bullet": False} for t in text.split("|")]
    return Shape(
        str(sid),
        "sp",
        f"Shape {sid}",
        paragraphs=paragraphs,
        bbox=[left, top, left + 10, top + 10],
        location="inside",
        **extra,
    )


def test_source_text_stays_searchable():
    assert plain("R&D uses snake_case and a*b [x]") == "R&D uses snake_case and a*b [x]"
    assert plain("see <Tenant> and <!-- powerpoint-note:end -->") == (
        "see \\<Tenant> and \\<!-- powerpoint-note&#58;end -->"
    )
    assert plain("a < b and 1<2") == "a < b and 1<2"
    assert item("# not a heading\n- not a list", 1) == "  - \\# not a heading\n    \\- not a list"


def test_alternative_text_written_by_tools_is_dropped():
    assert meaningful_alt("Architecture diagram of the hub network") != ""
    assert meaningful_alt("構成図: ハブとスポークの接続") != ""
    assert meaningful_alt("image.png") == ""
    assert meaningful_alt("グラフ\n\n自動的に生成された説明") == ""
    assert meaningful_alt("0 6 域 / こ 0 レ 巴 言 d 可 ・ ・ ・ / 「 E ソ - dd 江") == ""


def test_extracted_text_reads_top_to_bottom_and_skips_what_the_heading_says():
    slide = Slide(1, "256", "p", False, 1, title="Title")
    shapes = [
        text_shape(3, "lower|second line", 500),
        text_shape(2, "upper right", 100, 600),
        text_shape(4, "upper left", 110, 10),
        Shape(
            "1",
            "sp",
            "T",
            role="title",
            paragraphs=[{"text": "Title", "level": 0}],
            bbox=[0, 600, 9, 609],
            location="inside",
        ),
        Shape("5", "cxnSp", "Line", bbox=[0, 0, 5, 5], location="inside"),
        Shape("6", "pic", "Pic", bbox=[0, 700, 5, 705], location="inside", alt_text="Hub diagram"),
        Shape(
            "7",
            "table",
            "Tbl",
            bbox=[0, 300, 5, 305],
            location="partial",
            table=[["a|b", "c"], ["d"]],
        ),
    ]
    text = extracted_text(shapes, slide, DECK, WORDS)
    assert text.splitlines() == [
        "- upper left",
        "",
        "- upper right",
        "",
        "| 列 1 | 列 2 |",
        "| --- | --- |",
        "| a\\|b | c |",
        "| d |  |",
        "",
        "- lower",
        "- second line",
        "",
        "- 画像: Hub diagram",
    ]


def test_chart_and_partial_location_are_readable():
    slide = Slide(1, "256", "p", False, 1)
    chart = Shape(
        "8",
        "chart",
        "C",
        bbox=[0, 0, 5, 5],
        location="inside",
        details={
            "chart": {
                "type": "barChart",
                "title": "Sales",
                "series": [{"name": "2026", "categories": ["Q1", "Q2"], "values": ["10", "12"]}],
            }
        },
    )
    partial = text_shape(9, "clipped", 50)
    partial.location = "partial"
    text = extracted_text([chart, partial], slide, DECK, WORDS)
    assert "- グラフ（barChart）: Sales" in text
    assert "|  | 2026 |" in text and "| Q2 | 12 |" in text
    assert "- clipped （一部がスライド外）" in text


def metadata(**extra):
    return {
        "type": "PowerPoint",
        "schemaVersion": "5.0.0",
        "title": "2026-10-11",
        "description": "",
        "cover": "",
        "subject": "yes",
        "author": "Example User",
        "keywords": ["[[term]]", "1.0"],
        "categories": [],
        "comments": "",
        "sourceId": "demo:a.pptx",
        "sourceRoot": "demo",
        "sourceFileName": "a.pptx",
        "sourceFullPath": "C:\\path\\to\\a.pptx",
        "sourceCreated": "2026-06-21T05:44:56+09:00",
        "sourceModified": "",
        "contextStatus": "not-generated",
        "tags": [],
        "created": "2026-06-21T05:44:56+09:00",
        "updated": "2026-06-21T05:44:56+09:00",
        "noteId": "id",
        **extra,
    }


def test_frontmatter_order_groups_and_quoting():
    text = dump(metadata(custom="keep", contextSourceFingerprint="0123abcd"))
    lines = text.splitlines()
    assert lines[:5] == [
        "type: PowerPoint",
        'schemaVersion: "5.0.0"',
        'title: "2026-10-11"',
        'description: ""',
        'cover: ""',
    ]
    assert lines[-4:] == [
        "tags: []",
        'created: "2026-06-21T05:44:56+09:00"',
        'updated: "2026-06-21T05:44:56+09:00"',
        "noteId: id",
    ]
    assert 'subject: "yes"' in lines and "author: Example User" in lines
    assert "sourceFullPath: 'C:\\path\\to\\a.pptx'" in lines
    assert 'contextSourceFingerprint: "0123abcd"' in lines
    assert '- "[[term]]"' in lines and '- "1.0"' in lines
    assert (
        lines.index("# --- Additional properties ---")
        < lines.index("custom: keep")
        < lines.index("tags: []")
    )
    loaded = yaml.load(text, Loader=NoteLoader)
    assert loaded == metadata(custom="keep", contextSourceFingerprint="0123abcd")
    with pytest.raises(NoteError):
        dump({"type": "PowerPoint"})


def test_export_frontmatter_and_source_properties():
    core = {
        "core.xml": {
            "title": "",
            "creator": "A",
            "keywords": "x; y,z ;",
            "created": "2026-01-01T00:00:00Z",
        }
    }
    props = source_properties(core, "file-stem", "obsidian-link")
    assert props["title"] == "file-stem" and props["keywords"] == ["[[x]]", "[[y,z]]"]
    assert props["sourceCreated"] == "2026-01-01T09:00:00+09:00"
    assert source_properties(core, "s")["keywords"] == ["x", "y,z"]
    text = dump_export(
        {
            "type": "PowerPointExport",
            "schemaVersion": "1.0",
            **props,
            "sourceSnapshotSha256": "12e3",
        }
    )
    assert text.startswith('type: PowerPointExport\nschemaVersion: "1.0"\ntitle: "file-stem"\n')
    assert 'sourceSnapshotSha256: "12e3"' in text
    assert timestamp("2026-06-21", "x") == "2026-06-21"
    with pytest.raises(NoteError):
        timestamp("2026-06-21T05:44:56", "x")


def test_compose_and_split_round_trip_and_reject_foreign_files():
    text = compose(metadata(), "\n", "body line\n", "\n\nmine\n")
    assert text.count(BEGIN) == 1 and "\n# 2026-10-11\n\n" in text
    loaded, before, body, after = split_note(text)
    assert loaded["sourceId"] == "demo:a.pptx" and body == "body line\n" and after == "\n\nmine\n"
    assert compose(loaded, before, body, after) == text
    with pytest.raises(NoteError):
        split_note("no frontmatter")
    with pytest.raises(NoteError):
        split_note(text.replace("type: PowerPoint", "type: Excel"))
    with pytest.raises(NoteError):
        split_note(text.replace(END, ""))


def test_evidence_drops_empty_values_and_keys_ignore_positions(tmp_path, config):
    deck = read_deck(make_deck(tmp_path / "a.pptx", [(256, "One", "body")]), config["extraction"])
    slide = deck.slides[0]
    slide.shapes[1].details["colors"] = ["#FF0000", "accent1"]
    payload = slide_payload(slide, deck, "exclude")
    assert payload == {
        "title": "One",
        "shapes": [
            {"id": "1", "kind": "sp", "role": "title", "box": [10, 13, 210, 80], "text": ["One"]},
            {
                "id": "2",
                "kind": "sp",
                "box": [10, 133, 210, 200],
                "text": ["body"],
                "colors": ["#FF0000", "accent1"],
            },
        ],
    }
    text = payload_text(payload)
    assert " " not in text.replace("One", "")
    before = content_key(slide, text, 2400, deck)
    assert content_key(slide, text, 1200, deck) != before
    slide.number, slide.page_number = 9, 99
    assert payload_text(slide_payload(slide, deck, "exclude")) == text
    assert content_key(slide, text, 2400, deck) == before
    deck.height = 563  # the same slide parts drawn on another slide size
    assert content_key(slide, text, 2400, deck) != before


def test_scatter_series_keep_their_own_points():
    slide = Slide(1, "256", "p", False, 1)
    series = [
        {"name": "A", "x": ["1", "2"], "y": ["10", "20"]},
        {"name": "B", "x": ["5"], "y": ["50"], "size": ["3"]},
    ]
    chart = Shape(
        "8",
        "chart",
        "C",
        bbox=[0, 0, 5, 5],
        location="inside",
        details={"chart": {"type": "scatterChart", "title": "", "series": series}},
    )
    lines = extracted_text([chart], slide, DECK, WORDS).splitlines()
    assert lines[0] == "- グラフ（scatterChart）"
    first, second = lines.index("  - A"), lines.index("  - B")
    assert lines[first + 2 : first + 6] == [
        "| x | y |",
        "| --- | --- |",
        "| 1 | 10 |",
        "| 2 | 20 |",
    ]
    assert lines[second + 2 : second + 5] == [
        "| x | y | size |",
        "| --- | --- | --- |",
        "| 5 | 50 | 3 |",
    ]
    assert "| 5 | 10 |" not in lines


def test_sources_find_presentations_and_map_them_to_notes(tmp_path):
    root = tmp_path / "decks"
    (root / "archive").mkdir(parents=True)
    for name in ("a.pptx", "B.PPTM", "~$a.pptx", "notes.txt", "archive/old.pptx"):
        (root / name).write_text("x")
    spec = {
        "presentations": {
            "dir": str(root),
            "recursive": True,
            "include": ["*.pptx", "*.pptm"],
            "ignore": ["archive/**"],
        },
        "notes": {"dir": str(tmp_path / "notes")},
    }
    assert discover(spec) == ["a.pptx", "B.PPTM"]
    spec["presentations"]["ignore"] = []
    assert discover(spec) == ["a.pptx", "archive/old.pptx", "B.PPTM"]
    spec["presentations"]["recursive"] = False
    known = discover(spec)
    assert known == ["a.pptx", "B.PPTM"]
    assert resolve_file(spec, Path("b.pptm"), known) == "B.PPTM"
    assert resolve_file(spec, root / "a.pptx", known) == "a.pptx"
    assert note_path(spec, "archive/old.pptx") == tmp_path / "notes" / "archive" / "old.pptx.md"
    with pytest.raises(ConfigError):
        resolve_file(spec, Path("archive/old.pptx"), known)
    with pytest.raises(ConfigError):
        resolve_file(spec, tmp_path / "elsewhere.pptx", known)
    spec["presentations"]["dir"] = str(tmp_path / "absent")
    with pytest.raises(ConfigError):
        discover(spec)
