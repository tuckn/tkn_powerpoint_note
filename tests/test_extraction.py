from __future__ import annotations

from xml.etree.ElementTree import fromstring

import pytest
from conftest import A, P, R, relations, shape, write_deck

from powerpoint_note.extract import IDENTITY, bbox, location, read_deck
from powerpoint_note.io import file_digest
from powerpoint_note.models import NoteError
from powerpoint_note.selection import select


def test_order_metadata_notes_comments_group_and_read_only(deck_path, config):
    before = file_digest(deck_path)
    deck = read_deck(deck_path, config["extraction"])
    assert [s.slide_id for s in deck.slides] == ["300", "100", "200"]
    assert [s.hidden for s in deck.slides] == [False, True, False]
    first = deck.slides[0]
    assert first.title == "A & B"
    assert first.page_number == 7
    assert first.section == "Overview"
    assert first.notes == "Speaker evidence"
    assert first.comments[0]["author"] == "Example Author"
    assert first.comments[0]["text"] == "Review comment"
    shapes = {s.id: s for s in first.shapes}
    assert shapes["3"].location == "outside"
    assert shapes["4"].location == "partial"
    assert shapes["10"].bbox == [120, 120, 160, 160]
    assert file_digest(deck_path) == before
    assert [s.number for s in select(deck, config["selection"])] == [1, 3]


def test_selection_intersects_and_hidden_opt_in(deck_path, config):
    deck = read_deck(deck_path, config["extraction"])
    settings = {**config["selection"], "slides": "1-3", "sections": ["Design"]}
    assert [s.number for s in select(deck, settings)] == [3]
    settings["include_hidden"] = True
    assert [s.number for s in select(deck, settings)] == [2, 3]
    settings["slides"] = "3,1,3"
    settings["sections"] = []
    assert [s.number for s in select(deck, settings)] == [1, 3]


@pytest.mark.parametrize("spec", ["0", "4", "2-1", "1-4", "1,", "one", "1--2", ""])
def test_bad_ranges(deck_path, config, spec):
    deck = read_deck(deck_path, config["extraction"])
    with pytest.raises(NoteError):
        select(deck, {**config["selection"], "slides": spec})


def test_unknown_or_empty_selection(deck_path, config):
    deck = read_deck(deck_path, config["extraction"])
    for values in [{"sections": ["Absent"]}, {"slides": "2"}]:
        with pytest.raises(NoteError):
            select(deck, {**config["selection"], **values})


def test_notes_comments_disabled(deck_path, config):
    settings = {**config["extraction"], "include_notes": False, "include_comments": False}
    first = read_deck(deck_path, settings).slides[0]
    assert not first.notes and not first.comments


def test_rotation_and_unknown_position():
    xfrm = fromstring(
        f'<a:xfrm xmlns:a="{A}" rot="5400000"><a:off x="100" y="100"/><a:ext cx="200" cy="50"/></a:xfrm>'
    )
    assert bbox(xfrm, IDENTITY) == [175, 25, 225, 225]
    assert location(None, 1000, 750) == "unknown"
    assert location([-50, 0, 0, 10], 1000, 750) == "outside"


def test_inherited_placeholder_position(tmp_path, config):
    body = shape(1, "Title", role="title")
    import re

    body = re.sub(r"<a:xfrm>.*?</a:xfrm>", "", body)
    path = write_deck(
        tmp_path / "inherit.pptx",
        {
            "ppt/slides/slide2.xml": f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{body}</p:spTree></p:cSld></p:sld>',
            "ppt/slides/_rels/slide2.xml.rels": relations(
                [("l", "slideLayout", "../slideLayouts/slideLayout1.xml")]
            ),
            "ppt/slideLayouts/slideLayout1.xml": f'<p:sldLayout xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{shape(1, "", 20, 30, "title")}</p:spTree></p:cSld></p:sldLayout>',
        },
    )
    assert read_deck(path, config["extraction"]).slides[2].shapes[0].bbox == [20, 30, 220, 80]


def test_malformed_missing_parts_and_extension(tmp_path, config):
    path = tmp_path / "bad.pptx"
    path.write_text("not zip")
    with pytest.raises(NoteError):
        read_deck(path, config["extraction"])
    with pytest.raises(NoteError):
        read_deck(path.with_suffix(".ppt"), config["extraction"])
    path = write_deck(tmp_path / "missing.pptx", {"ppt/_rels/presentation.xml.rels": relations([])})
    with pytest.raises(NoteError):
        read_deck(path, config["extraction"])


def test_external_relationship_never_fetched(tmp_path, config):
    path = write_deck(
        tmp_path / "external.pptx",
        {
            "ppt/slides/_rels/slide2.xml.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="x" Type="{R}/image" Target="https://invalid.example/image.png" TargetMode="External"/></Relationships>'
        },
    )
    assert len(read_deck(path, config["extraction"]).slides) == 3


def test_table_chart_and_smartart(tmp_path, config):
    table = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    path = write_deck(
        tmp_path / "table.pptx",
        {
            "ppt/slides/slide2.xml": f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{table}</p:spTree></p:cSld></p:sld>'
        },
    )
    assert read_deck(path, config["extraction"]).slides[2].shapes[0].table == [["Cell"]]


def test_application_metadata_does_not_leak_unselected_title_vector(tmp_path, config):
    path = write_deck(
        tmp_path / "metadata.pptx",
        {
            "docProps/app.xml": '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>PowerPoint</Application><TitlesOfParts><vector><item>Hidden slide title</item></vector></TitlesOfParts></Properties>'
        },
    )
    metadata = read_deck(path, config["extraction"]).metadata["app.xml"]
    assert metadata == {"Application": "PowerPoint"}
