from __future__ import annotations

import re
from xml.etree.ElementTree import fromstring

import pytest
from conftest import A, P, R, make_deck, relations, shape, write_deck

from powerpoint_note.extract import IDENTITY, SLIDE_NUMBER, bbox, location, read_deck
from powerpoint_note.io import file_digest
from powerpoint_note.models import ConfigError, NoteError
from powerpoint_note.selection import format_slides, select


def visible(deck):
    return [s for s in deck.slides if not s.hidden]


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


def test_selection_intersects_ranges_and_sections(deck_path, config):
    deck = read_deck(deck_path, config["extraction"])
    assert [s.number for s in select(deck, visible(deck), None, None)] == [1, 3]
    assert [s.number for s in select(deck, visible(deck), "1-3", ["Design"])] == [3]
    assert [s.number for s in select(deck, deck.slides, "1-3", ["Design"])] == [2, 3]
    assert [s.number for s in select(deck, visible(deck), "3,1,3", None)] == [1, 3]


@pytest.mark.parametrize("spec", ["0", "4", "2-1", "1-4", "1,", "one", "1--2", ""])
def test_bad_ranges(deck_path, config, spec):
    deck = read_deck(deck_path, config["extraction"])
    with pytest.raises(ConfigError):
        select(deck, visible(deck), spec, None)


def test_unknown_section_or_only_hidden_slides(deck_path, config):
    deck = read_deck(deck_path, config["extraction"])
    with pytest.raises(ConfigError):
        select(deck, visible(deck), None, ["Absent"])
    with pytest.raises(ConfigError):
        select(deck, visible(deck), "2", None)


def test_slide_ranges_are_compact():
    assert format_slides([3, 4, 5, 6]) == "3-6"
    assert format_slides([7, 1, 2, 400, 2]) == "1-2,7,400"
    assert format_slides([]) == ""
    with pytest.raises(NoteError):
        format_slides([0])


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


def placeholder_shape(attributes: str, text: str = "", xfrm: str = "") -> str:
    position = f"<p:spPr>{xfrm}</p:spPr>" if xfrm else "<p:spPr/>"
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph {attributes}/>'
        f"</p:nvPr></p:nvSpPr>{position}<p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p>"
        "</p:txBody></p:sp>"
    )


def test_content_placeholder_inherits_position_from_master_body(tmp_path, config):
    """The standard Title and Content layout leaves the position to the master's body."""
    xfrm = '<a:xfrm><a:off x="40" y="150"/><a:ext cx="900" cy="500"/></a:xfrm>'
    small = '<a:xfrm><a:off x="1" y="1"/><a:ext cx="1" cy="1"/></a:xfrm>'
    on_slide = placeholder_shape('idx="1"', "Body text")
    on_layout = placeholder_shape('idx="1"')
    on_master = placeholder_shape('type="title"', xfrm=small) + placeholder_shape(
        'type="body" idx="1"', xfrm=xfrm
    )
    spaces = f'xmlns:p="{P}" xmlns:a="{A}"'
    path = write_deck(
        tmp_path / "inherit.pptx",
        {
            "ppt/slides/slide2.xml": f"<p:sld {spaces}><p:cSld><p:spTree>{on_slide}</p:spTree></p:cSld></p:sld>",
            "ppt/slides/_rels/slide2.xml.rels": relations(
                [("l", "slideLayout", "../slideLayouts/slideLayout1.xml")]
            ),
            "ppt/slideLayouts/slideLayout1.xml": f"<p:sldLayout {spaces}><p:cSld><p:spTree>{on_layout}</p:spTree></p:cSld></p:sldLayout>",
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels": relations(
                [("m", "slideMaster", "../slideMasters/slideMaster1.xml")]
            ),
            "ppt/slideMasters/slideMaster1.xml": f"<p:sldMaster {spaces}><p:cSld><p:spTree>{on_master}</p:spTree></p:cSld></p:sldMaster>",
        },
    )
    slide = read_deck(path, config["extraction"]).slides[2]
    assert slide.shapes[0].bbox == [40, 150, 940, 650]
    assert slide.shapes[0].location == "inside"
    assert not any("positions are unavailable" in w for w in slide.warnings)


def test_layout_position_wins_over_master(tmp_path, config):
    body = re.sub(r"<a:xfrm>.*?</a:xfrm>", "", shape(1, "Title", role="title"))
    path = write_deck(
        tmp_path / "layout.pptx",
        {
            "ppt/slides/slide2.xml": f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{body}</p:spTree></p:cSld></p:sld>',
            "ppt/slides/_rels/slide2.xml.rels": relations(
                [("l", "slideLayout", "../slideLayouts/slideLayout1.xml")]
            ),
            "ppt/slideLayouts/slideLayout1.xml": f'<p:sldLayout xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{shape(1, "", 20, 30, "title")}</p:spTree></p:cSld></p:sldLayout>',
        },
    )
    assert read_deck(path, config["extraction"]).slides[2].shapes[0].bbox == [20, 30, 220, 80]


def test_fingerprint_follows_the_slide_not_its_position(tmp_path, config):
    slides = [(256, "One", "first"), (257, "Two", "second")]
    before = read_deck(make_deck(tmp_path / "a.pptx", slides, numbered=True), config["extraction"])
    moved = [(300, "New", "inserted"), *slides]
    after = read_deck(make_deck(tmp_path / "b.pptx", moved, numbered=True), config["extraction"])
    prints = {s.slide_id: s.fingerprint for s in before.slides}
    assert {s.slide_id: s.fingerprint for s in after.slides if s.slide_id != "300"} == prints
    # The page number PowerPoint stores in the slide is not part of the text either.
    assert after.slides[1].shapes[-1].text == SLIDE_NUMBER
    edited = [(256, "One", "first"), (257, "Two", "changed")]
    changed = read_deck(make_deck(tmp_path / "c.pptx", edited, numbered=True), config["extraction"])
    assert changed.slides[0].fingerprint == prints["256"]
    assert changed.slides[1].fingerprint != prints["257"]


def test_comment_replies_keep_their_own_author_and_text(tmp_path, config):
    modern = (
        '<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" '
        f'xmlns:a="{A}"><p188:cm id="c1" authorId="a1" created="2026-01-01T00:00:00Z">'
        '<p188:replyLst><p188:reply id="r1" authorId="a2" created="2026-01-02T00:00:00Z">'
        "<p188:txBody><a:p><a:r><a:t>Agreed</a:t></a:r></a:p></p188:txBody></p188:reply>"
        "</p188:replyLst><p188:txBody><a:p><a:r><a:t>Please check</a:t></a:r></a:p>"
        "</p188:txBody></p188:cm></p188:cmLst>"
    )
    path = write_deck(
        tmp_path / "replies.pptx",
        {
            "ppt/comments/comment1.xml": modern,
            "ppt/authors.xml": '<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="a1" name="First"/><p188:author id="a2" name="Second"/></p188:authorLst>',
        },
    )
    comments = read_deck(path, config["extraction"]).slides[0].comments
    assert [(c["author"], c["text"]) for c in comments] == [
        ("First", "Please check"),
        ("Second", "Agreed"),
    ]
    assert comments[1]["reply_to"] == "c1"


def test_malformed_missing_parts_and_extension(tmp_path, config):
    path = tmp_path / "bad.pptx"
    path.write_text("not zip")
    with pytest.raises(NoteError):
        read_deck(path, config["extraction"])
    with pytest.raises(ConfigError):
        read_deck(path.with_suffix(".ppt"), config["extraction"])
    path = write_deck(tmp_path / "missing.pptx", {"ppt/_rels/presentation.xml.rels": relations([])})
    with pytest.raises(NoteError):
        read_deck(path, config["extraction"])


def test_macro_enabled_presentations_are_read_without_running_anything(tmp_path, config):
    path = write_deck(tmp_path / "macro.pptm", {"ppt/vbaProject.bin": "binary"})
    assert len(read_deck(path, config["extraction"]).slides) == 3


def test_external_relationship_never_fetched(tmp_path, config):
    path = write_deck(
        tmp_path / "external.pptx",
        {
            "ppt/slides/_rels/slide2.xml.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="x" Type="{R}/image" Target="https://invalid.example/image.png" TargetMode="External"/></Relationships>'
        },
    )
    assert len(read_deck(path, config["extraction"]).slides) == 3


def test_table_and_chart(tmp_path, config):
    table = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    chart = f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="{R}" r:id="c1"/></a:graphicData></a:graphic></p:graphicFrame>'
    C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    path = write_deck(
        tmp_path / "table.pptx",
        {
            "ppt/slides/slide2.xml": f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{table}{chart}</p:spTree></p:cSld></p:sld>',
            "ppt/slides/_rels/slide2.xml.rels": relations(
                [("c1", "chart", "../charts/chart1.xml")]
            ),
            "ppt/charts/chart1.xml": f'<c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>2026</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
        },
    )
    shapes = read_deck(path, config["extraction"]).slides[2].shapes
    assert shapes[0].table == [["Cell"]]
    assert shapes[1].kind == "chart"
    assert shapes[1].details["chart"] == {
        "type": "barChart",
        "title": "Sales",
        "series": [{"name": "2026", "categories": ["Q1"], "values": ["10"]}],
        "data_source": "cached chart values",
    }


def test_application_metadata_does_not_leak_unselected_title_vector(tmp_path, config):
    path = write_deck(
        tmp_path / "metadata.pptx",
        {
            "docProps/app.xml": '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>PowerPoint</Application><TitlesOfParts><vector><item>Hidden slide title</item></vector></TitlesOfParts></Properties>'
        },
    )
    metadata = read_deck(path, config["extraction"]).metadata["app.xml"]
    assert metadata == {"Application": "PowerPoint"}
