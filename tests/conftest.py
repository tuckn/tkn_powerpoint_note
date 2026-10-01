from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from powerpoint_note.config import resolve

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def relations(items: list[tuple[str, str, str]]) -> str:
    return (
        f'<Relationships xmlns="{REL}">'
        + "".join(
            f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{target}"/>'
            for rid, kind, target in items
        )
        + "</Relationships>"
    )


def shape(
    sid: int, text: str, x: int = 10, y: int = 10, role: str = "", w: int = 200, h: int = 50
) -> str:
    ph = f'<p:ph type="{role}"/>' if role else ""
    return f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Object {sid}"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>'


def write_deck(path: Path, extra: dict[str, str] | None = None) -> Path:
    root = f'<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" firstSlideNum="7">'
    root += '<p:sldIdLst><p:sldId id="300" r:id="r3"/><p:sldId id="100" r:id="r1"/><p:sldId id="200" r:id="r2"/></p:sldIdLst><p:sldSz cx="1000" cy="750"/>'
    root += '<p:extLst><p:ext uri="sections"><p14:sectionLst><p14:section name="Overview" id="one"><p14:sldIdLst><p14:sldId id="300"/></p14:sldIdLst></p14:section><p14:section name="Design" id="two"><p14:sldIdLst><p14:sldId id="100"/><p14:sldId id="200"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>'
    group = (
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="Group"/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="400" cy="200"/><a:chOff x="0" y="0"/><a:chExt cx="200" cy="100"/></a:xfrm></p:grpSpPr>'
        + shape(10, "Grouped detail", 10, 10, w=20, h=20)
        + "</p:grpSp>"
    )
    first = (
        shape(1, "A &amp; B", role="title")
        + shape(2, "Visible body", y=100)
        + shape(3, "Private off-canvas", 1200)
        + shape(4, "Partial", 900)
        + group
    )
    slides = {}
    for name, hidden, body in [
        ("slide3", "1", first),
        ("slide1", "0", shape(1, "Hidden", role="title")),
        ("slide2", "1", shape(1, "Final", role="title")),
    ]:
        slides[f"ppt/slides/{name}.xml"] = (
            f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}" show="{hidden}"><p:cSld><p:spTree>{body}</p:spTree></p:cSld></p:sld>'
        )
    contents = {
        "ppt/presentation.xml": root,
        "ppt/_rels/presentation.xml.rels": relations(
            [
                ("r3", "slide", "slides/slide3.xml"),
                ("r1", "slide", "slides/slide1.xml"),
                ("r2", "slide", "slides/slide2.xml"),
            ]
        ),
        "ppt/slides/_rels/slide3.xml.rels": relations(
            [
                ("n", "notesSlide", "../notesSlides/notesSlide1.xml"),
                ("c", "comments", "../comments/comment1.xml"),
            ]
        ),
        "ppt/notesSlides/notesSlide1.xml": f'<p:notes xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>{shape(1, "Speaker evidence", role="body")}{shape(2, "99", role="sldNum")}</p:spTree></p:cSld></p:notes>',
        "ppt/comments/comment1.xml": f'<p:cmLst xmlns:p="{P}"><p:cm authorId="0" dt="2026-01-01T00:00:00Z" idx="1"><p:text>Review comment</p:text></p:cm></p:cmLst>',
        "ppt/commentAuthors.xml": f'<p:cmAuthorLst xmlns:p="{P}"><p:cmAuthor id="0" name="Example Author"/></p:cmAuthorLst>',
        "docProps/core.xml": '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Example Deck</dc:title></cp:coreProperties>',
        **slides,
    }
    contents.update(extra or {})
    with ZipFile(path, "w", ZIP_DEFLATED) as package:
        for name, value in contents.items():
            package.writestr(name, value)
    return path


@pytest.fixture
def deck_path(tmp_path: Path) -> Path:
    return write_deck(tmp_path / "example.pptx")


@pytest.fixture
def config(tmp_path: Path) -> dict:
    return resolve(home=tmp_path / "home", cwd=tmp_path / "cwd")[0]
