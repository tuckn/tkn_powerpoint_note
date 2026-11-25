"""Read PowerPoint OOXML without Office, modifying files, or resolving external links."""

from __future__ import annotations

import hashlib
import math
import posixpath
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import unquote
from xml.etree.ElementTree import Element, ParseError, tostring
from zipfile import BadZipFile, ZipFile

from defusedxml.ElementTree import fromstring

from .io import digest
from .models import ConfigError, Deck, NoteError, Shape, Slide

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
RID = "{" + NS["r"] + "}id"
Matrix = tuple[float, float, float, float, float, float]
IDENTITY: Matrix = (1, 0, 0, 1, 0, 0)
FALSE = {"0", "false", "off"}
SLIDE_NUMBER = "‹#›"
# Master placeholders are matched by role; content placeholders inherit from the body.
MASTER_ROLES = {"title": "title", "ctrTitle": "title", "dt": "dt", "ftr": "ftr", "sldNum": "sldNum"}
SUFFIXES = {".pptx", ".pptm"}
CONNECTOR_ENDS = {
    "stCxn": ("from", "id"),
    "endCxn": ("to", "id"),
    "headEnd": ("start_arrow", "type"),
    "tailEnd": ("end_arrow", "type"),
}


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


class Package:
    def __init__(self, data: bytes, max_uncompressed_mb: int):
        self.zip = ZipFile(BytesIO(data))
        infos = self.zip.infolist()
        if len({x.filename for x in infos}) != len(infos):
            raise NoteError("Duplicate ZIP entries are not supported")
        if sum(x.file_size for x in infos) > max_uncompressed_mb * 1024 * 1024:
            raise NoteError("Uncompressed presentation exceeds extraction.max_uncompressed_mb")
        self.names = set(self.zip.namelist())
        self.cache: dict[str, Element] = {}
        self.digests: dict[str, str] = {}
        self.trees: dict[str, str] = {}

    def content_fingerprint(self) -> str:
        """Hash package content without ZIP bookkeeping or document properties."""
        result = hashlib.sha256()
        properties = {"docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"}
        for name in sorted(self.names):
            if name in properties or name.endswith("/"):
                continue
            result.update(name.encode("utf-8"))
            result.update(b"\0")
            result.update(self.zip.read(name))
            result.update(b"\0")
        return result.hexdigest()

    def part_digest(self, name: str) -> str:
        """Hash one part, ignoring the page number PowerPoint rewrites when slides move."""
        if name not in self.digests:
            data = self.zip.read(name)
            if b"slidenum" in data and name.endswith(".xml"):
                root = fromstring(data)
                for field in root.iter("{" + NS["a"] + "}fld"):
                    if field.get("type") == "slidenum":
                        for text in field.iter("{" + NS["a"] + "}t"):
                            text.text = ""
                data = tostring(root)
            self.digests[name] = digest(data)
        return self.digests[name]

    def slide_fingerprint(self, part: str) -> str:
        """Hash a slide with everything it draws: layout, master, theme, media, charts, notes.

        Part names are left out: PowerPoint renumbers `slideN.xml` and `notesSlideN.xml`
        when slides are inserted, moved or deleted, while their contents stay the same.
        """
        return self._tree_digest(part, ())

    def _tree_digest(self, part: str, trail: tuple[str, ...]) -> str:
        if part in self.trees:
            return self.trees[part]
        if part in trail or part not in self.names:
            return ""
        result = hashlib.sha256(self.part_digest(part).encode())
        kind = posixpath.basename(posixpath.dirname(part))
        for identifier, rel in sorted(self.rels(part).items()):
            if rel["type"] in {"slide", "notesMaster"}:
                continue  # back references and the notes page design
            if kind == "slideMasters" and rel["type"] == "slideLayout":
                continue  # sibling layouts do not draw this slide
            child = (
                "external:" + rel["target"]
                if rel["external"] == "True"
                else self._tree_digest(rel["target"], (*trail, part))
            )
            result.update(f"|{identifier}|{rel['type']}|{child}".encode())
        self.trees[part] = result.hexdigest()
        return self.trees[part]

    def xml(self, name: str) -> Element:
        if name not in self.names:
            raise NoteError(f"Missing required package part: {name}")
        if name not in self.cache:
            self.cache[name] = fromstring(self.zip.read(name))
        return self.cache[name]

    def rels(self, part: str) -> dict[str, dict[str, str]]:
        parent, name = posixpath.split(part)
        relpart = posixpath.join(parent, "_rels", name + ".rels")
        if relpart not in self.names:
            return {}
        result = {}
        for node in self.xml(relpart):
            target = node.get("Target", "")
            external = node.get("TargetMode") == "External"
            resolved = (
                target
                if external
                else posixpath.normpath(posixpath.join(parent, unquote(target))).lstrip("/")
            )
            if not external and (resolved == ".." or resolved.startswith("../")):
                raise NoteError("Package relationship escapes the package")
            result[node.get("Id", "")] = {
                "target": resolved,
                "type": node.get("Type", "").rsplit("/", 1)[-1],
                "external": str(external),
            }
        return result

    def related(self, part: str, kind: str) -> list[str]:
        return [
            v["target"]
            for v in self.rels(part).values()
            if v["type"] == kind and v["external"] == "False"
        ]


def multiply(m: Matrix, n: Matrix) -> Matrix:
    a, b, c, d, e, f = m
    g, h, i, j, k, offset_y = n
    return (
        a * g + c * h,
        b * g + d * h,
        a * i + c * j,
        b * i + d * j,
        a * k + c * offset_y + e,
        b * k + d * offset_y + f,
    )


def point(m: Matrix, x: float, y: float) -> tuple[float, float]:
    a, b, c, d, e, f = m
    return a * x + c * y + e, b * x + d * y + f


def values(xfrm: Element, tag: str, keys: tuple[str, str]) -> tuple[float, float]:
    node = xfrm.find("a:" + tag, NS)
    return (
        (0, 0) if node is None else (float(node.get(keys[0], "0")), float(node.get(keys[1], "0")))
    )


def orientation(xfrm: Element, x: float, y: float, w: float, h: float) -> Matrix:
    angle = math.radians(float(xfrm.get("rot", "0")) / 60000)
    sx = -1 if xfrm.get("flipH") in {"1", "true"} else 1
    sy = -1 if xfrm.get("flipV") in {"1", "true"} else 1
    rotate: Matrix = (
        math.cos(angle) * sx,
        math.sin(angle) * sx,
        -math.sin(angle) * sy,
        math.cos(angle) * sy,
        0,
        0,
    )
    cx, cy = x + w / 2, y + h / 2
    return multiply((1, 0, 0, 1, cx, cy), multiply(rotate, (1, 0, 0, 1, -cx, -cy)))


def bbox(xfrm: Element | None, transform: Matrix) -> list[int] | None:
    if xfrm is None:
        return None
    x, y = values(xfrm, "off", ("x", "y"))
    w, h = values(xfrm, "ext", ("cx", "cy"))
    m = multiply(transform, orientation(xfrm, x, y, w, h))
    corners = [point(m, px, py) for px, py in ((x, y), (x + w, y), (x, y + h), (x + w, y + h))]
    return [round(min(p[i] for p in corners)) for i in (0, 1)] + [
        round(max(p[i] for p in corners)) for i in (0, 1)
    ]


def location(box: list[int] | None, width: int, height: int) -> str:
    if box is None:
        return "unknown"
    x0, y0, x1, y1 = box
    if (
        x1 < 0
        or y1 < 0
        or x0 > width
        or y0 > height
        or (x1 == 0 and x0 < 0)
        or (y1 == 0 and y0 < 0)
    ):
        return "outside"
    if x0 < 0 or y0 < 0 or x1 > width or y1 > height:
        return "partial"
    return "inside"


def paragraphs(node: Element) -> list[dict[str, Any]]:
    result = []
    for p in node.findall(".//a:p", NS):
        parts: list[str] = []
        for child in p:
            kind = local(child.tag)
            if kind == "br":
                parts.append("\n")
            elif kind == "fld" and child.get("type") == "slidenum":
                parts.append(SLIDE_NUMBER)
            else:
                parts.extend(
                    e.text or "" if local(e.tag) == "t" else "\n"
                    for e in child.iter()
                    if local(e.tag) in {"t", "br"}
                )
        text = "".join(parts)
        if not text.strip():
            continue
        props = p.find("a:pPr", NS)
        level = int(props.get("lvl", "0")) if props is not None else 0
        bullet = props is not None and any(local(e.tag) in {"buChar", "buAutoNum"} for e in props)
        result.append({"text": text, "level": level, "bullet": bullet})
    return result


def placeholder(node: Element) -> Element | None:
    return node.find(".//p:nvPr/p:ph", NS)


def own_transform(node: Element) -> Element | None:
    own = node.find("p:spPr/a:xfrm", NS)
    return own if own is not None else node.find("p:xfrm", NS)


def placeholders(root: Element | None) -> list[tuple[Element, Element]]:
    if root is None:
        return []
    found = []
    for candidate in root.iter():
        if local(candidate.tag) in {"sp", "pic", "graphicFrame"}:
            ph = placeholder(candidate)
            if ph is not None:
                found.append((candidate, ph))
    return found


def inherited_transform(
    node: Element, layout: Element | None, master: Element | None
) -> Element | None:
    """Resolve slide -> layout (by idx, then type) -> master (by the layout's role)."""
    own = own_transform(node)
    if own is not None:
        return own
    ph = placeholder(node)
    if ph is None:
        return None
    role = ph.get("type", "obj")
    candidates = placeholders(layout)
    match = None
    if ph.get("idx") is not None:
        match = next((c for c in candidates if c[1].get("idx") == ph.get("idx")), None)
    if match is None:
        match = next((c for c in candidates if c[1].get("type", "obj") == role), None)
    if match is not None:
        found = own_transform(match[0])
        if found is not None:
            return found
        role = match[1].get("type", role)
    wanted = MASTER_ROLES.get(role, "body")
    for candidate, other in placeholders(master):
        if MASTER_ROLES.get(other.get("type", "obj"), "body") == wanted:
            found = own_transform(candidate)
            if found is not None:
                return found
    return None


CHART_FIELDS = {
    "tx": "name",
    "cat": "categories",
    "val": "values",
    "xVal": "x",
    "yVal": "y",
    "bubbleSize": "size",
}


def chart_data(pkg: Package, target: str) -> dict[str, Any]:
    chart = pkg.xml(target)
    plot = chart.find(".//c:plotArea", NS)
    kinds = (
        [local(e.tag) for e in plot if local(e.tag).endswith("Chart")] if plot is not None else []
    )
    series = []
    for ser in chart.findall(".//c:ser", NS):
        item: dict[str, Any] = {}
        for key, label in CHART_FIELDS.items():
            child = ser.find("c:" + key, NS)
            if child is not None:
                found = [e.text or "" for e in child.iter() if local(e.tag) == "v"]
                item[label] = "\n".join(found) if key == "tx" else found
        series.append(item)
    title = chart.find("c:chart/c:title", NS)
    return {
        "type": ", ".join(dict.fromkeys(kinds)),
        "title": "".join(e.text or "" for e in title.iter() if local(e.tag) == "t")
        if title is not None
        else "",
        "series": series,
        "data_source": "cached chart values",
    }


def read_shapes(
    pkg: Package,
    root: Element,
    part: str,
    width: int,
    height: int,
    layout: Element | None,
    master: Element | None,
) -> list[Shape]:
    rels = pkg.rels(part)
    result: list[Shape] = []

    def walk(parent: Element, transform: Matrix, groups: list[str]) -> None:
        for node in parent:
            kind = local(node.tag)
            if kind == "AlternateContent":
                choices = list(node)
                fallback = next((e for e in choices if local(e.tag) == "Fallback"), None)
                chosen = fallback if fallback is not None else (choices[0] if choices else None)
                if chosen is not None:
                    walk(chosen, transform, groups)
                continue
            if kind not in {"sp", "pic", "cxnSp", "graphicFrame", "grpSp"}:
                continue
            props = node.find(".//p:cNvPr", NS)
            shape_id = props.get("id", "") if props is not None else ""
            if kind == "grpSp":
                xfrm = node.find("p:grpSpPr/a:xfrm", NS)
                m = transform
                if xfrm is not None:
                    x, y = values(xfrm, "off", ("x", "y"))
                    w, h = values(xfrm, "ext", ("cx", "cy"))
                    ox, oy = values(xfrm, "chOff", ("x", "y"))
                    cw, ch = values(xfrm, "chExt", ("cx", "cy"))
                    sx, sy = w / cw if cw else 1, h / ch if ch else 1
                    m = multiply(
                        transform,
                        multiply(
                            orientation(xfrm, x, y, w, h), (sx, 0, 0, sy, x - ox * sx, y - oy * sy)
                        ),
                    )
                walk(node, m, groups + [shape_id])
                continue
            ph = placeholder(node)
            box = bbox(inherited_transform(node, layout, master), transform)
            shape = Shape(
                id=shape_id,
                kind=kind,
                name=props.get("name", "") if props is not None else "",
                role=ph.get("type", "obj") if ph is not None else "",
                bbox=box,
                location=location(box, width, height),
                alt_text=props.get("descr", props.get("title", "")) if props is not None else "",
            )
            textbody = node.find("p:txBody", NS)
            if textbody is not None:
                shape.paragraphs = paragraphs(textbody)
            if groups:
                shape.details["group_ids"] = groups
            table = node.find(".//a:tbl", NS)
            if table is not None:
                shape.kind = "table"
                shape.table = [
                    [
                        "\n".join(p["text"] for p in paragraphs(cell))
                        for cell in row.findall("a:tc", NS)
                    ]
                    for row in table.findall("a:tr", NS)
                ]
                shape.details["merged_cells"] = [
                    {
                        "row": i,
                        "column": j,
                        **{
                            k: v
                            for k, v in cell.attrib.items()
                            if k in {"gridSpan", "rowSpan", "hMerge", "vMerge"}
                        },
                    }
                    for i, row in enumerate(table.findall("a:tr", NS), 1)
                    for j, cell in enumerate(row.findall("a:tc", NS), 1)
                    if any(k in cell.attrib for k in {"gridSpan", "rowSpan", "hMerge", "vMerge"})
                ]
            colors = list(
                dict.fromkeys(
                    e.get("val", "")
                    if local(e.tag) == "schemeClr"
                    else "#" + e.get("val", e.get("lastClr", ""))
                    for e in node.iter()
                    if local(e.tag) in {"srgbClr", "schemeClr", "sysClr"}
                )
            )
            if colors:
                shape.details["colors"] = colors
            if kind == "cxnSp":
                ends = {local(e.tag): e for e in node.iter() if local(e.tag) in CONNECTOR_ENDS}
                shape.details["connection"] = {
                    label: ends[tag].get(attribute, "")
                    for tag, (label, attribute) in CONNECTOR_ENDS.items()
                    if tag in ends and ends[tag].get(attribute, "") not in ("", "none")
                }
            links = []
            for e in node.iter():
                for key, value in e.attrib.items():
                    if key.startswith("{" + NS["r"] + "}") and value in rels:
                        rel = rels[value]
                        if rel["external"] == "True":
                            links.append(
                                {"type": rel["type"], "target": rel["target"], "fetched": False}
                            )
                        elif rel["type"] == "chart":
                            shape.kind = "chart"
                            shape.details["chart"] = chart_data(pkg, rel["target"])
                        elif rel["type"] == "diagramData":
                            shape.kind = "smartart"
                            shape.details["smartart_text"] = [
                                x.text or ""
                                for x in pkg.xml(rel["target"]).iter()
                                if local(x.tag) == "t"
                            ]
            if links:
                shape.details["external_links"] = links
            result.append(shape)

    tree = root.find("p:cSld/p:spTree", NS)
    if tree is not None:
        walk(tree, IDENTITY, [])
    return result


def read_notes(pkg: Package, part: str) -> str:
    blocks: list[str] = []
    for target in pkg.related(part, "notesSlide"):
        root = pkg.xml(target)
        for shape in root.findall(".//p:sp", NS):
            ph = placeholder(shape)
            if ph is not None and ph.get("type") in {"sldImg", "sldNum", "dt", "hdr", "ftr"}:
                continue
            blocks.extend(p["text"] for p in paragraphs(shape))
    return "\n".join(blocks).strip()


def comment_text(node: Element) -> str:
    """Text of one comment or reply, without the replies nested below it."""
    found: list[str] = []

    def walk(element: Element) -> None:
        for child in element:
            kind = local(child.tag)
            if kind in {"replyLst", "reply"}:
                continue
            if kind in {"text", "t"}:
                found.append(child.text or "")
            walk(child)

    walk(node)
    return "\n".join(found).strip()


def read_comments(pkg: Package, part: str, authors: dict[str, str]) -> list[dict[str, Any]]:
    result = []
    for rel in pkg.rels(part).values():
        if "comment" not in rel["type"].lower() or rel["external"] == "True":
            continue
        root = pkg.xml(rel["target"])
        parents = {child: parent for parent in root.iter() for child in parent}
        for node in root.iter():
            kind = local(node.tag)
            if kind not in {"cm", "comment", "reply"}:
                continue
            text = comment_text(node)
            if not text:
                continue
            item = {
                "author": authors.get(node.get("authorId", ""), node.get("authorId", "")),
                "created": node.get("dt", node.get("created", "")),
                "text": text,
                "resolved": node.get("resolved"),
                "id": node.get("id", node.get("idx", "")),
            }
            if kind == "reply":
                parent = parents.get(node)
                while parent is not None and local(parent.tag) not in {"cm", "comment"}:
                    parent = parents.get(parent)
                item["reply_to"] = parent.get("id", "") if parent is not None else ""
            result.append(item)
    return result


def read_deck(path: Path, settings: dict[str, Any]) -> Deck:
    if path.suffix.lower() not in SUFFIXES:
        raise ConfigError(
            "Input must be a .pptx or .pptm file; convert legacy .ppt files in PowerPoint first"
        )
    if not path.is_file():
        raise ConfigError(f"Input file does not exist: {path}")
    limit = settings["max_file_mb"] * 1024 * 1024
    if path.stat().st_size > limit:
        raise NoteError("Presentation exceeds extraction.max_file_mb")
    with path.open("rb") as stream:
        data = stream.read(limit + 1)
    if len(data) > limit:
        raise NoteError("Presentation exceeds extraction.max_file_mb")
    pkg = None
    try:
        pkg = Package(data, settings["max_uncompressed_mb"])
        root = pkg.xml("ppt/presentation.xml")
        if root.tag != "{" + NS["p"] + "}presentation":
            raise NoteError("Unsupported PresentationML namespace; save as a standard .pptx")
        dimensions = root.find("p:sldSz", NS)
        if dimensions is None:
            raise NoteError("Presentation has no slide dimensions")
        width, height = int(dimensions.get("cx", "0")), int(dimensions.get("cy", "0"))
        if min(width, height) <= 0:
            raise NoteError("Invalid slide dimensions")
        sections = {}
        for section in root.iter():
            if local(section.tag) == "section":
                for sid in section.iter():
                    if local(sid.tag) == "sldId":
                        sections[sid.get("id", "")] = section.get("name", "")
        metadata: dict[str, Any] = {}
        for name in ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"):
            if name in pkg.names:
                metadata[posixpath.basename(name)] = {
                    e.get("name", local(e.tag)): "".join(e.itertext()).strip()
                    for e in pkg.xml(name)
                    if name != "docProps/app.xml" or len(e) == 0
                }
        authors = {}
        for name in sorted(pkg.names):
            if name.endswith(".xml") and ("author" in name.lower()):
                for e in pkg.xml(name).iter():
                    if e.get("id") is not None:
                        authors[e.get("id", "")] = e.get("name", e.get("displayName", ""))
        deck = Deck(
            digest(data),
            len(data),
            width,
            height,
            metadata,
            [],
            content_fingerprint=pkg.content_fingerprint(),
        )
        relationships = pkg.rels("ppt/presentation.xml")
        seen = set()
        for number, sid in enumerate(root.findall("p:sldIdLst/p:sldId", NS), 1):
            rel = relationships.get(sid.get(RID, ""))
            if not rel or rel["external"] == "True" or rel["type"] != "slide":
                raise NoteError(f"Invalid relationship for slide {number}")
            if sid.get("id", "") in seen:
                raise NoteError("Duplicate presentation slide IDs")
            seen.add(sid.get("id", ""))
            part = rel["target"]
            xml = pkg.xml(part)
            layout_parts = pkg.related(part, "slideLayout")
            master_parts = pkg.related(layout_parts[0], "slideMaster") if layout_parts else []
            layout = pkg.xml(layout_parts[0]) if layout_parts else None
            master = pkg.xml(master_parts[0]) if master_parts else None
            slide = Slide(
                number,
                sid.get("id", ""),
                part,
                xml.get("show", "1").lower() in FALSE,
                int(root.get("firstSlideNum", "1")) + number - 1,
                sections.get(sid.get("id", "")),
            )
            slide.shapes = read_shapes(pkg, xml, part, width, height, layout, master)
            slide.fingerprint = pkg.slide_fingerprint(part)
            slide.title = next(
                (
                    s.text
                    for s in slide.shapes
                    if s.role in {"title", "ctrTitle"} and s.location != "outside" and s.text
                ),
                "",
            )
            if settings["include_notes"]:
                slide.notes = read_notes(pkg, part)
            if settings["include_comments"]:
                slide.comments = read_comments(pkg, part, authors)
            if any(s.location == "unknown" for s in slide.shapes):
                slide.warnings.append(
                    "Some object positions are unavailable; these objects are retained."
                )
            if any(s.location == "partial" for s in slide.shapes):
                slide.warnings.append("Partially clipped objects retain their full extracted text.")
            if xml.find("p:timing", NS) is not None:
                slide.warnings.append("Animations and build order are not interpreted.")
            deck.slides.append(slide)
        if not deck.slides:
            raise NoteError("Presentation contains no slides")
        deck.warnings.append(
            "Text extraction does not interpret visual relationships, master/layout artwork, or media playback."
        )
        return deck
    except NoteError:
        raise
    except (BadZipFile, ParseError, KeyError, ValueError, TypeError, OSError) as exc:
        raise NoteError(f"Cannot parse PowerPoint package ({type(exc).__name__})") from exc
    finally:
        if pkg is not None:
            pkg.zip.close()
