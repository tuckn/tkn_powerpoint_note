"""Readable Markdown for extracted slides and AI context, and managed-note composition."""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import quote

import yaml

from .context_profiles import ContextProfile
from .frontmatter import PROXY_TYPE, NoteLoader, dump, dump_export
from .io import digest
from .models import Deck, NoteError, Shape, Slide

BEGIN = "<!-- powerpoint-note:begin -->"
END = "<!-- powerpoint-note:end -->"
MARKER = "powerpoint-note:"
MAP_HEADER = (
    "| Slide | ID | Title | Section | State | Text shapes | Pictures | Tables | Charts "
    "| Notes | Comments |\n"
    "| ---: | --- | --- | --- | --- | ---: | ---: | ---: | ---: | --- | ---: |"
)
# Phrases Office adds to alternative text it generated itself.
AUTO_ALT = ("自動的に生成された説明", "Description automatically generated", "AI 生成コンテンツ")
BLOCK_START = re.compile(r"^(\s*)(#{1,6}\s|>|[-+*]\s|\d+[.)]\s|```|~~~|---\s*$|===\s*$)")


def plain(text: str) -> str:
    """Source text as written, changed only where Markdown would hide or restructure it."""
    text = text.replace("\r", "").replace(MARKER, "powerpoint-note&#58;")
    # A bare `<word` starts an HTML tag that renderers hide.
    return re.sub(r"<(?=[A-Za-z/!?])", r"\\<", text)


def inline(text: str) -> str:
    return " / ".join(part.strip() for part in plain(text).splitlines() if part.strip())


def cell(text: str) -> str:
    return plain(text).replace("|", "\\|").replace("\n", "<br>")


def item(text: str, level: int = 0) -> str:
    """One list item; later lines continue the item instead of starting new blocks."""
    indent = "  " * min(level, 8)
    lines = [BLOCK_START.sub(r"\1\\\2", line) for line in plain(text).split("\n")]
    return indent + "- " + ("\n" + indent + "  ").join(lines)


def meaningful_alt(text: str) -> str:
    """Alternative text a person wrote; machine-generated captions and OCR debris are dropped."""
    text = text.strip()
    if not text or any(phrase in text for phrase in AUTO_ALT):
        return ""
    if re.fullmatch(r"[^\s/\\]+\.(?:png|jpe?g|gif|bmp|svg|emf|wmf|tiff?|webp)", text, re.I):
        return ""  # a file name is not a description
    tokens = text.split()
    if len(tokens) >= 8 and sorted(len(t) for t in tokens)[len(tokens) // 2] <= 2:
        return ""
    return text


def ordered(shapes: list[Shape], height: int) -> list[Shape]:
    """Titles first, then top to bottom and left to right; unknown positions keep file order."""
    band = max(height * 0.05, 1)

    def key(entry: tuple[int, Shape]) -> tuple[int, float, float, int]:
        index, shape = entry
        if shape.role in {"title", "ctrTitle"}:
            return (0, 0, 0, index)
        if shape.bbox is None:
            return (2, 0, 0, index)
        return (1, shape.bbox[1] // band, shape.bbox[0], index)

    return [shape for _, shape in sorted(enumerate(shapes), key=key)]


def chart_markdown(chart: dict[str, Any], words: dict[str, str]) -> list[str]:
    title = inline(chart.get("title", ""))
    kind = chart.get("type", "")
    label = words["chart"] + (f"（{kind}）" if kind else "") + (f": {title}" if title else "")
    series = [s for s in chart.get("series", []) if s.get("values") or s.get("y")]
    output = ["- " + label]
    if not series:
        return output
    categories = next((s.get("categories") or s.get("x") for s in series if s), None) or []
    rows = max(len(s.get("values") or s.get("y") or []) for s in series)
    header = [""] + [
        cell(s.get("name", "")) or f"{words['column']} {i}" for i, s in enumerate(series, 1)
    ]
    output += [
        "",
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * len(header)) + " |",
    ]
    for index in range(rows):
        values = [(s.get("values") or s.get("y") or []) for s in series]
        row = [cell(categories[index]) if index < len(categories) else ""]
        row += [cell(v[index]) if index < len(v) else "" for v in values]
        output.append("| " + " | ".join(row) + " |")
    output.append("")
    return output


def extracted_text(shapes: list[Shape], slide: Slide, deck: Deck, words: dict[str, str]) -> str:
    output: list[str] = []
    for shape in ordered(shapes, deck.height):
        block: list[str] = []
        note = (
            f"（{words['location_' + shape.location]}）"
            if shape.location in {"partial", "unknown"}
            else ""
        )
        is_title = shape.role in {"title", "ctrTitle"} and shape.text.strip() == slide.title.strip()
        if shape.paragraphs and not is_title:
            block += [item(p["text"], p["level"]) for p in shape.paragraphs]
            if note:
                block[0] += " " + note
        if shape.table:
            count = max(map(len, shape.table))
            rows = [r + [""] * (count - len(r)) for r in shape.table]
            # Generic column headings avoid inventing header semantics for a data-only first row.
            block += [
                "",
                "| " + " | ".join(f"{words['column']} {i}" for i in range(1, count + 1)) + " |",
                "| " + " | ".join("---" for _ in range(count)) + " |",
                *("| " + " | ".join(cell(c) for c in row) + " |" for row in rows),
                "",
            ]
            if shape.details.get("merged_cells"):
                block += ["> " + words["merged_cells"], ""]
        if shape.kind == "chart":
            block += chart_markdown(shape.details["chart"], words)
        if shape.kind == "smartart":
            texts = [t for t in shape.details.get("smartart_text", []) if t.strip()]
            block += ["- " + words["smartart"] + (":" if texts else "")]
            block += [item(t, 1) for t in texts]
        if shape.kind == "pic":
            alt = meaningful_alt(shape.alt_text)
            if alt:
                block.append(f"- {words['picture']}: {inline(alt)}")
        if block:
            output += [*block, ""]
    text = "\n".join(output)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def presentation_map(deck: Deck) -> str:
    rows = [MAP_HEADER]
    for slide in deck.slides:
        kinds = [s.kind for s in slide.shapes]
        rows.append(
            "| "
            + " | ".join(
                [
                    str(slide.number),
                    cell(slide.slide_id),
                    cell(inline(slide.title)) or "—",
                    cell(slide.section or "") or "—",
                    "hidden" if slide.hidden else "visible",
                    str(sum(bool(s.paragraphs) for s in slide.shapes)),
                    str(kinds.count("pic")),
                    str(kinds.count("table")),
                    str(kinds.count("chart")),
                    "yes" if slide.notes else "—",
                    str(len(slide.comments)),
                ]
            )
            + " |"
        )
    return "\n".join(rows)


def slide_context(profile: ContextProfile, data: dict[str, Any], image: str, stale: bool) -> str:
    sections = "\n\n".join(
        f"##### {inline(section['heading'])}\n\n{section['body']}" for section in data["sections"]
    )
    return profile.render(
        "slide",
        {
            "stale": profile.labels["stale"] if stale else "",
            "summary": data["summary"],
            "sections": sections,
            "uncertainties": "\n".join("- " + x for x in data["uncertainties"]),
            "image": image,
        },
    )


def render_body(
    deck: Deck,
    slides: list[Slide],
    profile: ContextProfile,
    *,
    off_slide: str,
    contexts: dict[str, dict[str, Any]],
    summary: dict[str, Any] | None,
    skipped: dict[str, str],
) -> str:
    """The generated part of a note. `contexts` maps slide IDs to data, image link and staleness."""
    words = profile.labels
    overview = ""
    if summary:
        points = [
            f"- {item_['text']}（{words['slide']} {', '.join(map(str, item_['slides']))}）"
            for item_ in summary["key_points"]
        ]
        overview = profile.render(
            "deck",
            {
                "summary": summary["summary"],
                "key_points": "\n".join(points),
                "uncertainties": "\n".join("- " + x for x in summary["uncertainties"]),
            },
        )
    rendered = []
    for slide in slides:
        context = contexts.get(slide.slide_id)
        context_text = (
            slide_context(profile, context["data"], context["image"], context["stale"])
            if context
            else ""
        )
        warnings = list(slide.warnings)
        if slide.slide_id in skipped:
            warnings.append(f"{words['skipped']}: {skipped[slide.slide_id]}")
        comments = "\n".join(
            item(f"{c['author']} ({c['created']}): {c['text']}", 1 if c.get("reply_to") else 0)
            for c in slide.comments
        )
        rendered.append(
            profile.render(
                "source",
                {
                    "number": str(slide.number),
                    "title": inline(slide.title) or "—",
                    "section": inline(slide.section) if slide.section else "",
                    "context": context_text,
                    "content": extracted_text(
                        [s for s in slide.shapes if s.location != "outside"], slide, deck, words
                    ),
                    "notes": "\n".join(
                        item(line) for line in slide.notes.splitlines() if line.strip()
                    ),
                    "comments": comments,
                    "outside": extracted_text(
                        [s for s in slide.shapes if s.location == "outside"], slide, deck, words
                    )
                    if off_slide == "append"
                    else "",
                    "warnings": "\n\n".join("> " + warning for warning in warnings),
                },
            )
        )
    body = profile.render(
        "note",
        {"map": presentation_map(deck), "overview": overview, "slides": "\n".join(rendered)},
    )
    return re.sub(r"\n{3,}", "\n\n", body)


def image_link(label: str, path: str) -> str:
    return f"![{label}]({quote(path, safe='/')})"


def read_frontmatter(text: str) -> tuple[dict[str, Any], str]:
    match = re.match(r"\A---\r?\n(.*?)\r?\n---\r?\n", text, re.DOTALL)
    if not match:
        raise NoteError("Existing file has no Frontmatter; it is not a note this tool manages")
    try:
        metadata = yaml.load(match[1], Loader=NoteLoader)
    except yaml.YAMLError as exc:
        raise NoteError("Invalid note Frontmatter; existing note preserved") from exc
    if not isinstance(metadata, dict):
        raise NoteError("Invalid note Frontmatter; existing note preserved")
    return metadata, text[match.end() :]


def split_note(text: str) -> tuple[dict[str, Any], str, str, str]:
    """Frontmatter, text before the generated block, the block, and text after it."""
    metadata, rest = read_frontmatter(text)
    if metadata.get("type") != PROXY_TYPE:
        raise NoteError("Existing file is not a PowerPoint proxy note; it was not changed")
    if rest.count(BEGIN) != 1 or rest.count(END) != 1 or rest.index(BEGIN) >= rest.index(END):
        raise NoteError("Missing, duplicated or reordered managed markers; existing note preserved")
    before, remainder = rest.split(BEGIN)
    body, after = remainder.split(END)
    return metadata, before, body.strip() + "\n", after


def body_digest(body: str) -> str:
    return digest(body.strip().encode("utf-8"))


def compose(metadata: dict[str, Any], before: str, body: str, after: str) -> str:
    if not before.strip():
        before = "\n# " + inline(str(metadata["title"])) + "\n\n"
    elif not before.startswith("\n"):
        before = "\n" + before
    return (
        "---\n"
        + dump(metadata)
        + "---\n"
        + before
        + BEGIN
        + "\n"
        + body.strip()
        + "\n"
        + END
        + after
    )


def compose_export(metadata: dict[str, Any], body: str) -> str:
    title = inline(str(metadata["title"]))
    return "---\n" + dump_export(metadata) + "---\n\n# " + title + "\n\n" + body.strip() + "\n"
