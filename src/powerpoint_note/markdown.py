"""Readable source/context rendering and protected managed-note updates."""

from __future__ import annotations

import html
import json
import re
from dataclasses import asdict
from typing import Any
from urllib.parse import quote

import yaml

from .context_profiles import ContextProfile
from .io import digest
from .models import Deck, NoteError, Shape, Slide

BEGIN = "<!-- powerpoint-note:begin -->"
END = "<!-- powerpoint-note:end -->"


def literal(text: str) -> str:
    return re.sub(r"([\\\x60*_\[\]|])", r"\\\1", html.escape(text, quote=False)).replace("\r", "")


def line(text: str) -> str:
    return literal(text).replace("\n", " / ")


def source_block(text: str) -> str:
    return "\n".join("> " + literal(t) for t in text.splitlines())


def slide_evidence(slide: Slide, off_slide: str) -> dict[str, Any]:
    data = asdict(slide)
    data["excluded_off_slide_count"] = (
        sum(s.location == "outside" for s in slide.shapes) if off_slide == "exclude" else 0
    )
    data["shapes"] = [
        asdict(s) for s in slide.shapes if off_slide == "append" or s.location != "outside"
    ]
    # Titles are already restricted to objects intersecting the slide.
    return data


def shape_markdown(shapes: list[Shape], words: dict[str, str]) -> str:
    output = []
    for shape in shapes:
        qualifier = (
            f" [{words['location_' + shape.location]}]"
            if shape.location in {"outside", "partial", "unknown"}
            else ""
        )
        if shape.paragraphs:
            output.append(f"**{line(shape.role or words['text'])} · #{line(shape.id)}{qualifier}**")
            for p in shape.paragraphs:
                indent = "  " * min(p["level"], 8)
                output.append(
                    indent + "- " + literal(p["text"]).replace("\n", "  \n" + indent + "  ")
                )
        if shape.table:
            count = max(map(len, shape.table))
            rows = [r + [""] * (count - len(r)) for r in shape.table]
            # Generic column headings avoid inventing header semantics for a data-only first row.
            output.append(
                "| " + " | ".join(f"{words['column']} {i}" for i in range(1, count + 1)) + " |"
            )
            output.append("| " + " | ".join("---" for _ in range(count)) + " |")
            output.extend("| " + " | ".join(line(c) for c in row) + " |" for row in rows)
            if shape.details.get("merged_cells"):
                output.append("> " + words["merged_cells"])
        if shape.kind == "chart":
            chart = shape.details["chart"]
            output.append(f"**{words['chart']} · #{line(shape.id)}{qualifier}**")
            output.append(source_block(json.dumps(chart, ensure_ascii=False, indent=2)))
        if shape.kind == "smartart":
            output.append(f"**{words['smartart']} · #{line(shape.id)}{qualifier}**")
            output.append(source_block("\n".join(shape.details.get("smartart_text", []))))
        if shape.kind in {"pic", "cxnSp"} or (not shape.text and not shape.table):
            details = f"; {line(shape.alt_text)}" if shape.alt_text else ""
            output.append(
                f"- {words['object']} #{line(shape.id)}: {line(shape.kind)}{qualifier}{details}"
            )
            if shape.kind == "cxnSp":
                output.append(
                    "  - "
                    + words["endpoints"]
                    + ": "
                    + line(json.dumps(shape.details.get("connections", []), ensure_ascii=False))
                )
        output.append("")
    return "\n".join(output).strip()


def render_body(
    deck: Deck,
    slides: list[Slide],
    evidence_link: str,
    image_root: str,
    config: dict[str, Any],
    profile: ContextProfile,
    contexts: dict[str, Any] | None = None,
) -> str:
    words = profile.labels
    overview = ""
    if contexts:
        summary = contexts["deck"]
        points = []
        for item in summary["key_points"]:
            cites = ", ".join(f"[{n}](#slide-{n})" for n in item["slides"])
            points.append(f"- {item['text']} ({words['slide']} {cites})")
        overview = profile.render(
            "deck",
            {
                "summary": summary["summary"],
                "key_points": "\n".join(points),
                "uncertainties": "\n".join("- " + x for x in summary["uncertainties"]),
            },
        )
    rendered_slides = []
    for slide in slides:
        image, context_text = "", ""
        if contexts:
            image_link = quote(f"{image_root}/slide-{slide.number:04d}.png", safe="/")
            image = f"![{words['slide']} {slide.number}]({image_link})"
            context = contexts["slides"][str(slide.number)]
            sections = "\n\n".join(
                f"#### {line(section['heading'])}\n\n{section['body']}"
                for section in context["sections"]
            )
            context_text = profile.render(
                "slide",
                {
                    "summary": context["summary"],
                    "sections": sections,
                    "uncertainties": "\n".join("- " + x for x in context["uncertainties"]),
                },
            )
        comments = "\n\n".join(
            f"- {line(c['author'])} · {line(c['created'])}\n{source_block(c['text'])}"
            for c in slide.comments
        )
        outside = (
            shape_markdown([s for s in slide.shapes if s.location == "outside"], words)
            if config["selection"]["off_slide"] == "append"
            else ""
        )
        rendered = profile.render(
            "source",
            {
                "number": str(slide.number),
                "title": line(slide.title) or "—",
                "slide_id": line(slide.slide_id),
                "page_number": str(slide.page_number)
                if slide.page_number is not None
                else words["unknown_value"],
                "hidden": words["hidden_yes" if slide.hidden else "hidden_no"],
                "section": line(slide.section) if slide.section else "",
                "image": image,
                "context": context_text,
                "content": shape_markdown(
                    [s for s in slide.shapes if s.location != "outside"], words
                ),
                "notes": source_block(slide.notes) if slide.notes else "",
                "comments": comments,
                "outside": outside,
                "warnings": "\n\n".join("> " + warning for warning in slide.warnings),
            },
        )
        # Anchors are application-owned source identifiers, independent of translated templates.
        rendered_slides.append(f'<a id="slide-{slide.number}"></a>\n' + rendered)
    return profile.render(
        "note",
        {
            "coverage": f"{', '.join(str(s.number) for s in slides)} / {len(deck.slides)}",
            "mode": words["ai"] if contexts else words["local"],
            "evidence": f"[{words['evidence']}]({quote(evidence_link, safe='/')})",
            "overview": overview,
            "slides": "\n".join(rendered_slides),
        },
    )


def split_note(text: str) -> tuple[dict[str, Any], str, str, str]:
    match = re.match(r"\A---\n(.*?)\n---\n", text, re.DOTALL)
    if not match:
        raise NoteError("Existing file is not a managed note; choose another --output")
    try:
        metadata = yaml.safe_load(match[1])
    except yaml.YAMLError as exc:
        raise NoteError("Invalid note Frontmatter; existing note preserved") from exc
    if (
        not isinstance(metadata, dict)
        or metadata.get("generator") != "tkn-powerpoint-note"
        or metadata.get("schemaVersion") != "1.0.0"
    ):
        raise NoteError("Unrecognized note generator/schema; choose another --output")
    rest = text[match.end() :]
    if rest.count(BEGIN) != 1 or rest.count(END) != 1 or rest.index(BEGIN) >= rest.index(END):
        raise NoteError("Missing, duplicated or reordered managed markers; existing note preserved")
    before, remainder = rest.split(BEGIN)
    body, after = remainder.split(END)
    return metadata, before, body.strip() + "\n", after


def compose(
    existing: str | None,
    metadata: dict[str, Any],
    body: str,
    profile: ContextProfile,
) -> str:
    if existing is not None:
        old, before, _, after = split_note(existing)
        old.update(metadata)
        metadata = old
    else:
        before = "\n"
        after = "\n\n" + profile.labels["personal"] + "\n"
    metadata["generatedSha256"] = digest(body.encode("utf-8"))
    return (
        "---\n"
        + yaml.safe_dump(metadata, allow_unicode=True, sort_keys=False).rstrip()
        + "\n---\n"
        + before
        + BEGIN
        + "\n"
        + body
        + END
        + after
    )
