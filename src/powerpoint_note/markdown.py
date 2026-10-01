"""Readable source/context rendering and protected managed-note updates."""

from __future__ import annotations

import html
import json
import re
from dataclasses import asdict
from importlib.resources import files
from typing import Any
from urllib.parse import quote

import yaml

from .io import digest
from .models import Deck, NoteError, Shape, Slide

BEGIN = "<!-- powerpoint-note:begin -->"
END = "<!-- powerpoint-note:end -->"


def labels(language: str) -> dict[str, str]:
    data: dict[str, dict[str, str]] = json.loads(
        files("powerpoint_note").joinpath("resources/labels.json").read_text("utf-8")
    )
    return data[language]


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


def shape_markdown(shapes: list[Shape]) -> str:
    output = []
    for shape in shapes:
        qualifier = (
            f" [{shape.location}]" if shape.location in {"outside", "partial", "unknown"} else ""
        )
        if shape.paragraphs:
            output.append(f"**{line(shape.role or 'text')} · #{line(shape.id)}{qualifier}**")
            for p in shape.paragraphs:
                indent = "  " * min(p["level"], 8)
                output.append(
                    indent + "- " + literal(p["text"]).replace("\n", "  \n" + indent + "  ")
                )
        if shape.table:
            count = max(map(len, shape.table))
            rows = [r + [""] * (count - len(r)) for r in shape.table]
            # Generic column headings avoid inventing header semantics for a data-only first row.
            output.append("| " + " | ".join(f"Column {i}" for i in range(1, count + 1)) + " |")
            output.append("| " + " | ".join("---" for _ in range(count)) + " |")
            output.extend("| " + " | ".join(line(c) for c in row) + " |" for row in rows)
            if shape.details.get("merged_cells"):
                output.append("> Merged cells are recorded in the evidence JSON.")
        if shape.kind == "chart":
            chart = shape.details["chart"]
            output.append(f"**Chart · #{line(shape.id)}{qualifier}**")
            output.append(source_block(json.dumps(chart, ensure_ascii=False, indent=2)))
        if shape.kind == "smartart":
            output.append(f"**SmartArt · #{line(shape.id)}{qualifier}**")
            output.append(source_block("\n".join(shape.details.get("smartart_text", []))))
        if shape.kind in {"pic", "cxnSp"} or (not shape.text and not shape.table):
            details = f"; {line(shape.alt_text)}" if shape.alt_text else ""
            output.append(f"- Object #{line(shape.id)}: {line(shape.kind)}{qualifier}{details}")
            if shape.kind == "cxnSp":
                output.append(
                    "  - Endpoints/style (no inferred semantics): "
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
    contexts: dict[str, Any] | None = None,
) -> str:
    words = labels(config["generation"]["language"])
    output = [
        f"{words['coverage']}: {', '.join(str(s.number) for s in slides)} / {len(deck.slides)}",
        "",
        words["ai"] if contexts else words["local"],
        "",
        f"[{words['evidence']}]({quote(evidence_link, safe='/')})",
        "",
    ]
    if contexts:
        summary = contexts["deck"]
        output.extend([f"## {words['overview']}", "", summary["summary"], ""])
        for item in summary["key_points"]:
            cites = ", ".join(f"[{n}](#slide-{n})" for n in item["slides"])
            output.append(f"- {item['text']} ({words['slide']} {cites})")
        if summary["uncertainties"]:
            output.extend(["", f"### {words['uncertainties']}", ""])
            output.extend("- " + x for x in summary["uncertainties"])
        output.append("")
    for slide in slides:
        output.extend(
            [
                f'<a id="slide-{slide.number}"></a>',
                f"## {words['slide']} {slide.number}: {line(slide.title) or '—'}",
                "",
                f"- Slide ID: {line(slide.slide_id)}; page number: {slide.page_number}; hidden: {str(slide.hidden).lower()}",
            ]
        )
        if slide.section:
            output.append(f"- Section: {line(slide.section)}")
        output.append("")
        if contexts:
            image_link = quote(f"{image_root}/slide-{slide.number:04d}.png", safe="/")
            output.extend(
                [
                    f"![{words['slide']} {slide.number}]({image_link})",
                    "",
                    f"### {words['context']}",
                    "",
                ]
            )
            context = contexts["slides"][str(slide.number)]
            output.extend([context["summary"], ""])
            for section in context["sections"]:
                output.extend([f"#### {line(section['heading'])}", "", section["body"], ""])
            if context["uncertainties"]:
                output.extend([f"#### {words['uncertainties']}", ""])
                output.extend("- " + item for item in context["uncertainties"])
                output.append("")
        output.extend(
            [
                f"### {words['content']}",
                "",
                shape_markdown([s for s in slide.shapes if s.location != "outside"]),
                "",
            ]
        )
        if slide.notes:
            output.extend([f"### {words['notes']}", "", source_block(slide.notes), ""])
        if slide.comments:
            output.extend([f"### {words['comments']}", ""])
            for comment in slide.comments:
                output.extend(
                    [
                        f"- {line(comment['author'])} · {line(comment['created'])}",
                        source_block(comment["text"]),
                        "",
                    ]
                )
        if config["selection"]["off_slide"] == "append":
            outside = [s for s in slide.shapes if s.location == "outside"]
            if outside:
                output.extend([f"### {words['outside']}", "", shape_markdown(outside), ""])
        for warning in slide.warnings:
            output.extend(["> " + warning, ""])
    return "\n".join(output).strip() + "\n"


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
    language: str,
) -> str:
    if existing is not None:
        old, before, _, after = split_note(existing)
        old.update(metadata)
        metadata = old
    else:
        before = "\n"
        after = "\n\n" + labels(language)["personal"] + "\n"
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
