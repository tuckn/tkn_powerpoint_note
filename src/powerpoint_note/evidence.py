"""Compact per-slide evidence for AI calls and the keys that decide reuse."""

from __future__ import annotations

import json
from typing import Any

from .io import digest
from .models import Deck, Shape, Slide

# Bump when the evidence sent to AI changes shape; existing explanations are then regenerated.
EVIDENCE_CONTRACT = "3"


def _box(shape: Shape, deck: Deck) -> list[int] | None:
    if shape.bbox is None:
        return None
    x0, y0, x1, y1 = shape.bbox
    return [
        round(x0 * 1000 / deck.width),
        round(y0 * 1000 / deck.height),
        round(x1 * 1000 / deck.width),
        round(y1 * 1000 / deck.height),
    ]


def shape_payload(shape: Shape, deck: Deck) -> dict[str, Any]:
    """One shape without empty values; names and repeated colors are left out."""
    details = shape.details
    item: dict[str, Any] = {
        "id": shape.id,
        "kind": shape.kind,
        "role": shape.role,
        "box": _box(shape, deck),
        "location": shape.location if shape.location != "inside" else "",
        "text": [
            p["text"] if not p["level"] else {"level": p["level"], "text": p["text"]}
            for p in shape.paragraphs
        ],
        "table": shape.table,
        "merged_cells": details.get("merged_cells"),
        "alt": shape.alt_text,
        "group": details.get("group_ids"),
        "colors": details.get("colors"),
        "connection": details.get("connection"),
        "chart": details.get("chart"),
        "smartart": details.get("smartart_text"),
        "links": [link["target"] for link in details.get("external_links", [])],
    }
    return {key: value for key, value in item.items() if value not in (None, "", [], {})}


def slide_payload(slide: Slide, deck: Deck, off_slide: str) -> dict[str, Any]:
    """Evidence for one slide. Positions in the deck are left out so reordering keeps reuse."""
    item: dict[str, Any] = {
        "title": slide.title,
        "section": slide.section,
        "hidden": slide.hidden,
        "shapes": [shape_payload(s, deck) for s in slide.shapes if s.location != "outside"],
        "off_slide": [shape_payload(s, deck) for s in slide.shapes if s.location == "outside"]
        if off_slide == "append"
        else [],
        "notes": slide.notes,
        "comments": [
            {key: value for key, value in comment.items() if value not in (None, "")}
            for comment in slide.comments
        ],
        "warnings": slide.warnings,
    }
    return {key: value for key, value in item.items() if value not in (None, "", [], {}, False)}


def payload_text(payload: Any) -> str:
    """Deterministic JSON without whitespace: every character counts against input limits."""
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def content_key(slide: Slide, payload: str, image_width: int, deck: Deck) -> str:
    """Identity of what AI would see for a slide: its drawing, its evidence and the image size."""
    return digest(
        payload_text(
            {
                "contract": EVIDENCE_CONTRACT,
                "fingerprint": slide.fingerprint,
                "payload": digest(payload.encode("utf-8")),
                "image_width": image_width,
                # The slide size shapes the picture even when no slide part changes.
                "slide_size": [deck.width, deck.height],
            }
        ).encode("utf-8")
    )


def generation_key(content: str, stage: str, profile: dict[str, Any], plan: dict[str, Any]) -> str:
    """Identity of one AI result: content plus the prompt, schema and connection that shape it."""
    return digest(
        payload_text(
            {
                "content": content,
                "stage": stage,
                "prompt": digest(profile["prompt"].encode("utf-8")),
                "schema": digest(payload_text(profile["schema"]).encode("utf-8")),
                "profile": profile["name"],
                "language": profile["language"],
                "provider": plan.get("provider"),
                "model": plan.get("model"),
                "reasoning_effort": plan.get("reasoning_effort"),
            }
        ).encode("utf-8")
    )
