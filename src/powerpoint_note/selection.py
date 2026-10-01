"""Slide selection always uses 1-based presentation order, not XML filenames."""

from __future__ import annotations

import re
from typing import Any

from .models import Deck, NoteError, Slide


def select(deck: Deck, settings: dict[str, Any]) -> list[Slide]:
    spec = settings["slides"]
    numbers: set[int] | None = None
    if spec is not None:
        numbers = set()
        for item in spec.split(","):
            match = re.fullmatch(r"\s*(\d+)(?:\s*-\s*(\d+))?\s*", item)
            if match is None:
                raise NoteError("slides must contain page numbers or inclusive ranges, e.g. 1,3-5")
            start, end = int(match[1]), int(match[2] or match[1])
            if not 1 <= start <= end <= len(deck.slides):
                raise NoteError(f"Slide range {item.strip()} is outside 1..{len(deck.slides)}")
            numbers.update(range(start, end + 1))
    sections = settings["sections"]
    known = {s.section for s in deck.slides if s.section is not None}
    missing = set(sections) - known
    if missing:
        raise NoteError(f"Unknown section(s): {', '.join(sorted(missing))}; inspect the deck first")
    selected = [
        s
        for s in deck.slides
        if (numbers is None or s.number in numbers)
        and (not sections or s.section in sections)
        and (settings["include_hidden"] or not s.hidden)
    ]
    if not selected:
        raise NoteError("No slides selected; check slide ranges, section names, and include_hidden")
    return selected
