"""Slide selection always uses 1-based presentation order, not XML filenames."""

from __future__ import annotations

import re
from typing import Any

from .models import ConfigError, Deck, NoteError, Slide

RANGE = re.compile(r"\s*(\d+)(?:\s*-\s*(\d+))?\s*")


def format_slides(value: Any) -> str:
    """Compact notation such as `3-6` or `1-2,7-400` for a list of positions."""
    if value in (None, "", []):
        return ""
    if not isinstance(value, list) or any(type(n) is not int or n < 1 for n in value):
        raise NoteError("Slide positions must be positive integers")
    ranges: list[list[int]] = []
    for number in sorted(set(value)):
        if ranges and number == ranges[-1][1] + 1:
            ranges[-1][1] = number
        else:
            ranges.append([number, number])
    return ",".join(str(a) if a == b else f"{a}-{b}" for a, b in ranges)


def select(
    deck: Deck, included: list[Slide], spec: str | None, sections: list[str] | None
) -> list[Slide]:
    """Slides of `included` that match `--slides` and `--section`; both must match if given."""
    numbers: set[int] | None = None
    if spec is not None:
        numbers = set()
        for item in spec.split(","):
            match = RANGE.fullmatch(item)
            if match is None:
                raise ConfigError("--slides takes positions or inclusive ranges, e.g. 1,3-5")
            start, end = int(match[1]), int(match[2] or match[1])
            if not 1 <= start <= end <= len(deck.slides):
                raise ConfigError(f"Slide range {item.strip()} is outside 1..{len(deck.slides)}")
            numbers.update(range(start, end + 1))
    sections = sections or []
    missing = set(sections) - {s.section for s in deck.slides if s.section is not None}
    if missing:
        raise ConfigError(
            f"Unknown section(s): {', '.join(sorted(missing))}; "
            "list them with presentation list-slides"
        )
    selected = [
        s
        for s in included
        if (numbers is None or s.number in numbers) and (not sections or s.section in sections)
    ]
    if not selected:
        raise ConfigError("No slides selected; check --slides, --section and hidden slides")
    return selected
