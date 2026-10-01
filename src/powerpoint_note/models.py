"""Serializable extraction models. Coordinates are slide-space EMUs."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


class NoteError(Exception):
    """An actionable input, configuration, or processing error."""


@dataclass
class Shape:
    id: str
    kind: str
    name: str
    role: str = ""
    paragraphs: list[dict[str, Any]] = field(default_factory=list)
    bbox: list[float] | None = None
    location: str = "unknown"
    alt_text: str = ""
    table: list[list[str]] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)

    @property
    def text(self) -> str:
        return "\n".join(p["text"] for p in self.paragraphs)


@dataclass
class Slide:
    number: int
    slide_id: str
    part: str
    hidden: bool
    page_number: int
    section: str | None = None
    title: str = ""
    shapes: list[Shape] = field(default_factory=list)
    notes: str = ""
    comments: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class Deck:
    source_sha256: str
    source_size: int
    width: int
    height: int
    metadata: dict[str, Any]
    slides: list[Slide]
    warnings: list[str] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)
