"""Files kept beside a note: its state record and reusable per-slide AI results."""

from __future__ import annotations

import json
import os
import re
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import state_root
from .io import atomic_write, json_text
from .models import NoteError

STATE_VERSION = 1
STATE_FILE = "state.json"
KEY_LENGTH = 16
KEY_NAME = re.compile(r"[0-9a-f]{16}")
RESULT_FILES = {"request.json", "response.json", "usage.json", "rejected.json", "slide.png"}


def assets_dir(note: Path) -> Path:
    """`img/<presentation file name>/` beside the note, like the Excel notes' `img/`."""
    return note.parent / "img" / note.name.removesuffix(".md")


def load_state(folder: Path) -> dict[str, Any]:
    path = folder / STATE_FILE
    if not path.is_file():
        return {}
    try:
        state = json.loads(path.read_text("utf-8"))
    except (OSError, ValueError) as exc:
        raise NoteError(f"Cannot read note state; restore or remove it: {path}") from exc
    if not isinstance(state, dict) or state.get("schemaVersion") != STATE_VERSION:
        raise NoteError(f"Unsupported note state; remove it to rebuild: {path}")
    return state


def save_state(folder: Path, state: dict[str, Any]) -> None:
    atomic_write(folder / STATE_FILE, json_text({"schemaVersion": STATE_VERSION, **state}))


def result_dir(folder: Path, stage: str, key: str) -> Path:
    return folder / ("slides" if stage == "slide" else "deck") / key[:KEY_LENGTH]


def cached(folder: Path, stage: str, key: str) -> dict[str, Any] | None:
    """A validated AI result saved by an earlier run, complete or interrupted."""
    target = result_dir(folder, stage, key)
    try:
        usage = json.loads((target / "usage.json").read_text("utf-8"))
        if usage.get("status") != "success" or usage.get("key") != key:
            return None
        data: dict[str, Any] = json.loads((target / "response.json").read_text("utf-8"))
        return {"data": data, "usage": usage}
    except (OSError, ValueError):
        return None


def _generated(folder: Path) -> bool:
    """A result folder this tool wrote: a key for a name and nothing but its own files inside."""
    if folder.is_symlink() or not folder.is_dir() or not KEY_NAME.fullmatch(folder.name):
        return False
    return all(
        child.is_file() and (child.name in RESULT_FILES or child.name.endswith(".tmp"))
        for child in folder.iterdir()
    )


def prune(folder: Path, keep: set[Path]) -> int:
    """Remove saved results that the note no longer refers to. Anything else is left alone."""
    removed = 0
    for group in ("slides", "deck"):
        parent = folder / group
        if parent.is_symlink() or not parent.is_dir():
            continue
        for child in parent.iterdir():
            if child in keep or not _generated(child):
                continue
            try:
                for item in child.iterdir():
                    item.unlink()
                child.rmdir()
                removed += 1
            except OSError:
                continue  # in use; a later run removes it
    return removed


@contextmanager
def note_lock(path: Path) -> Iterator[None]:
    # A lock left by an abruptly killed process is intentionally not stolen.
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        handle = path.open("x", encoding="utf-8")
    except FileExistsError as exc:
        raise NoteError(
            f"Another run may be active. Inspect the process before removing stale lock: {path}"
        ) from exc
    try:
        with handle:
            handle.write(str(os.getpid()))
        yield
    finally:
        path.unlink(missing_ok=True)


def backup(note: Path, text: str, stamp: str | None = None) -> Path:
    """Keep the replaced note under the user state folder, outside the notes folder."""
    stamp = stamp or datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    target = state_root() / "backups" / stamp / note.name
    counter = 1
    while target.exists():
        target = target.with_name(f"{note.name}.{counter}")
        counter += 1
    atomic_write(target, text)
    return target
