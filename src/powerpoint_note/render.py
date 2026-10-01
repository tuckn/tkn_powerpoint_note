"""Time-bounded Office automation in a child process; originals are never opened."""

from __future__ import annotations

import importlib
import json
import os
import subprocess
import sys
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring

from .io import digest
from .models import NoteError


def preflight() -> None:
    if os.name != "nt":
        raise NoteError("--context rendering requires Windows and desktop Microsoft PowerPoint")
    try:
        importlib.import_module("win32com.client")
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application\\CLSID"):
            pass
    except (ImportError, OSError) as exc:
        raise NoteError(
            "Check desktop PowerPoint installation and Python dependencies; reinstall with uv tool install . --reinstall"
        ) from exc


def check_renderable(data: bytes) -> None:
    with ZipFile(BytesIO(data)) as package:
        for name in package.namelist():
            if "vbaproject" in name.lower() or name.startswith("ppt/activeX/"):
                raise NoteError("Active content in this package cannot be rendered")
            if name.endswith(".rels"):
                for rel in fromstring(package.read(name)):
                    if rel.get("TargetMode") == "External" and not rel.get("Type", "").endswith(
                        "/hyperlink"
                    ):
                        raise NoteError(
                            "Rendering linked resources is disabled; embed external media first"
                        )


def render(
    source: Path, expected_hash: str, numbers: list[int], folder: Path, width: int, timeout: int
) -> list[Path]:
    preflight()
    data = source.read_bytes()
    if digest(data) != expected_hash:
        raise NoteError("Source changed before rendering; save the presentation and retry")
    check_renderable(data)
    with tempfile.TemporaryDirectory(prefix="powerpoint-note-render-") as temporary:
        temp = Path(temporary)
        snapshot = temp / "source.pptx"
        snapshot.write_bytes(data)
        request = temp / "request.json"
        request.write_text(
            json.dumps(
                {
                    "source": str(snapshot),
                    "numbers": numbers,
                    "folder": str(folder.resolve()),
                    "width": width,
                }
            ),
            encoding="utf-8",
        )
        try:
            process = subprocess.run(
                [sys.executable, "-m", "powerpoint_note.render", str(request)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                timeout=timeout,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired as exc:
            raise NoteError(
                "PowerPoint rendering timed out; previous note preserved. Close any source.pptx snapshot left open by Office, then retry."
            ) from exc
        if process.returncode:
            raise NoteError(
                "PowerPoint could not render the snapshot; check Office and any waiting dialog"
            )
    paths = [folder / f"slide-{n:04d}.png" for n in numbers]
    for path in paths:
        if not path.is_file() or path.read_bytes()[:8] != b"\x89PNG\r\n\x1a\n":
            raise NoteError(f"Missing or invalid rendered slide: {path.name}")
    return paths


def worker(request: dict[str, Any]) -> None:
    pythoncom = importlib.import_module("pythoncom")
    client = importlib.import_module("win32com.client")
    pythoncom.CoInitialize()
    presentation = app = None
    old_security = None
    try:
        app = client.DispatchEx("PowerPoint.Application")
        old_security = app.AutomationSecurity
        app.AutomationSecurity = 3
        presentation = app.Presentations.Open(
            request["source"], ReadOnly=True, Untitled=False, WithWindow=False
        )
        ratio = presentation.PageSetup.SlideHeight / presentation.PageSetup.SlideWidth
        for number in request["numbers"]:
            output = Path(request["folder"]) / f"slide-{number:04d}.png"
            presentation.Slides(number).Export(
                str(output), "PNG", request["width"], round(request["width"] * ratio)
            )
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        finally:
            if app is not None and old_security is not None:
                app.AutomationSecurity = old_security
            # PowerPoint may share the user's process even with DispatchEx: never Quit it.
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    worker(json.loads(Path(sys.argv[1]).read_text("utf-8")))
