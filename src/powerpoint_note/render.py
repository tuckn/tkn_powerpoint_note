"""Time-bounded Office automation in a child process; originals are never opened."""

from __future__ import annotations

import importlib
import json
import logging
import os
import subprocess
import sys
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring

from .models import NoteError

LOGGER = logging.getLogger("powerpoint_note")
PNG = b"\x89PNG\r\n\x1a\n"


class RenderError(NoteError):
    """PowerPoint could not draw the selected slides; text extraction is unaffected."""


def preflight() -> None:
    if os.name != "nt":
        raise RenderError("--context rendering requires Windows and desktop Microsoft PowerPoint")
    try:
        importlib.import_module("win32com.client")
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application\\CLSID"):
            pass
    except (ImportError, OSError) as exc:
        raise RenderError(
            "Check desktop PowerPoint installation and Python dependencies; "
            "reinstall with uv tool install . --reinstall"
        ) from exc


def check_renderable(data: bytes) -> None:
    with ZipFile(BytesIO(data)) as package:
        for name in package.namelist():
            if "vbaproject" in name.lower() or name.startswith("ppt/activeX/"):
                raise RenderError("Active content in this package cannot be rendered")
            if name.endswith(".rels"):
                for rel in fromstring(package.read(name)):
                    if rel.get("TargetMode") == "External" and not rel.get("Type", "").endswith(
                        "/hyperlink"
                    ):
                        raise RenderError(
                            "Rendering linked resources is disabled; embed external media first"
                        )


def render(data: bytes, targets: dict[int, Path], width: int, timeout: int) -> None:
    """Draw the given 1-based slide positions of a presentation snapshot as PNG files."""
    preflight()
    check_renderable(data)
    # A hung PowerPoint can keep the snapshot open; never let cleanup hide the real error.
    with tempfile.TemporaryDirectory(
        prefix="powerpoint-note-render-", ignore_cleanup_errors=True
    ) as temporary:
        temp = Path(temporary)
        snapshot = temp / "source.pptx"
        snapshot.write_bytes(data)
        request = temp / "request.json"
        request.write_text(
            json.dumps(
                {
                    "source": str(snapshot),
                    "targets": {str(n): str(path.resolve()) for n, path in targets.items()},
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
                errors="replace",
                timeout=timeout,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired as exc:
            raise RenderError(
                f"PowerPoint rendering exceeded generation.limits.max_render_seconds={timeout}. "
                "Close any source.pptx snapshot left open by Office, then retry."
            ) from exc
        if process.returncode:
            detail = process.stderr.strip()
            LOGGER.debug("PowerPoint render process output:\n%s", detail)
            reason = detail.splitlines()[-1] if detail else f"exit code {process.returncode}"
            raise RenderError(
                f"PowerPoint could not render the snapshot ({reason}); "
                "check Office and any waiting dialog"
            )
    for path in targets.values():
        if not path.is_file() or path.read_bytes()[:8] != PNG:
            raise RenderError(f"Missing or invalid rendered slide: {path.name}")


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
        for number, output in request["targets"].items():
            presentation.Slides(int(number)).Export(
                output, "PNG", request["width"], round(request["width"] * ratio)
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
