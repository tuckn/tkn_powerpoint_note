"""Human-readable stderr logs with conservative Windows ANSI support."""

from __future__ import annotations

import logging
import os
import sys
from typing import Any

SUCCESS = 25
logging.addLevelName(SUCCESS, "SUCCESS")


def supports_color(stream: Any) -> bool:
    try:
        if "NO_COLOR" in os.environ or os.environ.get("TERM") == "dumb" or not stream.isatty():
            return False
        if os.name == "nt":
            import ctypes
            import msvcrt

            handle = ctypes.c_void_p(msvcrt.get_osfhandle(stream.fileno()))
            mode = ctypes.c_ulong()
            kernel = ctypes.windll.kernel32
            return bool(
                kernel.GetConsoleMode(handle, ctypes.byref(mode))
                and kernel.SetConsoleMode(handle, mode.value | 4)
            )
        return True
    except (AttributeError, OSError, ValueError):
        return False


class ColorFormatter(logging.Formatter):
    def __init__(self, color: bool):
        super().__init__("[%(levelname)s] %(message)s")
        self.color = color

    def format(self, record: logging.LogRecord) -> str:
        value = super().format(record)
        color = (
            {SUCCESS: "\x1b[32m", logging.ERROR: "\x1b[31m", logging.CRITICAL: "\x1b[31m"}.get(
                record.levelno
            )
            if self.color
            else None
        )
        return f"{color}{value}\x1b[0m" if color else value


def configure(quiet: bool = False, verbose: bool = False) -> logging.Logger:
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(ColorFormatter(supports_color(sys.stderr)))
    logger = logging.getLogger("powerpoint_note")
    logger.handlers.clear()
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG if verbose else logging.ERROR if quiet else logging.INFO)
    logger.propagate = False
    return logger
