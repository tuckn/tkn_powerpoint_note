"""CLI-only rendering; pipeline and saved artifact contracts stay unchanged."""

from __future__ import annotations

import json
from collections.abc import Iterator
from typing import Any


def text_value(value: Any) -> str:
    if not isinstance(value, str):
        value = json.dumps(value, ensure_ascii=False)
    return "".join(
        {"\n": "\\n", "\r": "\\r", "\t": "\\t"}.get(char, f"\\u{ord(char):04x}")
        if ord(char) < 32 or ord(char) == 127
        else char
        for char in value
    )


def result_lines(value: Any, prefix: str = "") -> Iterator[str]:
    if isinstance(value, dict) and value:
        for key, item in value.items():
            yield from result_lines(item, f"{prefix}.{key}" if prefix else str(key))
    elif isinstance(value, list) and value:
        for index, item in enumerate(value):
            yield from result_lines(item, f"{prefix}[{index}]")
    else:
        yield f"{prefix}={text_value(value)}"


def emit(result: dict[str, Any], *, json_output: bool) -> None:
    if json_output:
        print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    else:
        print("\n".join(result_lines(result)))
