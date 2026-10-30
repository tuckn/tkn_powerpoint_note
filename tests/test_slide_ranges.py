from __future__ import annotations

import pytest

from powerpoint_note.frontmatter import dump, normalize
from powerpoint_note.models import NoteError
from powerpoint_note.selection import format_slides


@pytest.mark.parametrize(
    "value, expected",
    [
        ([3, 4, 5, 6], "3-6"),
        ([1, 2, *range(7, 401)], "1-2,7-400"),
        ([9, 3, 4, 4, 1, 7, 8], "1,3-4,7-9"),
        ("3 - 6, 1-2, 5, 10", "1-6,10"),
        ([], ""),
        ("", ""),
        ([4], "4"),
        ("1-1000000000", "1-1000000000"),
    ],
)
def test_compact_slide_ranges(value, expected):
    assert format_slides(value) == expected


@pytest.mark.parametrize("value", [[True], ["3"], [0], "6-3", "0", "1,,3", "all"])
def test_invalid_ranges_are_not_silently_changed(value):
    with pytest.raises(NoteError):
        format_slides(value)


def test_legacy_slide_properties_become_scalar_ranges():
    result = normalize(
        {
            "selectedSlides": [3, 4, 5, 6],
            "contextAnalyzedSlides": [3, 4, 5, 6],
            "contextOmittedSlides": [1, 2, *range(7, 401)],
        }
    )
    assert result["selectedSlides"] == result["contextAnalyzedSlides"] == "3-6"
    assert result["contextOmittedSlides"] == "1-2,7-400"
    assert "selectedSlides: 3-6" in dump(result)
