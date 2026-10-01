from pathlib import Path

import pytest

from powerpoint_note.io import atomic_write


def test_atomic_failure_preserves_old_output(tmp_path, monkeypatch):
    path = tmp_path / "note.md"
    path.write_text("old", encoding="utf-8")
    original = Path.replace

    def fail(self, target):
        if target == path:
            raise OSError("Publication failed")
        return original(self, target)

    monkeypatch.setattr(Path, "replace", fail)
    with pytest.raises(OSError):
        atomic_write(path, "new")
    assert path.read_text("utf-8") == "old"
    assert list(tmp_path.iterdir()) == [path]


def test_create_never_overwrites_racing_file(tmp_path, monkeypatch):
    import os

    path = tmp_path / "config.yaml"
    original = os.link

    def race(source, destination):
        path.write_text("human edit", encoding="utf-8")
        return original(source, destination)

    monkeypatch.setattr(os, "link", race)
    with pytest.raises(FileExistsError):
        atomic_write(path, "generated", overwrite=False)
    assert path.read_text("utf-8") == "human edit"
    assert list(tmp_path.iterdir()) == [path]
