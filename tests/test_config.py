from __future__ import annotations

from pathlib import Path

import pytest

from powerpoint_note.config import example, initialize, resolve
from powerpoint_note.models import NoteError


def save(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def test_layers_and_winners(tmp_path):
    home, cwd = tmp_path / "home", tmp_path / "cwd"
    g = save(
        home / ".tkn/powerpoint_note/config.yaml",
        'schema_version: "2.0.1"\ngeneration:\n  prompt_profile: default-en\n  max_slides: 2\n',
    )
    save(cwd / ".tkn/config.yaml", 'schema_version: "2.0.0"\ngeneration:\n  max_slides: 3\n')
    explicit = save(
        tmp_path / "extra.yaml", 'schema_version: "2.0.9"\ngeneration:\n  max_slides: 4\n'
    )
    config, report = resolve(explicit, {"generation": {"max_slides": 5}}, home=home, cwd=cwd)
    assert config["generation"]["max_slides"] == 5
    assert config["generation"]["prompt_profile"] == "default-en"
    assert report["winning_sources"]["generation.prompt_profile"] == str(g.resolve())
    assert report["winning_sources"]["generation.max_slides"] == "CLI"
    assert len(report["sources"]) == 3
    assert config["schema_version"] == "2.0.0"


@pytest.mark.parametrize(
    "text",
    [
        "{}",
        "schema_version: 1.0",
        "schema_version: '3.0.0'",
        "schema_version: '2.1.0'",
        "schema_version: '0.9.0'",
        "schema_version: '2.0.0'\nsources: {}",
        "schema_version: '2.0.0'\nselection:\n  include_hidden: yesplease",
        "schema_version: '2.0.0'\ngeneration:\n  max_slides: true",
        "schema_version: '2.0.0'\ngeneration:\n  max_slides: 0",
        "schema_version: '2.0.0'\ngeneration:\n  language: xx",
        "schema_version: '2.0.0'\nselection:\n  sections: [2]",
    ],
)
def test_invalid_config(tmp_path, text):
    p = save(tmp_path / "bad.yaml", text)
    with pytest.raises(NoteError):
        resolve(p, home=tmp_path / "home", cwd=tmp_path / "cwd")


def test_each_layer_validated_before_merge(tmp_path):
    save(
        tmp_path / "home/.tkn/powerpoint_note/config.yaml",
        'schema_version: "2.0.0"\ngeneration:\n  max_slides: invalid',
    )
    explicit = save(tmp_path / "good.yaml", 'schema_version: "2.0.0"\ngeneration:\n  max_slides: 2')
    with pytest.raises(NoteError):
        resolve(explicit, home=tmp_path / "home", cwd=tmp_path / "cwd")


def test_init_dry_run_idempotency_and_edit_protection(tmp_path):
    path = tmp_path / "absent/config.yaml"
    assert initialize(path, True)["status"] == "created"
    assert not path.parent.exists()
    assert initialize(path, False)["status"] == "created"
    assert path.read_text("utf-8").startswith('schema_version: "2.0.0"')
    assert initialize(path, False)["status"] == "unchanged"
    path.write_text(example() + "# my setting\n", encoding="utf-8")
    with pytest.raises(NoteError):
        initialize(path, False)


def test_explicit_missing_file(tmp_path):
    with pytest.raises(NoteError):
        resolve(tmp_path / "missing.yaml", home=tmp_path, cwd=tmp_path)
