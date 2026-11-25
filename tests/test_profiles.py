from __future__ import annotations

import json
import shutil
from importlib.resources import files
from pathlib import Path

import pytest

from powerpoint_note.context_profiles import load_profile, render_template
from powerpoint_note.models import NoteError


def copy_profile(tmp_path: Path, name: str = "custom") -> Path:
    packaged = Path(str(files("powerpoint_note").joinpath("context_profiles", "default-ja")))
    target = tmp_path / "profiles" / name
    shutil.copytree(packaged, target)
    return target


def settings(tmp_path: Path, name: str = "custom") -> dict:
    return {"prompt_profile": name, "prompt_profile_dirs": [str(tmp_path / "profiles")]}


@pytest.mark.parametrize(
    ("name", "language"), [("default-ja", "Japanese"), ("default-en", "English")]
)
def test_built_in_profiles_are_complete(name, language):
    profile = load_profile({"prompt_profile": name, "prompt_profile_dirs": []})
    assert (profile.version, profile.language) == ("2.0.0", language)
    assert language in profile.prompts["slide"] and "{{" not in profile.prompts["deck"]
    assert "slide_number" not in profile.schemas["slide"]["properties"]
    assert profile.source.startswith("package:")
    assert set(profile.resource_hashes) == {
        "prompt.md",
        "output.schema.json",
        "template.md",
        "deck-prompt.md",
        "deck-output.schema.json",
        "deck-template.md",
        "note-template.md",
        "source-template.md",
    }


def test_custom_profile_directory_wins_and_changes_provenance(tmp_path):
    folder = copy_profile(tmp_path)
    built_in = load_profile({"prompt_profile": "default-ja", "prompt_profile_dirs": []})
    custom = load_profile(settings(tmp_path))
    assert custom.source == str(folder.resolve())
    assert custom.provenance()["sha256"] == built_in.provenance()["sha256"]
    prompt = folder / "prompt.md"
    prompt.write_text(prompt.read_text("utf-8") + "Be brief.\n", encoding="utf-8")
    assert (
        load_profile(settings(tmp_path)).provenance()["sha256"] != built_in.provenance()["sha256"]
    )


def edit(folder: Path, name: str, old: str, new: str) -> None:
    path = folder / name
    text = path.read_text("utf-8")
    assert old in text
    path.write_text(text.replace(old, new), encoding="utf-8")


@pytest.mark.parametrize(
    ("name", "old", "new", "message"),
    [
        ("template.md", "version: 2.0.0", "version: 1.0.0", "unsupported template version"),
        ("template.md", "  stale:", "  renamed:", "labels are missing"),
        ("template.md", "{{summary}}", "{{summary}}{{summary}}", "each field must occur once"),
        ("source-template.md", "{{content}}", "{{unknown}}", "unknown template field"),
        ("note-template.md", "{{#overview}}", "", "unmatched template conditional"),
        ("prompt.md", "{{language}}", "the language", "expected one language placeholder"),
        ("template.md", "{{summary}}", "{{summary}}\n<!-- powerpoint-note:end -->", "reserved"),
    ],
)
def test_broken_profiles_fail_before_any_processing(tmp_path, name, old, new, message):
    edit(copy_profile(tmp_path), name, old, new)
    with pytest.raises(NoteError, match=message):
        load_profile(settings(tmp_path))


def test_schema_must_match_what_the_renderer_reads(tmp_path):
    folder = copy_profile(tmp_path)
    schema = json.loads((folder / "output.schema.json").read_text("utf-8"))
    schema["properties"]["extra"] = {"type": "string"}
    schema["required"].append("extra")
    (folder / "output.schema.json").write_text(json.dumps(schema), encoding="utf-8")
    with pytest.raises(NoteError, match="renderer contract"):
        load_profile(settings(tmp_path))
    (folder / "deck-prompt.md").unlink()
    with pytest.raises(NoteError, match="Incomplete"):
        load_profile(settings(tmp_path))
    with pytest.raises(NoteError, match="Unknown or invalid"):
        load_profile(settings(tmp_path, "absent"))
    with pytest.raises(NoteError, match="Invalid generation.prompt_profile"):
        load_profile(settings(tmp_path, "../escape"))


def test_templates_substitute_only_their_own_tokens():
    template = "A {{title}}\n{{#notes}}\nN: {{notes}}\n{{/notes}}\nend"
    assert render_template(template, {"title": "{{notes}}", "notes": ""}) == "A {{notes}}\n\nend\n"
    assert render_template(template, {"title": "T", "notes": "x"}) == "A T\n\nN: x\n\nend\n"
