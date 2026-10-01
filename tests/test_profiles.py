from __future__ import annotations

import json
from importlib.resources import files
from pathlib import Path

import pytest

from powerpoint_note import pipeline
from powerpoint_note.config import resolve
from powerpoint_note.context_profiles import RESOURCES, load_profile, render_template
from powerpoint_note.markdown import split_note
from powerpoint_note.models import NoteError


def copy_profile(root: Path, name: str = "default-ja") -> Path:
    target = root / name
    target.mkdir(parents=True)
    bundled = files("powerpoint_note").joinpath("context_profiles", name)
    for filename in RESOURCES:
        (target / filename).write_text(
            bundled.joinpath(filename).read_text("utf-8"), encoding="utf-8"
        )
    return target


@pytest.mark.parametrize(
    "name,language,heading",
    [
        ("default-ja", "Japanese", "抽出内容"),
        ("default-en", "English", "Extracted content"),
    ],
)
def test_profile_controls_prompts_and_local_note(
    config, deck_path, tmp_path, name, language, heading
):
    config["generation"]["prompt_profile"] = name
    profile = load_profile(config["generation"])
    assert profile.language == language
    assert all(
        language in prompt and "{{language}}" not in prompt for prompt in profile.prompts.values()
    )
    output = tmp_path / (name + ".md")
    result = pipeline.build(deck_path, output, config)
    metadata = split_note(output.read_text("utf-8"))[0]
    assert heading in output.read_text("utf-8")
    assert metadata["promptProfile"] == profile.provenance() == result["prompt_profile"]
    assert len(metadata["promptProfile"]["resources"]) == len(RESOURCES)
    assert pipeline.verify(output)["status"] == "verified"


@pytest.mark.parametrize("name", ["missing", "../default-ja", "/tmp", "C:\\tmp", ".", "Default-ja"])
def test_unknown_or_unsafe_profile_fails(config, name):
    config["generation"]["prompt_profile"] = name
    with pytest.raises(NoteError):
        load_profile(config["generation"])


def test_custom_profile_layer_resolution_and_no_partial_fallback(tmp_path):
    cwd = tmp_path / "cwd"
    profile_dir = copy_profile(cwd / "profiles")
    config_path = tmp_path / "settings.yaml"
    config_path.write_text(
        'schema_version: "2.0.0"\ngeneration:\n  profile_dirs: [profiles]\n', encoding="utf-8"
    )
    config, details = resolve(config_path, home=tmp_path / "home", cwd=cwd)
    assert config["generation"]["profile_dirs"] == [str((cwd / "profiles").resolve())]
    assert details["prompt_profile"]["source"] == str(profile_dir)
    (profile_dir / "deck-template.md").unlink()
    with pytest.raises(NoteError, match="Incomplete"):
        resolve(config_path, home=tmp_path / "home", cwd=cwd)


@pytest.mark.parametrize("filename", RESOURCES)
def test_every_selected_resource_invalidates_cache_and_preserves_notes(
    config, deck_path, tmp_path, filename
):
    profile_dir = copy_profile(tmp_path / "profiles")
    config["generation"]["profile_dirs"] = [str(profile_dir.parent)]
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    output.write_text(output.read_text("utf-8") + "\nPersonal notes.\n", encoding="utf-8")
    before = split_note(output.read_text("utf-8"))[0]["buildKey"]
    path = profile_dir / filename
    content = path.read_text("utf-8")
    if filename.endswith(".json"):
        schema = json.loads(content)
        schema["description"] = "Profile revision"
        content = json.dumps(schema)
    else:
        content += "\nProfile revision.\n"
    path.write_text(content, encoding="utf-8")
    assert pipeline.build(deck_path, output, config)["status"] == "updated"
    assert split_note(output.read_text("utf-8"))[0]["buildKey"] != before
    assert output.read_text("utf-8").endswith("\nPersonal notes.\n")
    assert pipeline.build(deck_path, output, config)["status"] == "unchanged"


def test_unselected_profile_does_not_invalidate_cache(config, deck_path, tmp_path):
    ja = copy_profile(tmp_path / "profiles")
    en = copy_profile(tmp_path / "profiles", "default-en")
    config["generation"]["profile_dirs"] = [str(ja.parent)]
    output = tmp_path / "note.md"
    pipeline.build(deck_path, output, config)
    (en / "note-template.md").write_text("invalid unused template", encoding="utf-8")
    assert pipeline.build(deck_path, output, config)["status"] == "unchanged"


@pytest.mark.parametrize(
    "filename,before,after",
    [
        ("note-template.md", "{{mode}}", "{{unknown}}"),
        ("note-template.md", "{{mode}}", ""),
        ("note-template.md", "{{mode}}", "{{mode}}\n{{mode}}"),
        ("note-template.md", "{{/overview}}", "{{/mode}}"),
        ("note-template.md", "{{overview}}", "{{#mode}}{{overview}}{{/mode}}"),
        ("template.md", "version: 1.0.0", "version: 9.0.0"),
        ("template.md", "language: Japanese", "language: null"),
        ("template.md", "  slide: スライド", "  slide: 10"),
        ("template.md", "{{summary}}", "<!-- powerpoint-note:begin -->\n{{summary}}"),
        ("prompt.md", "{{language}}", "{{lang}}"),
        ("output.schema.json", '"integer"', '"string"'),
        ("output.schema.json", '"type": "integer"', '"$ref": "https://invalid.example/schema"'),
        ("output.schema.json", '{\n      "type": "integer"\n    }', "true"),
    ],
)
def test_invalid_bundle_stops_before_office_ai_or_output(
    config, deck_path, tmp_path, monkeypatch, filename, before, after
):
    profile_dir = copy_profile(tmp_path / "profiles")
    config["generation"]["profile_dirs"] = [str(profile_dir.parent)]
    path = profile_dir / filename
    text = path.read_text("utf-8")
    assert before in text
    path.write_text(text.replace(before, after), encoding="utf-8")

    def forbidden(*args, **kwargs):
        pytest.fail("Office or AI setup must not occur for an invalid profile")

    monkeypatch.setattr(pipeline, "preflight", forbidden)
    monkeypatch.setattr(pipeline, "Generator", forbidden)
    output = tmp_path / "not-created/note.md"
    for context in (False, True):
        with pytest.raises(NoteError):
            pipeline.build(deck_path, output, config, context=context, dry_run=True)
    assert not output.parent.exists()


def test_template_does_not_interpret_source_tokens():
    template = "{{summary}}\n{{#uncertainties}}\nProblems: {{uncertainties}}\n{{/uncertainties}}"
    source = "Literal {{uncertainties}} and {{#summary}} in source text"
    assert render_template(template, {"summary": source, "uncertainties": ""}) == source + "\n"


def test_profile_change_cannot_overwrite_reviewed_content(config, deck_path, tmp_path):
    output = tmp_path / "reviewed.md"
    pipeline.build(deck_path, output, config)
    text = output.read_text("utf-8").replace("reviewStatus: unreviewed", "reviewStatus: reviewed")
    output.write_text(text, encoding="utf-8")
    config["generation"]["prompt_profile"] = "default-en"
    with pytest.raises(NoteError, match="Reviewed"):
        pipeline.build(deck_path, output, config)
    assert output.read_text("utf-8") == text


def test_old_language_setting_has_actionable_migration_error(tmp_path):
    path = tmp_path / "old.yaml"
    text = 'schema_version: "1.0.0"\ngeneration:\n  language: en\n'
    path.write_text(text, encoding="utf-8")
    with pytest.raises(NoteError, match="generation.prompt_profile default-ja/default-en"):
        resolve(path, home=tmp_path, cwd=tmp_path)
    assert path.read_text("utf-8") == text
