from __future__ import annotations

import json
from pathlib import Path

import pytest
import yaml
from test_cli import run_cli
from test_config import save

from powerpoint_note import pipeline
from powerpoint_note.config import example, resolve
from powerpoint_note.markdown import split_note
from powerpoint_note.models import NoteError


def settings(path: Path, generation: dict) -> Path:
    return save(path, yaml.safe_dump({"schema_version": "2.1.0", "generation": generation}))


def test_example_and_no_config_defaults(tmp_path):
    config, details = resolve(home=tmp_path / "home", cwd=tmp_path / "cwd")
    assert config["generation"]["bridge_profile"] == "codex-default"
    assert config["generation"]["prompt_profile"] == "default-ja"
    assert details["selected_generator"] is None
    path = save(tmp_path / "example.yaml", example())
    config, details = resolve(path, home=tmp_path / "home", cwd=tmp_path / "cwd")
    assert details["selected_generator"] == "my-codex-def"
    assert config["generation"]["generator_id"] == "my-codex-def"
    assert set(config["generation"]["generators"]) == {"my-codex-def", "my-codex-en"}


def test_layered_generator_fields_and_individual_cli_precedence(tmp_path):
    home, cwd = tmp_path / "home", tmp_path / "cwd"
    user = settings(
        home / ".tkn/powerpoint_note/config.yaml",
        {
            "default_generator": "general",
            "prompt_profile": "default-en",
            "generators": {
                "general": {
                    "bridge_profile": "user-bridge",
                    "prompt_profile": "default-ja",
                    "overrides": {"timeout_seconds": 300},
                }
            },
        },
    )
    project = settings(
        cwd / ".tkn/config.yaml",
        {
            "generators": {"general": {"overrides": {"timeout_seconds": 500}}},
        },
    )
    explicit = settings(
        tmp_path / "extra.yaml",
        {
            "generators": {"general": {"overrides": {"model": "fixture-model"}}},
        },
    )
    config, details = resolve(explicit, home=home, cwd=cwd)
    assert config["generation"]["prompt_profile"] == "default-ja"
    assert details["winning_sources"]["generation.prompt_profile"] == str(user.resolve())
    assert config["generation"]["overrides"] == {"model": "fixture-model", "timeout_seconds": 500}
    assert details["winning_sources"]["generation.overrides.timeout_seconds"] == str(
        project.resolve()
    )
    assert details["winning_sources"]["generation.overrides.model"] == str(explicit.resolve())
    config, details = resolve(
        explicit,
        {
            "generation": {"prompt_profile": "default-en", "bridge_profile": "cli-bridge"},
        },
        home=home,
        cwd=cwd,
    )
    assert config["generation"]["prompt_profile"] == "default-en"
    assert config["generation"]["bridge_profile"] == "cli-bridge"
    assert details["winning_sources"]["generation.prompt_profile"] == "CLI"
    assert details["selected_generator"] == "general"


def test_partial_generator_inherits_shared_settings_and_cli_selects(tmp_path):
    path = settings(
        tmp_path / "settings.yaml",
        {
            "default_generator": "japanese",
            "bridge_profile": "shared-bridge",
            "generators": {
                "japanese": {"prompt_profile": "default-ja"},
                "english": {"prompt_profile": "default-en"},
            },
        },
    )
    config, details = resolve(path, generator="english", home=tmp_path / "home", cwd=tmp_path)
    assert config["generation"]["bridge_profile"] == "shared-bridge"
    assert config["generation"]["prompt_profile"] == "default-en"
    assert details["selected_generator"] == "english"
    assert details["winning_sources"]["generation.generator_id"] == "CLI"
    with pytest.raises(NoteError, match="Unknown generator"):
        resolve(path, generator="missing", home=tmp_path / "home", cwd=tmp_path)


@pytest.mark.parametrize(
    "generation",
    [
        {"default_generator": "missing"},
        {"default_generator": ""},
        {"default_generator": False},
        {"generators": []},
        {"generators": {"bad name": {}}},
        {"generators": {"": {}}},
        {"generators": {5: {}}},
        {"generators": {"unused": []}},
        {"generators": {"unused": {"typo": 1}}},
        {"generators": {"unused": {"bridge_profile": " "}}},
        {"generators": {"unused": {"prompt_profile": False}}},
        {"generators": {"unused": {"overrides": []}}},
        {"generators": {"unused": {"overrides": {"typo": 1}}}},
        {"generators": {"unused": {"overrides": {"model": " "}}}},
        {"generators": {"unused": {"overrides": {"timeout_seconds": False}}}},
        {"generators": {"unused": {"overrides": {"timeout_seconds": -1}}}},
        {"generators": {"unused": {"overrides": {"timeout_seconds": float("inf")}}}},
        {"generators": {"unused": {"overrides": {"max_output_tokens": 0}}}},
        {"generators": {"unused": {"overrides": {"local_only": "yes"}}}},
    ],
)
def test_invalid_generators_fail_even_when_unselected(tmp_path, generation):
    path = settings(tmp_path / "bad.yaml", generation)
    with pytest.raises(NoteError):
        resolve(path, home=tmp_path / "home", cwd=tmp_path / "cwd")


def test_invalid_generator_layer_is_not_hidden_by_later_values(tmp_path):
    settings(
        tmp_path / "home/.tkn/powerpoint_note/config.yaml",
        {
            "generators": {"named": {"overrides": {"timeout_seconds": True}}},
        },
    )
    explicit = settings(
        tmp_path / "good.yaml",
        {
            "generators": {"named": {"overrides": {"timeout_seconds": 500}}},
        },
    )
    with pytest.raises(NoteError, match="timeout_seconds"):
        resolve(explicit, home=tmp_path / "home", cwd=tmp_path / "cwd")


def test_cli_generator_changes_note_language_and_is_inspectable(deck_path, tmp_path):
    path = save(tmp_path / "settings.yaml", example())
    output = tmp_path / "english.md"
    shown = run_cli("config", "list", "--config", path, "--generator", "my-codex-en", cwd=tmp_path)
    assert shown.returncode == 0, shown.stderr
    assert "selected_generator=my-codex-en" in shown.stdout.splitlines()
    assert "config.generation.prompt_profile=default-en" in shown.stdout.splitlines()
    assert any(
        line.startswith("sources[") and line.endswith(f"].path={path.resolve()}")
        for line in shown.stdout.splitlines()
    )
    assert f"winning_sources.generation.prompt_profile={path.resolve()}" in shown.stdout.splitlines()
    machine = run_cli("config", "list", "--config", path, "--generator", "my-codex-en", "--json", cwd=tmp_path)
    assert machine.returncode == 0, machine.stderr
    report = json.loads(machine.stdout)
    assert report["selected_generator"] == "my-codex-en"
    assert report["config"]["generation"]["prompt_profile"] == "default-en"
    exported = run_cli(
        "export",
        deck_path,
        "--config",
        path,
        "--generator",
        "my-codex-en",
        "--output",
        output,
        cwd=tmp_path,
    )
    assert exported.returncode == 0, exported.stderr
    assert "Extracted content" in output.read_text("utf-8")
    assert json.loads(exported.stdout)["generation_generator"] == "my-codex-en"
    assert split_note(output.read_text("utf-8"))[0]["generationGenerator"] == "my-codex-en"
    assert pipeline.verify(output)["status"] == "verified"
    planned_output = tmp_path / "absent/planned.md"
    planned = run_cli(
        "export",
        deck_path,
        "--config",
        path,
        "--generator",
        "my-codex-en",
        "--output",
        planned_output,
        "--dry-run",
        cwd=tmp_path,
    )
    assert planned.returncode == 0, planned.stderr
    assert json.loads(planned.stdout)["ai_calls"] == 0
    assert not planned_output.parent.exists()
    bad = run_cli(
        "export",
        deck_path,
        "--config",
        path,
        "--generator",
        "missing",
        "--output",
        planned_output,
        cwd=tmp_path,
    )
    assert bad.returncode == 2
    assert not planned_output.parent.exists()


def test_generator_cache_scope_and_provenance(deck_path, tmp_path, monkeypatch):
    from test_pipeline import FakeGenerator, fake_render

    FakeGenerator.calls = []
    monkeypatch.setattr(pipeline, "Generator", FakeGenerator)
    monkeypatch.setattr(pipeline, "preflight", lambda: None)
    monkeypatch.setattr(pipeline, "render", fake_render)
    generation = {
        "default_generator": "selected",
        "generators": {
            "selected": {"prompt_profile": "default-ja", "overrides": {"model": "first"}},
            "unused": {"prompt_profile": "default-en"},
        },
    }
    path = settings(tmp_path / "settings.yaml", generation)
    config = resolve(path, home=tmp_path / "home", cwd=tmp_path)[0]
    output = tmp_path / "context.md"
    pipeline.build(deck_path, output, config, context=True)
    first_calls = len(FakeGenerator.calls)
    output.write_text(output.read_text("utf-8") + "\nMy handwritten note.\n", encoding="utf-8")
    generation["generators"]["unused"]["overrides"] = {"model": "changed-unused"}
    settings(path, generation)
    config = resolve(path, home=tmp_path / "home", cwd=tmp_path)[0]
    assert pipeline.build(deck_path, output, config, context=True)["status"] == "unchanged"
    assert len(FakeGenerator.calls) == first_calls
    generation["generators"]["selected"]["overrides"]["model"] = "second"
    settings(path, generation)
    config = resolve(path, home=tmp_path / "home", cwd=tmp_path)[0]
    result = pipeline.build(deck_path, output, config, context=True)
    assert result["status"] == "updated" and len(FakeGenerator.calls) == first_calls * 2
    assert "My handwritten note." in output.read_text("utf-8")
    metadata = split_note(output.read_text("utf-8"))[0]
    manifest = json.loads(
        (output.parent / metadata["evidencePath"] / "manifest.json").read_text("utf-8")
    )
    assert manifest["generation_generator"] == metadata["generationGenerator"] == "selected"
    tampered = output.read_text("utf-8").replace(
        "generationGenerator: selected", "generationGenerator: unused"
    )
    output.write_text(tampered, encoding="utf-8")
    with pytest.raises(NoteError, match="generation settings provenance"):
        pipeline.verify(output)
