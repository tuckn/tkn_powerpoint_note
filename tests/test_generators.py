from __future__ import annotations

import json
import re
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
    return save(path, yaml.safe_dump({"schema_version": "3.0.0", "generation": generation}))


def test_example_and_no_config_defaults(tmp_path):
    config, details = resolve(home=tmp_path / "home", cwd=tmp_path / "cwd")
    assert config["generation"]["genai_profile"] == "codex-default"
    assert config["generation"]["prompt_profile"] == "default-ja"
    assert config["genai_profiles"]["codex-default"] == {"provider": "codex", "model": None}
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
                    "genai_profile": "user-connection",
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
            "generation": {"prompt_profile": "default-en", "genai_profile": "cli-connection"},
        },
        home=home,
        cwd=cwd,
    )
    assert config["generation"]["prompt_profile"] == "default-en"
    assert config["generation"]["genai_profile"] == "cli-connection"
    assert details["winning_sources"]["generation.prompt_profile"] == "CLI"
    assert details["selected_generator"] == "general"


def test_partial_generator_inherits_shared_settings_and_cli_selects(tmp_path):
    path = settings(
        tmp_path / "settings.yaml",
        {
            "default_generator": "japanese",
            "genai_profile": "shared-connection",
            "generators": {
                "japanese": {"prompt_profile": "default-ja"},
                "english": {"prompt_profile": "default-en"},
            },
        },
    )
    config, details = resolve(path, generator="english", home=tmp_path / "home", cwd=tmp_path)
    assert config["generation"]["genai_profile"] == "shared-connection"
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
        {"generators": {"unused": {"genai_profile": " "}}},
        {"generators": {"unused": {"bridge_profile": "codex-default"}}},
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
        "--json",
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
        "--json",
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


def test_genai_profiles_merge_by_name_and_field_with_file_relative_paths(tmp_path):
    home, cwd = tmp_path / "home", tmp_path / "cwd"
    user = save(
        home / ".tkn/powerpoint_note/config.yaml",
        yaml.safe_dump(
            {
                "schema_version": "3.0.0",
                "genai_profiles": {
                    "custom": {
                        "provider": "claude-code",
                        "model": "base",
                        "cli": {"executable": "./tools/claude.exe"},
                    },
                    "named": {"provider": "github-copilot", "cli": {"executable": "copilot"}},
                },
            }
        ),
    )
    explicit = save(
        tmp_path / "other/config.yaml",
        yaml.safe_dump(
            {
                "schema_version": "3.0.0",
                "genai_profiles": {"custom": {"model": "override"}},
                "generation": {"genai_profile": "custom"},
            }
        ),
    )
    config, details = resolve(explicit, home=home, cwd=cwd)
    custom = config["genai_profiles"]["custom"]
    assert custom["provider"] == "claude-code" and custom["model"] == "override"
    assert custom["cli"]["executable"] == str((user.parent / "tools/claude.exe").resolve())
    assert config["genai_profiles"]["named"]["cli"]["executable"] == "copilot"
    assert config["genai_profiles"]["codex-default"]["provider"] == "codex"
    assert config["generation"]["genai_profile"] == "custom"
    winners = details["winning_sources"]
    assert winners["genai_profiles.custom.model"] == str(explicit.resolve())
    assert winners["genai_profiles.custom.cli.executable"] == str(user.resolve())
    assert winners["genai_profiles.codex-default.provider"] == "built-in"


def test_invalid_genai_profile_layer_is_not_hidden_by_later_values(tmp_path):
    save(
        tmp_path / "home/.tkn/powerpoint_note/config.yaml",
        yaml.safe_dump(
            {"schema_version": "3.0.0", "genai_profiles": {"custom": {"timeout_seconds": "60"}}}
        ),
    )
    explicit = save(
        tmp_path / "good.yaml",
        yaml.safe_dump(
            {"schema_version": "3.0.0", "genai_profiles": {"custom": {"timeout_seconds": 60}}}
        ),
    )
    with pytest.raises(NoteError, match="timeout_seconds"):
        resolve(explicit, home=tmp_path / "home", cwd=tmp_path / "cwd")


def test_old_bridge_settings_fail_with_migration_steps(tmp_path):
    path = save(
        tmp_path / "old.yaml",
        yaml.safe_dump(
            {
                "schema_version": "2.1.0",
                "generation": {"generators": {"named": {"bridge_profile": "codex-default"}}},
            }
        ),
    )
    before = path.read_bytes()
    with pytest.raises(NoteError, match="rename bridge_profile to genai_profile"):
        resolve(path, home=tmp_path / "home", cwd=tmp_path / "cwd")
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    ("generation", "where"),
    [
        ({"bridge_profile": "codex-default"}, "generation.bridge_profile"),
        (
            {"generators": {"named": {"bridge_profile": "codex-default"}}},
            "generation.generators.named.bridge_profile",
        ),
    ],
)
def test_bridge_profile_in_current_schema_names_its_replacement(tmp_path, generation, where):
    path = save(
        tmp_path / "new.yaml",
        yaml.safe_dump({"schema_version": "3.0.0", "generation": generation}),
    )
    with pytest.raises(NoteError, match=f"{re.escape(where)} was renamed to genai_profile"):
        resolve(path, home=tmp_path / "home", cwd=tmp_path / "cwd")


def test_unknown_override_points_to_genai_profiles(tmp_path):
    path = save(
        tmp_path / "override.yaml",
        yaml.safe_dump({"schema_version": "3.0.0", "generation": {"overrides": {"provider": "x"}}}),
    )
    with pytest.raises(NoteError, match="unknown key provider; define other connection settings"):
        resolve(path, home=tmp_path / "home", cwd=tmp_path / "cwd")


PRICING = {
    "currency": "JPY",
    "pricing_date": "2026-09-17",
    "input_per_million": 300.0,
    "output_per_million": 1900.0,
}


def test_nested_pricing_merges_by_field_and_is_checked_after_merging(tmp_path):
    home, cwd = tmp_path / "home", tmp_path / "cwd"

    def layer(path, connection):
        document = {"schema_version": "3.0.0", "genai_profiles": {"azure": connection}}
        return save(path, yaml.safe_dump(document))

    layer(
        home / ".tkn/powerpoint_note/config.yaml",
        {
            "provider": "azure-openai",
            "model": "m",
            "azure": {"endpoint": "https://example.invalid/openai/v1"},
            "pricing": {"m": PRICING},
        },
    )
    explicit = layer(tmp_path / "price.yaml", {"pricing": {"m": {"input_per_million": 318.64}}})
    config, _ = resolve(explicit, home=home, cwd=cwd)
    assert config["genai_profiles"]["azure"]["pricing"]["m"] == {
        **PRICING,
        "input_per_million": 318.64,
    }
    # A partial entry must be complete after merging; a wrong type fails in its own layer.
    incomplete = layer(tmp_path / "part.yaml", {"pricing": {"other": {"input_per_million": 1.0}}})
    with pytest.raises(NoteError, match="genai_profiles.azure: invalid connection settings"):
        resolve(incomplete, home=home, cwd=cwd)
    wrong = layer(tmp_path / "wrong.yaml", {"pricing": {"m": {"input_per_million": "318"}}})
    with pytest.raises(NoteError, match="genai_profiles.azure: invalid value for pricing"):
        resolve(wrong, home=home, cwd=cwd)
