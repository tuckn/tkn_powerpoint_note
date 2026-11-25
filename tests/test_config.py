from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from powerpoint_note.config import example, initialize, resolve
from powerpoint_note.models import ConfigError


def write(path: Path, data: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump({"schema_version": "4.0.0", **data}), encoding="utf-8")
    return path


def source(root: Path, **extra) -> dict:
    return {
        "presentations": {"dir": str(root / "decks")},
        "notes": {"dir": str(root / "notes")},
        **extra,
    }


def test_built_in_defaults_need_no_file(tmp_path):
    config, details = resolve(home=tmp_path / "home", cwd=tmp_path / "cwd")
    assert config["sources"] == {}
    assert config["generation"]["limits"] == {
        "max_slides": 10,
        "max_input_chars": 120000,
        "image_width": 2400,
        "max_render_seconds": 600,
    }
    assert config["generation"]["generator_id"] is None
    assert config["generation"]["prompt_profile"] == "default-ja"
    assert sorted(config["genai_profiles"]) == [
        "antigravity-default",
        "claude-default",
        "codex-default",
        "copilot-default",
    ]
    assert details["loaded_sources"] == []
    assert details["winning_sources"]["generation.limits.max_slides"] == "built-in"


def test_layers_merge_by_field_and_report_where_values_came_from(tmp_path):
    home, cwd = tmp_path / "home", tmp_path / "cwd"
    user = write(
        home / ".tkn/powerpoint_note/config.yaml",
        {"generation": {"limits": {"max_slides": 20, "image_width": 1200}}},
    )
    local = write(cwd / ".tkn/config.yaml", {"generation": {"limits": {"max_slides": 30}}})
    explicit = write(tmp_path / "explicit.yaml", {"extraction": {"include_hidden": True}})
    config, details = resolve(
        explicit, {"generation": {"limits": {"max_input_chars": 5}}}, home=home, cwd=cwd
    )
    limits = config["generation"]["limits"]
    assert (limits["max_slides"], limits["image_width"], limits["max_input_chars"]) == (30, 1200, 5)
    assert config["extraction"]["include_hidden"] is True
    origins = details["winning_sources"]
    assert origins["generation.limits.max_slides"] == str(local.resolve())
    assert origins["generation.limits.image_width"] == str(user.resolve())
    assert origins["generation.limits.max_input_chars"] == "CLI"
    assert [item["path"] for item in details["loaded_sources"]] == [
        str(user.resolve()),
        str(local.resolve()),
        str(explicit.resolve()),
    ]


def test_sources_get_defaults_and_a_later_file_replaces_the_whole_list(tmp_path):
    home, cwd = tmp_path / "home", tmp_path / "cwd"
    cwd.mkdir()
    write(home / ".tkn/powerpoint_note/config.yaml", {"sources": {"old": source(tmp_path)}})
    explicit = write(
        tmp_path / "explicit.yaml",
        {"sources": {"new": {"presentations": {"dir": "decks"}, "notes": {"dir": "notes"}}}},
    )
    config, _ = resolve(explicit, home=home, cwd=cwd)
    assert list(config["sources"]) == ["new"]
    spec = config["sources"]["new"]
    assert spec["presentations"] == {
        "dir": str((cwd / "decks").resolve()),
        "recursive": False,
        "include": ["*.pptx", "*.pptm"],
        "ignore": [],
    }
    assert spec["notes"]["frontmatter_term_format"] == "obsidian-link"
    with pytest.raises(ConfigError, match="Unknown source"):
        resolve(explicit, source="old", home=home, cwd=cwd)


def test_generator_comes_from_cli_then_source_then_default(tmp_path):
    generators = {
        "ja": {"prompt_profile": "default-ja"},
        "en": {"prompt_profile": "default-en", "overrides": {"model": "special"}},
    }
    path = write(
        tmp_path / "config.yaml",
        {
            "sources": {
                "plain": source(tmp_path),
                "english": source(tmp_path, generation={"generator": "en"}),
            },
            "generation": {"default_generator": "ja", "generators": generators},
        },
    )
    kwargs = {"home": tmp_path / "home", "cwd": tmp_path}
    assert resolve(path, source="plain", **kwargs)[1]["selected_generator"] == "ja"
    config, details = resolve(path, source="english", **kwargs)
    assert details["selected_generator"] == "en"
    assert details["selected_prompt_profile"] == "default-en"
    assert config["generation"]["overrides"] == {"model": "special"}
    assert (
        resolve(path, generator="ja", source="english", **kwargs)[1]["selected_generator"] == "ja"
    )
    profile = {"generation": {"prompt_profile": "default-ja"}}
    assert (
        resolve(path, profile, source="english", **kwargs)[0]["generation"]["prompt_profile"]
        == "default-ja"
    )
    with pytest.raises(ConfigError, match="Unknown generator"):
        resolve(path, generator="missing", **kwargs)


@pytest.mark.parametrize(
    ("data", "message"),
    [
        ({"selection": {"slides": "1"}}, "unknown key selection"),
        ({"generation": {"profile_dirs": []}}, "unknown key generation.profile_dirs"),
        ({"generation": {"max_slides": 5}}, "unknown key generation.max_slides"),
        ({"generation": {"limits": {"max_slides": 0}}}, "must be positive"),
        ({"generation": {"limits": {"image_width": 100}}}, "600..8000"),
        ({"generation": {"limits": {"max_slides": "10"}}}, "invalid type"),
        ({"extraction": {"off_slide": "keep"}}, "exclude or append"),
        ({"sources": {"x": {"notes": {"dir": "n"}}}}, "presentations must be a mapping"),
        ({"sources": {"x": {"presentations": {}, "notes": {"dir": "n"}}}}, "dir is required"),
        ({"sources": {"bad id": {}}}, "expected a nonempty name"),
        ({"genai_profiles": {"codex-default": {"bogus": 1}}}, "unknown key bogus"),
        ({"generation": {"overrides": {"api_key": "x"}}}, "unknown key api_key"),
    ],
)
def test_unknown_keys_and_bad_values_fail_before_processing(tmp_path, data, message):
    path = write(tmp_path / "config.yaml", data)
    with pytest.raises(ConfigError, match=message):
        resolve(path, home=tmp_path / "home", cwd=tmp_path)


def test_other_schema_versions_are_rejected_and_missing_explicit_file_fails(tmp_path):
    path = tmp_path / "old.yaml"
    path.write_text('schema_version: "3.0.0"\n', encoding="utf-8")
    with pytest.raises(ConfigError, match="unsupported schema_version"):
        resolve(path, home=tmp_path / "home", cwd=tmp_path)
    with pytest.raises(ConfigError, match="does not exist"):
        resolve(tmp_path / "missing.yaml", home=tmp_path / "home", cwd=tmp_path)


def test_config_init_writes_once_and_protects_edits(tmp_path):
    target = tmp_path / "nested" / "config.yaml"
    assert initialize(target, dry_run=True)["status"] == "created" and not target.exists()
    result = initialize(target, dry_run=False)
    assert result == {
        "status": "created",
        "command": "config init",
        "dryRun": False,
        "configPath": str(target.resolve()),
    }
    assert target.read_text("utf-8") == example()
    assert initialize(target, dry_run=False)["status"] == "unchanged"
    # The written example must itself be a valid configuration.
    config, details = resolve(target, home=tmp_path / "home", cwd=tmp_path)
    assert list(config["sources"]) == ["personal-powerpoint"]
    assert details["selected_generator"] == "my-codex-def"
    target.write_text(example() + "# mine\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="edited"):
        initialize(target, dry_run=False)
