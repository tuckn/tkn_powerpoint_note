from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
from unittest.mock import Mock

import pytest
from conftest import make_deck

from powerpoint_note import __version__, cli
from powerpoint_note.logging_utils import SUCCESS, ColorFormatter, configure, supports_color
from powerpoint_note.output import result_lines

SLIDES = [(256, "One", "first body"), (257, "Two", "second body")]


def run_cli(*args, cwd):
    # Keep user-level app and Bridge configuration files out of subprocess tests.
    home = str(cwd / ".test-home")
    return subprocess.run(
        [sys.executable, "-m", "powerpoint_note", *map(str, args)],
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env={**os.environ, "USERPROFILE": home, "HOME": home},
    )


def test_help_version_and_export_end_to_end(tmp_path):
    help_result = run_cli("--help", cwd=tmp_path)
    assert help_result.returncode == 0
    for command in ("export", "pull", "status", "presentation", "config"):
        assert command in help_result.stdout
    version = run_cli("--version", cwd=tmp_path)
    assert version.returncode == 0 and version.stdout.strip() == __version__
    deck = make_deck(tmp_path / "deck.pptx", SLIDES)
    built = run_cli("export", deck, cwd=tmp_path)
    assert built.returncode == 0, built.stderr
    lines = built.stdout.splitlines()
    assert "status=success" in lines and "exported=1" in lines and "command=export" in lines
    assert f"results[0].outputPath={tmp_path / 'deck.pptx.md'}" in lines
    assert "[SUCCESS] success" in built.stderr
    again = run_cli("--json", "export", deck, cwd=tmp_path)
    assert again.returncode == 1
    payload = json.loads(again.stdout)
    assert payload["status"] == "error" and payload["results"][0]["status"] == "export-error"


def test_pull_status_and_listing_through_the_command_line(source, tmp_path):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    config = source["config_path"]
    status = run_cli("status", "--source", "demo", "--config", config, "--json", cwd=tmp_path)
    assert status.returncode == 0, status.stderr
    assert json.loads(status.stdout)["statusCounts"] == {"would-create": 1}
    assert not source["notes"].exists()
    pulled = run_cli("--config", config, "pull", "--all-sources", "--json", cwd=tmp_path)
    assert pulled.returncode == 0, pulled.stderr
    result = json.loads(pulled.stdout)
    assert (result["command"], result["sources"], result["changed"]) == ("pull", ["demo"], 1)
    assert result["results"][0]["notePath"] == str(source["notes"] / "deck.pptx.md")
    one = run_cli("pull", "--source", "demo", "deck.pptx", "--config", config, cwd=tmp_path)
    assert one.returncode == 0 and "statusCounts.unchanged=1" in one.stdout.splitlines()
    listed = run_cli(
        "presentation",
        "list-slides",
        "--presentation",
        source["decks"] / "deck.pptx",
        "--json",
        cwd=tmp_path,
    )
    slides = json.loads(listed.stdout)
    assert slides["totalSlides"] == 2 and slides["slides"][1]["title"] == "Two"
    shown = run_cli("config", "list", "--config", config, "--source", "demo", cwd=tmp_path)
    assert shown.returncode == 0
    assert f"config.sources.demo.notes.dir={source['notes']}" in shown.stdout.splitlines()
    assert "[INFO] Showing resolved configuration" in shown.stderr


@pytest.mark.parametrize(
    ("args", "status", "code"),
    [
        (["export"], "argument-error", 2),
        (["nonsense"], "argument-error", 2),
        (["pull"], "argument-error", 2),
        (["pull", "--source", "a", "--all-sources"], "argument-error", 2),
        (["export", "x.pptx", "--slides", "1"], "argument-error", 2),
        (["-q", "-v", "config", "list"], "argument-error", 2),
        (["export", "absent.pptx"], "config-error", 3),
        (["pull", "--source", "absent"], "config-error", 3),
        (["pull", "--all-sources"], "config-error", 3),
        (["config", "list", "--config", "absent.yaml"], "config-error", 3),
        (["presentation", "list-slides", "--presentation", "absent.pptx"], "config-error", 3),
    ],
)
def test_failures_keep_the_format_and_use_distinct_exit_codes(
    tmp_path, monkeypatch, capsys, args, status, code
):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("USERPROFILE", str(tmp_path / "home"))
    monkeypatch.setenv("HOME", str(tmp_path / "home"))
    assert cli.main([*args, "--json"]) == code
    result = json.loads(capsys.readouterr().out)
    assert result["status"] == status and result["message"]
    assert cli.main(args) == code
    assert f"status={status}" in capsys.readouterr().out.splitlines()


def test_interrupt_unexpected_errors_and_targets(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(
        cli, "resolve", lambda *args, **kwargs: (_ for _ in ()).throw(KeyboardInterrupt())
    )
    assert cli.main(["--json", "config", "list"]) == 130
    assert json.loads(capsys.readouterr().out) == {
        "status": "interrupted",
        "command": "config list",
        "message": "Interrupted",
    }

    def broken(*args, **kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr(cli, "resolve", broken)
    assert cli.main(["export", str(tmp_path / "deck.pptx"), "--json"]) == 1
    result = json.loads(capsys.readouterr().out)
    assert result == {
        "status": "error",
        "command": "export",
        "message": "boom",
        "inputPath": str(tmp_path / "deck.pptx"),
    }


def test_config_init_through_the_command_line(tmp_path):
    target = tmp_path / "config.yaml"
    created = run_cli("config", "init", "--output", target, cwd=tmp_path)
    assert created.returncode == 0 and f"configPath={target}" in created.stdout.splitlines()
    target.write_text("schema_version: '4.0.0'\n", encoding="utf-8")
    edited = run_cli("config", "init", "--output", target, "--json", cwd=tmp_path)
    assert edited.returncode == 3 and json.loads(edited.stdout)["status"] == "config-error"


def test_key_value_lines_escape_only_control_characters():
    lines = list(
        result_lines(
            {"path": "C:\\a\\b", "items": [], "nested": {"ok": True, "none": None}, "text": "a\nb"}
        )
    )
    assert lines == [
        "path=C:\\a\\b",
        "items=[]",
        "nested.ok=true",
        "nested.none=null",
        "text=a\\nb",
    ]


def test_log_format_color_and_levels(monkeypatch):
    stream = Mock()
    stream.isatty.return_value = False
    assert not supports_color(stream)
    record = logging.LogRecord("x", SUCCESS, "", 0, "done", (), None)
    assert ColorFormatter(False).format(record) == "[SUCCESS] done"
    assert ColorFormatter(True).format(record) == "\x1b[32m[SUCCESS] done\x1b[0m"
    assert configure(quiet=True).level == logging.ERROR
    assert configure(verbose=True).level == logging.DEBUG
    logger = configure(no_color=True)
    assert logger.level == logging.INFO and not logger.handlers[0].formatter.color
    monkeypatch.setenv("NO_COLOR", "")
    stream.isatty.return_value = True
    assert not supports_color(stream)


def test_all_sources_reports_an_unreachable_source_and_continues(source, tmp_path):
    make_deck(source["decks"] / "deck.pptx", SLIDES)
    config = source["config_path"]
    extra = (
        "  absent:\n    presentations:\n"
        f"      dir: '{tmp_path / 'nowhere'}'\n"
        f"    notes:\n      dir: '{tmp_path / 'notes2'}'\n"
    )
    text = config.read_text("utf-8").replace("genai_profiles:", extra + "genai_profiles:")
    config.write_text(text, encoding="utf-8")
    pulled = run_cli("pull", "--all-sources", "--config", config, "--json", cwd=tmp_path)
    assert pulled.returncode == 1, pulled.stderr
    result = json.loads(pulled.stdout)
    assert result["statusCounts"] == {"created": 1, "error": 1}
    failed = next(r for r in result["results"] if r["status"] == "error")
    assert failed["sourceRoot"] == "absent" and "does not exist" in failed["message"]
    single = run_cli("pull", "--source", "absent", "--config", config, "--json", cwd=tmp_path)
    assert single.returncode == 3
