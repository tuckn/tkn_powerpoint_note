from __future__ import annotations

import json
import logging
import subprocess
import sys
from unittest.mock import Mock

import pytest

from powerpoint_note import __version__, cli
from powerpoint_note.logging_utils import SUCCESS, ColorFormatter, configure, supports_color
from powerpoint_note.output import result_lines


def run_cli(*args, cwd):
    return subprocess.run(
        [sys.executable, "-m", "powerpoint_note", *map(str, args)],
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )


def test_entrypoint_help_version_and_e2e(deck_path, tmp_path):
    help_result = run_cli("--help", cwd=tmp_path)
    assert help_result.returncode == 0
    assert "export" in help_result.stdout and "build" not in help_result.stdout
    version = run_cli("--version", cwd=tmp_path)
    assert version.returncode == 0 and version.stdout.strip() == __version__
    output = tmp_path / "result.md"
    built = run_cli("export", deck_path, "--output", output, cwd=tmp_path)
    assert built.returncode == 0, built.stderr
    assert "status=created" in built.stdout.splitlines()
    assert f"output={output}" in built.stdout.splitlines()
    assert "[SUCCESS]" in built.stderr
    verified = run_cli("verify", output, "--json", cwd=tmp_path)
    assert verified.returncode == 0
    assert json.loads(verified.stdout)["status"] == "verified"
    viewed = run_cli("inspect", deck_path, cwd=tmp_path)
    assert viewed.returncode == 0
    assert f"source={deck_path}" in viewed.stdout.splitlines()
    assert any(line.startswith("slides[0].title=") for line in viewed.stdout.splitlines())
    assert "\x1b[" not in viewed.stdout


def test_cli_config_list_readonly_and_error(tmp_path):
    listing = run_cli("config", "list", cwd=tmp_path)
    assert listing.returncode == 0, listing.stderr
    assert listing.stderr == "[INFO] Showing resolved configuration\n"
    assert "effective_schema_version=2.1.0" in listing.stdout.splitlines()
    assert "status=success" in listing.stdout.splitlines()
    assert "config.selection.sections=[]" in listing.stdout.splitlines()
    assert any(
        line.startswith("winning_sources.generation.max_slides=")
        for line in listing.stdout.splitlines()
    )
    machine = run_cli("config", "list", "--json", cwd=tmp_path)
    assert machine.returncode == 0
    assert machine.stderr == "[INFO] Showing resolved configuration\n"
    assert json.loads(machine.stdout)["effective_schema_version"] == "2.1.0"
    quiet = run_cli("config", "list", "--quiet", cwd=tmp_path)
    assert quiet.returncode == 0 and quiet.stderr == ""
    bad = run_cli("export", "missing.pptx", cwd=tmp_path)
    assert bad.returncode == 2 and "status=failed" in bad.stdout.splitlines()
    assert f"source_path={tmp_path / 'missing.pptx'}" in bad.stdout.splitlines()
    bad_json = run_cli("--json", "export", "missing.pptx", cwd=tmp_path)
    assert bad_json.returncode == 2 and json.loads(bad_json.stdout)["status"] == "failed"
    assert "Traceback" not in bad.stderr


def test_result_lines_preserve_windows_paths_and_flatten_nested_values():
    lines = list(result_lines({"path": r"C:\Users\ExampleUser\profiles", "items": [{"active": True}], "note": "a\nb"}))
    assert lines == [
        r"path=C:\Users\ExampleUser\profiles",
        "items[0].active=true",
        r"note=a\nb",
    ]


@pytest.mark.parametrize(
    "args", [["--quiet", "--verbose", "config", "list"], ["--quiet", "config", "list", "--verbose"]]
)
def test_quiet_verbose_mutually_exclusive(tmp_path, args):
    result = run_cli(*args, cwd=tmp_path)
    assert result.returncode != 0
    assert "status=failed" in result.stdout.splitlines()


def test_removed_build_and_unsupported_sync_commands(tmp_path):
    for command in ("build", "pull", "push"):
        result = run_cli(command, "--help", cwd=tmp_path)
        assert result.returncode == 2
    old_config = run_cli("config", "show", cwd=tmp_path)
    assert old_config.returncode == 2


@pytest.mark.parametrize("args", [
    ("--json", "config", "list"),
    ("config", "--json", "list"),
    ("config", "list", "--json"),
])
def test_json_at_every_command_level(tmp_path, args):
    result = run_cli(*args, cwd=tmp_path)
    assert result.returncode == 0
    assert json.loads(result.stdout)["status"] == "success"
    assert "[INFO]" in result.stderr and "[INFO]" not in result.stdout


def test_parse_error_and_interrupt_honor_json(tmp_path, monkeypatch, capsys):
    bad = run_cli("config", "--json", "missing", cwd=tmp_path)
    assert bad.returncode == 2 and json.loads(bad.stdout)["status"] == "failed"
    monkeypatch.setattr(cli, "resolve", lambda *args, **kwargs: (_ for _ in ()).throw(KeyboardInterrupt()))
    assert cli.main(["--json", "config", "list"]) == 130
    captured = capsys.readouterr()
    assert json.loads(captured.out)["error"] == "interrupted"
    assert cli.main(["config", "list"]) == 130
    assert "error=interrupted" in capsys.readouterr().out.splitlines()


def test_runtime_error_keeps_format_and_target(tmp_path, monkeypatch, capsys):
    def broken(*args, **kwargs):
        raise RuntimeError("test failure")

    monkeypatch.setattr(cli, "resolve", broken)
    assert cli.main(["--json", "inspect", str(tmp_path / "deck.pptx")]) == 1
    result = json.loads(capsys.readouterr().out)
    assert result["status"] == "failed"
    assert result["source_path"] == str(tmp_path / "deck.pptx")


def test_config_init_dry_run_and_output_mode(tmp_path):
    target = tmp_path / "new" / "config.yaml"
    text = run_cli("config", "init", "--dry-run", "--output", target, cwd=tmp_path)
    machine = run_cli("--json", "config", "init", "--dry-run", "--output", target, cwd=tmp_path)
    assert text.returncode == machine.returncode == 0
    assert f"path={target}" in text.stdout.splitlines()
    assert json.loads(machine.stdout)["path"] == str(target)
    assert not target.exists() and not target.parent.exists()


def test_quiet_verbose_change_logs_only(tmp_path):
    normal = run_cli("config", "list", cwd=tmp_path)
    quiet = run_cli("--quiet", "config", "list", cwd=tmp_path)
    verbose = run_cli("config", "list", "--verbose", cwd=tmp_path)
    assert normal.stdout == quiet.stdout == verbose.stdout
    assert quiet.stderr == ""


def test_log_format_color_and_levels(monkeypatch):
    assert logging.getLevelName(SUCCESS) == "SUCCESS" and logging.INFO < SUCCESS < logging.WARNING
    for level, color in [
        (SUCCESS, "\x1b[32m"),
        (logging.ERROR, "\x1b[31m"),
        (logging.CRITICAL, "\x1b[31m"),
    ]:
        record = logging.LogRecord("test", level, "", 1, "done", (), None)
        assert ColorFormatter(True).format(record).startswith(color)
        assert "\x1b" not in ColorFormatter(False).format(record)
    stream = Mock()
    stream.isatty.return_value = False
    assert not supports_color(stream)
    stream.isatty.return_value = True
    monkeypatch.setenv("NO_COLOR", "")
    assert not supports_color(stream)
    monkeypatch.delenv("NO_COLOR")
    monkeypatch.setenv("TERM", "dumb")
    assert not supports_color(stream)


def test_log_filtering_and_stderr(capsys):
    logger = configure(quiet=True)
    logger.info("hidden")
    logger.error("shown")
    capture = capsys.readouterr()
    assert capture.out == "" and "shown" in capture.err and "hidden" not in capture.err
    logger = configure(verbose=True)
    logger.debug("detail")
    assert "[DEBUG] detail" in capsys.readouterr().err


def test_cli_profile_selection_and_removed_language_option(deck_path, tmp_path):
    help_result = run_cli("export", "--help", cwd=tmp_path)
    assert "--prompt-profile" in help_result.stdout and "--language" not in help_result.stdout
    output = tmp_path / "english.md"
    result = run_cli(
        "export", deck_path, "--output", output, "--prompt-profile", "default-en", cwd=tmp_path
    )
    assert result.returncode == 0, result.stderr
    assert "Extracted content" in output.read_text("utf-8")
    bad = run_cli("export", deck_path, "--language", "en", cwd=tmp_path)
    assert bad.returncode == 2
