from __future__ import annotations

import json
import logging
import subprocess
import sys
from unittest.mock import Mock

import pytest

from powerpoint_note import __version__
from powerpoint_note.logging_utils import SUCCESS, ColorFormatter, configure, supports_color


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
    assert help_result.returncode == 0 and "build" in help_result.stdout
    version = run_cli("--version", cwd=tmp_path)
    assert version.returncode == 0 and version.stdout.strip() == __version__
    output = tmp_path / "result.md"
    built = run_cli("build", deck_path, "--output", output, cwd=tmp_path)
    assert built.returncode == 0, built.stderr
    assert json.loads(built.stdout)["status"] == "created"
    assert "[SUCCESS]" in built.stderr
    verified = run_cli("verify", output, cwd=tmp_path)
    assert verified.returncode == 0
    assert json.loads(verified.stdout)["status"] == "verified"


def test_cli_config_show_readonly_and_error(tmp_path):
    show = run_cli("config", "show", cwd=tmp_path)
    assert show.returncode == 0 and json.loads(show.stdout)["effective_schema_version"] == "2.0.0"
    bad = run_cli("build", "missing.pptx", cwd=tmp_path)
    assert bad.returncode != 0 and json.loads(bad.stdout)["status"] == "failed"
    assert "Traceback" not in bad.stderr


@pytest.mark.parametrize(
    "args", [["--quiet", "--verbose", "config", "show"], ["--quiet", "config", "show", "--verbose"]]
)
def test_quiet_verbose_mutually_exclusive(tmp_path, args):
    result = run_cli(*args, cwd=tmp_path)
    assert result.returncode != 0
    assert json.loads(result.stdout)["status"] == "failed"


def test_no_sources_or_sync_commands(tmp_path):
    for command in ("pull", "push"):
        result = run_cli(command, cwd=tmp_path)
        assert result.returncode != 0


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
    help_result = run_cli("build", "--help", cwd=tmp_path)
    assert "--prompt-profile" in help_result.stdout and "--language" not in help_result.stdout
    output = tmp_path / "english.md"
    result = run_cli(
        "build", deck_path, "--output", output, "--prompt-profile", "default-en", cwd=tmp_path
    )
    assert result.returncode == 0, result.stderr
    assert "Extracted content" in output.read_text("utf-8")
    bad = run_cli("build", deck_path, "--language", "en", cwd=tmp_path)
    assert bad.returncode == 2
