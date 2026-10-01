"""The public, file-oriented CLI. Progress on stderr; JSON results except config show."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, NoReturn

from . import __version__
from .config import initialize, resolve, user_config
from .logging_utils import SUCCESS, configure
from .models import NoteError
from .pipeline import build, default_output, inspect, verify


class Parser(argparse.ArgumentParser):
    def error(self, message: str) -> NoReturn:
        raise NoteError(message)


def common(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--config",
        type=Path,
        default=argparse.SUPPRESS,
        help="Additional config file; overrides user and working-directory settings.",
    )
    logs = parser.add_mutually_exclusive_group()
    logs.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        default=argparse.SUPPRESS,
        help="Only errors on stderr.",
    )
    logs.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        default=argparse.SUPPRESS,
        help="Include diagnostics on stderr.",
    )


def selectors(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--slides", help="1-based positions/ranges, e.g. 1,3-5. Default: all visible slides."
    )
    parser.add_argument(
        "--section",
        action="append",
        help="Exact section name; repeat for multiple names. Intersects --slides.",
    )
    parser.add_argument(
        "--include-hidden",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Include hidden slides (default: excluded).",
    )
    parser.add_argument(
        "--off-slide",
        choices=["exclude", "append"],
        help="Fully outside objects: exclude (default) or separate supplement.",
    )
    parser.add_argument(
        "--notes",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Extract speaker notes (default: yes).",
    )
    parser.add_argument(
        "--comments",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Extract comments (default: yes).",
    )


def make_parser() -> argparse.ArgumentParser:
    parser = Parser(
        prog="tkn-powerpoint-note",
        description="Create PowerPoint Markdown notes; add --context for visual interpretation. Originals are read-only.",
    )
    parser.add_argument("--version", action="version", version=__version__)
    common(parser)
    subs = parser.add_subparsers(dest="command", required=True)
    view = subs.add_parser(
        "inspect", help="Read selected slide titles, sections and object counts."
    )
    common(view)
    view.add_argument("source", type=Path)
    selectors(view)
    create = subs.add_parser("export", help="Export a Markdown note and evidence; AI is opt-in.")
    common(create)
    create.add_argument("source", type=Path)
    create.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Destination .md; default: ~/.tkn/powerpoint_note/data/<source-path-id>/<source>.md",
    )
    selectors(create)
    create.add_argument(
        "--context",
        action="store_true",
        help="Use PowerPoint images and GenAI Bridge; selected images/text leave this app for the configured provider.",
    )
    create.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate/preview only. No persistent files, Office launch, authentication, network or AI calls.",
    )
    create.add_argument(
        "--force",
        action="store_true",
        help="Back up and replace edited/reviewed generated content; preserves handwriting outside markers.",
    )
    create.add_argument(
        "--refresh",
        action="store_true",
        help="Regenerate context even for unchanged input; requires --context and can incur charges.",
    )
    create.add_argument(
        "--generator", help="Named generation settings; overrides generation.default_generator."
    )
    create.add_argument(
        "--bridge-profile", help="GenAI Bridge connection profile (default: codex-default)."
    )
    create.add_argument(
        "--prompt-profile",
        help="Content profile: default-ja (default), default-en, or a complete custom bundle.",
    )
    create.add_argument(
        "--max-slides",
        type=int,
        help="Maximum selected slides for --context (default: 10). Fails rather than truncating.",
    )
    check = subs.add_parser(
        "verify",
        help="Verify note, evidence and source hashes; does not assess AI semantic accuracy.",
    )
    common(check)
    check.add_argument("note", type=Path)
    check.add_argument("--source", type=Path, help="Source location if the presentation moved.")
    config = subs.add_parser("config", help="Create or inspect defaults and setting sources.")
    common(config)
    config_sub = config.add_subparsers(dest="config_command", required=True)
    init = config_sub.add_parser(
        "init", help="Write example user config; edited files are protected."
    )
    common(init)
    init.add_argument("--output", type=Path, help="Explicit alternative config output.")
    init.add_argument("--dry-run", action="store_true", help="Preview without creating files.")
    show = config_sub.add_parser("show", help="Read effective settings and winning sources.")
    common(show)
    show.add_argument("--generator", help="Inspect effective settings for a named generator.")
    show.add_argument("--json", action="store_true", help="Print the full result as JSON.")
    return parser


def overrides(args: argparse.Namespace) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    mapping = {
        "slides": ("selection", "slides"),
        "section": ("selection", "sections"),
        "include_hidden": ("selection", "include_hidden"),
        "off_slide": ("selection", "off_slide"),
        "notes": ("extraction", "include_notes"),
        "comments": ("extraction", "include_comments"),
        "bridge_profile": ("generation", "bridge_profile"),
        "prompt_profile": ("generation", "prompt_profile"),
        "max_slides": ("generation", "max_slides"),
    }
    for arg, (group, key) in mapping.items():
        value = getattr(args, arg, None)
        if value is not None:
            result.setdefault(group, {})[key] = value
    return result


def config_lines(value: Any, prefix: str = "") -> list[str]:
    """Flatten config data into copyable key=value lines without escaping path separators."""
    if isinstance(value, dict):
        if not value:
            return [f"{prefix}={{}}"]
        return [
            line
            for key, item in value.items()
            for line in config_lines(item, f"{prefix}.{key}" if prefix else key)
        ]
    if isinstance(value, list):
        if not value:
            return [f"{prefix}=[]"]
        return [
            line
            for index, item in enumerate(value)
            for line in config_lines(item, f"{prefix}[{index}]")
        ]
    if isinstance(value, str):
        display = value.replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t")
    else:
        display = json.dumps(value, ensure_ascii=False)
    return [f"{prefix}={display}"]


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    logger = configure()
    try:
        args = make_parser().parse_args(argv)
        if getattr(args, "quiet", False) and getattr(args, "verbose", False):
            raise NoteError("--quiet and --verbose cannot be combined")
        logger = configure(getattr(args, "quiet", False), getattr(args, "verbose", False))
        if args.command == "config" and args.config_command == "init":
            result = initialize((args.output or user_config()).expanduser().resolve(), args.dry_run)
        elif args.command == "verify":
            result = verify(
                args.note.expanduser().resolve(),
                args.source.expanduser().resolve() if args.source else None,
            )
        else:
            config, details = resolve(
                getattr(args, "config", None),
                overrides(args),
                generator=getattr(args, "generator", None),
            )
            if args.command == "config":
                result = {"config": config, **details}
            elif args.command == "inspect":
                result = inspect(args.source.expanduser().resolve(), config)
            else:
                result = build(
                    args.source,
                    args.output or default_output(args.source),
                    config,
                    context=args.context,
                    dry_run=args.dry_run,
                    force=args.force,
                    refresh=args.refresh,
                )
        logger.log(SUCCESS, "%s", result.get("status", "success"))
        if args.command == "config" and args.config_command == "show" and not args.json:
            print("\n".join(config_lines(result)))
        else:
            print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
        return 0
    except NoteError as exc:
        logger.error("%s", exc)
        print(json.dumps({"status": "failed", "error": str(exc)}, ensure_ascii=False))
        return 2
    except (OSError, ValueError, KeyError, TypeError) as exc:
        logger.error("Cannot complete command (%s): %s", type(exc).__name__, exc)
        logger.debug("Diagnostic traceback", exc_info=True)
        print(json.dumps({"status": "failed", "error": str(exc)}, ensure_ascii=False))
        return 2
    except Exception as exc:
        logger.error("Processing failed (%s); use --verbose for diagnostics", type(exc).__name__)
        logger.debug("Diagnostic traceback", exc_info=True)
        print(json.dumps({"status": "failed", "error": type(exc).__name__}, ensure_ascii=False))
        return 1
