"""The public CLI. Logs on stderr; one result on stdout."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Any, NoReturn

from . import __version__
from .config import initialize, resolve, user_config
from .logging_utils import SUCCESS, configure
from .models import ConfigError
from .output import emit
from .pipeline import Options, export, list_slides, pull, summarize

EXIT_OK, EXIT_FAILED, EXIT_ARGUMENT, EXIT_CONFIG, EXIT_INTERRUPTED = 0, 1, 2, 3, 130


class ArgumentError(Exception):
    """Invalid command line (exit code 2)."""


class Parser(argparse.ArgumentParser):
    def error(self, message: str) -> NoReturn:
        raise ArgumentError(message)


def common(parser: argparse.ArgumentParser) -> None:
    """Options accepted before or after every command."""
    parser.add_argument(
        "--config",
        type=Path,
        default=argparse.SUPPRESS,
        help="Additional config file; overrides user and working-directory settings.",
    )
    for flags, text in (
        (("-q", "--quiet"), "Only errors on stderr."),
        (("-v", "--verbose"), "Include diagnostics on stderr."),
        (("--no-color",), "Write logs without color."),
        (("--json",), "Print one JSON result instead of key=value lines."),
    ):
        parser.add_argument(*flags, action="store_true", default=argparse.SUPPRESS, help=text)


def generation(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--context",
        action="store_true",
        help="Explain slides with AI from PowerPoint images; selected images and text are "
        "sent to the configured provider.",
    )
    parser.add_argument(
        "--slides", help="With --context: 1-based positions or ranges to explain, e.g. 1,3-5."
    )
    parser.add_argument(
        "--section",
        action="append",
        help="With --context: exact section name to explain; repeatable. Intersects --slides.",
    )
    parser.add_argument("--generator", help="Named generation settings in generation.generators.")
    parser.add_argument("--profile", help="Prompt profile, e.g. default-ja or default-en.")
    parser.add_argument(
        "--max-slides",
        type=int,
        help="Slides sent to AI per presentation in this run (default: 10). Never truncates.",
    )
    parser.add_argument(
        "--include-hidden",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Include hidden slides in the note (default: excluded).",
    )
    parser.add_argument(
        "--off-slide",
        choices=["exclude", "append"],
        help="Objects fully outside the slide: exclude (default) or list them separately.",
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
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would change. No files, PowerPoint, network or AI calls.",
    )


def make_parser() -> argparse.ArgumentParser:
    parser = Parser(
        prog="tkn-powerpoint-note",
        description="Turn PowerPoint presentations into Markdown notes; --context adds AI "
        "explanations of slide images. Presentations are never modified.",
        allow_abbrev=False,
    )
    parser.add_argument("--version", action="version", version=__version__)
    common(parser)
    subs = parser.add_subparsers(dest="command", required=True)

    create = subs.add_parser("export", help="Write Markdown beside a presentation or folder.")
    common(create)
    create.add_argument("input_path", type=Path, metavar="PATH", help="A .pptx/.pptm or a folder.")
    create.add_argument(
        "-o", "--output", type=Path, help="Destination .md for one file (default: <file>.md)."
    )
    generation(create)
    create.add_argument("--force", action="store_true", help="Replace an existing Markdown file.")

    fetch = subs.add_parser("pull", help="Create or update the proxy notes of a source.")
    common(fetch)
    scope = fetch.add_mutually_exclusive_group(required=True)
    scope.add_argument("--source", help="Source ID in sources.")
    scope.add_argument("--all-sources", action="store_true", help="Every configured source.")
    fetch.add_argument(
        "presentation",
        nargs="?",
        type=Path,
        help="One presentation, relative to or inside presentations.dir.",
    )
    generation(fetch)
    fetch.add_argument(
        "--force",
        action="store_true",
        help="Back up and replace notes whose generated part was edited by hand.",
    )

    state = subs.add_parser("status", help="Show what pull would change; writes nothing.")
    common(state)
    state.add_argument("--source", required=True, help="Source ID in sources.")

    deck = subs.add_parser("presentation", help="Inspect a saved presentation.")
    common(deck)
    deck_sub = deck.add_subparsers(dest="presentation_command", required=True)
    slides = deck_sub.add_parser("list-slides", help="List slides, sections and hidden state.")
    common(slides)
    slides.add_argument("--presentation", type=Path, required=True, help="A .pptx or .pptm file.")

    config = subs.add_parser("config", help="Create or inspect settings.")
    common(config)
    config_sub = config.add_subparsers(dest="config_command", required=True)
    init = config_sub.add_parser("init", help="Write the example user config; never overwrites.")
    common(init)
    init.add_argument("--output", type=Path, help="Write the example to another path.")
    init.add_argument("--dry-run", action="store_true", help="Preview without creating files.")
    listing = config_sub.add_parser(
        "list", help="List effective settings and where they came from."
    )
    common(listing)
    listing.add_argument("--generator", help="Show the settings of a named generator.")
    listing.add_argument("--source", help="Show the settings used for a source.")
    return parser


def overrides(args: argparse.Namespace) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    mapping = {
        "include_hidden": ("extraction", "include_hidden"),
        "off_slide": ("extraction", "off_slide"),
        "notes": ("extraction", "include_notes"),
        "comments": ("extraction", "include_comments"),
        "profile": ("generation", "prompt_profile"),
    }
    for arg, (group, key) in mapping.items():
        value = getattr(args, arg, None)
        if value is not None:
            result.setdefault(group, {})[key] = value
    if getattr(args, "max_slides", None) is not None:
        result.setdefault("generation", {})["limits"] = {"max_slides": args.max_slides}
    return result


def command_name(args: argparse.Namespace | None) -> str | None:
    if args is None:
        return None
    command: str | None = getattr(args, "command", None)
    sub = getattr(args, "config_command", None) or getattr(args, "presentation_command", None)
    return f"{command} {sub}" if command and sub else command


def error_targets(
    args: argparse.Namespace | None, exc: BaseException | None = None
) -> dict[str, Any]:
    """Paths and IDs the failed command referred to, separate from the message."""
    targets: dict[str, Any] = {}
    if args is not None:
        for attribute, key in (
            ("config", "configPath"),
            ("input_path", "inputPath"),
            ("output", "outputPath"),
            ("presentation", "presentationPath"),
            ("source", "sourceId"),
        ):
            value = getattr(args, attribute, None)
            if value is not None:
                targets[key] = str(value)
    if isinstance(exc, OSError) and exc.filename:
        targets["targetPath"] = str(exc.filename)
    return targets


def execute(args: argparse.Namespace) -> tuple[dict[str, Any], int]:
    explicit = getattr(args, "config", None)
    if args.command == "config" and args.config_command == "init":
        target = (args.output or user_config()).expanduser().resolve()
        return initialize(target, args.dry_run), EXIT_OK
    if args.command == "config":
        config, details = resolve(explicit, generator=args.generator, source=args.source)
        return {"status": "success", "command": "config list", "config": config, **details}, EXIT_OK
    if args.command == "presentation":
        config, _ = resolve(explicit)
        return list_slides(args.presentation, config), EXIT_OK
    if args.command == "status":
        options = Options(dry_run=True)
        config, _ = resolve(explicit, source=args.source)
        result = summarize("status", pull(args.source, config, options), options)
        return result, EXIT_FAILED if result["errors"] else EXIT_OK
    options = Options(
        context=args.context,
        force=args.force,
        dry_run=args.dry_run,
        slides=args.slides,
        sections=args.section or [],
    )
    if not args.context and (
        args.slides or args.section or args.generator or args.profile or args.max_slides
    ):
        raise ArgumentError(
            "--slides, --section, --generator, --profile and --max-slides require --context"
        )
    if args.command == "export":
        config, _ = resolve(explicit, overrides(args), generator=args.generator)
        result = export(args.input_path, args.output, config, options)
        return result, EXIT_FAILED if result["errors"] else EXIT_OK
    if args.presentation is not None and args.all_sources:
        raise ArgumentError("A presentation path requires --source, not --all-sources")
    # Settings shared by every source fail here, as configuration errors.
    config, _ = resolve(explicit, overrides(args), generator=args.generator)
    sources = sorted(config["sources"]) if args.all_sources else [args.source]
    if not sources:
        raise ConfigError("No sources are configured; add one under sources in config.yaml")
    rows: list[dict[str, Any]] = []
    for source in sources:
        try:
            config, _ = resolve(explicit, overrides(args), generator=args.generator, source=source)
            rows.extend(pull(source, config, options, args.presentation))
        except ConfigError as exc:
            if not args.all_sources:
                raise
            # One unreachable source must not hide what the others did.
            logging.getLogger("powerpoint_note").error("%s: %s", source, exc)
            rows.append({"status": "error", "sourceRoot": source, "message": str(exc)})
    result = {**summarize("pull", rows, options), "sources": sources}
    return result, EXIT_FAILED if result["errors"] else EXIT_OK


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    tokens = argv[: argv.index("--")] if "--" in argv else argv
    json_output = "--json" in tokens
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    quiet = "--quiet" in tokens or "-q" in tokens
    verbose = "--verbose" in tokens or "-v" in tokens
    logger = configure(quiet, verbose, no_color="--no-color" in tokens)
    args: argparse.Namespace | None = None

    def fail(status: str, message: str, code: int, exc: BaseException | None = None) -> int:
        emit(
            {
                "status": status,
                "command": command_name(args),
                "message": message,
                **error_targets(args, exc),
            },
            json_output=json_output,
        )
        return code

    try:
        args = make_parser().parse_args(argv)
        if quiet and verbose:
            raise ArgumentError("--quiet and --verbose cannot be combined")
        result, code = execute(args)
        if command_name(args) == "config list":
            logger.info("Showing resolved configuration")
        else:
            level = SUCCESS if code == EXIT_OK else logging.ERROR
            logger.log(level, "%s", result.get("status", "success"))
        emit(result, json_output=json_output)
        return code
    except ArgumentError as exc:
        logger.error("%s", exc)
        return fail("argument-error", str(exc), EXIT_ARGUMENT)
    except ConfigError as exc:
        logger.error("%s", exc)
        return fail("config-error", str(exc), EXIT_CONFIG, exc)
    except KeyboardInterrupt:
        logger.error("Interrupted; completed notes and saved AI results are kept.")
        return fail("interrupted", "Interrupted", EXIT_INTERRUPTED)
    except Exception as exc:
        # Processing failures and anything unexpected: one result, details with --verbose.
        logger.error("%s: %s", type(exc).__name__, exc)
        logger.debug("Diagnostic traceback", exc_info=True)
        return fail("error", str(exc) or type(exc).__name__, EXIT_FAILED, exc)
