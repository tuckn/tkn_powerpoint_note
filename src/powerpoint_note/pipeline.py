"""Build immutable evidence bundles, then atomically publish a protected Markdown note."""

from __future__ import annotations

import json
import logging
import os
import uuid
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from . import __version__
from .context import Generator
from .context_profiles import load_profile
from .extract import read_deck
from .frontmatter import SCHEMA_VERSION, normalize, now_iso, profile_properties, source_properties
from .io import atomic_write, digest, file_digest, fingerprint, json_text
from .markdown import compose, render_body, slide_evidence, split_note
from .models import Deck, NoteError, Slide
from .render import check_renderable, preflight, render
from .selection import select

LOGGER = logging.getLogger("powerpoint_note")


def default_output(source: Path) -> Path:
    identity = digest(os.path.normcase(str(source.resolve())).encode())[:16]
    return Path.home() / ".tkn/powerpoint_note/data" / identity / (source.name + ".md")


def inspect(source: Path, config: dict[str, Any]) -> dict[str, Any]:
    deck = read_deck(source, config["extraction"])
    slides = select(deck, config["selection"])
    return {
        "status": "success",
        "source_sha256": deck.source_sha256,
        "total_slides": len(deck.slides),
        "selected_slides": len(slides),
        "sections": list(dict.fromkeys(s.section for s in deck.slides if s.section is not None)),
        "slides": [
            {
                "number": s.number,
                "slide_id": s.slide_id,
                "title": s.title,
                "section": s.section,
                "hidden": s.hidden,
                "page_number": s.page_number,
                "objects": len(s.shapes),
                "off_slide_objects": sum(x.location == "outside" for x in s.shapes),
                "has_notes": bool(s.notes),
                "comment_count": len(s.comments),
            }
            for s in slides
        ],
        "warnings": deck.warnings,
    }


def safe_child(parent: Path, relative: str) -> Path:
    path = (parent / relative).resolve()
    if path == parent.resolve() or not path.is_relative_to(parent.resolve()):
        raise NoteError("Evidence path escapes its bundle")
    return path


def validate_bundle(folder: Path) -> dict[str, Any]:
    try:
        manifest: dict[str, Any] = json.loads((folder / "manifest.json").read_text("utf-8"))
        if manifest.get("schema_version") != "1.0.0" or manifest.get("status") != "complete":
            raise NoteError("Evidence bundle is incomplete or has an unsupported schema")
        if (
            not isinstance(manifest.get("files"), dict)
            or not {"body.md", "evidence.json"} <= manifest["files"].keys()
        ):
            raise NoteError("Evidence manifest lacks required files")
        for name, expected in manifest["files"].items():
            path = safe_child(folder, name)
            if not path.is_file() or file_digest(path) != expected:
                raise NoteError(f"Evidence is missing or modified: {name}")
        return manifest
    except (OSError, ValueError, TypeError, KeyError) as exc:
        raise NoteError(
            "Cannot read evidence bundle; restore it or rebuild to a different output"
        ) from exc


def verify(note: Path, source: Path | None = None) -> dict[str, Any]:
    metadata, _, body, _ = split_note(note.read_text("utf-8-sig"))
    if digest(body.encode()) != metadata.get("generatedSha256"):
        raise NoteError("Generated Markdown was edited")
    folder = safe_child(note.parent, metadata["evidencePath"])
    manifest = validate_bundle(folder)
    if (folder / "body.md").read_text("utf-8") != body:
        raise NoteError("Note and evidence body disagree")
    if manifest["source_sha256"] != metadata.get("sourceSha256") or manifest[
        "build_key"
    ] != metadata.get("buildKey"):
        raise NoteError("Note and evidence provenance disagree")
    if (manifest.get("generation_generator") or "") != (metadata.get("generationGenerator") or ""):
        raise NoteError("Note and evidence generation settings provenance disagree")
    provenance = manifest.get("prompt_profile")
    profile_matches = (
        provenance == metadata.get("promptProfile")
        if metadata["schemaVersion"] == "1.0.0"
        else isinstance(provenance, dict)
        and all(metadata.get(key) == value for key, value in profile_properties(provenance).items())
    )
    if not profile_matches:
        raise NoteError("Note and evidence prompt profile provenance disagree")
    original = source or Path(metadata["sourceFile"])
    if not original.is_file() or file_digest(original) != metadata.get("sourceSha256"):
        raise NoteError("Source is missing or changed since this note was built")
    return {
        "status": "verified",
        "note": str(note),
        "source_sha256": metadata["sourceSha256"],
        "selected_slides": metadata["selectedSlides"],
        "context": metadata["context"],
        "semantic_accuracy": "not assessed",
    }


def _existing(output: Path, source: Path, force: bool) -> tuple[str | None, dict[str, Any]]:
    if not output.exists():
        return None, {}
    text = output.read_text("utf-8-sig")
    metadata, _, body, _ = split_note(text)
    if Path(metadata.get("sourceFile", "")).resolve() != source.resolve():
        raise NoteError("Output belongs to another source; choose another --output")
    if digest(body.encode()) != metadata.get("generatedSha256") and not force:
        raise NoteError("Generated content was edited; use another --output or --force with backup")
    return text, metadata


def _evidence(deck: Deck, slides: list[Slide], config: dict[str, Any]) -> dict[str, Any]:
    return {
        "schema_version": "1.0.0",
        "source_sha256": deck.source_sha256,
        "source_size": deck.source_size,
        "dimensions_emu": [deck.width, deck.height],
        "metadata": deck.metadata,
        "total_slides": len(deck.slides),
        "selected_slides": [s.number for s in slides],
        "selection": config["selection"],
        "warnings": deck.warnings,
        "slides": [slide_evidence(s, config["selection"]["off_slide"]) for s in slides],
    }


def build(
    source: Path,
    output: Path,
    config: dict[str, Any],
    *,
    context: bool = False,
    dry_run: bool = False,
    force: bool = False,
    refresh: bool = False,
) -> dict[str, Any]:
    source, output = source.expanduser().resolve(), output.expanduser().resolve()
    if output == source or output.suffix.lower() != ".md":
        raise NoteError("Output must be a separate .md file")
    if refresh and not context:
        raise NoteError("--refresh requires --context")
    profile = load_profile(config["generation"])
    LOGGER.info("Reading presentation and selecting slides.")
    deck = read_deck(source, config["extraction"])
    slides = select(deck, config["selection"])
    evidence = _evidence(deck, slides, config)
    existing, old = _existing(output, source, force)
    # Validate migrations and source dates before any writes, rendering or AI calls.
    previous = normalize(old) if existing is not None else {}
    source_fields = source_properties(deck.metadata)
    generator = None
    if context:
        if len(slides) > config["generation"]["max_slides"]:
            raise NoteError(
                f"Selected {len(slides)} slides exceeds generation.max_slides={config['generation']['max_slides']}; select a smaller range or explicitly raise --max-slides"
            )
        for slide in evidence["slides"]:
            if len(json_text({"slide": slide})) > config["generation"]["max_input_chars"]:
                raise NoteError(
                    f"Slide {slide['number']} exceeds generation.max_input_chars; no AI was called"
                )
        preflight()
        check_renderable(source.read_bytes())
        generator = Generator(config["generation"], content_profile=profile)
    build_key = fingerprint(
        {
            "source": deck.source_sha256,
            "evidence": evidence,
            "generation": {
                key: value
                for key, value in config["generation"].items()
                if key not in {"generators", "default_generator"}
            }
            if context
            else {
                "prompt_profile": profile.name,
                "generator_id": config["generation"].get("generator_id"),
            },
            "context": context,
            "version": __version__,
            "note_schema_version": SCHEMA_VERSION,
            "prompt_profile": profile.provenance(),
            "connection": generator.plan if generator else None,
        }
    )
    if old.get("buildKey") == build_key and not refresh and not force:
        verify(output, source)
        return {
            "status": "unchanged",
            "dry_run": dry_run,
            "output": str(output),
            "selected_slides": [s.number for s in slides],
            "generation_generator": config["generation"].get("generator_id"),
            "ai_calls": 0,
        }
    if old.get("reviewStatus") in {"reviewed", "accepted"} and not force:
        raise NoteError(
            "Reviewed note is protected; choose another --output or use --force with backup"
        )
    base = output.with_name(output.name + ".assets")
    folder = base / build_key
    if refresh:
        folder = base / (build_key + "-" + uuid.uuid4().hex[:12])
    cached = folder.is_dir() and not refresh
    if cached:
        validate_bundle(folder)
    status = "updated" if existing is not None else "created"
    result: dict[str, Any] = {
        "status": status,
        "dry_run": dry_run,
        "output": str(output),
        "evidence": str(folder),
        "selected_slides": [s.number for s in slides],
        "total_slides": len(deck.slides),
        "context": context,
        "prompt_profile": profile.provenance(),
        "generation_generator": config["generation"].get("generator_id"),
        "ai_calls": (len(slides) + 1 if context and not cached else 0),
        "excluded_hidden": sum(s.hidden for s in deck.slides)
        if not config["selection"]["include_hidden"]
        else 0,
    }
    if generator:
        result["connection"] = generator.plan
    if dry_run:
        result["planned_ai_calls"] = result["ai_calls"]
        result["ai_calls"] = 0
        result["planned_counts"] = {status: 1, "reused_bundles": int(cached)}
        return result
    output.parent.mkdir(parents=True, exist_ok=True)
    lock = output.with_name(output.name + ".lock")
    try:
        handle = lock.open("x", encoding="utf-8")
    except FileExistsError as exc:
        raise NoteError(
            "A note lock already exists; check for a running export before removing the stale lock"
        ) from exc
    staging: Path | None = None
    try:
        with handle:
            handle.write(str(os.getpid()))
        if (output.read_text("utf-8-sig") if output.exists() else None) != existing:
            raise NoteError("Output changed during validation; retry")
        if not cached:
            base.mkdir(parents=True, exist_ok=True)
            staging = base / ("pending-" + uuid.uuid4().hex)
            staging.mkdir()
            atomic_write(staging / "evidence.json", json_text(evidence))
            contexts = None
            if generator:
                LOGGER.info("Rendering %d selected slides.", len(slides))
                images = render(
                    source,
                    deck.source_sha256,
                    [s.number for s in slides],
                    staging,
                    config["generation"]["image_width"],
                    config["generation"]["render_timeout_seconds"],
                )
                analyses = {}
                for slide, image in zip(evidence["slides"], images, strict=True):
                    LOGGER.info("Generating context for slide %d.", slide["number"])
                    analyses[str(slide["number"])] = generator.generate(
                        "slide", {"slide": slide}, image, staging, f"slide-{slide['number']:04d}"
                    )
                summary = generator.generate(
                    "deck", {"slides": list(analyses.values())}, None, staging, "deck"
                )
                contexts = {"slides": analyses, "deck": summary}
                atomic_write(staging / "context.json", json_text(contexts))
            relative = folder.relative_to(output.parent).as_posix()
            body = render_body(
                deck, slides, relative + "/evidence.json", relative, config, profile, contexts
            )
            atomic_write(staging / "body.md", body)
            manifest = {
                "schema_version": "1.0.0",
                "status": "complete",
                "build_key": build_key,
                "source_sha256": deck.source_sha256,
                "generator_version": __version__,
                "prompt_profile": profile.provenance(),
                "generation_generator": config["generation"].get("generator_id"),
                "created_at": datetime.now(UTC).isoformat(),
                "files": {p.name: file_digest(p) for p in staging.iterdir() if p.is_file()},
            }
            atomic_write(staging / "manifest.json", json_text(manifest))
            validate_bundle(staging)
            if file_digest(source) != deck.source_sha256:
                raise NoteError("Source changed during export; previous note preserved")
            staging.rename(folder)
            staging = None
        else:
            body = (folder / "body.md").read_text("utf-8")
        now = now_iso()
        metadata = {
            "type": "powerpoint",
            "schemaVersion": SCHEMA_VERSION,
            **source_fields,
            "title": source_fields["title"] or source.stem,
            "description": previous.get("description", ""),
            "cover": previous.get("cover", ""),
            "generator": "tkn-powerpoint-note",
            "generatorVersion": __version__,
            **profile_properties(profile.provenance()),
            "generationGenerator": config["generation"].get("generator_id"),
            "sourceFile": str(source),
            "sourceSha256": deck.source_sha256,
            "selectedSlides": [s.number for s in slides],
            "context": context,
            "buildKey": build_key,
            "evidencePath": folder.relative_to(output.parent).as_posix(),
            "reviewStatus": "unreviewed",
            "tags": previous.get("tags", []),
            "created": previous.get("created", "") if existing is not None else now,
            "updated": now,
            "noteId": previous.get("noteId") or str(uuid.uuid4()),
        }
        text = compose(existing, metadata, body, profile)
        split_note(text)
        if (output.read_text("utf-8-sig") if output.exists() else None) != existing or file_digest(
            source
        ) != deck.source_sha256:
            raise NoteError("Source or note changed before publication; previous note preserved")
        if force and existing is not None:
            backup = output.with_name(output.name + ".backup-" + digest(existing.encode())[:16])
            if not backup.exists():
                atomic_write(backup, existing)
            result["backup"] = str(backup)
        atomic_write(output, text, overwrite=existing is not None)
        result["verified"] = verify(output, source)["status"]
        return result
    except BaseException:
        if staging is not None and staging.exists():
            failed = base / ("failed-" + staging.name.removeprefix("pending-"))
            staging.rename(failed)
            LOGGER.error("Previous note preserved. Failed-run evidence: %s", failed)
        raise
    finally:
        lock.unlink(missing_ok=True)
