"""Turn presentations into notes: plan per slide, reuse saved AI results, then publish."""

from __future__ import annotations

import logging
import os
import uuid
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .context import Generator, generator_label
from .context_profiles import ContextProfile, load_profile
from .evidence import content_key, payload_text, slide_payload
from .extract import SUFFIXES, read_deck
from .frontmatter import (
    CONTEXT,
    EXPORT_SCHEMA,
    EXPORT_TYPE,
    PROXY_SCHEMA,
    PROXY_TYPE,
    now_iso,
    source_properties,
    timestamp,
)
from .io import atomic_write, digest
from .markdown import (
    body_digest,
    compose,
    compose_export,
    image_link,
    read_frontmatter,
    render_body,
    split_note,
)
from .models import ConfigError, Deck, NoteError, Slide
from .render import check_renderable, preflight, render
from .selection import format_slides, select
from .sources import discover, missing_sources, note_path, resolve_file
from .store import assets_dir, backup, cached, load_state, note_lock, prune, result_dir, save_state

LOGGER = logging.getLogger("powerpoint_note")
CHANGED = {"created", "updated", "exported"}
FAILED = {"error", "export-error", "protected"}


@dataclass
class Options:
    context: bool = False
    force: bool = False
    dry_run: bool = False
    slides: str | None = None
    sections: list[str] = field(default_factory=list)


@dataclass
class Target:
    source: Path
    note: Path
    proxy: bool
    identity: dict[str, Any] = field(default_factory=dict)
    term_format: str = "plain"


@dataclass
class Plan:
    """Everything known about one presentation before anything is written or sent."""

    deck: Deck
    included: list[Slide]
    payloads: dict[str, str]
    contents: dict[str, str]
    records: dict[str, dict[str, Any]]
    summary: dict[str, Any] | None
    fresh: list[tuple[Slide, str]] = field(default_factory=list)
    # Slides identical to one in `fresh` (dividers, duplicates): one AI call serves them all.
    twins: list[tuple[Slide, str]] = field(default_factory=list)
    skipped: dict[str, str] = field(default_factory=dict)
    reused: int = 0
    warnings: list[str] = field(default_factory=list)
    # Slides this run may send to AI (`--slides` / `--section`; all of them otherwise).
    selected: set[str] = field(default_factory=set)

    def stale(self, slide_id: str) -> bool:
        return bool(self.records[slide_id]["contentKey"] != self.contents[slide_id])

    def current(self) -> list[Slide]:
        return [
            s for s in self.included if s.slide_id in self.records and not self.stale(s.slide_id)
        ]

    def summary_keys(self) -> list[list[Any]]:
        return [[s.number, self.records[s.slide_id]["key"]] for s in self.current()]

    def summary_stale(self) -> bool:
        """The saved summary cites slide numbers or explanations that have since changed."""
        return self.summary is not None and self.summary.get("slideKeys") != self.summary_keys()


def _record(hit: dict[str, Any], key: str, content: str) -> dict[str, Any]:
    usage = hit["usage"]
    return {
        "key": key,
        "contentKey": content,
        "profile": usage.get("prompt_profile", ""),
        "generator": generator_label(usage),
        "generatedAt": timestamp(usage.get("finished_at"), "generatedAt"),
    }


def _context_properties(plan: Plan) -> dict[str, Any]:
    """Frontmatter that describes every explanation kept in the note, not only this run."""
    records, summary = plan.records, plan.summary
    with_context = [s for s in plan.included if s.slide_id in records]
    stale = [s.number for s in with_context if plan.stale(s.slide_id)]
    summary_stale = plan.summary_stale()
    if not with_context and not summary:
        return {"contextStatus": "not-generated"}
    if stale or summary_stale:
        status = "stale"
    elif len(with_context) == len(plan.included):
        status = "current"
    else:
        status = "partial"
    profiles = {r["profile"] for r in records.values()} | (
        {summary["profile"]} if summary else set()
    )
    times = [r["generatedAt"] for r in records.values() if r["generatedAt"]]
    return {
        "contextStatus": status,
        "contextProfiles": sorted(p for p in profiles if p),
        "contextAnalyzedSlides": format_slides([s.number for s in with_context]),
        "contextOmittedSlides": format_slides(
            [s.number for s in plan.included if s.slide_id not in records]
        ),
        "contextStaleSlides": format_slides(stale),
        "contextSourceFingerprint": summary["sourceFingerprint"]
        if summary
        else plan.deck.content_fingerprint,
        "contextGeneratedAt": summary["generatedAt"] if summary else max(times, default=""),
    }


def _body(target: Target, plan: Plan, profile: ContextProfile, off_slide: str) -> str:
    assets = assets_dir(target.note)
    contexts: dict[str, dict[str, Any]] = {}
    for slide in plan.included:
        record = plan.records.get(slide.slide_id)
        hit = cached(assets, "slide", record["key"]) if record else None
        if hit is None:
            plan.records.pop(slide.slide_id, None)
            continue
        image = result_dir(assets, "slide", hit["usage"]["key"]) / "slide.png"
        contexts[slide.slide_id] = {
            "data": hit["data"],
            "image": image_link(
                f"{profile.labels['slide']} {slide.number}",
                image.relative_to(target.note.parent).as_posix(),
            )
            if image.is_file()
            else "",
            "stale": plan.stale(slide.slide_id),
        }
    summary = None
    if plan.summary:
        hit = cached(assets, "deck", plan.summary["key"])
        if hit is None:
            plan.summary = None
        else:
            summary = hit["data"]
    return render_body(
        plan.deck,
        plan.included,
        profile,
        off_slide=off_slide,
        contexts=contexts,
        summary=summary,
        summary_stale=plan.summary_stale(),
        skipped=plan.skipped,
    )


def _summary_decision(
    plan: Plan, generator: Generator, assets: Path, limits: dict[str, Any]
) -> dict[str, Any]:
    """What to do about the presentation summary, before or after the slides are explained."""
    pending = {slide.slide_id: key for slide, key in [*plan.fresh, *plan.twins]}
    entries: list[tuple[Slide, str]] = []
    for slide in plan.included:
        sid = slide.slide_id
        if sid in pending:
            entries.append((slide, pending[sid]))
        elif (
            sid in plan.records
            and not plan.stale(sid)
            # A record whose saved result is gone cannot be summarized or shown.
            and cached(assets, "slide", plan.records[sid]["key"]) is not None
        ):
            entries.append((slide, plan.records[sid]["key"]))
    if not entries:
        return {"action": "none"}
    keys = [[slide.number, key] for slide, key in entries]
    key = generator.key("deck", digest(payload_text(keys).encode("utf-8")))
    decision: dict[str, Any] = {"key": key, "keys": keys, "slides": [s for s, _ in entries]}
    if cached(assets, "deck", key) is not None:
        return {**decision, "action": "reuse"}
    if any(slide.slide_id not in plan.selected for slide, _ in entries):
        # Explanations of slides outside this run's selection are never sent anywhere.
        return {
            **decision,
            "action": "skip",
            "reason": "Presentation summary not refreshed: it would send explanations of slides "
            "outside --slides/--section. Run --context without them to refresh it.",
        }
    analyses = []
    for slide, slide_key in entries:
        saved = cached(assets, "slide", slide_key)
        if saved is None:
            return {**decision, "action": "call"}  # known only after the slides are explained
        analyses.append({"slide": slide.number, "title": slide.title, **saved["data"]})
    payload = payload_text({"slides": analyses})
    if len(payload) > limits["max_input_chars"]:
        return {
            **decision,
            "action": "skip",
            "reason": f"Presentation summary skipped: slide explanations need {len(payload):,} "
            f"characters; generation.limits.max_input_chars={limits['max_input_chars']:,}",
        }
    return {**decision, "action": "call", "payload": payload}


def _generate(
    target: Target,
    plan: Plan,
    generator: Generator,
    limits: dict[str, Any],
    progress: dict[str, int],
) -> int:
    """Render and analyze the slides without a saved result, then the summary if it changed."""
    assets = assets_dir(target.note)
    deck = plan.deck
    generated = 0
    if plan.fresh:
        data = target.source.read_bytes()
        if digest(data) != deck.source_sha256:
            raise NoteError("Presentation changed while it was being read; save it and retry")
        images = {
            slide.number: result_dir(assets, "slide", key) / "slide.png"
            for slide, key in plan.fresh
        }
        for image in images.values():
            image.parent.mkdir(parents=True, exist_ok=True)
        LOGGER.info("Rendering %d slide(s) with PowerPoint.", len(images))
        render(data, images, limits["image_width"], limits["max_render_seconds"])
        for slide, key in plan.fresh:
            LOGGER.info("Generating context for slide %d.", slide.number)
            folder = result_dir(assets, "slide", key)
            try:
                generator.generate(
                    "slide", plan.payloads[slide.slide_id], images[slide.number], folder, key
                )
            finally:
                # Counted once the provider was called, even when the call or its output fails.
                progress["aiCalls"] = generator.calls
            generated += 1
            hit = cached(assets, "slide", key)
            if hit is None:
                raise NoteError(f"Saved AI result for slide {slide.number} cannot be read back")
            plan.records[slide.slide_id] = _record(hit, key, plan.contents[slide.slide_id])
        for slide, key in plan.twins:
            hit = cached(assets, "slide", key)
            if hit is not None:
                plan.records[slide.slide_id] = _record(hit, key, plan.contents[slide.slide_id])
    decision = _summary_decision(plan, generator, assets, limits)
    if decision["action"] == "none":
        plan.summary = None
        return generated
    if decision["action"] == "skip":
        # An earlier summary stays in the note and is marked as out of date.
        plan.warnings.append(decision["reason"])
        return generated
    key = decision["key"]
    if decision["action"] == "call":
        if "payload" not in decision:
            plan.warnings.append("Presentation summary skipped: a saved slide result is missing")
            return generated
        LOGGER.info(
            "Generating the presentation summary from %d slide(s).", len(decision["slides"])
        )
        try:
            generator.generate(
                "deck",
                decision["payload"],
                None,
                result_dir(assets, "deck", key),
                key,
                allowed_slides={s.number for s in decision["slides"]},
            )
        finally:
            progress["aiCalls"] = generator.calls
    hit = cached(assets, "deck", key)
    if hit is None:
        raise NoteError("Saved AI summary cannot be read back")
    plan.summary = {
        **_record(hit, key, ""),
        "slideKeys": decision["keys"],
        "sourceFingerprint": deck.content_fingerprint,
    }
    return generated


def process(
    target: Target,
    config: dict[str, Any],
    options: Options,
    progress: dict[str, int] | None = None,
) -> dict[str, Any]:
    """Create or update one note. Raises NoteError; the caller turns it into a result row."""
    progress = progress if progress is not None else {"aiCalls": 0}
    extraction, generation = config["extraction"], config["generation"]
    limits = generation["limits"]
    profile = load_profile(generation)
    source, note = target.source, target.note
    if note.suffix.lower() != ".md" or note == source:
        raise ConfigError("Output must be a separate .md file")
    LOGGER.info("Reading %s", source.name)
    deck = read_deck(source, extraction)
    # Fail on unusable document properties before anything is rendered or sent.
    document = source_properties(deck.metadata, source.stem, target.term_format)
    included = [s for s in deck.slides if extraction["include_hidden"] or not s.hidden]
    if not included:
        raise NoteError("Every slide is hidden; set extraction.include_hidden to include them")
    chosen = (
        select(deck, included, options.slides, options.sections) if options.context else included
    )
    assets = assets_dir(note)
    identity = (
        _note_identity(target.identity["sourceId"])
        if target.proxy
        else os.path.normcase(str(source))
    )
    state = load_state(assets)
    if state and state.get("identity") != identity:
        # Results saved for another presentation must never appear in this note.
        state = {}
    existing = note.read_text("utf-8-sig") if note.exists() else None
    old_meta: dict[str, Any] = {}
    before, after, old_body = "\n", "\n\n" + profile.labels["personal"] + "\n", None
    if existing is not None and not target.proxy:
        if not options.force:
            raise NoteError(f"Output already exists; use --force to replace it: {note}")
        try:
            written_by_export = read_frontmatter(existing)[0].get("type") == EXPORT_TYPE
        except NoteError:
            written_by_export = False
        if not written_by_export:
            raise NoteError(
                f"Output exists and was not written by export; choose another --output: {note}"
            )
    if existing is not None and target.proxy:
        old_meta, before, old_body, after = split_note(existing)
        if _note_identity(str(old_meta.get("sourceId", ""))) != identity:
            raise NoteError("The existing note belongs to another presentation; it was not changed")

    payloads = {
        s.slide_id: payload_text(slide_payload(s, deck, extraction["off_slide"])) for s in included
    }
    contents = {
        s.slide_id: content_key(s, payloads[s.slide_id], limits["image_width"], deck)
        for s in included
    }
    plan = Plan(
        deck,
        included,
        payloads,
        contents,
        {sid: dict(r) for sid, r in state.get("slides", {}).items() if sid in contents},
        dict(state["deck"]) if state.get("deck") else None,
        selected={s.slide_id for s in chosen},
    )
    off_slide = extraction["off_slide"]
    edited = False
    if old_body is not None and body_digest(old_body) != state.get("generatedSha256"):
        # Without a matching record, the note is untouched only if it equals what we would write.
        edited = body_digest(_body(target, plan, profile, off_slide)) != body_digest(old_body)
    if edited and not options.force:
        return {
            "status": "protected",
            "message": "The generated part of the note was edited by hand; "
            "use --force to back it up and replace it.",
        }

    generator = None
    summary_call = False
    if options.context:
        generator = Generator(generation, config["genai_profiles"], profile)
        for slide in chosen:
            sid = slide.slide_id
            key = generator.key("slide", contents[sid])
            hit = cached(assets, "slide", key)
            if hit is not None:
                plan.records[sid] = _record(hit, key, contents[sid])
                plan.reused += 1
            elif len(payloads[sid]) > limits["max_input_chars"]:
                plan.skipped[sid] = (
                    f"evidence needs {len(payloads[sid]):,} characters; "
                    f"generation.limits.max_input_chars={limits['max_input_chars']:,}"
                )
            elif any(key == other for _, other in plan.fresh):
                plan.twins.append((slide, key))
                plan.reused += 1
            else:
                plan.fresh.append((slide, key))
        if len(plan.fresh) > limits["max_slides"]:
            raise NoteError(
                f"{len(plan.fresh)} slides need AI; generation.limits.max_slides="
                f"{limits['max_slides']}. Select fewer with --slides or raise --max-slides. "
                "No AI was called."
            )
        summary_call = _summary_decision(plan, generator, assets, limits)["action"] == "call"
        if plan.fresh:
            preflight()
            check_renderable(source.read_bytes())

    planned = len(plan.fresh) + int(summary_call)
    row: dict[str, Any] = {
        "slides": len(included),
        "aiCalls": 0,
        "plannedAiCalls": planned,
        "generatedSlides": 0,
        "reusedSlides": plan.reused,
        "skippedSlides": [
            f"{s.number}: {plan.skipped[s.slide_id]}"
            for s in included
            if s.slide_id in plan.skipped
        ],
    }
    would = (
        ("would-update" if existing is not None else "would-create")
        if target.proxy
        else "would-export"
    )
    if options.dry_run and planned:
        return {"status": would, **row}

    def finish() -> dict[str, Any]:
        generated = 0
        if generator is not None and (not options.dry_run or not planned):
            # With nothing left to send, this only attaches results saved earlier.
            generated = _generate(target, plan, generator, limits, progress)
        body = _body(target, plan, profile, off_slide)
        properties = _context_properties(plan)
        if target.proxy:
            metadata = {k: v for k, v in old_meta.items() if k not in CONTEXT}
            metadata.update(
                type=PROXY_TYPE,
                schemaVersion=PROXY_SCHEMA,
                **document,
                **target.identity,
                sourceFullPath=str(source),
                **properties,
            )
            stamp = now_iso()
            for key, default in (("description", ""), ("cover", ""), ("tags", [])):
                metadata.setdefault(key, default)
            metadata.setdefault("created", stamp)
            metadata["noteId"] = metadata.get("noteId") or str(uuid.uuid4())
            metadata["updated"] = old_meta.get("updated") or stamp
            text = compose(metadata, before, body, after)
            if existing == text:
                status = "unchanged"
            else:
                metadata["updated"] = stamp
                text = compose(metadata, before, body, after)
                status = "updated" if existing is not None else "created"
        else:
            metadata = {
                "type": EXPORT_TYPE,
                "schemaVersion": EXPORT_SCHEMA,
                **document,
                "sourceFileName": source.name,
                "sourceFullPath": str(source),
                "sourceSnapshotSha256": deck.source_sha256,
                "generatedAt": now_iso(),
                "generationMethod": "context" if plan.records or plan.summary else "text",
                **properties,
            }
            text = compose_export(metadata, body)
            status = "exported"
        result = {
            "status": status,
            **row,
            "aiCalls": progress["aiCalls"],
            "plannedAiCalls": 0 if not options.dry_run else planned,
            "generatedSlides": generated,
            "contextStatus": properties["contextStatus"],
            "staleSlides": properties.get("contextStaleSlides", ""),
            "warnings": plan.warnings,
        }
        if options.dry_run:
            return {**result, "status": would if status != "unchanged" else status}
        if file_changed(note, existing) or digest(source.read_bytes()) != deck.source_sha256:
            raise NoteError(
                "Presentation or note changed during the run; the note was not replaced"
            )
        if status != "unchanged":
            if edited and existing is not None:
                result["backupPath"] = str(backup(note, existing))
            # A note that appeared meanwhile (an editor, a sync tool) is never overwritten.
            atomic_write(note, text, overwrite=existing is not None)
        if target.proxy or plan.records or plan.summary:
            new_state = {
                "identity": identity,
                "source": {
                    "path": str(source),
                    "sha256": deck.source_sha256,
                    "fingerprint": deck.content_fingerprint,
                },
                "generatedSha256": body_digest(body),
                "slides": plan.records,
                "deck": plan.summary,
            }
            if new_state != {k: v for k, v in state.items() if k != "schemaVersion"}:
                save_state(assets, new_state)
            keep = {result_dir(assets, "slide", r["key"]) for r in plan.records.values()}
            if plan.summary:
                keep.add(result_dir(assets, "deck", plan.summary["key"]))
            prune(assets, keep)
        return result

    if options.dry_run:
        return finish()
    with note_lock(note.with_name("." + note.name + ".lock")):
        return finish()


def _note_identity(source_id: str) -> str:
    """Source IDs are exact; the file path after `:` follows Windows and ignores case."""
    root, _, relative = source_id.partition(":")
    return f"{root}:{relative.lower()}"


def file_changed(note: Path, existing: str | None) -> bool:
    return (note.read_text("utf-8-sig") if note.exists() else None) != existing


def _attempt(target: Target, config: dict[str, Any], options: Options) -> dict[str, Any]:
    labels = (
        {"sourcePath": target.identity["sourceFileName"], "notePath": str(target.note)}
        if target.proxy
        else {"inputPath": str(target.source), "outputPath": str(target.note)}
    )
    if target.proxy:
        labels = {"sourceRoot": target.identity["sourceRoot"], **labels}
        labels["presentationPath"] = str(target.source)
    progress = {"aiCalls": 0}
    try:
        return {**labels, **process(target, config, options, progress)}
    except (NoteError, OSError) as exc:
        LOGGER.error("%s: %s", target.source.name, exc)
        return {
            **labels,
            "status": "error" if target.proxy else "export-error",
            "message": str(exc),
            # Calls made before the failure used quota and stay visible.
            "aiCalls": progress["aiCalls"],
        }


def summarize(command: str, rows: list[dict[str, Any]], options: Options) -> dict[str, Any]:
    counts = Counter(row["status"] for row in rows)
    errors = sum(counts[status] for status in FAILED)
    return {
        "status": "error" if errors else "success",
        "command": command,
        "dryRun": options.dry_run,
        # Rows that stand for a presentation; a source that could not be read has none.
        "files": sum("presentationPath" in row or "inputPath" in row for row in rows),
        "changed": sum(counts[status] for status in CHANGED),
        "errors": errors,
        "statusCounts": dict(sorted(counts.items())),
        "aiCalls": sum(row.get("aiCalls", 0) for row in rows),
        "plannedAiCalls": sum(row.get("plannedAiCalls", 0) for row in rows),
        "results": rows,
    }


def pull(
    source_id: str, config: dict[str, Any], options: Options, file: Path | None = None
) -> list[dict[str, Any]]:
    """Create or update the proxy notes of one source. Notes are never deleted."""
    spec = config["sources"][source_id]
    known = discover(spec)
    selected = [resolve_file(spec, file, known)] if file is not None else known
    rows = []
    for relative in selected:
        target = Target(
            Path(spec["presentations"]["dir"]) / relative,
            note_path(spec, relative),
            proxy=True,
            identity={
                "sourceId": f"{source_id}:{relative}",
                "sourceRoot": source_id,
                "sourceFileName": relative,
            },
            term_format=spec["notes"]["frontmatter_term_format"],
        )
        rows.append(_attempt(target, config, options))
    if file is None:
        rows.extend(missing_sources(spec, source_id, known))
    return rows


def export(
    path: Path, output: Path | None, config: dict[str, Any], options: Options
) -> dict[str, Any]:
    """Write `<file>.md` beside each presentation, or one file to --output."""
    # Keep paths as the user wrote them: folders reached through links stay as given.
    path = Path(os.path.abspath(path.expanduser()))
    if path.is_dir():
        if output is not None:
            raise ConfigError("--output cannot be used when the input is a folder")
        sources = sorted(
            p
            for p in path.rglob("*")
            if p.is_file() and p.suffix.lower() in SUFFIXES and not p.name.startswith("~$")
        )
        if not sources:
            raise ConfigError(f"No .pptx or .pptm files found in: {path}")
    elif path.is_file():
        if path.suffix.lower() not in SUFFIXES:
            raise ConfigError(f"Input must be a .pptx or .pptm file or a folder: {path}")
        if output is not None and output.suffix.lower() != ".md":
            raise ConfigError(f"--output must be a .md file: {output}")
        sources = [path]
    else:
        raise ConfigError(f"Input does not exist: {path}")
    rows = []
    for source in sources:
        note = (
            Path(os.path.abspath(output.expanduser()))
            if output is not None
            else source.with_name(source.name + ".md")
        )
        rows.append(_attempt(Target(source, note, proxy=False), config, options))
    summary = summarize("export", rows, options)
    counts = summary["statusCounts"]
    return {
        **{k: v for k, v in summary.items() if k not in {"changed", "results"}},
        "exported": counts.get("exported", 0),
        "planned": counts.get("would-export", 0),
        "results": rows,
    }


def list_slides(path: Path, config: dict[str, Any]) -> dict[str, Any]:
    path = Path(os.path.abspath(path.expanduser()))
    deck = read_deck(path, config["extraction"])
    return {
        "status": "success",
        "command": "presentation list-slides",
        "presentationPath": str(path),
        "totalSlides": len(deck.slides),
        "hiddenSlides": sum(s.hidden for s in deck.slides),
        "sections": list(dict.fromkeys(s.section for s in deck.slides if s.section is not None)),
        "slides": [
            {
                "number": s.number,
                "slideId": s.slide_id,
                "title": s.title,
                "section": s.section,
                "hidden": s.hidden,
                "shapes": len(s.shapes),
                "offSlideShapes": sum(x.location == "outside" for x in s.shapes),
                "hasNotes": bool(s.notes),
                "comments": len(s.comments),
            }
            for s in deck.slides
        ],
        "warnings": deck.warnings,
    }
