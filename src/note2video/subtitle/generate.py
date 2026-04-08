from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from note2video.subtitle.ass import build_ass


class SubtitleGenerationError(RuntimeError):
    """Raised when subtitle generation fails."""


@dataclass
class SubtitleSegment:
    index: int
    page: int
    start_ms: int
    end_ms: int
    text: str


def generate_subtitles(input_json: str, output_dir: str) -> dict[str, Any]:
    input_path = Path(input_json)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    project_dir = _resolve_project_dir(input_path, output_dir)
    manifest_path = project_dir / "manifest.json"
    subtitles_dir = project_dir / "subtitles"
    logs_dir = project_dir / "logs"

    subtitles_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)

    scripts = _load_scripts(input_path)
    manifest = _load_manifest(manifest_path)
    timing_segments = _load_timing_segments(project_dir, manifest)
    word_timings = _load_word_timing_segments(project_dir, manifest)
    if timing_segments is not None:
        segments = _build_segments_from_timings(timing_segments, word_timings=word_timings)
    else:
        durations = _load_slide_durations(manifest)
        segments = _build_segments(scripts=scripts, durations=durations)
    srt_path = subtitles_dir / "subtitles.srt"
    ass_path = subtitles_dir / "subtitles.ass"
    json_path = subtitles_dir / "subtitles.json"

    srt_path.write_text(_render_srt(segments), encoding="utf-8")
    # Always generate an ASS version too; it enables fade/scale/outline/shadow/highlight later.
    ass_path.write_text(build_ass(segments=[segment.__dict__ for segment in segments]), encoding="utf-8")
    json_path.write_text(
        json.dumps(
            {"segments": [segment.__dict__ for segment in segments]},
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    outputs = manifest.setdefault("outputs", {})
    outputs["subtitle"] = "subtitles/subtitles.srt"
    outputs["subtitle_ass"] = "subtitles/subtitles.ass"
    outputs["subtitle_json"] = "subtitles/subtitles.json"
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")

    (logs_dir / "subtitle.log").write_text(
        f"segments: {len(segments)}\nslides: {len(scripts)}\n",
        encoding="utf-8",
    )

    return {
        "subtitle": str(srt_path),
        "subtitle_ass": str(ass_path),
        "subtitle_json": str(json_path),
        "segment_count": len(segments),
        "slide_count": len(scripts),
    }


def _resolve_project_dir(input_path: Path, output_dir: str) -> Path:
    if output_dir and output_dir != "./dist":
        return Path(output_dir)
    return input_path.parent.parent


def _load_scripts(input_path: Path) -> list[dict[str, Any]]:
    payload = json.loads(input_path.read_text(encoding="utf-8"))
    slides = payload.get("slides", [])
    if not isinstance(slides, list):
        raise ValueError("Script file is missing a valid 'slides' array.")
    return slides


def _load_manifest(manifest_path: Path) -> dict[str, Any]:
    if not manifest_path.exists():
        raise SubtitleGenerationError("manifest.json is required before generating subtitles.")
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def _load_slide_durations(manifest: dict[str, Any]) -> dict[int, int]:
    durations: dict[int, int] = {}
    for slide in manifest.get("slides", []):
        page = int(slide["page"])
        duration_ms = int(slide.get("duration_ms", 0))
        durations[page] = max(duration_ms, 300)
    return durations


def _load_timing_segments(project_dir: Path, manifest: dict[str, Any]) -> list[dict[str, Any]] | None:
    timings_rel = manifest.get("outputs", {}).get("timings")
    if not timings_rel:
        return None
    timings_path = project_dir / timings_rel
    if not timings_path.exists():
        return None
    payload = json.loads(timings_path.read_text(encoding="utf-8"))
    segments = payload.get("segments")
    if not isinstance(segments, list):
        return None
    return segments


def _load_word_timing_segments(project_dir: Path, manifest: dict[str, Any]) -> dict[tuple[int, int], list[dict[str, Any]]]:
    """
    Load word-level timing from word_timings.json if available.

    Returns a dict keyed by (page, sentence_index) → list of word dicts
    with {text, offset_ms, duration_ms}.
    """
    word_timings_rel = manifest.get("outputs", {}).get("word_timings")
    if not word_timings_rel:
        return {}
    word_timings_path = project_dir / word_timings_rel
    if not word_timings_path.exists():
        return {}
    payload = json.loads(word_timings_path.read_text(encoding="utf-8"))
    segments = payload.get("segments")
    if not isinstance(segments, list):
        return {}
    # Index by (page, sentence_index) for O(1) lookup when building ASS lines.
    result: dict[tuple[int, int], list[dict[str, Any]]] = {}
    for seg in segments:
        words = seg.get("words") or []
        if isinstance(words, list) and words:
            key = (int(seg["page"]), int(seg["sentence_index"]))
            result[key] = words
    return result


def _build_segments(
    *,
    scripts: list[dict[str, Any]],
    durations: dict[int, int],
) -> list[SubtitleSegment]:
    segments: list[SubtitleSegment] = []
    cursor_ms = 0
    index = 1

    for slide in scripts:
        page = int(slide["page"])
        text = slide.get("script", "")
        sentences = _split_sentences(text)
        slide_duration = durations.get(page, 300)

        if not sentences:
            segments.append(
                SubtitleSegment(
                    index=index,
                    page=page,
                    start_ms=cursor_ms,
                    end_ms=cursor_ms + slide_duration,
                    text="",
                )
            )
            cursor_ms += slide_duration
            index += 1
            continue

        weights = [max(len(sentence.replace("\n", "")), 1) for sentence in sentences]
        total_weight = sum(weights)
        segment_durations = _allocate_durations(total=slide_duration, weights=weights)

        sentence_start = cursor_ms
        for sentence, duration in zip(sentences, segment_durations):
            sentence_end = sentence_start + duration
            segments.append(
                SubtitleSegment(
                    index=index,
                    page=page,
                    start_ms=sentence_start,
                    end_ms=sentence_end,
                    text=sentence,
                )
            )
            sentence_start = sentence_end
            index += 1

        cursor_ms += slide_duration

    return segments


def _build_segments_from_timings(
    timing_segments: list[dict[str, Any]],
    *,
    word_timings: dict[tuple[int, int], list[dict[str, Any]]],
) -> list[SubtitleSegment]:
    segments: list[SubtitleSegment] = []
    for item in timing_segments:
        page = int(item["page"])
        sentence_index = int(item.get("sentence_index", 0) or 0)
        words = word_timings.get((page, sentence_index), []) or []
        segments.append(
            SubtitleSegment(
                index=int(item["index"]),
                page=page,
                start_ms=int(item["start_ms"]),
                end_ms=int(item["end_ms"]),
                text=str(item.get("text", "")),
            )
        )
        if words:
            # Attach words to the last segment's __dict__ so build_ass can read it.
            segments[-1].__dict__["words"] = words
    return segments


def _split_sentences(text: str) -> list[str]:
    normalized = text.replace("\r", "\n").strip()
    if not normalized:
        return []

    raw_parts = re.split(r"(?<=[。！？!?；;])|\n+", normalized)
    parts = [part.strip() for part in raw_parts if part and part.strip()]
    return parts


def _allocate_durations(*, total: int, weights: list[int]) -> list[int]:
    if not weights:
        return []

    total_weight = sum(weights)
    durations = [max(int(total * weight / total_weight), 1) for weight in weights]
    diff = total - sum(durations)
    durations[-1] += diff
    return durations


def _render_srt(segments: list[SubtitleSegment]) -> str:
    blocks: list[str] = []
    for segment in segments:
        blocks.append(
            "\n".join(
                [
                    str(segment.index),
                    f"{_format_timestamp(segment.start_ms)} --> {_format_timestamp(segment.end_ms)}",
                    segment.text,
                ]
            )
        )
    return "\n\n".join(blocks) + "\n"


def _format_timestamp(ms: int) -> str:
    hours, rem = divmod(ms, 3_600_000)
    minutes, rem = divmod(rem, 60_000)
    seconds, millis = divmod(rem, 1000)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d},{millis:03d}"
