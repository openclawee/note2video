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


def generate_subtitles(
    input_json: str,
    output_dir: str,
    *,
    subtitle_size: int | None = None,
    max_lines: int = 2,
) -> dict[str, Any]:
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
    wrap = _SubtitleWrapConfig.from_subtitle_size(subtitle_size, max_lines=max_lines)
    timing_segments = _load_timing_segments(project_dir, manifest)
    if timing_segments is not None:
        segments = _build_segments_from_timings(timing_segments, wrap=wrap)
    else:
        durations = _load_slide_durations(manifest)
        segments = _build_segments(scripts=scripts, durations=durations, wrap=wrap)
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


def _build_segments(
    *,
    scripts: list[dict[str, Any]],
    durations: dict[int, int],
    wrap: "_SubtitleWrapConfig",
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
            sentence = _wrap_subtitle_text(sentence, wrap=wrap)
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
    wrap: "_SubtitleWrapConfig",
) -> list[SubtitleSegment]:
    segments: list[SubtitleSegment] = []
    for item in timing_segments:
        page = int(item["page"])
        text = _wrap_subtitle_text(str(item.get("text", "")), wrap=wrap)
        segments.append(
            SubtitleSegment(
                index=int(item["index"]),
                page=page,
                start_ms=int(item["start_ms"]),
                end_ms=int(item["end_ms"]),
                text=text,
            )
        )
    return segments


def _split_sentences(text: str) -> list[str]:
    normalized = text.replace("\r", "\n").strip()
    if not normalized:
        return []

    raw_parts = re.split(r"(?<=[。！？!?；;])|\n+", normalized)
    parts = [part.strip() for part in raw_parts if part and part.strip()]
    return parts


class _SubtitleWrapConfig:
    def __init__(self, *, max_chars_per_line: int, max_lines: int) -> None:
        self.max_chars_per_line = int(max_chars_per_line)
        self.max_lines = int(max_lines)

    @staticmethod
    def from_subtitle_size(subtitle_size: int | None, *, max_lines: int) -> "_SubtitleWrapConfig":
        """
        Heuristic mapping from font size (1080p) to max characters per line.

        Baseline: 48px font ≈ 18 chars/line (tuned for 1920x1080 with typical margins).
        """
        try:
            size = int(subtitle_size) if subtitle_size is not None else 0
        except Exception:
            size = 0

        if size <= 0:
            return _SubtitleWrapConfig(max_chars_per_line=18, max_lines=max_lines)

        # Inverse scale around the baseline: bigger font → fewer chars per line.
        est = int(round(18 * 48 / max(size, 8)))
        est = max(10, min(30, est))
        return _SubtitleWrapConfig(max_chars_per_line=est, max_lines=max_lines)


def _wrap_subtitle_text(text: str, *, wrap: _SubtitleWrapConfig) -> str:
    """
    Wrap a subtitle sentence into at most `max_lines` lines.

    We use a simple character-count heuristic so it works for both CJK and English
    without relying on font metrics. This is tuned for 1080p output.
    """
    t = (text or "").replace("\r", "\n").strip()
    if not t:
        return ""
    # Preserve explicit newlines from upstream.
    if "\n" in t:
        return "\n".join(line.strip() for line in t.splitlines() if line.strip())
    max_chars_per_line = int(wrap.max_chars_per_line)
    max_lines = int(wrap.max_lines)
    if max_chars_per_line <= 0 or max_lines <= 1:
        return t
    if len(t) <= max_chars_per_line:
        return t

    # Prefer splitting near the middle using punctuation / spaces.
    preferred_breaks = "，,、：:；;。！？!?"
    target = min(len(t) - 1, max_chars_per_line)
    best = -1
    best_score = 10**9
    for i, ch in enumerate(t[:-1], start=1):
        if ch not in preferred_breaks and ch != " ":
            continue
        # Lower score is better: close to target and not too early.
        score = abs(i - target)
        if score < best_score:
            best = i
            best_score = score

    if best <= 0:
        best = max_chars_per_line

    first = t[:best].rstrip()
    second = t[best:].lstrip()
    if not second:
        return first

    if max_lines == 2:
        # If still too long, hard-wrap the second line once.
        if len(second) > max_chars_per_line:
            second = second[:max_chars_per_line].rstrip() + "…"
        return first + "\n" + second

    # Generic multi-line wrapping fallback.
    lines: list[str] = [first]
    rest = second
    while rest and len(lines) < max_lines:
        if len(rest) <= max_chars_per_line:
            lines.append(rest)
            rest = ""
            break
        lines.append(rest[:max_chars_per_line].rstrip())
        rest = rest[max_chars_per_line:].lstrip()
    if rest:
        lines[-1] = (lines[-1].rstrip() + "…").rstrip()
    return "\n".join(lines)


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
