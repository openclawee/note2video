from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from note2video.subtitle.ass import build_ass
from note2video.subtitle.wrap import (
    SubtitleWrapLayout,
    subtitle_wrap_layout_from_canvas,
    wrap_subtitle_text,
)
from note2video.text_segmentation import split_sentences
from note2video.video_canvas import canvas_size

_TRAILING_DISPLAY_PUNCT = "。！？!?；;：:，,、.…"


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
    wrap_ctx = _subtitle_wrap_context(manifest)
    timing_segments = _load_timing_segments(project_dir, manifest)
    if timing_segments is not None:
        segments = _build_segments_from_timings(timing_segments, wrap_ctx=wrap_ctx)
    else:
        durations = _load_slide_durations(manifest)
        segments = _build_segments(scripts=scripts, durations=durations, wrap_ctx=wrap_ctx)
    srt_path = subtitles_dir / "subtitles.srt"
    ass_path = subtitles_dir / "subtitles.ass"
    json_path = subtitles_dir / "subtitles.json"

    srt_path.write_text(_render_srt(segments), encoding="utf-8")
    # Always generate an ASS version too; it enables fade/scale/outline/shadow/highlight later.
    ass_path.write_text(
        build_ass(
            segments=[segment.__dict__ for segment in segments],
            font=wrap_ctx["font"],
            font_size=int(wrap_ctx["font_size"]),
            play_res_x=int(wrap_ctx["canvas_w"]),
            play_res_y=int(wrap_ctx["canvas_h"]),
            margin_l=int(wrap_ctx["margin_lr"]),
            margin_r=int(wrap_ctx["margin_lr"]),
            margin_v=int(wrap_ctx["margin_v"]),
            outline=int(wrap_ctx["outline"]),
        ),
        encoding="utf-8",
    )
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


def _subtitle_wrap_context(manifest: dict[str, Any]) -> dict[str, Any]:
    ratio = str(manifest.get("ratio") or "16:9").strip() or "16:9"
    resolution = str(manifest.get("resolution") or "1080p").strip().lower() or "1080p"
    canvas_w, canvas_h = canvas_size(ratio=ratio, resolution=resolution)
    try:
        fs = int(manifest.get("subtitle_size", 48) or 48)
    except Exception:
        fs = 48
    fs = max(8, fs)
    try:
        outline = int(manifest.get("subtitle_outline", 1) or 1)
    except Exception:
        outline = 1
    outline = max(0, outline)
    font = str(manifest.get("subtitle_font") or "").strip()
    scale_w = float(canvas_w) / 1920.0
    margin_lr = max(24, int(round(80 * scale_w)))
    margin_v = max(int(round(fs * 1.1)), int(round(60 * (float(canvas_h) / 1080.0))))
    layout = subtitle_wrap_layout_from_canvas(
        canvas_w=canvas_w,
        canvas_h=canvas_h,
        font_size=fs,
        margin_l=margin_lr,
        margin_r=margin_lr,
        outline=outline,
        max_lines=4,
    )
    return {
        "layout": layout,
        "font": font,
        "font_size": fs,
        "canvas_w": canvas_w,
        "canvas_h": canvas_h,
        "margin_lr": margin_lr,
        "margin_v": margin_v,
        "outline": outline,
    }


def _build_segments(
    *,
    scripts: list[dict[str, Any]],
    durations: dict[int, int],
    wrap_ctx: dict[str, Any],
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
            sentence = _to_display_subtitle_text(sentence, wrap_ctx=wrap_ctx)
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
    wrap_ctx: dict[str, Any],
) -> list[SubtitleSegment]:
    segments: list[SubtitleSegment] = []
    for item in timing_segments:
        page = int(item["page"])
        text = _to_display_subtitle_text(str(item.get("text", "")), wrap_ctx=wrap_ctx)
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
    return split_sentences(text)


def _to_display_subtitle_text(text: str, *, wrap_ctx: dict[str, Any]) -> str:
    layout: SubtitleWrapLayout | None = wrap_ctx.get("layout")  # type: ignore[assignment]
    font = str(wrap_ctx.get("font") or "")
    wrapped = wrap_subtitle_text(text, layout=layout, font_name=font)
    return _strip_trailing_display_punct(wrapped)


def _strip_trailing_display_punct(text: str) -> str:
    t = (text or "").rstrip()
    while t and t[-1] in _TRAILING_DISPLAY_PUNCT:
        t = t[:-1].rstrip()
    return t


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
