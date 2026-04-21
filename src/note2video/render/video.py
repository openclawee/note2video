from __future__ import annotations

import json
import subprocess
import wave
from contextlib import closing
from pathlib import Path
from typing import Any

from note2video.subtitle.ass import build_ass
from note2video.subtitle.wrap import subtitle_wrap_layout_from_canvas, wrap_subtitle_text


class RenderError(RuntimeError):
    """Raised when final video rendering fails."""


def _normalize_ratio(value: str | None) -> str:
    raw = str(value or "").strip().lower()
    if not raw:
        return "16:9"
    raw = raw.replace("：", ":").replace("x", ":")
    raw = raw.replace(" ", "")
    if raw in {"16:9", "9:16", "1:1"}:
        return raw
    raise RenderError(f"Unsupported ratio: {value!r}. Use 16:9, 9:16, or 1:1.")


def _ratio_canvas_size(ratio: str) -> tuple[int, int]:
    r = _normalize_ratio(ratio)
    if r == "16:9":
        return 1920, 1080
    if r == "9:16":
        return 1080, 1920
    return 1080, 1080


def _normalize_resolution(value: str | None) -> str:
    raw = str(value or "").strip().lower()
    if not raw:
        return "1080p"
    if raw in {"720p", "1080p", "1440p"}:
        return raw
    raise RenderError(f"Unsupported resolution: {value!r}. Use 720p, 1080p, or 1440p.")


def _scale_canvas_size(base_w: int, base_h: int, resolution: str) -> tuple[int, int]:
    normalized = _normalize_resolution(resolution)
    scale_map = {
        "720p": 2.0 / 3.0,
        "1080p": 1.0,
        "1440p": 4.0 / 3.0,
    }
    scale = scale_map[normalized]
    return int(round(base_w * scale)), int(round(base_h * scale))


def _normalize_fps(value: int | str | None) -> int:
    try:
        fps = int(value)
    except (TypeError, ValueError):
        fps = 30
    if fps < 1 or fps > 120:
        raise RenderError(f"Unsupported fps: {value!r}. Use an integer between 1 and 120.")
    return fps


def _quality_encode_options(value: str | None) -> tuple[str, str]:
    raw = str(value or "").strip().lower() or "standard"
    if raw == "standard":
        return "medium", "23"
    if raw == "high":
        return "slow", "19"
    raise RenderError(f"Unsupported quality: {value!r}. Use standard or high.")


def _load_subtitle_segments_wrapped(
    *,
    subtitle_json_path: Path,
    canvas_w: int,
    canvas_h: int,
    subtitle_size: int | None,
    subtitle_font: str | None,
    subtitle_outline: int | None,
) -> list[dict[str, Any]]:
    payload = json.loads(subtitle_json_path.read_text(encoding="utf-8"))
    base_segments = payload.get("segments") or []
    if not isinstance(base_segments, list):
        return []
    try:
        fs = int(subtitle_size) if subtitle_size is not None else 48
    except Exception:
        fs = 48
    fs = max(8, fs)
    try:
        outline = int(subtitle_outline) if subtitle_outline is not None else 1
    except Exception:
        outline = 1
    outline = max(0, outline)
    scale_w = float(canvas_w) / 1920.0
    margin_lr = max(24, int(round(80 * scale_w)))
    layout = subtitle_wrap_layout_from_canvas(
        canvas_w=canvas_w,
        canvas_h=canvas_h,
        font_size=fs,
        margin_l=margin_lr,
        margin_r=margin_lr,
        outline=outline,
        max_lines=4,
    )
    font = str(subtitle_font or "").strip()
    out: list[dict[str, Any]] = []
    for seg in base_segments:
        if not isinstance(seg, dict):
            continue
        seg2 = dict(seg)
        seg2["text"] = wrap_subtitle_text(
            str(seg2.get("text", "") or ""),
            layout=layout,
            font_name=font,
        )
        out.append(seg2)
    return out


def _render_srt_from_segments(segments: list[dict[str, Any]]) -> str:
    def _fmt(ms: int) -> str:
        ms = int(ms)
        if ms < 0:
            ms = 0
        hours, rem = divmod(ms, 3_600_000)
        minutes, rem = divmod(rem, 60_000)
        seconds, millis = divmod(rem, 1000)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d},{millis:03d}"

    blocks: list[str] = []
    idx = 1
    for seg in segments:
        try:
            start_ms = int(seg.get("start_ms", 0))
            end_ms = int(seg.get("end_ms", 0))
        except Exception:
            continue
        if end_ms <= start_ms:
            continue
        text = str(seg.get("text", "") or "").replace("\r", "")
        blocks.append("\n".join([str(idx), f"{_fmt(start_ms)} --> {_fmt(end_ms)}", text]))
        idx += 1
    return "\n\n".join(blocks) + "\n"


def render_video(
    project_dir: str,
    output_path: str | None = None,
    *,
    ratio: str | None = None,
    resolution: str | None = None,
    fps: int = 30,
    quality: str = "standard",
    bgm_path: str | None = None,
    bgm_volume: float = 0.18,
    narration_volume: float = 1.0,
    bgm_fade_in_s: float = 0.0,
    bgm_fade_out_s: float = 0.0,
    subtitle_color: str | None = None,
    subtitle_fade_in_ms: int | None = None,
    subtitle_fade_out_ms: int | None = None,
    subtitle_scale_from: int | None = None,
    subtitle_scale_to: int | None = None,
    subtitle_outline: int | None = None,
    subtitle_shadow: int | None = None,
    subtitle_font: str | None = None,
    subtitle_size: int | None = None,
    subtitle_y_ratio: float | None = None,
    avatar_video: str | None = None,
    avatar_pos: str | None = None,
    avatar_scale: float | None = None,
    avatar_key: str | None = None,
    avatar_key_color: str | None = None,
    avatar_key_similarity: float | None = None,
    avatar_key_blend: float | None = None,
    avatar_x_ratio: float | None = None,
    avatar_y_ratio: float | None = None,
) -> dict[str, Any]:
    root = Path(project_dir)
    manifest_path = root / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"manifest.json not found in: {root}")

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    slides = manifest.get("slides", [])
    if not slides:
        raise RenderError("No slides found in manifest.")

    normalized_ratio = _normalize_ratio(ratio or manifest.get("ratio"))
    normalized_resolution = _normalize_resolution(resolution or manifest.get("resolution"))
    normalized_fps = _normalize_fps(fps or manifest.get("fps"))
    quality_preset, crf = _quality_encode_options(quality or manifest.get("quality"))
    base_canvas_w, base_canvas_h = _ratio_canvas_size(normalized_ratio)
    canvas_w, canvas_h = _scale_canvas_size(base_canvas_w, base_canvas_h, normalized_resolution)

    try:
        outline_for_wrap = int(subtitle_outline) if subtitle_outline is not None else int(manifest.get("subtitle_outline", 1) or 1)
    except Exception:
        outline_for_wrap = 1
    outline_for_wrap = max(0, outline_for_wrap)
    font_for_wrap = (str(subtitle_font).strip() if subtitle_font else "") or str(manifest.get("subtitle_font") or "").strip()

    audio_path = root / manifest.get("outputs", {}).get("merged_audio", "audio/merged.wav")
    if not audio_path.exists():
        raise RenderError("Merged audio file is missing. Run 'voice' first.")

    video_dir = root / "video"
    audio_dir = root / "audio"
    logs_dir = root / "logs"
    video_dir.mkdir(parents=True, exist_ok=True)
    audio_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)

    target_video = Path(output_path) if output_path else video_dir / "output.mp4"
    concat_file = video_dir / "slides.ffconcat"
    temp_video = video_dir / "video_only.mp4"
    subtitle_path = root / manifest.get("outputs", {}).get("subtitle", "subtitles/subtitles.srt")
    subtitle_json_path = root / manifest.get("outputs", {}).get("subtitle_json", "subtitles/subtitles.json")

    avatar_path = _resolve_avatar_video_path(
        root=root,
        manifest=manifest,
        avatar_video=avatar_video,
    )
    avatar_position = _normalize_avatar_pos(avatar_pos or manifest.get("avatar_pos") or "bl")
    avatar_scale_ratio = _normalize_avatar_scale(avatar_scale if avatar_scale is not None else manifest.get("avatar_scale"))
    key_mode = _normalize_avatar_key_mode(avatar_key or manifest.get("avatar_key") or "auto")
    key_color = str(avatar_key_color or manifest.get("avatar_key_color") or "#00ff00").strip() or "#00ff00"
    key_similarity = _normalize_key_float(avatar_key_similarity if avatar_key_similarity is not None else manifest.get("avatar_key_similarity"), default=0.15)
    key_blend = _normalize_key_float(avatar_key_blend if avatar_key_blend is not None else manifest.get("avatar_key_blend"), default=0.02)
    x_ratio = _normalize_ratio_float(avatar_x_ratio if avatar_x_ratio is not None else manifest.get("avatar_x_ratio"))
    y_ratio = _normalize_ratio_float(avatar_y_ratio if avatar_y_ratio is not None else manifest.get("avatar_y_ratio"))

    # If advanced subtitle effects are requested, generate an ASS and burn that in.
    use_ass_effects = False
    if any(
        v is not None
        for v in (
            subtitle_fade_in_ms,
            subtitle_fade_out_ms,
            subtitle_scale_from,
            subtitle_scale_to,
            subtitle_outline,
            subtitle_shadow,
        )
    ):
        use_ass_effects = True

    # Wrap subtitles based on font size at render time.
    # This keeps text layout consistent even if subtitles were generated earlier with different settings.
    if subtitle_json_path.exists():
        wrapped_segments = _load_subtitle_segments_wrapped(
            subtitle_json_path=subtitle_json_path,
            canvas_w=canvas_w,
            canvas_h=canvas_h,
            subtitle_size=subtitle_size,
            subtitle_font=font_for_wrap or None,
            subtitle_outline=outline_for_wrap,
        )
        if wrapped_segments:
            subtitles_dir = root / "subtitles"
            subtitles_dir.mkdir(parents=True, exist_ok=True)
            if use_ass_effects:
                # ASS will be generated below using wrapped segments.
                pass
            else:
                wrapped_srt = subtitles_dir / "subtitles.wrapped.srt"
                wrapped_srt.write_text(_render_srt_from_segments(wrapped_segments), encoding="utf-8")
                subtitle_path = wrapped_srt

    if use_ass_effects and subtitle_json_path.exists():
        segments = _load_subtitle_segments_wrapped(
            subtitle_json_path=subtitle_json_path,
            canvas_w=canvas_w,
            canvas_h=canvas_h,
            subtitle_size=subtitle_size,
            subtitle_font=font_for_wrap or None,
            subtitle_outline=outline_for_wrap,
        )
        ass_out = root / "subtitles" / "subtitles.effects.ass"
        ass_out.parent.mkdir(parents=True, exist_ok=True)
        # Scale ASS play resolution and margins to match the actual output canvas,
        # so multi-line subtitles don't drift out of bounds on 9:16 / 1:1.
        base_w = 1920
        scale_w = float(canvas_w) / float(base_w)
        margin_lr = max(24, int(round(80 * scale_w)))
        fs = int(subtitle_size if subtitle_size is not None else 48)
        # Larger font / more lines → push baseline upward with a bigger bottom margin.
        margin_v = max(int(round(fs * 1.1)), int(round(60 * (float(canvas_h) / 1080.0))))
        ass_text = build_ass(
            segments=segments,
            base_color=subtitle_color or "#FFFFFF",
            fade_in_ms=subtitle_fade_in_ms if subtitle_fade_in_ms is not None else 80,
            fade_out_ms=subtitle_fade_out_ms if subtitle_fade_out_ms is not None else 120,
            scale_from=subtitle_scale_from if subtitle_scale_from is not None else 100,
            scale_to=subtitle_scale_to if subtitle_scale_to is not None else 104,
            outline=subtitle_outline if subtitle_outline is not None else 1,
            shadow=subtitle_shadow if subtitle_shadow is not None else 0,
            font=subtitle_font or "",
            font_size=subtitle_size if subtitle_size is not None else 48,
            play_res_x=canvas_w,
            play_res_y=canvas_h,
            margin_l=margin_lr,
            margin_r=margin_lr,
            margin_v=margin_v,
            subtitle_y_ratio=subtitle_y_ratio,
        )
        ass_out.write_text(ass_text, encoding="utf-8")
        subtitle_path = ass_out

    _write_ffconcat(root=root, slides=slides, concat_file=concat_file)
    ffmpeg = _get_ffmpeg_path()

    # Optional: mix background music with narration audio, or adjust narration gain
    # even when no BGM is present so the volume knob always affects the output.
    mix_used = False
    mix_info = ""
    mux_audio_path = audio_path
    narration_gain = float(narration_volume)
    if bgm_path:
        bgm = Path(bgm_path)
        if not bgm.exists():
            raise RenderError(f"BGM file not found: {bgm}")
        narration_duration_s = _read_wav_duration_s(audio_path)
        fade_in = max(0.0, float(bgm_fade_in_s or 0.0))
        fade_out = max(0.0, float(bgm_fade_out_s or 0.0))
        fade_out_start = max(0.0, narration_duration_s - fade_out)
        # Write AAC audio for muxing (smaller/faster than WAV for intermediate).
        mixed_audio = audio_dir / "mixed.m4a"
        _run_ffmpeg(
            [
                ffmpeg,
                "-y",
                "-i",
                str(audio_path),
                "-stream_loop",
                "-1",
                "-i",
                str(bgm),
                "-filter_complex",
                (
                    f"[0:a]volume={narration_gain:.3f}[nar];"
                    f"[1:a]volume={bgm_volume:.3f},"
                    f"afade=t=in:st=0:d={fade_in:.3f},"
                    f"afade=t=out:st={fade_out_start:.3f}:d={fade_out:.3f}"
                    f"[bgm];"
                    # amix defaults to normalize=1, which rescales inputs toward full scale and
                    # largely cancels small BGM volume factors. normalize=0 keeps linear sum gain.
                    f"[nar][bgm]amix=inputs=2:duration=first:dropout_transition=2:normalize=0,"
                    f"alimiter=limit=0.98:level=0[a]"
                ),
                "-map",
                "[a]",
                "-c:a",
                "aac",
                "-b:a",
                "192k",
                "-shortest",
                str(mixed_audio),
            ]
        )
        mix_used = True
        mix_info = (
            f"bgm: {bgm}\n"
            f"bgm_volume: {bgm_volume}\n"
            f"narration_volume: {narration_gain}\n"
            f"bgm_fade_in_s: {fade_in}\n"
            f"bgm_fade_out_s: {fade_out}\n"
            f"narration_duration_s: {narration_duration_s:.3f}\n"
        )
        mux_audio_path = mixed_audio
    elif abs(narration_gain - 1.0) > 1e-9:
        adjusted_audio = audio_dir / "narration_adjusted.m4a"
        _run_ffmpeg(
            [
                ffmpeg,
                "-y",
                "-i",
                str(audio_path),
                "-filter:a",
                f"volume={narration_gain:.3f}",
                "-c:a",
                "aac",
                "-b:a",
                "192k",
                str(adjusted_audio),
            ]
        )
        mix_info = f"narration_volume: {narration_gain}\n"
        mux_audio_path = adjusted_audio

    _run_ffmpeg(
        [
            ffmpeg,
            "-y",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            str(concat_file),
            "-vf",
            # Standardize output canvas for predictable subtitle sizing.
            # Keep aspect ratio and pad to avoid stretching.
            f"scale={canvas_w}:{canvas_h}:force_original_aspect_ratio=decrease,"
            f"pad={canvas_w}:{canvas_h}:(ow-iw)/2:(oh-ih)/2,"
            "setsar=1,"
            f"fps={normalized_fps},format=yuv420p",
            "-r",
            str(normalized_fps),
            "-c:v",
            "libx264",
            "-preset",
            quality_preset,
            "-crf",
            crf,
            "-pix_fmt",
            "yuv420p",
            str(temp_video),
        ]
    )

    subtitle_burned = False
    mux_command = [
        ffmpeg,
        "-y",
        "-i",
        str(temp_video),
        "-i",
        str(mux_audio_path),
        "-c:v",
        "copy",
        "-c:a",
        "aac",
        "-shortest",
    ]

    def _mux_with_filters(*, subtitle_ok: bool) -> None:
        nonlocal subtitle_burned
        filters, vmap = _build_video_filtergraph(
            canvas_w=canvas_w,
            canvas_h=canvas_h,
            subtitle_path=(subtitle_path if subtitle_ok and subtitle_path.exists() else None),
            subtitle_color=subtitle_color,
            subtitle_font=subtitle_font,
            subtitle_size=subtitle_size,
            avatar_path=avatar_path if avatar_path and avatar_path.exists() else None,
            avatar_pos=avatar_position,
            avatar_scale=avatar_scale_ratio,
            avatar_key=key_mode,
            avatar_key_color=key_color,
            avatar_key_similarity=key_similarity,
            avatar_key_blend=key_blend,
            avatar_x_ratio=x_ratio,
            avatar_y_ratio=y_ratio,
        )
        cmd = [
            ffmpeg,
            "-y",
            "-i",
            str(temp_video),
            "-i",
            str(mux_audio_path),
        ]
        if avatar_path and avatar_path.exists():
            cmd += ["-i", str(avatar_path)]
        cmd += [
            "-filter_complex",
            filters,
            "-map",
            vmap,
            "-map",
            "1:a",
            "-c:v",
            "libx264",
            "-preset",
            quality_preset,
            "-crf",
            crf,
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "aac",
            "-shortest",
            str(target_video),
        ]
        _run_ffmpeg(cmd)
        subtitle_burned = bool(subtitle_ok and subtitle_path.exists())

    if subtitle_path.exists():
        try:
            _mux_with_filters(subtitle_ok=True)
        except RenderError:
            # Fallback: no subtitles and no avatar overlay.
            _run_ffmpeg(mux_command + [str(target_video)])
    else:
        # Fast path: no subtitle and no avatar overlay.
        if avatar_path and avatar_path.exists():
            _mux_with_filters(subtitle_ok=False)
        else:
            _run_ffmpeg(mux_command + [str(target_video)])

    outputs = manifest.setdefault("outputs", {})
    manifest["ratio"] = normalized_ratio
    manifest["resolution"] = normalized_resolution
    manifest["fps"] = normalized_fps
    manifest["quality"] = quality
    outputs["video"] = _relative_path(root, target_video)
    outputs["video_subtitles_burned"] = subtitle_burned
    if avatar_path and avatar_path.exists():
        outputs["avatar_video"] = _relative_path(root, avatar_path)
        manifest["avatar_pos"] = avatar_position
        manifest["avatar_scale"] = avatar_scale_ratio
        manifest["avatar_key"] = key_mode
        manifest["avatar_key_color"] = key_color
        manifest["avatar_key_similarity"] = key_similarity
        manifest["avatar_key_blend"] = key_blend
        if x_ratio is not None:
            manifest["avatar_x_ratio"] = x_ratio
        if y_ratio is not None:
            manifest["avatar_y_ratio"] = y_ratio
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    _cleanup_render_intermediates(temp_video=temp_video, concat_file=concat_file)

    (logs_dir / "render.log").write_text(
        f"video: {target_video}\n"
        f"ratio: {normalized_ratio}\n"
        f"resolution: {normalized_resolution}\n"
        f"fps: {normalized_fps}\n"
        f"quality: {quality}\n"
        f"canvas: {canvas_w}x{canvas_h}\n"
        f"subtitles_burned: {subtitle_burned}\n"
        f"avatar: {avatar_path or ''}\n"
        f"avatar_pos: {avatar_position}\n"
        f"avatar_scale: {avatar_scale_ratio}\n"
        f"avatar_key: {key_mode}\n"
        f"avatar_key_color: {key_color}\n"
        f"avatar_key_similarity: {key_similarity}\n"
        f"avatar_key_blend: {key_blend}\n"
        f"avatar_x_ratio: {'' if x_ratio is None else x_ratio}\n"
        f"avatar_y_ratio: {'' if y_ratio is None else y_ratio}\n"
        f"audio: {mux_audio_path}\n"
        f"mixed: {mix_used}\n"
        f"{mix_info}",
        encoding="utf-8",
    )

    return {
        "video": str(target_video),
        "subtitles_burned": subtitle_burned,
        "slide_count": len(slides),
        "mixed_audio": str(mux_audio_path) if mix_used else None,
    }


def _write_ffconcat(*, root: Path, slides: list[dict[str, Any]], concat_file: Path) -> None:
    lines = ["ffconcat version 1.0"]

    for slide in slides:
        image_rel = slide.get("image", "")
        duration_ms = int(slide.get("duration_ms", 0))
        if not image_rel:
            continue
        image_path = root / image_rel
        if not image_path.exists():
            raise RenderError(f"Slide image missing: {image_path}")

        duration_s = max(duration_ms, 300) / 1000
        lines.append(f"file '{_ffconcat_escape(str(image_path.resolve()))}'")
        lines.append(f"duration {duration_s:.3f}")

    last_image = root / slides[-1].get("image", "")
    if last_image.exists():
        lines.append(f"file '{_ffconcat_escape(str(last_image.resolve()))}'")

    concat_file.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _get_ffmpeg_path() -> str:
    try:
        import imageio_ffmpeg
    except ImportError as exc:  # pragma: no cover
        raise RenderError("imageio-ffmpeg is required for video rendering.") from exc
    return imageio_ffmpeg.get_ffmpeg_exe()


def _run_ffmpeg(command: list[str]) -> None:
    # Force a stable decode on Windows; ffmpeg may emit non-GBK bytes.
    result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if result.returncode != 0:
        raise RenderError(result.stderr.strip() or "ffmpeg command failed.")


def _read_wav_duration_s(path: Path) -> float:
    try:
        with closing(wave.open(str(path), "rb")) as wf:
            frames = wf.getnframes()
            rate = wf.getframerate()
            if rate <= 0:
                return 0.0
            return float(frames) / float(rate)
    except Exception as exc:
        raise RenderError(f"Unable to read WAV duration: {path} ({exc})") from exc


def _ass_primary_colour_from_rgb_hex(value: str) -> str:
    """
    ASS colour format is &HAABBGGRR (hex).
    Input supports #RRGGBB / RRGGBB.
    """
    raw = (value or "").strip()
    if raw.startswith("#"):
        raw = raw[1:]
    if len(raw) != 6 or any(ch not in "0123456789abcdefABCDEF" for ch in raw):
        raise RenderError(f"Invalid subtitle_color: {value!r}. Expected #RRGGBB.")
    rr = raw[0:2]
    gg = raw[2:4]
    bb = raw[4:6]
    # AA=00 (opaque), then BB GG RR
    return f"&H00{bb}{gg}{rr}"


def _build_subtitle_filter(
    subtitle_path: Path,
    *,
    subtitle_color: str | None = None,
    subtitle_font: str | None = None,
    subtitle_size: int | None = None,
) -> str:
    value = str(subtitle_path.resolve()).replace("\\", "/")
    value = value.replace(":", "\\:").replace("'", r"\'")
    style_parts: list[str] = []
    if subtitle_color:
        primary = _ass_primary_colour_from_rgb_hex(subtitle_color)
        style_parts.append(f"PrimaryColour={primary}")
        # Keep outline for readability by default.
        style_parts.append("Outline=1")
        style_parts.append("Shadow=0")

    font = (subtitle_font or "").strip()
    if font:
        # ASS style: FontName supports font family names. Keep it simple and avoid quoting;
        # we escape the comma separators for ffmpeg's force_style.
        style_parts.append(f"FontName={font}")

    if subtitle_size is not None:
        try:
            size = int(subtitle_size)
        except Exception:
            size = 0
        if size > 0:
            style_parts.append(f"FontSize={size}")

    if style_parts:
        # ffmpeg expects commas inside force_style to be escaped.
        style = "\\,".join(style_parts)
        return f"subtitles='{value}':force_style='{style}'"
    return f"subtitles='{value}'"


def _normalize_avatar_pos(value: str) -> str:
    raw = str(value or "").strip().lower()
    mapping = {
        "bl": "bl",
        "bottom-left": "bl",
        "left-bottom": "bl",
        "lb": "bl",
        "br": "br",
        "bottom-right": "br",
        "right-bottom": "br",
        "rb": "br",
        "tl": "tl",
        "top-left": "tl",
        "lt": "tl",
        "tr": "tr",
        "top-right": "tr",
        "rt": "tr",
    }
    return mapping.get(raw, "bl")


def _normalize_avatar_scale(value: Any) -> float:
    try:
        v = float(value)
    except Exception:
        v = 0.25
    # Keep reasonable bounds.
    return max(0.05, min(0.8, v))


def _normalize_avatar_key_mode(value: str) -> str:
    raw = str(value or "").strip().lower()
    if raw in {"", "auto"}:
        return "auto"
    if raw in {"none", "off", "false", "0"}:
        return "none"
    if raw in {"green", "blue", "custom"}:
        return raw
    return "auto"


def _normalize_key_float(value: Any, *, default: float) -> float:
    try:
        v = float(value)
    except Exception:
        v = default
    return max(0.0, min(1.0, v))


def _normalize_ratio_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        v = float(value)
    except Exception:
        return None
    if v < 0.0:
        v = 0.0
    if v > 1.0:
        v = 1.0
    return v


def _hex_to_0xrrggbb(value: str) -> str:
    raw = (value or "").strip()
    if raw.startswith("#"):
        raw = raw[1:]
    raw = raw.lower()
    if len(raw) == 3:
        raw = "".join([ch * 2 for ch in raw])
    if len(raw) != 6 or any(ch not in "0123456789abcdef" for ch in raw):
        return "0x00ff00"
    return f"0x{raw}"


def _resolve_avatar_video_path(*, root: Path, manifest: dict[str, Any], avatar_video: str | None) -> Path | None:
    if avatar_video is not None and str(avatar_video).strip():
        p = Path(str(avatar_video)).expanduser()
        if not p.is_absolute():
            p = (root / p).resolve()
        return p
    rel = manifest.get("outputs", {}).get("avatar_video") if isinstance(manifest.get("outputs", {}), dict) else None
    if rel:
        return (root / str(rel)).resolve()
    default = root / "avatar" / "avatar.mp4"
    return default if default.exists() else None


def _build_video_filtergraph(
    *,
    canvas_w: int,
    canvas_h: int,
    subtitle_path: Path | None,
    subtitle_color: str | None,
    subtitle_font: str | None,
    subtitle_size: int | None,
    avatar_path: Path | None,
    avatar_pos: str,
    avatar_scale: float,
    avatar_key: str,
    avatar_key_color: str,
    avatar_key_similarity: float,
    avatar_key_blend: float,
    avatar_x_ratio: float | None,
    avatar_y_ratio: float | None,
) -> tuple[str, str]:
    """
    Build a filter_complex graph that optionally overlays an avatar video (PiP)
    and optionally burns subtitles.

    Inputs:
    - 0:v: base video (slides video_only.mp4)
    - 2:v: avatar video (if present)
    """
    chain = ""
    last = "[0:v]"

    if avatar_path is not None:
        margin = max(12, int(round(canvas_w * 0.03)))
        w = int(round(canvas_w * float(avatar_scale)))
        w = max(64, min(canvas_w, w))
        # Keep aspect ratio; ensure even dims by using -2 for height.
        key_color = _hex_to_0xrrggbb(avatar_key_color)
        if avatar_key in {"auto", "green"}:
            key_color = "0x00ff00"
        elif avatar_key == "blue":
            key_color = "0x0000ff"
        if avatar_key != "none":
            chain += f"[2:v]colorkey={key_color}:{avatar_key_similarity:.3f}:{avatar_key_blend:.3f},format=rgba,scale={w}:-2[av];"
        else:
            chain += f"[2:v]scale={w}:-2[av];"
        if avatar_x_ratio is not None or avatar_y_ratio is not None:
            xr = 0.0 if avatar_x_ratio is None else float(avatar_x_ratio)
            yr = 0.0 if avatar_y_ratio is None else float(avatar_y_ratio)
            # Position refers to the overlay's top-left corner relative to available space.
            x = f"{margin}+(main_w-overlay_w-2*{margin})*{xr:.3f}"
            y = f"{margin}+(main_h-overlay_h-2*{margin})*{yr:.3f}"
        else:
            if avatar_pos == "bl":
                x, y = f"{margin}", f"main_h-overlay_h-{margin}"
            elif avatar_pos == "br":
                x, y = f"main_w-overlay_w-{margin}", f"main_h-overlay_h-{margin}"
            elif avatar_pos == "tl":
                x, y = f"{margin}", f"{margin}"
            else:  # tr
                x, y = f"main_w-overlay_w-{margin}", f"{margin}"
        chain += f"{last}[av]overlay=x={x}:y={y}:shortest=1[v0];"
        last = "[v0]"

    if subtitle_path is not None and subtitle_path.exists():
        sub = _build_subtitle_filter(
            subtitle_path,
            subtitle_color=subtitle_color,
            subtitle_font=subtitle_font,
            subtitle_size=subtitle_size,
        )
        chain += f"{last}{sub}[vout]"
        return chain, "[vout]"

    # No subtitles: still need to label output for mapping.
    chain += f"{last}null[vout]"
    return chain, "[vout]"


def _load_subtitle_segments_for_ass(
    *,
    subtitle_json_path: Path,
) -> list[dict[str, Any]]:
    # Backward-compatible alias: kept for older imports/tests.
    payload = json.loads(subtitle_json_path.read_text(encoding="utf-8"))
    base_segments = payload.get("segments") or []
    if not isinstance(base_segments, list):
        return []
    return [seg for seg in base_segments if isinstance(seg, dict)]


def _cleanup_render_intermediates(*, temp_video: Path, concat_file: Path) -> None:
    for path in (temp_video, concat_file):
        try:
            if path.exists():
                path.unlink()
        except OSError:
            # Keep the final render successful even if cleanup fails.
            pass


def _ffconcat_escape(path: str) -> str:
    return path.replace("'", r"'\''")


def _relative_path(root: Path, path: Path) -> str:
    try:
        return str(path.relative_to(root)).replace("\\", "/")
    except ValueError:
        return str(path)
