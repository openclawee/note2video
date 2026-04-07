from __future__ import annotations

import json
import subprocess
import wave
from contextlib import closing
from pathlib import Path
from typing import Any


class RenderError(RuntimeError):
    """Raised when final video rendering fails."""


def render_video(
    project_dir: str,
    output_path: str | None = None,
    *,
    bgm_path: str | None = None,
    bgm_volume: float = 0.18,
    narration_volume: float = 1.0,
    bgm_fade_in_s: float = 0.0,
    bgm_fade_out_s: float = 0.0,
    subtitle_color: str | None = None,
) -> dict[str, Any]:
    root = Path(project_dir)
    manifest_path = root / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"manifest.json not found in: {root}")

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    slides = manifest.get("slides", [])
    if not slides:
        raise RenderError("No slides found in manifest.")

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

    _write_ffconcat(root=root, slides=slides, concat_file=concat_file)
    ffmpeg = _get_ffmpeg_path()

    # Optional: mix background music with narration audio.
    mix_used = False
    mix_info = ""
    mux_audio_path = audio_path
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
                    f"[0:a]volume={narration_volume:.3f}[nar];"
                    f"[1:a]volume={bgm_volume:.3f},"
                    f"afade=t=in:st=0:d={fade_in:.3f},"
                    f"afade=t=out:st={fade_out_start:.3f}:d={fade_out:.3f}"
                    f"[bgm];"
                    f"[nar][bgm]amix=inputs=2:duration=first:dropout_transition=2,alimiter=limit=0.98[a]"
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
            f"narration_volume: {narration_volume}\n"
            f"bgm_fade_in_s: {fade_in}\n"
            f"bgm_fade_out_s: {fade_out}\n"
            f"narration_duration_s: {narration_duration_s:.3f}\n"
        )
        mux_audio_path = mixed_audio

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
            "fps=30,format=yuv420p",
            "-r",
            "30",
            "-c:v",
            "libx264",
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

    if subtitle_path.exists():
        subtitle_filter = _build_subtitle_filter(subtitle_path, subtitle_color=subtitle_color)
        mux_with_subtitles = [
            ffmpeg,
            "-y",
            "-i",
            str(temp_video),
            "-i",
            str(mux_audio_path),
            "-vf",
            subtitle_filter,
            "-c:v",
            "libx264",
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "aac",
            "-shortest",
            str(target_video),
        ]
        try:
            _run_ffmpeg(mux_with_subtitles)
            subtitle_burned = True
        except RenderError:
            _run_ffmpeg(mux_command + [str(target_video)])
    else:
        _run_ffmpeg(mux_command + [str(target_video)])

    outputs = manifest.setdefault("outputs", {})
    outputs["video"] = _relative_path(root, target_video)
    outputs["video_subtitles_burned"] = subtitle_burned
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    _cleanup_render_intermediates(temp_video=temp_video, concat_file=concat_file)

    (logs_dir / "render.log").write_text(
        f"video: {target_video}\nsubtitles_burned: {subtitle_burned}\n"
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


def _build_subtitle_filter(subtitle_path: Path, *, subtitle_color: str | None = None) -> str:
    value = str(subtitle_path.resolve()).replace("\\", "/")
    value = value.replace(":", "\\:").replace("'", r"\'")
    if subtitle_color:
        primary = _ass_primary_colour_from_rgb_hex(subtitle_color)
        # Keep defaults but set primary colour; outline kept for readability.
        # ffmpeg expects commas inside force_style to be escaped.
        style = f"PrimaryColour={primary}\\,Outline=1\\,Shadow=0"
        return f"subtitles='{value}':force_style='{style}'"
    return f"subtitles='{value}'"


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
