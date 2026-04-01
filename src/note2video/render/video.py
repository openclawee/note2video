from __future__ import annotations

import json
import subprocess
from pathlib import Path
from typing import Any


class RenderError(RuntimeError):
    """Raised when final video rendering fails."""


def render_video(project_dir: str, output_path: str | None = None) -> dict[str, Any]:
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
    logs_dir = root / "logs"
    video_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)

    target_video = Path(output_path) if output_path else video_dir / "output.mp4"
    concat_file = video_dir / "slides.ffconcat"
    temp_video = video_dir / "video_only.mp4"
    subtitle_path = root / manifest.get("outputs", {}).get("subtitle", "subtitles/subtitles.srt")

    _write_ffconcat(root=root, slides=slides, concat_file=concat_file)
    ffmpeg = _get_ffmpeg_path()

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
        str(audio_path),
        "-c:v",
        "copy",
        "-c:a",
        "aac",
        "-shortest",
    ]

    if subtitle_path.exists():
        subtitle_filter = _build_subtitle_filter(subtitle_path)
        mux_with_subtitles = [
            ffmpeg,
            "-y",
            "-i",
            str(temp_video),
            "-i",
            str(audio_path),
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
        f"video: {target_video}\nsubtitles_burned: {subtitle_burned}\n",
        encoding="utf-8",
    )

    return {
        "video": str(target_video),
        "subtitles_burned": subtitle_burned,
        "slide_count": len(slides),
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
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        raise RenderError(result.stderr.strip() or "ffmpeg command failed.")


def _build_subtitle_filter(subtitle_path: Path) -> str:
    value = str(subtitle_path.resolve()).replace("\\", "/")
    value = value.replace(":", "\\:").replace("'", r"\'")
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
