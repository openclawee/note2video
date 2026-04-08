from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from note2video.parser.extract import extract_project
from note2video.render.video import render_video
from note2video.subtitle.generate import generate_subtitles
from note2video.tts.voice import generate_voice_assets, list_available_voices


@dataclass(frozen=True)
class BuildRequest:
    input_file: str
    out_dir: str
    pages: str = "all"
    tts_provider: str = "pyttsx3"
    voice_id: str = ""
    tts_rate: float = 1.0
    bgm_path: str | None = None
    bgm_volume: float = 0.18
    bgm_fade_in_s: float = 0.0
    bgm_fade_out_s: float = 0.0
    narration_volume: float = 1.0
    subtitle_color: str | None = None
    subtitle_highlight_mode: str | None = None
    subtitle_highlight_color: str | None = None
    subtitle_fade_in_ms: int = 80
    subtitle_fade_out_ms: int = 120
    subtitle_scale_from: int = 100
    subtitle_scale_to: int = 104
    subtitle_outline: int = 1
    subtitle_shadow: int = 0
    subtitle_font: str | None = None
    subtitle_size: int | None = None


@dataclass(frozen=True)
class ExtractRequest:
    input_file: str
    out_dir: str
    pages: str = "all"


@dataclass(frozen=True)
class VoiceRequest:
    input_json: str
    out_dir: str
    tts_provider: str = "pyttsx3"
    voice_id: str = ""
    tts_rate: float = 1.0


@dataclass(frozen=True)
class VoicesRequest:
    tts_provider: str = "edge"
    keyword: str = ""


@dataclass(frozen=True)
class SubtitleRequest:
    input_json: str
    out_dir: str


@dataclass(frozen=True)
class RenderRequest:
    project_dir: str
    output_path: str | None = None
    bgm_path: str | None = None
    bgm_volume: float = 0.18
    bgm_fade_in_s: float = 0.0
    bgm_fade_out_s: float = 0.0
    narration_volume: float = 1.0
    subtitle_color: str | None = None
    subtitle_highlight_mode: str | None = None
    subtitle_highlight_color: str | None = None
    subtitle_fade_in_ms: int = 80
    subtitle_fade_out_ms: int = 120
    subtitle_scale_from: int = 100
    subtitle_scale_to: int = 104
    subtitle_outline: int = 1
    subtitle_shadow: int = 0
    subtitle_font: str | None = None
    subtitle_size: int | None = None


class PipelineServiceError(RuntimeError):
    """Raised when pipeline orchestration fails."""


class PipelineService:
    def build(self, request: BuildRequest) -> dict[str, Any]:
        return run_build_pipeline(request)

    def extract(self, request: ExtractRequest) -> Any:
        return extract_project(request.input_file, request.out_dir, pages=request.pages)

    def voice(self, request: VoiceRequest) -> dict[str, Any]:
        return generate_voice_assets(
            request.input_json,
            request.out_dir,
            provider_name=request.tts_provider,
            voice_id=request.voice_id,
            tts_rate=request.tts_rate,
            minimax_base_url=None,
        )

    def voices(self, request: VoicesRequest) -> list[dict[str, Any]]:
        return list_available_voices(
            provider_name=request.tts_provider,
            keyword=request.keyword,
            minimax_base_url=None,
        )

    def subtitle(self, request: SubtitleRequest) -> dict[str, Any]:
        return generate_subtitles(request.input_json, request.out_dir)

    def render(self, request: RenderRequest) -> dict[str, Any]:
        return render_video(
            request.project_dir,
            request.output_path,
            bgm_path=request.bgm_path,
            bgm_volume=float(request.bgm_volume),
            bgm_fade_in_s=float(request.bgm_fade_in_s),
            bgm_fade_out_s=float(request.bgm_fade_out_s),
            narration_volume=float(request.narration_volume),
            subtitle_color=request.subtitle_color,
            subtitle_highlight_mode=request.subtitle_highlight_mode,
            subtitle_highlight_color=request.subtitle_highlight_color,
            subtitle_fade_in_ms=int(request.subtitle_fade_in_ms),
            subtitle_fade_out_ms=int(request.subtitle_fade_out_ms),
            subtitle_scale_from=int(request.subtitle_scale_from),
            subtitle_scale_to=int(request.subtitle_scale_to),
            subtitle_outline=int(request.subtitle_outline),
            subtitle_shadow=int(request.subtitle_shadow),
            subtitle_font=request.subtitle_font,
            subtitle_size=request.subtitle_size,
        )


def run_build_pipeline(
    request: BuildRequest,
    *,
    extract_project_fn: Callable[..., Any] = extract_project,
    generate_voice_assets_fn: Callable[..., dict[str, Any]] = generate_voice_assets,
    generate_subtitles_fn: Callable[..., dict[str, Any]] = generate_subtitles,
    render_video_fn: Callable[..., dict[str, Any]] = render_video,
) -> dict[str, Any]:
    out_dir = Path(request.out_dir)
    manifest = extract_project_fn(request.input_file, str(out_dir), pages=request.pages)
    script_path = out_dir / "scripts" / "script.json"

    voice_result = generate_voice_assets_fn(
        str(script_path),
        str(out_dir),
        provider_name=request.tts_provider or "pyttsx3",
        voice_id=request.voice_id,
        tts_rate=request.tts_rate,
        minimax_base_url=None,
    )
    subtitle_result = generate_subtitles_fn(str(script_path), str(out_dir))
    render_kwargs = {
        "bgm_path": request.bgm_path,
        "bgm_volume": float(request.bgm_volume),
        "bgm_fade_in_s": float(request.bgm_fade_in_s),
        "bgm_fade_out_s": float(request.bgm_fade_out_s),
        "narration_volume": float(request.narration_volume),
        "subtitle_color": request.subtitle_color,
    }
    advanced_kwargs = {
        "subtitle_highlight_mode": request.subtitle_highlight_mode,
        "subtitle_highlight_color": request.subtitle_highlight_color,
        "subtitle_fade_in_ms": int(request.subtitle_fade_in_ms),
        "subtitle_fade_out_ms": int(request.subtitle_fade_out_ms),
        "subtitle_scale_from": int(request.subtitle_scale_from),
        "subtitle_scale_to": int(request.subtitle_scale_to),
        "subtitle_outline": int(request.subtitle_outline),
        "subtitle_shadow": int(request.subtitle_shadow),
        "subtitle_font": request.subtitle_font,
        "subtitle_size": request.subtitle_size,
    }
    try:
        render_result = render_video_fn(str(out_dir), **render_kwargs, **advanced_kwargs)
    except TypeError:
        # Backward-compatibility for tests/mocks that still use an older render signature.
        render_result = render_video_fn(str(out_dir), **render_kwargs)

    return {
        "output_dir": str(out_dir),
        "artifacts": {
            "manifest": "manifest.json",
            "notes": "notes/notes.json",
            "script": "scripts/script.json",
            "audio_dir": "audio",
            "merged_audio": "audio/merged.wav",
            "subtitle": "subtitles/subtitles.srt",
            "subtitle_json": "subtitles/subtitles.json",
            "video": render_result["video"],
        },
        "slide_count": int(getattr(manifest, "slide_count", 0)),
        "segment_count": int(subtitle_result["segment_count"]),
        "voice_provider": str(voice_result["provider"]),
        "subtitles_burned": bool(render_result["subtitles_burned"]),
        "mixed_audio": bool(render_result.get("mixed_audio")),
    }


def run_extract_pipeline(
    request: ExtractRequest,
    *,
    extract_project_fn: Callable[..., Any] = extract_project,
) -> dict[str, Any]:
    manifest = extract_project_fn(request.input_file, request.out_dir, pages=request.pages)
    return {
        "output_dir": str(Path(request.out_dir)),
        "artifacts": dict(getattr(manifest, "outputs", {})),
        "slide_count": int(getattr(manifest, "slide_count", 0)),
    }


def run_voice_pipeline(
    request: VoiceRequest,
    *,
    generate_voice_assets_fn: Callable[..., dict[str, Any]] = generate_voice_assets,
) -> dict[str, Any]:
    result = generate_voice_assets_fn(
        request.input_json,
        request.out_dir,
        provider_name=request.tts_provider,
        voice_id=request.voice_id,
        tts_rate=request.tts_rate,
        minimax_base_url=None,
    )
    return {
        "input": request.input_json,
        "output_dir": str(Path(request.out_dir)),
        "slide_count": int(result["slide_count"]),
        "provider": str(result["provider"]),
        "voice": str(result["voice"]),
        "tts_rate": float(result["tts_rate"]),
        "artifacts": {
            "audio_dir": "audio",
            "merged_audio": "audio/merged.wav",
            "timings": "audio/timings.json",
        },
    }


def run_voices_pipeline(
    request: VoicesRequest,
    *,
    list_available_voices_fn: Callable[..., list[dict[str, Any]]] = list_available_voices,
) -> dict[str, Any]:
    voices = list_available_voices_fn(
        provider_name=request.tts_provider,
        keyword=request.keyword,
        minimax_base_url=None,
    )
    return {
        "provider": request.tts_provider,
        "count": len(voices),
        "voices": voices,
    }


def run_subtitle_pipeline(
    request: SubtitleRequest,
    *,
    generate_subtitles_fn: Callable[..., dict[str, Any]] = generate_subtitles,
) -> dict[str, Any]:
    result = generate_subtitles_fn(request.input_json, request.out_dir)
    return {
        "input": request.input_json,
        "output_dir": str(Path(request.out_dir)),
        "slide_count": int(result["slide_count"]),
        "segment_count": int(result["segment_count"]),
        "artifacts": {
            "subtitle": "subtitles/subtitles.srt",
            "subtitle_json": "subtitles/subtitles.json",
        },
    }


def run_render_pipeline(
    request: RenderRequest,
    *,
    render_video_fn: Callable[..., dict[str, Any]] = render_video,
) -> dict[str, Any]:
    result = render_video_fn(
        request.project_dir,
        request.output_path,
        bgm_path=request.bgm_path,
        bgm_volume=float(request.bgm_volume),
        bgm_fade_in_s=float(request.bgm_fade_in_s),
        bgm_fade_out_s=float(request.bgm_fade_out_s),
        narration_volume=float(request.narration_volume),
        subtitle_color=request.subtitle_color,
        subtitle_highlight_mode=request.subtitle_highlight_mode,
        subtitle_highlight_color=request.subtitle_highlight_color,
        subtitle_fade_in_ms=int(request.subtitle_fade_in_ms),
        subtitle_fade_out_ms=int(request.subtitle_fade_out_ms),
        subtitle_scale_from=int(request.subtitle_scale_from),
        subtitle_scale_to=int(request.subtitle_scale_to),
        subtitle_outline=int(request.subtitle_outline),
        subtitle_shadow=int(request.subtitle_shadow),
        subtitle_font=request.subtitle_font,
        subtitle_size=request.subtitle_size,
    )
    return {
        "input": request.project_dir,
        "output_dir": str(Path(request.project_dir)),
        "slide_count": int(result["slide_count"]),
        "artifacts": {"video": str(result["video"])},
        "subtitles_burned": bool(result["subtitles_burned"]),
    }


# Backward-friendly aliases with simpler names for adapters.
def build_pipeline(request: BuildRequest) -> dict[str, Any]:
    return run_build_pipeline(request)


def extract_pipeline(request: ExtractRequest) -> dict[str, Any]:
    return run_extract_pipeline(request)


def voice_pipeline(request: VoiceRequest) -> dict[str, Any]:
    return run_voice_pipeline(request)


def voices_pipeline(request: VoicesRequest) -> dict[str, Any]:
    return run_voices_pipeline(request)


def subtitle_pipeline(request: SubtitleRequest) -> dict[str, Any]:
    return run_subtitle_pipeline(request)


def render_pipeline(request: RenderRequest) -> dict[str, Any]:
    return run_render_pipeline(request)
