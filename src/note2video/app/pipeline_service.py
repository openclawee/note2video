from __future__ import annotations

from dataclasses import dataclass
import json
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
    ratio: str = "16:9"
    resolution: str = "1080p"
    fps: int = 30
    quality: str = "standard"
    tts_provider: str = "pyttsx3"
    voice_id: str = ""
    tts_rate: float = 1.0
    script_text: str | None = None
    script_file: str | None = None
    bgm_path: str | None = None
    bgm_volume: float = 0.18
    bgm_fade_in_s: float = 0.0
    bgm_fade_out_s: float = 0.0
    narration_volume: float = 1.0
    subtitle_color: str | None = None
    subtitle_fade_in_ms: int = 80
    subtitle_fade_out_ms: int = 120
    subtitle_scale_from: int = 100
    subtitle_scale_to: int = 104
    subtitle_outline: int = 1
    subtitle_shadow: int = 0
    subtitle_font: str | None = None
    subtitle_size: int | None = None
    subtitle_y_ratio: float | None = None


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
    ratio: str = "16:9"
    resolution: str = "1080p"
    fps: int = 30
    quality: str = "standard"
    bgm_path: str | None = None
    bgm_volume: float = 0.18
    bgm_fade_in_s: float = 0.0
    bgm_fade_out_s: float = 0.0
    narration_volume: float = 1.0
    subtitle_color: str | None = None
    subtitle_fade_in_ms: int = 80
    subtitle_fade_out_ms: int = 120
    subtitle_scale_from: int = 100
    subtitle_scale_to: int = 104
    subtitle_outline: int = 1
    subtitle_shadow: int = 0
    subtitle_font: str | None = None
    subtitle_size: int | None = None
    subtitle_y_ratio: float | None = None


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
            ratio=request.ratio,
            bgm_path=request.bgm_path,
            bgm_volume=float(request.bgm_volume),
            bgm_fade_in_s=float(request.bgm_fade_in_s),
            bgm_fade_out_s=float(request.bgm_fade_out_s),
            narration_volume=float(request.narration_volume),
            subtitle_color=request.subtitle_color,
            subtitle_fade_in_ms=int(request.subtitle_fade_in_ms),
            subtitle_fade_out_ms=int(request.subtitle_fade_out_ms),
            subtitle_scale_from=int(request.subtitle_scale_from),
            subtitle_scale_to=int(request.subtitle_scale_to),
            subtitle_outline=int(request.subtitle_outline),
            subtitle_shadow=int(request.subtitle_shadow),
            subtitle_font=request.subtitle_font,
            subtitle_size=request.subtitle_size,
            subtitle_y_ratio=request.subtitle_y_ratio,
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
    manifest = extract_project_fn(
        request.input_file,
        str(out_dir),
        request.pages,
        ratio=request.ratio,
        resolution=request.resolution,
        fps=int(request.fps),
        quality=request.quality,
    )
    script_path = out_dir / "scripts" / "script.json"
    _merge_build_options_into_manifest(
        out_dir / "manifest.json",
        subtitle_font=request.subtitle_font,
        subtitle_size=request.subtitle_size,
        subtitle_outline=request.subtitle_outline,
    )

    script_override = _resolve_script_override_text(request)
    if script_override is not None and str(script_override).strip():
        slide_count = int(getattr(manifest, "slide_count", 0) or 0)
        manifest_path = out_dir / "manifest.json"
        slides_meta = _load_manifest_slides_meta(manifest_path) if manifest_path.exists() else None
        _write_script_override(
            script_path=script_path,
            slides_meta=slides_meta,
            slide_count=slide_count,
            script_text=str(script_override),
        )

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
        "ratio": request.ratio,
        "resolution": request.resolution,
        "fps": int(request.fps),
        "quality": request.quality,
        "bgm_path": request.bgm_path,
        "bgm_volume": float(request.bgm_volume),
        "bgm_fade_in_s": float(request.bgm_fade_in_s),
        "bgm_fade_out_s": float(request.bgm_fade_out_s),
        "narration_volume": float(request.narration_volume),
        "subtitle_color": request.subtitle_color,
    }
    advanced_kwargs = {
        "subtitle_fade_in_ms": int(request.subtitle_fade_in_ms),
        "subtitle_fade_out_ms": int(request.subtitle_fade_out_ms),
        "subtitle_scale_from": int(request.subtitle_scale_from),
        "subtitle_scale_to": int(request.subtitle_scale_to),
        "subtitle_outline": int(request.subtitle_outline),
        "subtitle_shadow": int(request.subtitle_shadow),
        "subtitle_font": request.subtitle_font,
        "subtitle_size": request.subtitle_size,
        "subtitle_y_ratio": request.subtitle_y_ratio,
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


def _resolve_script_override_text(request: BuildRequest) -> str | None:
    """
    Determine script override input.

    Priority:
    - explicit script_text
    - script_file contents (utf-8)
    """
    if request.script_text is not None and str(request.script_text).strip():
        return str(request.script_text)
    if request.script_file is not None and str(request.script_file).strip():
        path = Path(str(request.script_file))
        return path.read_text(encoding="utf-8")
    return None


def _merge_build_options_into_manifest(
    manifest_path: Path,
    *,
    subtitle_font: str | None,
    subtitle_size: int | None,
    subtitle_outline: int | None,
) -> None:
    """Persist subtitle styling hints so `subtitle` can match final burn-in layout."""
    if not manifest_path.exists():
        return
    try:
        m = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return
    if not isinstance(m, dict):
        return
    changed = False
    if subtitle_font is not None and str(subtitle_font).strip():
        m["subtitle_font"] = str(subtitle_font).strip()
        changed = True
    if subtitle_size is not None:
        m["subtitle_size"] = int(subtitle_size)
        changed = True
    elif "subtitle_size" not in m:
        m["subtitle_size"] = 48
        changed = True
    if subtitle_outline is not None:
        m["subtitle_outline"] = int(subtitle_outline)
        changed = True
    if changed:
        manifest_path.write_text(json.dumps(m, indent=2, ensure_ascii=False), encoding="utf-8")


def _load_manifest_slides_meta(manifest_path: Path) -> list[dict[str, Any]] | None:
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    slides = payload.get("slides")
    if not isinstance(slides, list):
        return None
    out: list[dict[str, Any]] = []
    for s in slides:
        if not isinstance(s, dict):
            continue
        try:
            page = int(s.get("page"))
        except Exception:
            continue
        out.append({"page": page, "title": str(s.get("title") or "")})
    return out or None


def _write_script_override(
    *,
    script_path: Path,
    slides_meta: list[dict[str, Any]] | None,
    slide_count: int,
    script_text: str,
) -> None:
    """
    Write a replacement `scripts/script.json` from user-provided script content.

    Accepted formats (auto-detected):
    - JSON object with {"slides":[{"page":1,"title":"...","script":"..."}...]}
    - "all.txt" format blocks starting with `--- Slide 001 ... ---`
    - Plain text: if it splits into N blocks (blank-line separated) and N == slide_count,
      map blocks sequentially to pages; otherwise use it as slide 1 script.
    """
    script_path.parent.mkdir(parents=True, exist_ok=True)
    raw = (script_text or "").replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    if not raw.strip():
        return

    # 1) JSON passthrough if user provided a full script.json payload.
    if raw.lstrip().startswith("{"):
        try:
            payload = json.loads(raw)
        except Exception:
            payload = None
        if isinstance(payload, dict) and isinstance(payload.get("slides"), list):
            script_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
            return

    pages: list[int]
    titles: dict[int, str] = {}
    if slides_meta:
        pages = [int(x["page"]) for x in slides_meta if "page" in x]
        for x in slides_meta:
            try:
                titles[int(x["page"])] = str(x.get("title") or "")
            except Exception:
                pass
    else:
        pages = list(range(1, max(1, int(slide_count or 1)) + 1))

    by_page: dict[int, str] = {}

    # 2) Parse `--- Slide NNN ... ---` blocks (same as scripts/all.txt).
    import re

    header_re = re.compile(r"^\s*---\s*Slide\s+(\d+)\b.*---\s*$")
    current_page: int | None = None
    buf: list[str] = []
    seen_any_header = False
    lines = raw.split("\n")
    for line in lines:
        m = header_re.match(line)
        if m:
            seen_any_header = True
            if current_page is not None:
                by_page[current_page] = "\n".join(buf).strip("\n")
            buf = []
            try:
                current_page = int(m.group(1))
            except Exception:
                current_page = None
            continue
        if current_page is not None:
            buf.append(line)
    if current_page is not None:
        by_page[current_page] = "\n".join(buf).strip("\n")

    if not seen_any_header:
        # 3) Plain text heuristic: blank-line separated blocks.
        blocks = [b.strip("\n") for b in re.split(r"\n{2,}", raw) if b.strip()]
        if slide_count and len(blocks) == int(slide_count):
            for idx, page in enumerate(pages[: len(blocks)]):
                by_page[int(page)] = blocks[idx].strip()
        else:
            # Fallback: treat as slide 1 script.
            if pages:
                by_page[int(pages[0])] = raw.strip()

    slides_out: list[dict[str, Any]] = []
    for page in pages:
        p = int(page)
        slides_out.append(
            {
                "page": p,
                "title": titles.get(p, ""),
                "script": str(by_page.get(p, "") or ""),
            }
        )

    script_path.write_text(
        json.dumps({"slides": slides_out}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


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
        ratio=request.ratio,
        resolution=request.resolution,
        fps=int(request.fps),
        quality=request.quality,
        bgm_path=request.bgm_path,
        bgm_volume=float(request.bgm_volume),
        bgm_fade_in_s=float(request.bgm_fade_in_s),
        bgm_fade_out_s=float(request.bgm_fade_out_s),
        narration_volume=float(request.narration_volume),
        subtitle_color=request.subtitle_color,
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
