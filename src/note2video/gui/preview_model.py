from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from note2video.video_canvas import canvas_size, normalize_ratio, normalize_resolution


@dataclass(frozen=True)
class PreviewCue:
    index: int
    page: int
    start_ms: int | None
    end_ms: int | None
    text: str


@dataclass(frozen=True)
class PreviewData:
    project_dir: Path
    page: int
    available_pages: tuple[int, ...]
    page_count: int
    canvas_w: int
    canvas_h: int
    title: str
    image_path: Path | None
    cues: tuple[PreviewCue, ...]
    active_cue_index: int
    active_text: str
    cue_count: int
    text_source: str
    status_text: str


def load_preview_data(
    *,
    project_dir: str | Path,
    page: int,
    ratio: str | None,
    resolution: str | None,
    sample_text: str,
    cue_index: int = 0,
) -> PreviewData:
    root = Path(project_dir)
    canvas_w, canvas_h = _preview_canvas_size(ratio=ratio, resolution=resolution)

    manifest = _load_json_dict(root / "manifest.json") if root.exists() else None
    slides = _load_slides(root, manifest)
    subtitle_map = _load_subtitle_map(root, manifest)
    script_map = _load_script_map(root, manifest)

    available = tuple(sorted({*slides.keys(), *subtitle_map.keys(), *script_map.keys()}))
    if not available:
        available = (1,)
    current_page = _nearest_page(available, page)

    slide = slides.get(current_page, {})
    image_path = _existing_path(_candidate_image_path(root, slide, current_page))
    title = str(slide.get("title") or "")

    cues = subtitle_map.get(current_page, ())
    text_source = "subtitle"
    if not cues:
        cues = script_map.get(current_page, ())
        text_source = "script"
    if not cues:
        sample = str(sample_text or "").strip()
        if sample:
            cues = (_single_cue(page=current_page, text=sample),)
            text_source = "sample"
        else:
            text_source = "empty"

    active_cue_index, active_text = _select_active_cue(cues, cue_index)
    status_text = _status_text(
        project_dir=root,
        manifest=manifest,
        image_path=image_path,
        text_source=text_source,
        cue_count=len(cues),
        active_cue_index=active_cue_index,
    )
    return PreviewData(
        project_dir=root,
        page=current_page,
        available_pages=available,
        page_count=len(available),
        canvas_w=canvas_w,
        canvas_h=canvas_h,
        title=title,
        image_path=image_path,
        cues=tuple(cues),
        active_cue_index=active_cue_index,
        active_text=active_text,
        cue_count=len(cues),
        text_source=text_source,
        status_text=status_text,
    )


def _preview_canvas_size(*, ratio: str | None, resolution: str | None) -> tuple[int, int]:
    try:
        return canvas_size(ratio=normalize_ratio(ratio), resolution=normalize_resolution(resolution))
    except Exception:
        return canvas_size(ratio="16:9", resolution="1080p")


def _load_json_dict(path: Path) -> dict[str, Any] | None:
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    return payload if isinstance(payload, dict) else None


def _resolve_output_path(root: Path, manifest: dict[str, Any] | None, key: str, default_rel: str) -> Path:
    rel = default_rel
    if isinstance(manifest, dict):
        outputs = manifest.get("outputs")
        if isinstance(outputs, dict):
            raw = outputs.get(key)
            if isinstance(raw, str) and raw.strip():
                rel = raw.strip()
    return root / rel


def _load_slides(root: Path, manifest: dict[str, Any] | None) -> dict[int, dict[str, Any]]:
    out: dict[int, dict[str, Any]] = {}
    if isinstance(manifest, dict):
        raw_slides = manifest.get("slides")
        if isinstance(raw_slides, list):
            for item in raw_slides:
                if not isinstance(item, dict):
                    continue
                try:
                    page = int(item.get("page"))
                except Exception:
                    continue
                out[page] = {
                    "page": page,
                    "title": str(item.get("title") or ""),
                    "image": str(item.get("image") or "").strip(),
                }
    if out:
        return out

    slides_dir = root / "slides"
    if not slides_dir.is_dir():
        return out
    for path in sorted(slides_dir.glob("*.png")):
        try:
            page = int(path.stem)
        except Exception:
            continue
        out[page] = {"page": page, "title": "", "image": f"slides/{path.name}"}
    return out


def _load_script_map(root: Path, manifest: dict[str, Any] | None) -> dict[int, tuple[PreviewCue, ...]]:
    path = _resolve_output_path(root, manifest, "script", "scripts/script.json")
    payload = _load_json_dict(path)
    if not isinstance(payload, dict):
        return {}
    raw_slides = payload.get("slides")
    if not isinstance(raw_slides, list):
        return {}

    out: dict[int, tuple[PreviewCue, ...]] = {}
    for item in raw_slides:
        if not isinstance(item, dict):
            continue
        try:
            page = int(item.get("page"))
        except Exception:
            continue
        text = str(item.get("script") or "").strip()
        if text:
            out[page] = (_single_cue(page=page, text=text),)
    return out


def _load_subtitle_map(root: Path, manifest: dict[str, Any] | None) -> dict[int, tuple[PreviewCue, ...]]:
    path = _resolve_output_path(root, manifest, "subtitle_json", "subtitles/subtitles.json")
    payload = _load_json_dict(path)
    if not isinstance(payload, dict):
        return {}
    raw_segments = payload.get("segments")
    if not isinstance(raw_segments, list):
        return {}

    grouped: dict[int, list[PreviewCue]] = {}
    for pos, item in enumerate(raw_segments, start=1):
        if not isinstance(item, dict):
            continue
        try:
            page = int(item.get("page"))
        except Exception:
            continue
        text = str(item.get("text") or "").strip()
        if not text:
            continue
        grouped.setdefault(page, []).append(
            PreviewCue(
                index=_as_int(item.get("index"), default=pos),
                page=page,
                start_ms=_as_optional_int(item.get("start_ms")),
                end_ms=_as_optional_int(item.get("end_ms")),
                text=text,
            )
        )

    out: dict[int, tuple[PreviewCue, ...]] = {}
    for page, cues in grouped.items():
        out[page] = tuple(sorted(cues, key=_cue_sort_key))
    return out


def _single_cue(*, page: int, text: str) -> PreviewCue:
    return PreviewCue(index=1, page=page, start_ms=None, end_ms=None, text=text)


def _as_int(value: Any, *, default: int) -> int:
    try:
        return int(value)
    except Exception:
        return default


def _as_optional_int(value: Any) -> int | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return int(value)
    except Exception:
        return None


def _cue_sort_key(cue: PreviewCue) -> tuple[int, int, int, str]:
    return (
        cue.index,
        cue.start_ms if cue.start_ms is not None else 10**15,
        cue.end_ms if cue.end_ms is not None else 10**15,
        cue.text,
    )


def _select_active_cue(cues: tuple[PreviewCue, ...], cue_index: int) -> tuple[int, str]:
    if not cues:
        return -1, ""
    try:
        requested = int(cue_index)
    except Exception:
        requested = 0
    index = min(max(requested, 0), len(cues) - 1)
    return index, cues[index].text


def _candidate_image_path(root: Path, slide: dict[str, Any], page: int) -> Path:
    raw = str(slide.get("image") or "").strip()
    if raw:
        return root / raw
    return root / "slides" / f"{page:03d}.png"


def _existing_path(path: Path) -> Path | None:
    return path if path.exists() else None


def _nearest_page(available_pages: tuple[int, ...], requested_page: int) -> int:
    if not available_pages:
        return 1
    try:
        page = int(requested_page)
    except Exception:
        return available_pages[0]
    if page in available_pages:
        return page
    return min(available_pages, key=lambda candidate: (abs(candidate - page), candidate))


def _status_text(
    *,
    project_dir: Path,
    manifest: dict[str, Any] | None,
    image_path: Path | None,
    text_source: str,
    cue_count: int,
    active_cue_index: int,
) -> str:
    if not project_dir.exists():
        return "输出目录不存在，当前显示示例文本。"
    if manifest is None and image_path is None and text_source in {"sample", "empty"}:
        return "还没有可预览的工件，先运行 Extract 或 Build。"
    if text_source == "subtitle":
        if cue_count > 0 and active_cue_index >= 0:
            return f"文本来源：subtitles/subtitles.json · 当前页字幕 {active_cue_index + 1}/{cue_count}"
        return "文本来源：subtitles/subtitles.json"
    if text_source == "script":
        return "文本来源：scripts/script.json"
    if text_source == "sample":
        return "未找到字幕或脚本工件，当前显示示例文本。"
    return "当前页暂无可预览文本。"
