from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

from note2video.schemas.manifest import Manifest, SlideRecord

MSO_PLACEHOLDER = 14
PP_PLACEHOLDER_BODY = 2
PP_PLACEHOLDER_SLIDE_NUMBER = 13
PP_PLACEHOLDER_HEADER = 14
PP_PLACEHOLDER_FOOTER = 15
PP_PLACEHOLDER_DATE = 16


class PowerPointUnavailableError(RuntimeError):
    """Raised when PowerPoint automation is unavailable on this machine."""


def extract_project(input_file: str, output_dir: str, pages: str | None = None) -> Manifest:
    """Extract slide images and speaker notes into the project workspace."""
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() != ".pptx":
        raise ValueError("Only .pptx input is supported.")
    if sys.platform != "win32":
        raise RuntimeError("Real PPT extraction currently requires Windows PowerPoint.")

    out_dir = Path(output_dir)
    slides_dir = out_dir / "slides"
    notes_dir = out_dir / "notes"
    notes_raw_dir = notes_dir / "raw"
    notes_speaker_dir = notes_dir / "speaker"
    scripts_dir = out_dir / "scripts"
    scripts_txt_dir = scripts_dir / "txt"
    logs_dir = out_dir / "logs"
    log_file = logs_dir / "build.log"

    slides_dir.mkdir(parents=True, exist_ok=True)
    notes_dir.mkdir(parents=True, exist_ok=True)
    notes_raw_dir.mkdir(parents=True, exist_ok=True)
    notes_speaker_dir.mkdir(parents=True, exist_ok=True)
    scripts_dir.mkdir(parents=True, exist_ok=True)
    scripts_txt_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)

    slide_data = _extract_with_powerpoint(input_path, slides_dir)
    selected_pages = _parse_page_selection(pages, total_slides=len(slide_data))
    if selected_pages is not None:
        slide_data = [item for item in slide_data if item["page"] in selected_pages]

    slides: list[SlideRecord] = []
    for item in slide_data:
        raw_notes = item["raw_notes"]
        speaker_notes = _to_speaker_notes(raw_notes)
        script = _to_script(speaker_notes)
        slides.append(
            SlideRecord(
                page=item["page"],
                title=item["title"],
                image=item["image"],
                raw_notes=raw_notes,
                script=script,
            )
        )

    manifest = Manifest(
        project_name=input_path.stem,
        input_file=str(input_path),
        slide_count=len(slides),
        outputs={
            "notes": "notes/notes.json",
            "notes_raw_txt_dir": "notes/raw",
            "notes_speaker_txt_dir": "notes/speaker",
            "script": "scripts/script.json",
            "script_txt_dir": "scripts/txt",
            "notes_all_txt": "notes/all.txt",
            "script_all_txt": "scripts/all.txt",
            "manifest": "manifest.json",
            "log": "logs/build.log",
        },
        slides=slides,
    )

    notes_payload = {
        "source": str(input_path),
        "slide_count": len(slides),
        "slides": [
            {
                "page": slide.page,
                "title": slide.title,
                "image": slide.image,
                "raw_notes": slide.raw_notes,
                "speaker_notes": _to_speaker_notes(slide.raw_notes),
            }
            for slide in slides
        ],
    }
    script_payload = {
        "slides": [
            {
                "page": slide.page,
                "title": slide.title,
                "script": slide.script,
            }
            for slide in slides
        ]
    }

    (notes_dir / "notes.json").write_text(
        json.dumps(notes_payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (scripts_dir / "script.json").write_text(
        json.dumps(script_payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (out_dir / "manifest.json").write_text(
        json.dumps(manifest.to_dict(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    _write_text_exports(
        slides=slides,
        notes_raw_dir=notes_raw_dir,
        notes_speaker_dir=notes_speaker_dir,
        scripts_txt_dir=scripts_txt_dir,
        notes_all_file=notes_dir / "all.txt",
        script_all_file=scripts_dir / "all.txt",
    )
    log_file.write_text(
        _format_extract_log(input_path=input_path, slide_count=len(slides), pages=pages),
        encoding="utf-8",
    )

    return manifest


def _extract_with_powerpoint(input_path: Path, slides_dir: Path) -> list[dict[str, Any]]:
    """Use PowerPoint COM automation to export slide images and notes."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:  # pragma: no cover - environment specific
        raise PowerPointUnavailableError(
            "pywin32 is required for PowerPoint extraction on Windows."
        ) from exc

    pythoncom.CoInitialize()
    app = None
    presentation = None

    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        # Some PowerPoint versions reject hiding the application window here.
        app.Visible = 1
        presentation = app.Presentations.Open(
            str(input_path.resolve()),
            ReadOnly=1,
            Untitled=0,
            WithWindow=0,
        )

        records: list[dict[str, Any]] = []
        for index, slide in enumerate(presentation.Slides, start=1):
            image_name = f"{index:03d}.png"
            image_path = slides_dir / image_name
            slide.Export(str(image_path.resolve()), "PNG")

            records.append(
                {
                    "page": index,
                    "title": _extract_slide_title(slide),
                    "image": f"slides/{image_name}",
                    "raw_notes": _extract_slide_notes(slide),
                }
            )

        return records
    except Exception as exc:  # pragma: no cover - requires PowerPoint runtime
        raise PowerPointUnavailableError(
            "PowerPoint automation failed. Make sure Microsoft PowerPoint is "
            f"installed and can open this file. Original error: {exc}"
        ) from exc
    finally:
        if presentation is not None:
            presentation.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()


def _parse_page_selection(pages: str | None, *, total_slides: int) -> set[int] | None:
    if not pages or pages == "all":
        return None

    selected: set[int] = set()
    for chunk in pages.split(","):
        token = chunk.strip()
        if not token:
            continue
        if "-" in token:
            start_text, end_text = token.split("-", 1)
            start = int(start_text.strip())
            end = int(end_text.strip())
            if start <= 0 or end <= 0 or end < start:
                raise ValueError(f"Invalid page range: {token}")
            selected.update(range(start, end + 1))
            continue

        page = int(token)
        if page <= 0:
            raise ValueError(f"Invalid page number: {token}")
        selected.add(page)

    out_of_bounds = [page for page in sorted(selected) if page > total_slides]
    if out_of_bounds:
        raise ValueError(
            f"Requested pages out of range: {out_of_bounds}. Total slides: {total_slides}"
        )
    return selected


def _format_extract_log(*, input_path: Path, slide_count: int, pages: str | None) -> str:
    requested_pages = pages or "all"
    return (
        "Note2Video extract run\n"
        f"input_file: {input_path}\n"
        f"requested_pages: {requested_pages}\n"
        f"exported_slides: {slide_count}\n"
    )


def _write_text_exports(
    *,
    slides: list[SlideRecord],
    notes_raw_dir: Path,
    notes_speaker_dir: Path,
    scripts_txt_dir: Path,
    notes_all_file: Path,
    script_all_file: Path,
) -> None:
    notes_all_chunks: list[str] = []
    script_all_chunks: list[str] = []

    for slide in slides:
        file_name = f"{slide.page:03d}.txt"
        raw_text = slide.raw_notes
        speaker_text = _to_speaker_notes(slide.raw_notes)
        script_text = slide.script

        (notes_raw_dir / file_name).write_text(raw_text, encoding="utf-8")
        (notes_speaker_dir / file_name).write_text(speaker_text, encoding="utf-8")
        (scripts_txt_dir / file_name).write_text(script_text, encoding="utf-8")

        notes_all_chunks.append(_format_all_text_chunk(slide.page, slide.title, speaker_text))
        script_all_chunks.append(_format_all_text_chunk(slide.page, slide.title, script_text))

    notes_all_file.write_text("\n\n".join(notes_all_chunks), encoding="utf-8")
    script_all_file.write_text("\n\n".join(script_all_chunks), encoding="utf-8")


def _format_all_text_chunk(page: int, title: str, text: str) -> str:
    header = f"--- Slide {page:03d}"
    if title:
        header += f" | {title}"
    header += " ---"
    return f"{header}\n{text}"


def _extract_slide_title(slide: Any) -> str:
    try:
        if slide.Shapes.HasTitle:
            title_shape = slide.Shapes.Title
            return _normalize_text(title_shape.TextFrame.TextRange.Text)
    except Exception:
        return ""
    return ""


def _extract_slide_notes(slide: Any) -> str:
    notes: list[str] = []

    try:
        for shape in slide.NotesPage.Shapes:
            if not _is_notes_body_placeholder(shape):
                continue
            text = _read_shape_text(shape)
            if text:
                notes.append(text)
    except Exception:
        return ""

    joined = "\n".join(notes)
    return _normalize_text(joined)


def _read_shape_text(shape: Any) -> str:
    try:
        if not shape.HasTextFrame:
            return ""
        if not shape.TextFrame.HasText:
            return ""
        text = shape.TextFrame.TextRange.Text
    except Exception:
        return ""

    normalized = _normalize_text(text)
    if normalized.lower() == "click to add notes":
        return ""
    return normalized


def _is_notes_body_placeholder(shape: Any) -> bool:
    try:
        if shape.Type != MSO_PLACEHOLDER:
            return False

        placeholder = shape.PlaceholderFormat
        placeholder_type = placeholder.Type
    except Exception:
        return False

    if placeholder_type in {
        PP_PLACEHOLDER_SLIDE_NUMBER,
        PP_PLACEHOLDER_HEADER,
        PP_PLACEHOLDER_FOOTER,
        PP_PLACEHOLDER_DATE,
    }:
        return False

    return placeholder_type == PP_PLACEHOLDER_BODY


def _normalize_text(text: str) -> str:
    lines = [line.strip() for line in text.replace("\r", "\n").split("\n")]
    filtered = [line for line in lines if line]
    return "\n".join(filtered)


def _to_speaker_notes(text: str) -> str:
    normalized = _normalize_text(text)
    if not normalized:
        return ""

    lines = []
    for line in normalized.split("\n"):
        cleaned = line.strip()
        if not cleaned:
            continue
        if _is_meta_line(cleaned):
            continue
        lines.append(cleaned)
    return "\n".join(lines)


def _to_script(text: str) -> str:
    speaker_notes = _to_speaker_notes(text)
    if not speaker_notes:
        return ""

    script = speaker_notes
    script = script.replace("AI", "AI")
    script = re.sub(r"[ \t]+", " ", script)
    script = re.sub(r"\n{2,}", "\n", script)
    script = re.sub(r"([。！？；])(?=[^\n])", r"\1\n", script)
    script = re.sub(r"\n{2,}", "\n", script)

    cleaned_lines = []
    for line in script.split("\n"):
        cleaned = line.strip(" ，,")
        if cleaned:
            cleaned_lines.append(cleaned)
    return "\n".join(cleaned_lines)


def _is_meta_line(text: str) -> bool:
    meta_prefixes = (
        "备注：",
        "注：",
        "说明：",
        "tips:",
        "tip:",
        "note:",
    )
    lower_text = text.lower()
    if any(lower_text.startswith(prefix) for prefix in meta_prefixes):
        return True
    if text.startswith("（") and text.endswith("）"):
        return True
    if text.startswith("(") and text.endswith(")"):
        return True
    return False
