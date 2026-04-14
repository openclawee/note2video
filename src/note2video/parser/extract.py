from __future__ import annotations

import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path
from typing import Any

from note2video.schemas.manifest import Manifest, SlideRecord

MSO_PLACEHOLDER = 14
PP_PLACEHOLDER_BODY = 2
PP_PLACEHOLDER_SLIDE_NUMBER = 13
PP_PLACEHOLDER_HEADER = 14
PP_PLACEHOLDER_FOOTER = 15
PP_PLACEHOLDER_DATE = 16

PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

XML_NS = {
    "p": PML_NS,
    "a": DML_NS,
    "rel": REL_NS,
}


class PowerPointUnavailableError(RuntimeError):
    """Raised when PowerPoint automation is unavailable on this machine."""


def extract_project(input_file: str, output_dir: str, pages: str | None = None) -> Manifest:
    """Extract slide images and speaker notes into the project workspace."""
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    suffix = input_path.suffix.lower()
    if suffix not in {".pptx", ".pdf"}:
        raise ValueError("Only .pptx and .pdf inputs are supported.")

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

    if suffix == ".pdf":
        total_pages = _count_pdf_pages(input_path)
        selected_pages = _parse_page_selection(pages, total_slides=total_pages)
        slide_data, extractor = _extract_pdf_slide_data(
            input_path,
            slides_dir,
            selected_pages=selected_pages,
        )
    else:
        # Prefer counting slides without exporting images, so page selection doesn't generate all slide PNGs.
        # If counting fails (e.g. tests using a placeholder file), fall back to extraction-driven counting.
        try:
            total_slides = _count_slides_openxml(input_path)
            selected_pages = _parse_page_selection(pages, total_slides=total_slides)
            slide_data, extractor = _extract_slide_data(input_path, slides_dir, selected_pages=selected_pages)
        except Exception:
            slide_data, extractor = _extract_slide_data(input_path, slides_dir, selected_pages=None)
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
        _format_extract_log(
            input_path=input_path,
            slide_count=len(slides),
            pages=pages,
            extractor=extractor,
        ),
        encoding="utf-8",
    )

    return manifest


def _count_pdf_pages(input_path: Path) -> int:
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise PowerPointUnavailableError("pypdf is required to read PDF page counts.") from exc
    reader = PdfReader(str(input_path))
    return int(len(reader.pages))


def _extract_pdf_slide_data(
    input_path: Path,
    slides_dir: Path,
    *,
    selected_pages: set[int] | None,
) -> tuple[list[dict[str, Any]], str]:
    """
    Convert a PDF to per-page PNG slides.

    Notes are not available in PDFs, so raw_notes are empty and script is empty.
    """
    total_pages = _count_pdf_pages(input_path)
    pages = list(range(1, total_pages + 1))
    if selected_pages is not None:
        pages = [p for p in pages if p in selected_pages]

    # Prefer a pure-Python renderer (PyMuPDF) so users don't need to install Poppler.
    # PyMuPDF is AGPL; ensure this is acceptable for your distribution.
    pymupdf_ok = False
    try:
        import fitz  # PyMuPDF

        pymupdf_ok = True
    except Exception:
        pymupdf_ok = False

    if pymupdf_ok:
        import fitz  # type: ignore[no-redef]

        dpi = float(os.environ.get("NOTE2VIDEO_PDF_RENDER_DPI", "150").strip() or "150")
        zoom = max(0.5, min(6.0, dpi / 72.0))
        doc = fitz.open(str(input_path))
        try:
            for page in pages:
                p = doc.load_page(page - 1)
                mat = fitz.Matrix(zoom, zoom)
                pix = p.get_pixmap(matrix=mat, alpha=False)
                dest = slides_dir / f"{page:03d}.png"
                pix.save(str(dest))
        finally:
            try:
                doc.close()
            except Exception:
                pass
        extractor = "pdf-pymupdf"
    else:
        pdftoppm = _find_pdftoppm()
        if pdftoppm:
            # Render to a temp directory, then copy/rename into slides_dir as 001.png etc.
            work_dir = Path(tempfile.mkdtemp(prefix="note2video-pdf-"))
            try:
                prefix = work_dir / "page"
                dpi = os.environ.get("NOTE2VIDEO_PDF_RENDER_DPI", "150").strip() or "150"
                _run_tool(
                    [pdftoppm, "-png", "-r", dpi, str(input_path.resolve()), str(prefix)],
                    timeout=int(os.environ.get("NOTE2VIDEO_PDFTOPPM_TIMEOUT", "180")),
                    label="pdftoppm (pdf to png)",
                )
                rendered = sorted(work_dir.glob("page-*.png"), key=lambda p: _pdftoppm_page_index(p.name))
                by_index = {(_pdftoppm_page_index(p.name)): p for p in rendered}
                for page in pages:
                    src = by_index.get(page)
                    dest = slides_dir / f"{page:03d}.png"
                    if src and src.exists():
                        shutil.copy2(src, dest)
                    else:
                        _write_placeholder_slide_png(dest, page=page, title="")
            finally:
                shutil.rmtree(work_dir, ignore_errors=True)
            extractor = "pdf-pdftoppm"
        else:
            # Fallback: generate placeholder PNGs so the pipeline can still run.
            for page in pages:
                _write_placeholder_slide_png(slides_dir / f"{page:03d}.png", page=page, title="PDF (placeholder)")
            extractor = "pdf-placeholder"

    records: list[dict[str, Any]] = []
    for page in pages:
        image_name = f"{page:03d}.png"
        records.append(
            {
                "page": page,
                "title": "",
                "image": f"slides/{image_name}",
                "raw_notes": "",
            }
        )
    return records, extractor


def _extract_slide_data(
    input_path: Path,
    slides_dir: Path,
    *,
    selected_pages: set[int] | None,
) -> tuple[list[dict[str, Any]], str]:
    if sys.platform == "win32":
        try:
            return _extract_with_powerpoint(input_path, slides_dir, selected_pages=selected_pages), "powerpoint-com"
        except PowerPointUnavailableError:
            # Fallback keeps extraction available even when COM automation fails.
            return _extract_with_openxml(input_path, slides_dir, selected_pages=selected_pages), "openxml-fallback"
    if _should_try_libreoffice_export():
        try:
            return _extract_with_libreoffice(input_path, slides_dir, selected_pages=selected_pages), "libreoffice"
        except PowerPointUnavailableError:
            pass
    return _extract_with_openxml(input_path, slides_dir, selected_pages=selected_pages), "openxml"


def _extract_with_powerpoint(
    input_path: Path,
    slides_dir: Path,
    *,
    selected_pages: set[int] | None = None,
) -> list[dict[str, Any]]:
    """Use PowerPoint COM automation to export slide images and notes."""
    try:
        import pythoncom
        import win32com.client
        import win32gui
    except ImportError as exc:  # pragma: no cover - environment specific
        raise PowerPointUnavailableError(
            "pywin32 is required for PowerPoint extraction on Windows."
        ) from exc

    pythoncom.CoInitialize()
    app = None
    presentation = None

    def _hide_window() -> None:
        """把 PowerPoint 主窗口移到屏幕外，防止闪烁，同时保持渲染正常。"""
        try:
            hwnd = win32gui.FindWindow("PPTFrameClass", None)
            if hwnd:
                # SWP_NOACTIVATE=0x0004：不激活窗口；把窗口移到屏幕左边很远的位置
                win32gui.SetWindowPos(hwnd, None, -32000, 0, 0, 0, 0x0004 | 0x0001)
        except Exception:
            pass

    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = 1  # 必须可见，Export 才能正确渲染
        app.DisplayAlerts = 2  # ppAlertsNone，不弹对话框
        presentation = app.Presentations.Open(
            str(input_path.resolve()),
            ReadOnly=1,
            Untitled=0,
            WithWindow=0,
        )
        # 打开后立即藏窗口
        _hide_window()

        records: list[dict[str, Any]] = []
        for index, slide in enumerate(presentation.Slides, start=1):
            if selected_pages is not None and index not in selected_pages:
                continue
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


def _extract_with_openxml(
    input_path: Path,
    slides_dir: Path,
    *,
    selected_pages: set[int] | None = None,
) -> list[dict[str, Any]]:
    try:
        records: list[dict[str, Any]] = []
        for item in _read_openxml_slide_records(input_path):
            index = item["page"]
            if selected_pages is not None and index not in selected_pages:
                continue
            image_name = f"{index:03d}.png"
            image_path = slides_dir / image_name
            _write_placeholder_slide_png(image_path, page=index, title=item["title"])
            records.append(
                {
                    "page": index,
                    "title": item["title"],
                    "image": f"slides/{image_name}",
                    "raw_notes": item["raw_notes"],
                }
            )
        return records
    except zipfile.BadZipFile as exc:
        raise PowerPointUnavailableError(f"Invalid .pptx package: {input_path}") from exc
    except KeyError as exc:
        raise PowerPointUnavailableError(f"Missing .pptx part: {exc}") from exc
    except ET.ParseError as exc:
        raise PowerPointUnavailableError(f"Malformed .pptx XML: {exc}") from exc
    except OSError as exc:
        raise PowerPointUnavailableError(f"Unable to read .pptx file: {exc}") from exc
    except ValueError as exc:
        raise PowerPointUnavailableError(str(exc)) from exc


def _list_slide_paths(pptx_file: zipfile.ZipFile) -> list[str]:
    slide_paths = [
        name
        for name in pptx_file.namelist()
        if name.startswith("ppt/slides/slide") and name.endswith(".xml")
    ]
    return sorted(slide_paths, key=_slide_sort_key)


def _slide_sort_key(path: str) -> tuple[int, str]:
    match = re.search(r"slide(\d+)\.xml$", path)
    if match:
        return int(match.group(1)), path
    return 999999, path


def _extract_openxml_slide_title(slide_root: ET.Element) -> str:
    title_candidates: list[str] = []
    fallback_candidates: list[str] = []

    for shape in slide_root.findall(".//p:sp", XML_NS):
        text = _extract_openxml_shape_text(shape)
        if not text:
            continue
        fallback_candidates.append(text)
        ph = shape.find("./p:nvSpPr/p:nvPr/p:ph", XML_NS)
        if ph is not None and ph.get("type") in {"title", "ctrTitle"}:
            title_candidates.append(text)

    if title_candidates:
        return title_candidates[0]
    if fallback_candidates:
        return fallback_candidates[0]
    return ""


def _extract_openxml_slide_notes(pptx_file: zipfile.ZipFile, slide_path: str) -> str:
    notes_path = _resolve_notes_slide_path(pptx_file, slide_path)
    if not notes_path:
        return ""
    notes_root = ET.fromstring(pptx_file.read(notes_path))

    lines: list[str] = []
    for shape in notes_root.findall(".//p:sp", XML_NS):
        ph = shape.find("./p:nvSpPr/p:nvPr/p:ph", XML_NS)
        if ph is not None and ph.get("type") in {"dt", "ftr", "hdr", "sldNum"}:
            continue
        text = _extract_openxml_shape_text(shape)
        if text:
            lines.append(text)
    return _normalize_text("\n".join(lines))


def _read_openxml_slide_records(input_path: Path) -> list[dict[str, Any]]:
    """Parse slide titles and speaker notes from the .pptx package (no raster export)."""
    with zipfile.ZipFile(input_path, "r") as pptx_file:
        slide_paths = _list_slide_paths(pptx_file)
        if not slide_paths:
            raise ValueError("No slides found in .pptx package.")

        records: list[dict[str, Any]] = []
        for index, slide_path in enumerate(slide_paths, start=1):
            slide_root = ET.fromstring(pptx_file.read(slide_path))
            title = _extract_openxml_slide_title(slide_root)
            raw_notes = _extract_openxml_slide_notes(pptx_file, slide_path)
            records.append({"page": index, "title": title, "raw_notes": raw_notes})
        return records


def _libreoffice_export_disabled() -> bool:
    flag = os.environ.get("NOTE2VIDEO_USE_LIBREOFFICE", "").strip().lower()
    return flag in {"0", "false", "no", "off"}


def _find_soffice() -> str | None:
    env_path = os.environ.get("NOTE2VIDEO_LIBREOFFICE", "").strip()
    if env_path:
        candidate = Path(env_path)
        if candidate.is_file():
            return str(candidate.resolve())
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return None


def _find_pdftoppm() -> str | None:
    return shutil.which("pdftoppm")


def _should_try_libreoffice_export() -> bool:
    if sys.platform == "win32":
        return False
    if _libreoffice_export_disabled():
        return False
    return _find_soffice() is not None and _find_pdftoppm() is not None


def _run_tool(args: list[str], *, timeout: int, label: str) -> None:
    try:
        result = subprocess.run(
            args,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        raise PowerPointUnavailableError(f"{label} timed out after {timeout}s.") from exc
    except OSError as exc:
        raise PowerPointUnavailableError(f"{label} could not be executed: {exc}") from exc

    if result.returncode != 0:
        stderr = (result.stderr or "").strip()
        stdout = (result.stdout or "").strip()
        detail = stderr or stdout or "no output"
        raise PowerPointUnavailableError(
            f"{label} failed (exit {result.returncode}): {detail[:1200]}"
        )


def _libreoffice_convert_to_pdf(input_path: Path, work_dir: Path, soffice: str) -> Path:
    """Run LibreOffice headless to produce a PDF next to the .pptx stem."""
    _run_tool(
        [
            soffice,
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            "--convert-to",
            "pdf",
            "--outdir",
            str(work_dir),
            str(input_path.resolve()),
        ],
        timeout=int(os.environ.get("NOTE2VIDEO_LIBREOFFICE_TIMEOUT", "300")),
        label="LibreOffice (pptx to pdf)",
    )

    expected = work_dir / f"{input_path.stem}.pdf"
    if expected.exists():
        return expected

    pdfs = sorted(work_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not pdfs:
        raise PowerPointUnavailableError(
            "LibreOffice finished but no PDF was produced. "
            "Check that LibreOffice Impress can open this .pptx."
        )
    return pdfs[0]


def _pdftoppm_split_pngs(pdf_path: Path, work_dir: Path, pdftoppm: str) -> list[Path]:
    prefix = work_dir / "slide"
    dpi = os.environ.get("NOTE2VIDEO_PDF_RENDER_DPI", "150").strip() or "150"
    _run_tool(
        [pdftoppm, "-png", "-r", dpi, str(pdf_path), str(prefix)],
        timeout=int(os.environ.get("NOTE2VIDEO_PDFTOPPM_TIMEOUT", "180")),
        label="pdftoppm (pdf to png)",
    )

    pattern = f"{prefix.name}-*.png"
    pages = sorted(
        work_dir.glob(pattern),
        key=lambda p: _pdftoppm_page_index(p.name),
    )
    if pages:
        return pages

    fallback = sorted(work_dir.glob("*.png"), key=lambda p: _pdftoppm_page_index(p.name))
    if not fallback:
        raise PowerPointUnavailableError(
            "pdftoppm produced no PNG pages. Check Poppler (pdftoppm) installation."
        )
    return fallback


def _pdftoppm_page_index(filename: str) -> int:
    match = re.search(r"-(\d+)\.png$", filename, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return 0


def _apply_png_sequence_to_slides(
    meta: list[dict[str, Any]],
    png_paths: list[Path],
    slides_dir: Path,
    *,
    selected_pages: set[int] | None = None,
) -> list[dict[str, Any]]:
    """Copy raster pages in order; pad missing slides with placeholders."""
    records: list[dict[str, Any]] = []
    for index, row in enumerate(meta):
        page = int(row["page"])
        if selected_pages is not None and page not in selected_pages:
            continue
        title = row["title"]
        raw_notes = row["raw_notes"]
        image_name = f"{page:03d}.png"
        dest = slides_dir / image_name
        # Prefer page-based indexing for stability even when selected_pages skips items.
        page_idx = page - 1
        if 0 <= page_idx < len(png_paths):
            shutil.copy2(png_paths[page_idx], dest)
        else:
            _write_placeholder_slide_png(dest, page=page, title=title)
        records.append(
            {
                "page": page,
                "title": title,
                "image": f"slides/{image_name}",
                "raw_notes": raw_notes,
            }
        )
    return records


def _extract_with_libreoffice(
    input_path: Path,
    slides_dir: Path,
    *,
    selected_pages: set[int] | None = None,
) -> list[dict[str, Any]]:
    soffice = _find_soffice()
    pdftoppm = _find_pdftoppm()
    if not soffice or not pdftoppm:
        raise PowerPointUnavailableError("LibreOffice export requires soffice and pdftoppm on PATH.")

    try:
        meta = _read_openxml_slide_records(input_path)
    except zipfile.BadZipFile as exc:
        raise PowerPointUnavailableError(f"Invalid .pptx package: {input_path}") from exc
    except KeyError as exc:
        raise PowerPointUnavailableError(f"Missing .pptx part: {exc}") from exc
    except ET.ParseError as exc:
        raise PowerPointUnavailableError(f"Malformed .pptx XML: {exc}") from exc
    except OSError as exc:
        raise PowerPointUnavailableError(f"Unable to read .pptx file: {exc}") from exc
    except ValueError as exc:
        raise PowerPointUnavailableError(str(exc)) from exc

    work_dir = Path(tempfile.mkdtemp(prefix="note2video-lo-"))
    try:
        pdf_path = _libreoffice_convert_to_pdf(input_path, work_dir, soffice)
        png_paths = _pdftoppm_split_pngs(pdf_path, work_dir, pdftoppm)

        if len(png_paths) < len(meta):
            raise PowerPointUnavailableError(
                f"PDF has fewer pages ({len(png_paths)}) than slides in the deck ({len(meta)}). "
                "Try opening and re-saving the file in PowerPoint or LibreOffice."
            )
        if len(png_paths) > len(meta):
            png_paths = png_paths[: len(meta)]

        return _apply_png_sequence_to_slides(meta, png_paths, slides_dir, selected_pages=selected_pages)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _count_slides_openxml(input_path: Path) -> int:
    """Fast slide count without exporting images; used for page selection parsing."""
    with zipfile.ZipFile(input_path, "r") as pptx_file:
        slide_paths = _list_slide_paths(pptx_file)
        if not slide_paths:
            raise ValueError("No slides found in .pptx package.")
        return len(slide_paths)


def _resolve_notes_slide_path(pptx_file: zipfile.ZipFile, slide_path: str) -> str | None:
    rels_path = _to_rels_path(slide_path)
    if rels_path not in pptx_file.namelist():
        return None

    rels_root = ET.fromstring(pptx_file.read(rels_path))
    for relationship in rels_root.findall(".//rel:Relationship", XML_NS):
        if relationship.get("Type") != f"{DOC_REL_NS}/notesSlide":
            continue
        target = relationship.get("Target", "")
        if not target:
            continue
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(slide_path), target))
        return resolved
    return None


def _to_rels_path(part_path: str) -> str:
    parent = posixpath.dirname(part_path)
    base = posixpath.basename(part_path)
    return f"{parent}/_rels/{base}.rels"


def _extract_openxml_shape_text(shape: ET.Element) -> str:
    lines: list[str] = []
    for paragraph in shape.findall(".//a:p", XML_NS):
        chunks = [node.text or "" for node in paragraph.findall(".//a:t", XML_NS)]
        content = "".join(chunks).strip()
        if content:
            lines.append(content)
    return _normalize_text("\n".join(lines))


def _write_placeholder_slide_png(image_path: Path, *, page: int, title: str) -> None:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError as exc:  # pragma: no cover - dependency is declared in pyproject
        raise PowerPointUnavailableError("Pillow is required to generate slide placeholders.") from exc

    image = Image.new("RGB", (1280, 720), color=(245, 247, 255))
    drawer = ImageDraw.Draw(image)
    drawer.rectangle((50, 50, 1230, 670), outline=(80, 90, 120), width=4)

    def _load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
        # Try common Windows fonts for better readability; fall back to default bitmap font.
        candidates = [
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\segoeui.ttf",
            r"C:\Windows\Fonts\msyh.ttc",
        ]
        for p in candidates:
            try:
                if Path(p).exists():
                    return ImageFont.truetype(p, size=size)
            except Exception:
                continue
        try:
            return ImageFont.load_default()
        except Exception:
            return ImageFont.load_default()

    font_big = _load_font(64)
    font_small = _load_font(28)
    drawer.text((90, 120), f"SLIDE {page:03d}", fill=(15, 20, 35), font=font_big)
    if title:
        drawer.text((90, 220), str(title)[:140], fill=(35, 45, 70), font=font_small)
    drawer.text(
        (90, 620),
        "Placeholder image (install PyMuPDF or Poppler pdftoppm for real PDF rendering)",
        fill=(50, 55, 65),
        font=_load_font(18),
    )
    image.save(image_path, format="PNG")


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


def _format_extract_log(
    *,
    input_path: Path,
    slide_count: int,
    pages: str | None,
    extractor: str,
) -> str:
    requested_pages = pages or "all"
    return (
        "Note2Video extract run\n"
        f"input_file: {input_path}\n"
        f"requested_pages: {requested_pages}\n"
        f"extractor: {extractor}\n"
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
