from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any


class ComposeError(RuntimeError):
    """Raised when PPTX compose fails."""


@dataclass(frozen=True)
class ComposeStats:
    slide_count: int
    applied_text_fields: int
    ignored_text_fields: int
    applied_images: int
    ignored_images: int
    applied_notes: int


def compose_pptx_from_template(
    *,
    template_pptx: str,
    params_json: str,
    output_pptx: str,
    assets_base_dir: str | None = None,
) -> ComposeStats:
    """
    Compose a PPTX by duplicating the single slide in `template_pptx` N times and applying per-page fields.

    Params JSON format (loose/forgiving):
    {
      "pages": [
        {"fields": {"title": "..."}, "images": {"hero_image": "C:/a.png"}, "notes": "..."},
        ...
      ]
    }

    Behavior:
    - Missing shape names are ignored (loose mode).
    - Only string-like values are written (values are coerced via str()).
    - This implementation uses PowerPoint COM on Windows for best fidelity.
    """
    template_path = Path(template_pptx)
    if not template_path.exists():
        raise FileNotFoundError(f"Template file not found: {template_path}")
    if template_path.suffix.lower() != ".pptx":
        raise ValueError("template_pptx must be a .pptx file.")

    try:
        payload = json.loads(Path(params_json).read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        # Most common issue on Windows: unescaped backslashes in paths, e.g. "C:\Users\..."
        raise ValueError(
            "Invalid params.json (not valid JSON). "
            "If you used Windows paths, please escape backslashes (\\\\) or use forward slashes (/). "
            f"Details: {exc}"
        ) from exc
    pages = payload.get("pages")
    if not isinstance(pages, list) or not pages:
        raise ValueError("params_json must contain a non-empty 'pages' array.")

    out_path = Path(output_pptx)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    base_dir = Path(assets_base_dir) if assets_base_dir else None

    if sys.platform != "win32":
        raise ComposeError("compose is currently implemented via PowerPoint COM and requires Windows.")

    return _compose_pptx_powerpoint_com(
        template_path=template_path,
        pages=pages,
        out_path=out_path,
        assets_base_dir=base_dir,
    )


def _compose_pptx_powerpoint_com(
    *,
    template_path: Path,
    pages: list[Any],
    out_path: Path,
    assets_base_dir: Path | None,
) -> ComposeStats:
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover - environment specific
        raise ComposeError("pywin32 is required for PPTX compose on Windows.") from exc

    pythoncom.CoInitialize()
    app = None
    pres = None
    pres_out = None

    # Office constants (avoid win32com.constants import to keep tests easy to mock).
    msoTrue = -1
    msoFalse = 0
    # Placeholder shape type in tests uses 14; we keep it consistent with extract.py usage.
    msoPlaceholder = 14
    ppPlaceholderBody = 2

    def _resolve_image_path(raw: str) -> str:
        p = Path(str(raw).strip().strip('"'))
        if not p.is_absolute() and assets_base_dir is not None:
            p = (assets_base_dir / p).resolve()
        return str(p)

    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = 1
        app.DisplayAlerts = 2  # ppAlertsNone

        pres = app.Presentations.Open(
            str(template_path.resolve()),
            ReadOnly=1,
            Untitled=0,
            WithWindow=0,
        )

        # Create a writable copy and operate on it so the template remains untouched.
        if out_path.exists():
            try:
                out_path.unlink()
            except OSError:
                pass
        pres.SaveCopyAs(str(out_path.resolve()))
        pres.Close()
        pres = None

        pres_out = app.Presentations.Open(
            str(out_path.resolve()),
            ReadOnly=0,
            Untitled=0,
            WithWindow=0,
        )

        if int(getattr(pres_out, "Slides", []).Count) < 1:  # pragma: no cover - sanity
            raise ComposeError("Template presentation has no slides.")

        # Ensure slide count == len(pages) by duplicating slide 1.
        target_n = len(pages)
        while int(pres_out.Slides.Count) < target_n:
            # Duplicate the first slide and move it to the end.
            dup_range = pres_out.Slides(1).Duplicate()
            # SlideRange.Item(1) for the actual slide in many Office versions.
            dup_slide = dup_range.Item(1) if hasattr(dup_range, "Item") else dup_range
            dup_slide.MoveTo(int(pres_out.Slides.Count))

        # If template has more than 1 slide, trim extras from the end.
        while int(pres_out.Slides.Count) > target_n:
            pres_out.Slides(int(pres_out.Slides.Count)).Delete()

        applied_text = ignored_text = 0
        applied_images = ignored_images = 0
        applied_notes = 0

        for i, page in enumerate(pages, start=1):
            if not isinstance(page, dict):
                continue
            slide = pres_out.Slides(i)
            fields = page.get("fields") if isinstance(page.get("fields"), dict) else {}
            images = page.get("images") if isinstance(page.get("images"), dict) else {}
            notes = page.get("notes")

            # Text fields (loose mode).
            for key, value in dict(fields).items():
                name = str(key)
                desired = str(value if value is not None else "")
                shape = _find_shape_by_name(slide.Shapes, name)
                if shape is None:
                    ignored_text += 1
                    continue
                if not getattr(shape, "HasTextFrame", msoFalse):
                    ignored_text += 1
                    continue
                tf = getattr(shape, "TextFrame", None)
                if tf is None:
                    ignored_text += 1
                    continue
                tr = getattr(tf, "TextRange", None)
                if tr is None:
                    ignored_text += 1
                    continue
                tr.Text = desired
                applied_text += 1

            # Images (loose mode).
            for key, raw_path in dict(images).items():
                name = str(key)
                shape = _find_shape_by_name(slide.Shapes, name)
                if shape is None:
                    ignored_images += 1
                    continue
                path = _resolve_image_path(str(raw_path))
                if not Path(path).exists():
                    # Missing file: ignore in loose mode.
                    ignored_images += 1
                    continue
                try:
                    left = float(getattr(shape, "Left"))
                    top = float(getattr(shape, "Top"))
                    width = float(getattr(shape, "Width"))
                    height = float(getattr(shape, "Height"))
                except Exception:
                    ignored_images += 1
                    continue

                # Replace by deleting the placeholder/image shape and inserting a picture at same rect.
                # This is the most compatible approach across Office versions.
                try:
                    shape.Delete()
                except Exception:
                    pass
                pic = slide.Shapes.AddPicture(
                    FileName=str(Path(path).resolve()),
                    LinkToFile=msoFalse,
                    SaveWithDocument=msoTrue,
                    Left=left,
                    Top=top,
                    Width=width,
                    Height=height,
                )
                try:
                    pic.Name = name
                except Exception:
                    pass
                applied_images += 1

            # Speaker notes (optional; loose mode).
            if notes is not None and str(notes).strip():
                if _try_set_slide_notes(slide, str(notes), placeholder_shape_type=msoPlaceholder, body_type=ppPlaceholderBody):
                    applied_notes += 1

        pres_out.Save()
        return ComposeStats(
            slide_count=target_n,
            applied_text_fields=applied_text,
            ignored_text_fields=ignored_text,
            applied_images=applied_images,
            ignored_images=ignored_images,
            applied_notes=applied_notes,
        )
    except ComposeError:
        raise
    except Exception as exc:  # pragma: no cover - depends on PowerPoint runtime
        raise ComposeError(f"PowerPoint compose failed: {exc}") from exc
    finally:
        try:
            if pres_out is not None:
                pres_out.Close()
        except Exception:
            pass
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _find_shape_by_name(shapes, name: str):
    try:
        count = int(getattr(shapes, "Count"))
    except Exception:
        return None
    for i in range(1, count + 1):
        try:
            shape = shapes(i)
        except Exception:
            continue
        try:
            if str(getattr(shape, "Name", "")) == name:
                return shape
        except Exception:
            continue
    return None


def _try_set_slide_notes(slide, text: str, *, placeholder_shape_type: int, body_type: int) -> bool:
    """
    Best-effort: set speaker notes in the notes body placeholder.

    Returns True if we found a body placeholder and wrote text.
    """
    try:
        notes_page = slide.NotesPage
        shapes = notes_page.Shapes
        count = int(getattr(shapes, "Count", 0))
    except Exception:
        return False
    for i in range(1, count + 1):
        try:
            shape = shapes(i)
        except Exception:
            continue
        try:
            if int(getattr(shape, "Type", 0)) != int(placeholder_shape_type):
                continue
            pf = getattr(shape, "PlaceholderFormat", None)
            if pf is None:
                continue
            if int(getattr(pf, "Type", -1)) != int(body_type):
                continue
            if not getattr(shape, "HasTextFrame", 0):
                continue
            tf = shape.TextFrame
            tr = tf.TextRange
            tr.Text = text
            return True
        except Exception:
            continue
    return False

