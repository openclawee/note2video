from __future__ import annotations

import json
import mimetypes
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = {
    "p": PML_NS,
    "a": DML_NS,
    "r": DOC_REL_NS,
    "rel": PKG_REL_NS,
    "ct": CONTENT_TYPES_NS,
}
IMAGE_REL_TYPE = f"{DOC_REL_NS}/image"
NOTES_REL_TYPE = f"{DOC_REL_NS}/notesSlide"
SLIDE_REL_TYPE = f"{DOC_REL_NS}/slide"
NOTES_MASTER_REL_TYPE = f"{DOC_REL_NS}/notesMaster"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

ET.register_namespace("a", DML_NS)
ET.register_namespace("p", PML_NS)
ET.register_namespace("r", DOC_REL_NS)


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
    Compose a PPTX by duplicating the first template slide N times and applying per-page fields.

    Params JSON format (loose/forgiving):
    {
      "pages": [
        {"fields": {"title": "..."}, "images": {"hero_image": "C:/a.png"}, "notes": "..."},
        ...
      ]
    }

    Accepted compatibility aliases:
    - Top-level `slides` may be used instead of `pages`.
    - Per-page `speaker_notes` or `script` may be used when `notes` is absent.

    Behavior:
    - Missing shape names are ignored (loose mode).
    - Only string-like values are written (values are coerced via str()).
    - Windows prefers PowerPoint COM for best fidelity.
    - Other platforms compose by editing the PPTX package directly via OpenXML.
    """
    template_path = Path(template_pptx)
    if not template_path.exists():
        raise FileNotFoundError(f"Template file not found: {template_path}")
    if template_path.suffix.lower() != ".pptx":
        raise ValueError("template_pptx must be a .pptx file.")

    try:
        payload = json.loads(Path(params_json).read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(
            "Invalid params.json (not valid JSON). "
            "If you used Windows paths, please escape backslashes (\\\\) or use forward slashes (/). "
            f"Details: {exc}"
        ) from exc
    pages = _resolve_compose_pages(payload)
    if not isinstance(pages, list) or not pages:
        raise ValueError("params_json must contain a non-empty 'pages' array.")

    out_path = Path(output_pptx)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    base_dir = Path(assets_base_dir) if assets_base_dir else None

    if sys.platform == "win32":
        try:
            return _compose_pptx_powerpoint_com(
                template_path=template_path,
                pages=pages,
                out_path=out_path,
                assets_base_dir=base_dir,
            )
        except ComposeError:
            # PowerPoint COM provides the best fidelity on Windows, but it may be unavailable
            # (not installed, locked-down environment, automation disabled). Fall back to
            # OpenXML editing so `compose` can still work in CI/headless setups.
            return _compose_pptx_openxml(
                template_path=template_path,
                pages=pages,
                out_path=out_path,
                assets_base_dir=base_dir,
            )
    return _compose_pptx_openxml(
        template_path=template_path,
        pages=pages,
        out_path=out_path,
        assets_base_dir=base_dir,
    )


def _resolve_compose_pages(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None
    pages = payload.get("pages")
    if pages is not None:
        return pages
    return payload.get("slides")


def _resolve_page_notes(page: Any) -> str | None:
    if not isinstance(page, dict):
        return None
    for key in ("notes", "speaker_notes", "script"):
        value = page.get(key)
        if value is None:
            continue
        text = str(value)
        if text.strip():
            return text
    return None


def _resolve_image_path(raw: str, assets_base_dir: Path | None) -> Path:
    path = Path(str(raw).strip().strip('"'))
    if not path.is_absolute() and assets_base_dir is not None:
        path = (assets_base_dir / path).resolve()
    return path


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

    msoTrue = -1
    msoFalse = 0
    msoPlaceholder = 14
    ppPlaceholderBody = 2

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

        target_n = len(pages)
        while int(pres_out.Slides.Count) < target_n:
            dup_range = pres_out.Slides(1).Duplicate()
            dup_slide = dup_range.Item(1) if hasattr(dup_range, "Item") else dup_range
            dup_slide.MoveTo(int(pres_out.Slides.Count))

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
            notes = _resolve_page_notes(page)

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

            for key, raw_path in dict(images).items():
                name = str(key)
                shape = _find_shape_by_name(slide.Shapes, name)
                if shape is None:
                    ignored_images += 1
                    continue
                path = _resolve_image_path(str(raw_path), assets_base_dir)
                if not path.exists():
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

                try:
                    shape.Delete()
                except Exception:
                    pass
                pic = slide.Shapes.AddPicture(
                    FileName=str(path.resolve()),
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


def _compose_pptx_openxml(
    *,
    template_path: Path,
    pages: list[Any],
    out_path: Path,
    assets_base_dir: Path | None,
) -> ComposeStats:
    work_dir = Path(tempfile.mkdtemp(prefix="note2video-compose-"))
    try:
        try:
            with zipfile.ZipFile(template_path, "r") as archive:
                archive.extractall(work_dir)
        except zipfile.BadZipFile as exc:
            raise ComposeError(f"Invalid .pptx package: {template_path}") from exc

        content_types_path = work_dir / "[Content_Types].xml"
        presentation_path = work_dir / "ppt" / "presentation.xml"
        presentation_rels_path = work_dir / "ppt" / "_rels" / "presentation.xml.rels"
        for required_path in (content_types_path, presentation_path, presentation_rels_path):
            if not required_path.exists():
                raise ComposeError(f"Invalid .pptx package: missing {required_path.relative_to(work_dir)}")

        try:
            content_types_root = ET.parse(content_types_path).getroot()
            presentation_root = ET.parse(presentation_path).getroot()
            presentation_rels_root = ET.parse(presentation_rels_path).getroot()
        except ET.ParseError as exc:
            raise ComposeError(f"Malformed .pptx XML: {exc}") from exc

        base_slide_path = _resolve_primary_slide_path(presentation_root, presentation_rels_root)
        if not base_slide_path:
            raise ComposeError("Template presentation has no slides.")

        base_slide_root = _read_xml_part(work_dir, base_slide_path)
        base_slide_rels_root = _read_xml_part(work_dir, _to_rels_path(base_slide_path), required=False)
        if base_slide_rels_root is None:
            base_slide_rels_root = _new_relationships_root()
        base_notes_path = _resolve_related_part_path(base_slide_path, base_slide_rels_root, NOTES_REL_TYPE)
        base_notes_root = _read_xml_part(work_dir, base_notes_path, required=False) if base_notes_path else None

        media_dir = work_dir / "ppt" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)

        applied_text = ignored_text = 0
        applied_images = ignored_images = 0
        applied_notes = 0
        slide_parts: list[str] = []
        notes_parts: list[str] = []
        wrote_any_notes = False

        for slide_index, raw_page in enumerate(pages, start=1):
            page = raw_page if isinstance(raw_page, dict) else {}
            slide_root = deepcopy(base_slide_root)
            slide_rels_root = deepcopy(base_slide_rels_root)
            notes_value = _resolve_page_notes(page)
            wants_notes = notes_value is not None

            # If the template slide has no notes part, we can still create one on demand.
            notes_root = (
                deepcopy(base_notes_root)
                if base_notes_root is not None
                else (_build_notes_slide_root(str(notes_value)) if wants_notes else None)
            )

            text_applied, text_ignored = _apply_text_fields_openxml(slide_root, page.get("fields"))
            applied_text += text_applied
            ignored_text += text_ignored

            image_applied, image_ignored = _apply_images_openxml(
                slide_root,
                slide_rels_root,
                page.get("images"),
                media_dir=media_dir,
                assets_base_dir=assets_base_dir,
                slide_index=slide_index,
            )
            applied_images += image_applied
            ignored_images += image_ignored

            if notes_root is not None:
                if wants_notes:
                    # If we created a new notes root, it already contains the desired text.
                    if base_notes_root is not None:
                        if _try_set_slide_notes_openxml(notes_root, str(notes_value)):
                            applied_notes += 1
                    else:
                        applied_notes += 1
                    wrote_any_notes = True
                notes_part = f"ppt/notesSlides/notesSlide{slide_index}.xml"
                _set_relationship_target(
                    slide_rels_root,
                    rel_type=NOTES_REL_TYPE,
                    target=f"../notesSlides/notesSlide{slide_index}.xml",
                )
                _write_xml_part(work_dir, notes_part, notes_root)
                notes_parts.append(notes_part)
            else:
                _remove_relationships_by_type(slide_rels_root, NOTES_REL_TYPE)

            slide_part = f"ppt/slides/slide{slide_index}.xml"
            slide_parts.append(slide_part)
            _write_xml_part(work_dir, slide_part, slide_root)
            _write_xml_part(work_dir, _to_rels_path(slide_part), slide_rels_root)

        _rebuild_presentation_slides(presentation_root, presentation_rels_root, slide_parts)
        if wrote_any_notes:
            notes_master_part = _ensure_notes_master_parts(
                work_dir=work_dir,
                presentation_root=presentation_root,
                presentation_rels_root=presentation_rels_root,
                content_types_root=content_types_root,
            )
            for notes_part in notes_parts:
                _write_notes_slide_rels(
                    work_dir=work_dir,
                    notes_part=notes_part,
                    notes_master_part=notes_master_part,
                )
        _rebuild_content_types(content_types_root, slide_parts, notes_parts, media_dir)

        _write_xml_part(work_dir, "ppt/presentation.xml", presentation_root)
        _write_xml_part(work_dir, "ppt/_rels/presentation.xml.rels", presentation_rels_root)
        _write_xml_part(work_dir, "[Content_Types].xml", content_types_root)

        if out_path.exists():
            out_path.unlink()
        _write_zip_from_directory(work_dir, out_path)

        return ComposeStats(
            slide_count=len(slide_parts),
            applied_text_fields=applied_text,
            ignored_text_fields=ignored_text,
            applied_images=applied_images,
            ignored_images=ignored_images,
            applied_notes=applied_notes,
        )
    except ComposeError:
        raise
    except ET.ParseError as exc:
        raise ComposeError(f"Malformed .pptx XML: {exc}") from exc
    except OSError as exc:
        raise ComposeError(f"Unable to compose .pptx file: {exc}") from exc
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def _apply_text_fields_openxml(slide_root: ET.Element, fields: Any) -> tuple[int, int]:
    if not isinstance(fields, dict):
        return 0, 0
    sp_tree = _find_sp_tree(slide_root)
    if sp_tree is None:
        return 0, len(fields)

    applied = ignored = 0
    for key, value in dict(fields).items():
        target = _find_named_slide_element(sp_tree, str(key))
        if target is None or _local_name(target[1].tag) != "sp":
            ignored += 1
            continue
        if _set_text_body(target[1], str(value if value is not None else "")):
            applied += 1
        else:
            ignored += 1
    return applied, ignored


def _apply_images_openxml(
    slide_root: ET.Element,
    slide_rels_root: ET.Element,
    images: Any,
    *,
    media_dir: Path,
    assets_base_dir: Path | None,
    slide_index: int,
) -> tuple[int, int]:
    if not isinstance(images, dict):
        return 0, 0
    sp_tree = _find_sp_tree(slide_root)
    if sp_tree is None:
        return 0, len(images)

    applied = ignored = 0
    for key, raw_path in dict(images).items():
        name = str(key)
        target = _find_named_slide_element(sp_tree, name)
        if target is None:
            ignored += 1
            continue

        image_path = _resolve_image_path(str(raw_path), assets_base_dir)
        if not image_path.exists() or not image_path.is_file():
            ignored += 1
            continue

        if _image_content_type(image_path) is None:
            ignored += 1
            continue

        _, element, c_nvpr = target
        xfrm = _extract_shape_transform(element)
        if c_nvpr is None or xfrm is None:
            ignored += 1
            continue

        copied_name = _copy_media_asset(image_path, media_dir, slide_index=slide_index, shape_name=name)
        rel_id = _append_relationship(slide_rels_root, rel_type=IMAGE_REL_TYPE, target=f"../media/{copied_name}")
        sp_tree[target[0]] = _build_picture_element(c_nvpr, xfrm, rel_id)
        applied += 1
    return applied, ignored


def _try_set_slide_notes_openxml(notes_root: ET.Element, text: str) -> bool:
    for shape in notes_root.findall("./p:cSld/p:spTree/p:sp", XML_NS):
        ph = shape.find("./p:nvSpPr/p:nvPr/p:ph", XML_NS)
        if ph is None or ph.get("type") != "body":
            continue
        return _set_text_body(shape, text)
    return False


def _build_notes_slide_root(text: str) -> ET.Element:
    """
    Create a minimal notes slide XML that contains a body placeholder.
    This lets us write speaker notes even when the template slide has no notes part.
    """
    notes = ET.Element(_qn(PML_NS, "notes"))
    c_sld = ET.SubElement(notes, _qn(PML_NS, "cSld"))
    sp_tree = ET.SubElement(c_sld, _qn(PML_NS, "spTree"))

    nv_grp = ET.SubElement(sp_tree, _qn(PML_NS, "nvGrpSpPr"))
    ET.SubElement(nv_grp, _qn(PML_NS, "cNvPr"), {"id": "1", "name": ""})
    ET.SubElement(nv_grp, _qn(PML_NS, "cNvGrpSpPr"))
    ET.SubElement(nv_grp, _qn(PML_NS, "nvPr"))

    # A minimal required group properties container.
    grp_sp_pr = ET.SubElement(sp_tree, _qn(PML_NS, "grpSpPr"))
    ET.SubElement(grp_sp_pr, _qn(DML_NS, "xfrm"))

    sp = ET.SubElement(sp_tree, _qn(PML_NS, "sp"))
    nv_sp_pr = ET.SubElement(sp, _qn(PML_NS, "nvSpPr"))
    ET.SubElement(nv_sp_pr, _qn(PML_NS, "cNvPr"), {"id": "2", "name": "Notes Placeholder 1"})
    ET.SubElement(nv_sp_pr, _qn(PML_NS, "cNvSpPr"))
    nv_pr = ET.SubElement(nv_sp_pr, _qn(PML_NS, "nvPr"))
    ET.SubElement(nv_pr, _qn(PML_NS, "ph"), {"type": "body", "idx": "1"})

    ET.SubElement(sp, _qn(PML_NS, "spPr"))
    tx_body = ET.SubElement(sp, _qn(PML_NS, "txBody"))
    ET.SubElement(tx_body, _qn(DML_NS, "bodyPr"))
    ET.SubElement(tx_body, _qn(DML_NS, "lstStyle"))
    # Write the text as one paragraph per line.
    for paragraph_text in str(text or "").splitlines() or [""]:
        tx_body.append(_build_text_paragraph(paragraph_text))

    clr_map_ovr = ET.SubElement(notes, _qn(PML_NS, "clrMapOvr"))
    ET.SubElement(clr_map_ovr, _qn(DML_NS, "masterClrMapping"))

    return notes


def _ensure_notes_master_parts(
    *,
    work_dir: Path,
    presentation_root: ET.Element,
    presentation_rels_root: ET.Element,
    content_types_root: ET.Element,
) -> str:
    """
    Ensure the PPTX package contains a notes master and presentation references to it.

    PowerPoint may ignore notesSlides unless a notesMaster exists and notesSlide rels
    point to that master.
    """
    notes_master_part = "ppt/notesMasters/notesMaster1.xml"
    notes_master_path = work_dir / Path(notes_master_part)
    if not notes_master_path.exists():
        _write_xml_part(work_dir, notes_master_part, _build_notes_master_root())

    # Content type override for notes master.
    part_name = f"/{notes_master_part}"
    has_override = False
    for override in content_types_root.findall("./ct:Override", XML_NS):
        if override.get("PartName") == part_name:
            has_override = True
            break
    if not has_override:
        override = ET.SubElement(content_types_root, _qn(CONTENT_TYPES_NS, "Override"))
        override.set("PartName", part_name)
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml")

    # Relationship from presentation.xml.rels to notes master.
    rel_id = None
    for rel in presentation_rels_root.findall("./rel:Relationship", XML_NS):
        if rel.get("Type") == NOTES_MASTER_REL_TYPE:
            rel_id = rel.get("Id")
            break
    if rel_id is None:
        rel_id = _append_relationship(
            presentation_rels_root,
            rel_type=NOTES_MASTER_REL_TYPE,
            target=posixpath.relpath(notes_master_part, posixpath.dirname("ppt/presentation.xml")),
        )

    # presentation.xml needs <p:notesMasterIdLst><p:notesMasterId r:id="..."/></...>
    notes_master_id_lst = presentation_root.find("./p:notesMasterIdLst", XML_NS)
    if notes_master_id_lst is None:
        notes_master_id_lst = ET.Element(_qn(PML_NS, "notesMasterIdLst"))
        # Insert near the top but after sldIdLst if present.
        sld_id_lst = presentation_root.find("./p:sldIdLst", XML_NS)
        if sld_id_lst is not None:
            idx = list(presentation_root).index(sld_id_lst) + 1
            presentation_root.insert(idx, notes_master_id_lst)
        else:
            presentation_root.insert(0, notes_master_id_lst)
    # Ensure at least one notesMasterId exists.
    existing = notes_master_id_lst.find("./p:notesMasterId", XML_NS)
    if existing is None:
        node = ET.SubElement(notes_master_id_lst, _qn(PML_NS, "notesMasterId"))
        node.set(f"{{{DOC_REL_NS}}}id", rel_id)

    return notes_master_part


def _build_notes_master_root() -> ET.Element:
    master = ET.Element(_qn(PML_NS, "notesMaster"))
    c_sld = ET.SubElement(master, _qn(PML_NS, "cSld"))
    sp_tree = ET.SubElement(c_sld, _qn(PML_NS, "spTree"))

    nv_grp = ET.SubElement(sp_tree, _qn(PML_NS, "nvGrpSpPr"))
    ET.SubElement(nv_grp, _qn(PML_NS, "cNvPr"), {"id": "1", "name": ""})
    ET.SubElement(nv_grp, _qn(PML_NS, "cNvGrpSpPr"))
    ET.SubElement(nv_grp, _qn(PML_NS, "nvPr"))

    grp_sp_pr = ET.SubElement(sp_tree, _qn(PML_NS, "grpSpPr"))
    ET.SubElement(grp_sp_pr, _qn(DML_NS, "xfrm"))

    # Minimal color mapping override.
    clr_map_ovr = ET.SubElement(master, _qn(PML_NS, "clrMapOvr"))
    ET.SubElement(clr_map_ovr, _qn(DML_NS, "masterClrMapping"))
    return master


def _write_notes_slide_rels(*, work_dir: Path, notes_part: str, notes_master_part: str) -> None:
    rels_part = _to_rels_path(notes_part)
    root = _new_relationships_root()
    _append_relationship(
        root,
        rel_type=NOTES_MASTER_REL_TYPE,
        target=posixpath.relpath(notes_master_part, posixpath.dirname(notes_part)),
    )
    _write_xml_part(work_dir, rels_part, root)


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
    """Best-effort: set speaker notes in the notes body placeholder."""
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
            shape.TextFrame.TextRange.Text = text
            return True
        except Exception:
            continue
    return False


def _find_sp_tree(slide_root: ET.Element) -> ET.Element | None:
    return slide_root.find("./p:cSld/p:spTree", XML_NS)


def _find_named_slide_element(sp_tree: ET.Element, name: str) -> tuple[int, ET.Element, ET.Element | None] | None:
    for index, child in enumerate(list(sp_tree)):
        c_nvpr = _find_non_visual_properties(child)
        if c_nvpr is not None and str(c_nvpr.get("name", "")) == name:
            return index, child, c_nvpr
    return None


def _find_non_visual_properties(element: ET.Element) -> ET.Element | None:
    tag = _local_name(element.tag)
    if tag == "sp":
        return element.find("./p:nvSpPr/p:cNvPr", XML_NS)
    if tag == "pic":
        return element.find("./p:nvPicPr/p:cNvPr", XML_NS)
    return None


def _extract_shape_transform(element: ET.Element) -> ET.Element | None:
    return element.find("./p:spPr/a:xfrm", XML_NS)


def _set_text_body(shape: ET.Element, text: str) -> bool:
    tx_body = shape.find("./p:txBody", XML_NS)
    if tx_body is None:
        return False

    body_pr = deepcopy(tx_body.find("./a:bodyPr", XML_NS))
    if body_pr is None:
        body_pr = ET.Element(_qn(DML_NS, "bodyPr"))
    lst_style = deepcopy(tx_body.find("./a:lstStyle", XML_NS))
    if lst_style is None:
        lst_style = ET.Element(_qn(DML_NS, "lstStyle"))
    tx_body.clear()
    tx_body.append(body_pr)
    tx_body.append(lst_style)
    for paragraph_text in text.splitlines() or [""]:
        tx_body.append(_build_text_paragraph(paragraph_text))
    return True


def _build_text_paragraph(text: str) -> ET.Element:
    paragraph = ET.Element(_qn(DML_NS, "p"))
    if not text:
        return paragraph
    run = ET.SubElement(paragraph, _qn(DML_NS, "r"))
    node = ET.SubElement(run, _qn(DML_NS, "t"))
    if text != text.strip() or "  " in text:
        node.set(XML_SPACE, "preserve")
    node.text = text
    return paragraph


def _build_picture_element(c_nvpr: ET.Element, xfrm: ET.Element, rel_id: str) -> ET.Element:
    pic = ET.Element(_qn(PML_NS, "pic"))
    nv_pic_pr = ET.SubElement(pic, _qn(PML_NS, "nvPicPr"))
    nv_pic_pr.append(deepcopy(c_nvpr))
    ET.SubElement(nv_pic_pr, _qn(PML_NS, "cNvPicPr"))
    ET.SubElement(nv_pic_pr, _qn(PML_NS, "nvPr"))

    blip_fill = ET.SubElement(pic, _qn(PML_NS, "blipFill"))
    blip = ET.SubElement(blip_fill, _qn(DML_NS, "blip"))
    blip.set(f"{{{DOC_REL_NS}}}embed", rel_id)
    stretch = ET.SubElement(blip_fill, _qn(DML_NS, "stretch"))
    ET.SubElement(stretch, _qn(DML_NS, "fillRect"))

    sp_pr = ET.SubElement(pic, _qn(PML_NS, "spPr"))
    sp_pr.append(deepcopy(xfrm))
    prst_geom = ET.SubElement(sp_pr, _qn(DML_NS, "prstGeom"))
    prst_geom.set("prst", "rect")
    ET.SubElement(prst_geom, _qn(DML_NS, "avLst"))
    return pic


def _resolve_primary_slide_path(presentation_root: ET.Element, presentation_rels_root: ET.Element) -> str | None:
    sld_id = presentation_root.find("./p:sldIdLst/p:sldId", XML_NS)
    rel_id = sld_id.get(f"{{{DOC_REL_NS}}}id") if sld_id is not None else None
    return _resolve_related_part_path("ppt/presentation.xml", presentation_rels_root, SLIDE_REL_TYPE, rel_id=rel_id)


def _resolve_related_part_path(source_part: str, rels_root: ET.Element, rel_type: str, *, rel_id: str | None = None) -> str | None:
    for relationship in rels_root.findall("./rel:Relationship", XML_NS):
        if rel_id is not None and relationship.get("Id") != rel_id:
            continue
        if relationship.get("Type") != rel_type:
            continue
        target = relationship.get("Target", "")
        if not target:
            continue
        return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))
    return None


def _read_xml_part(work_dir: Path, part_path: str | None, *, required: bool = True) -> ET.Element | None:
    if not part_path:
        return None
    xml_path = work_dir / Path(part_path)
    if not xml_path.exists():
        if required:
            raise ComposeError(f"Invalid .pptx package: missing {part_path}")
        return None
    try:
        return ET.parse(xml_path).getroot()
    except ET.ParseError as exc:
        raise ComposeError(f"Malformed .pptx XML: {exc}") from exc


def _write_xml_part(work_dir: Path, part_path: str, root: ET.Element) -> None:
    xml_path = work_dir / Path(part_path)
    xml_path.parent.mkdir(parents=True, exist_ok=True)
    ET.ElementTree(root).write(xml_path, encoding="utf-8", xml_declaration=True)


def _to_rels_path(part_path: str) -> str:
    parent = posixpath.dirname(part_path)
    base = posixpath.basename(part_path)
    return f"{parent}/_rels/{base}.rels"


def _new_relationships_root() -> ET.Element:
    return ET.Element(_qn(PKG_REL_NS, "Relationships"))


def _append_relationship(rels_root: ET.Element, *, rel_type: str, target: str) -> str:
    rel_id = f"rId{_next_relationship_index(rels_root)}"
    relationship = ET.SubElement(rels_root, _qn(PKG_REL_NS, "Relationship"))
    relationship.set("Id", rel_id)
    relationship.set("Type", rel_type)
    relationship.set("Target", target)
    return rel_id


def _set_relationship_target(rels_root: ET.Element, *, rel_type: str, target: str) -> None:
    for relationship in rels_root.findall("./rel:Relationship", XML_NS):
        if relationship.get("Type") == rel_type:
            relationship.set("Target", target)
            return
    _append_relationship(rels_root, rel_type=rel_type, target=target)


def _remove_relationships_by_type(rels_root: ET.Element, rel_type: str) -> None:
    for relationship in list(rels_root.findall("./rel:Relationship", XML_NS)):
        if relationship.get("Type") == rel_type:
            rels_root.remove(relationship)


def _next_relationship_index(rels_root: ET.Element) -> int:
    highest = 0
    for relationship in rels_root.findall("./rel:Relationship", XML_NS):
        rel_id = relationship.get("Id", "")
        match = re.fullmatch(r"rId(\d+)", rel_id)
        if match:
            highest = max(highest, int(match.group(1)))
    return highest + 1


def _rebuild_presentation_slides(presentation_root: ET.Element, presentation_rels_root: ET.Element, slide_parts: list[str]) -> None:
    sld_id_lst = presentation_root.find("./p:sldIdLst", XML_NS)
    existing_ids = [
        int(node.get("id", "255"))
        for node in presentation_root.findall("./p:sldIdLst/p:sldId", XML_NS)
        if str(node.get("id", "")).isdigit()
    ]
    if sld_id_lst is None:
        sld_id_lst = ET.Element(_qn(PML_NS, "sldIdLst"))
        presentation_root.insert(0, sld_id_lst)
    else:
        for child in list(sld_id_lst):
            sld_id_lst.remove(child)
    next_slide_id = max(existing_ids or [255]) + 1

    _remove_relationships_by_type(presentation_rels_root, SLIDE_REL_TYPE)
    next_rel_index = _next_relationship_index(presentation_rels_root)
    for slide_part in slide_parts:
        rel_id = f"rId{next_rel_index}"
        next_rel_index += 1
        relationship = ET.SubElement(presentation_rels_root, _qn(PKG_REL_NS, "Relationship"))
        relationship.set("Id", rel_id)
        relationship.set("Type", SLIDE_REL_TYPE)
        relationship.set("Target", posixpath.relpath(slide_part, posixpath.dirname("ppt/presentation.xml")))

        slide_id = ET.SubElement(sld_id_lst, _qn(PML_NS, "sldId"))
        slide_id.set("id", str(next_slide_id))
        slide_id.set(f"{{{DOC_REL_NS}}}id", rel_id)
        next_slide_id += 1


def _rebuild_content_types(
    content_types_root: ET.Element,
    slide_parts: list[str],
    notes_parts: list[str],
    media_dir: Path,
) -> None:
    for override in list(content_types_root.findall("./ct:Override", XML_NS)):
        part_name = override.get("PartName", "")
        if part_name.startswith("/ppt/slides/slide") or part_name.startswith("/ppt/notesSlides/notesSlide"):
            content_types_root.remove(override)

    for slide_part in slide_parts:
        override = ET.SubElement(content_types_root, _qn(CONTENT_TYPES_NS, "Override"))
        override.set("PartName", f"/{slide_part}")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")

    for notes_part in notes_parts:
        override = ET.SubElement(content_types_root, _qn(CONTENT_TYPES_NS, "Override"))
        override.set("PartName", f"/{notes_part}")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")

    existing_defaults = {
        node.get("Extension", "").lower()
        for node in content_types_root.findall("./ct:Default", XML_NS)
    }
    media_files = sorted(media_dir.iterdir()) if media_dir.exists() else []
    for media_file in media_files:
        extension = media_file.suffix.lower().lstrip(".")
        if not extension or extension in existing_defaults:
            continue
        content_type = _image_content_type(media_file)
        if content_type is None:
            continue
        default = ET.SubElement(content_types_root, _qn(CONTENT_TYPES_NS, "Default"))
        default.set("Extension", extension)
        default.set("ContentType", content_type)
        existing_defaults.add(extension)


def _copy_media_asset(image_path: Path, media_dir: Path, *, slide_index: int, shape_name: str) -> str:
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", shape_name).strip("_") or "image"
    suffix = image_path.suffix.lower()
    candidate = f"compose-{slide_index:03d}-{safe_name}{suffix}"
    destination = media_dir / candidate
    dedupe = 1
    while destination.exists():
        candidate = f"compose-{slide_index:03d}-{safe_name}-{dedupe}{suffix}"
        destination = media_dir / candidate
        dedupe += 1
    shutil.copy2(image_path, destination)
    return candidate


def _image_content_type(path: Path) -> str | None:
    guessed, _ = mimetypes.guess_type(path.name)
    if guessed and guessed.startswith("image/"):
        return guessed
    return None


def _write_zip_from_directory(source_dir: Path, out_path: Path) -> None:
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(source_dir).as_posix())


def _qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def _local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]

