from __future__ import annotations

import argparse
import math
import os
import shutil
import tempfile
from pathlib import Path


def _export_slide_png(pres, *, slide_index: int, out_path: Path, width: int | None) -> None:
    # PowerPoint Slide.Export signature: Export(FileName, FilterName, ScaleWidth, ScaleHeight)
    # If ScaleWidth/ScaleHeight omitted, PowerPoint chooses default export size.
    if width and width > 0:
        pres.Slides(slide_index).Export(str(out_path), "PNG", int(width))
    else:
        pres.Slides(slide_index).Export(str(out_path), "PNG")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Export PPTX slides to PNG and stitch every N slides into one vertical image (1xN) via PowerPoint COM."
    )
    parser.add_argument("pptx", help="Path to .pptx")
    parser.add_argument("--chunk", type=int, default=4, help="Slides per stitched image (default: 4)")
    parser.add_argument("--out-dir", default="", help="Output directory (default: <pptx_dir>/pdf_chunks_images)")
    parser.add_argument("--prefix", default="", help="Output filename prefix (default: PPTX stem)")
    parser.add_argument("--export-width", type=int, default=0, help="Optional slide export width in pixels (0 = default)")
    args = parser.parse_args(argv)

    pptx_path = Path(args.pptx).expanduser().resolve()
    if not pptx_path.exists() or pptx_path.suffix.lower() != ".pptx":
        raise SystemExit(f"Invalid PPTX: {pptx_path}")

    chunk = int(args.chunk)
    if chunk <= 0:
        raise SystemExit("--chunk must be > 0")

    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else (pptx_path.parent / "pdf_chunks_images")
    out_dir.mkdir(parents=True, exist_ok=True)

    prefix = (args.prefix or pptx_path.stem).strip() or "slides"
    export_width = int(args.export_width or 0)

    from PIL import Image

    # COM automation must run in STA on Windows.
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    temp_root = Path(tempfile.mkdtemp(prefix="note2video_pptx_export_"))
    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = 1
        pres = app.Presentations.Open(str(pptx_path), WithWindow=False)
        total = int(pres.Slides.Count)
        if total <= 0:
            raise SystemExit("Presentation has no slides.")

        groups = int(math.ceil(total / chunk))
        created: list[Path] = []

        for gi in range(groups):
            start = gi * chunk + 1
            end = min((gi + 1) * chunk, total)
            out_name = f"{prefix}_{start:03d}-{end:03d}.png"
            out_path = out_dir / out_name
            print(f"[{gi+1}/{groups}] export+stitch slides {start}-{end} -> {out_path}")

            # Export slides to temp PNGs
            chunk_dir = temp_root / f"{start:03d}-{end:03d}"
            chunk_dir.mkdir(parents=True, exist_ok=True)
            slide_pngs: list[Path] = []
            for idx in range(start, end + 1):
                p = chunk_dir / f"{idx:03d}.png"
                if p.exists():
                    try:
                        p.unlink()
                    except OSError:
                        pass
                _export_slide_png(pres, slide_index=idx, out_path=p, width=export_width if export_width > 0 else None)
                slide_pngs.append(p)

            # Load and stitch vertically (1xN)
            images = [Image.open(str(p)).convert("RGB") for p in slide_pngs if p.exists()]
            if not images:
                raise RuntimeError(f"Failed to export slides {start}-{end}.")

            max_w = max(im.width for im in images)
            total_h = sum(im.height for im in images)
            canvas = Image.new("RGB", (max_w, total_h), (255, 255, 255))
            y = 0
            for im in images:
                # Center align horizontally if widths differ
                x = (max_w - im.width) // 2
                canvas.paste(im, (x, y))
                y += im.height
                im.close()

            if out_path.exists():
                try:
                    out_path.unlink()
                except OSError:
                    pass
            canvas.save(str(out_path), format="PNG", optimize=True)
            created.append(out_path)

        print(f"pptx: {pptx_path}")
        print(f"slides: {total}")
        print(f"chunk: {chunk}")
        print(f"out_dir: {out_dir}")
        for p in created:
            size = p.stat().st_size if p.exists() else -1
            print(f"- {p} ({size} bytes)")
    finally:
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
        try:
            shutil.rmtree(temp_root, ignore_errors=True)
        except Exception:
            pass

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

