from __future__ import annotations

import argparse
import math
import os
from pathlib import Path


def _export_presentation_to_pdf(pres, *, out_path: Path) -> None:
    # PowerPoint constants (avoid importing win32com constants module)
    ppFixedFormatTypePDF = 2
    ppFixedFormatIntentPrint = 2
    ppPrintOutputSlides = 1
    ppPrintHandoutVerticalFirst = 1
    ppPrintAll = 1

    # NOTE: On some PowerPoint builds, PrintOptions.Ranges may be ignored by ExportAsFixedFormat,
    # causing the whole deck to be exported. To reliably export a subset, we create a temporary
    # presentation containing only the desired slides, then export that whole temp deck.
    pres.ExportAsFixedFormat(
        Path=str(out_path),
        FixedFormatType=ppFixedFormatTypePDF,
        Intent=ppFixedFormatIntentPrint,
        FrameSlides=True,
        HandoutOrder=ppPrintHandoutVerticalFirst,
        OutputType=ppPrintOutputSlides,
        PrintHiddenSlides=ppPrintAll,
        PrintRange=None,
        IncludeDocProperties=False,
        KeepIRMSettings=True,
        DocStructureTags=True,
        BitmapMissingFonts=True,
        UseISO19005_1=True,
    )


def _create_temp_presentation_with_slides(app, src_pres, *, start: int, end: int):
    # Create a new empty presentation and copy selected slides.
    temp = app.Presentations.Add()
    # Remove the default slide if it exists (some templates add one).
    try:
        while temp.Slides.Count:
            temp.Slides(1).Delete()
    except Exception:
        pass

    for idx in range(start, end + 1):
        src_pres.Slides(idx).Copy()
        # Paste at end
        temp.Slides.Paste(temp.Slides.Count + 1 if temp.Slides.Count else 1)
    return temp


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Split a PPTX into PDFs (N slides per file) via PowerPoint COM.")
    parser.add_argument("pptx", help="Path to .pptx")
    parser.add_argument("--chunk", type=int, default=4, help="Slides per PDF (default: 4)")
    parser.add_argument("--out-dir", default="", help="Output directory (default: <pptx_dir>/pdf_chunks)")
    parser.add_argument("--prefix", default="", help="Output filename prefix (default: PPTX stem)")
    args = parser.parse_args(argv)

    pptx_path = Path(args.pptx).expanduser().resolve()
    if not pptx_path.exists() or pptx_path.suffix.lower() != ".pptx":
        raise SystemExit(f"Invalid PPTX: {pptx_path}")

    chunk = int(args.chunk)
    if chunk <= 0:
        raise SystemExit("--chunk must be > 0")

    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else (pptx_path.parent / "pdf_chunks")
    out_dir.mkdir(parents=True, exist_ok=True)

    prefix = (args.prefix or pptx_path.stem).strip()
    if not prefix:
        prefix = "slides"

    # COM automation must run in STA on Windows.
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

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
        for i in range(groups):
            start = i * chunk + 1
            end = min((i + 1) * chunk, total)
            out_name = f"{prefix}_{start:03d}-{end:03d}.pdf"
            out_path = out_dir / out_name
            print(f"[{i+1}/{groups}] export slides {start}-{end} -> {out_path}")
            if out_path.exists():
                try:
                    out_path.unlink()
                except OSError:
                    pass
            temp_pres = None
            try:
                temp_pres = _create_temp_presentation_with_slides(app, pres, start=start, end=end)
                _export_presentation_to_pdf(temp_pres, out_path=out_path)
            finally:
                try:
                    if temp_pres is not None:
                        temp_pres.Close()
                except Exception:
                    pass
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

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

