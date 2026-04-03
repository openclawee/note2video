---
name: note2video
description: Convert a PPTX (speaker notes) into narrated video assets via the note2video CLI.
metadata: {"openclaw":{"os":["win32","linux","darwin"],"requires":{"anyBins":["py","python"],"bins":["git"]},"homepage":"https://github.com/openclawee/note2video"}}
---

## When to use

Use this skill when the user wants to turn a `.pptx` into:

- extracted slide images + notes + scripts (`extract`)
- or a full narrated video pipeline (`build`)

**Windows** uses PowerPoint COM when available for slide rasterization. **Linux / macOS** use LibreOffice (`soffice`) plus `pdftoppm` (Poppler) when those binaries are on `PATH`; otherwise slide images fall back to OpenXML placeholders.

## Preconditions (one-time per workspace)

Run these commands from the repo root (the folder that contains `pyproject.toml`).

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -e .
```

Quick sanity check:

```powershell
python -m note2video --help
python -m note2video extract --help
```

## Extract assets from a PPTX

```powershell
.\.venv\Scripts\Activate.ps1
python -m note2video extract "C:\path\to\deck.pptx" --out ".\dist\deck01"
```

## Build full video (end-to-end)

```powershell
.\.venv\Scripts\Activate.ps1
python -m note2video build "C:\path\to\deck.pptx" --out ".\dist\deck01"
```

## Notes / troubleshooting

- If you see `No module named note2video`, you are not using the venv Python. Re-run `.\.venv\Scripts\Activate.ps1` and try again.
- If PowerPoint export fails, confirm PowerPoint is installed and can open the file, and re-run from an interactive Windows session (not a headless server).

