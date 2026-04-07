---
name: note2video
description: Use the note2video CLI to turn PowerPoint (.pptx) speaker notes into per-slide images, scripts, voice-over, subtitles, and MP4; requires a one-time local pip install.
metadata: {"openclaw":{"emoji":"🎬","homepage":"https://github.com/openclawee/note2video","os":["win32","linux","darwin"],"requires":{"anyBins":["python","python3"]}}}
---

# Note2Video

Guide for agents to use **note2video**: a CLI tool that turns PPTX into an assets directory or a full MP4.

## Important: OpenClaw will not install this package automatically

**This skill does not include the Python source package.** OpenClaw typically will **not** run `pip install` for you (sandboxing, approvals, or missing install flows). Treat installation as something the **user runs in their local terminal**, or only run it via `exec` **when the user explicitly agrees and the environment allows network + writes**.

**Before running any `extract` / `build`**, the agent should:

1. If `exec` is allowed, **probe first**:  
   `python -m note2video --help`  
   or  
   `python3 -m note2video --help`
2. If it fails (e.g. `No module named note2video`, missing `python`, non-zero exit code): **do not assume it is installed.**  
   - Send the **User install (one-time)** block below and ask the user to run it in **their own terminal** (if the sandbox forbids installs, do not force-install in the sandbox).  
   - Only run `pip install` via `exec` when the **user explicitly requests it** and the environment allows **network + writes** to the chosen venv.

## User install (one-time) — copy/paste as-is

Choose one method. Requires **Python 3.10+**. Always use a **venv** to avoid polluting system Python.

### Windows（PowerShell）

```powershell
cd $env:USERPROFILE
python -m venv note2video-venv
.\note2video-venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install "git+https://github.com/openclawee/note2video.git"
python -m note2video --help
```

### Linux / macOS（bash）

```bash
cd ~
python3 -m venv note2video-venv
source note2video-venv/bin/activate
python -m pip install -U pip
python -m pip install "git+https://github.com/openclawee/note2video.git"
python -m note2video --help
```

After installation, always use the **Python inside that venv**, for example:

- Windows: `~\note2video-venv\Scripts\python.exe -m note2video ...`
- Unix-like: `~/note2video-venv/bin/python -m note2video ...`

### Single-repo mode (skill inside a cloned note2video repo)

Repo root = the parent directory of `skills/` (where `pyproject.toml` exists):

```bash
python -m venv .venv
# After activating the venv:
python -m pip install -U pip && python -m pip install -e .
python -m note2video --help
```

## When to use this skill

- The user has a `.pptx` and needs **extract** (export per-slide images + notes + script) or **build** (generate an MP4 end-to-end).

## Commands (after install; Python invocation as above)

```bash
python -m note2video extract "/path/to/deck.pptx" --out "./dist/deck01"
python -m note2video build "/path/to/deck.pptx" --out "./dist/deck01"
```

Add `--json` for machine-readable output.

## Platforms and slide images

- **Windows**: when Office is installed, prefer **PowerPoint COM** to export real PNGs.
- **Linux / macOS**: if **`soffice` or `libreoffice`** and **`pdftoppm`** are both in `PATH`, export real PNGs; otherwise use OpenXML **placeholder images**.  
  Set `NOTE2VIDEO_USE_LIBREOFFICE=0` to force placeholders.

## Linux system dependencies (only needed for real slide images)

```bash
sudo apt install -y libreoffice-nogui poppler-utils
```

## Docker (Linux image with LibreOffice + Poppler)

```bash
docker build -t note2video <path to the repo containing Dockerfile>
docker run --rm -v "$PWD:/work" -w /work note2video build ./deck.pptx --out ./dist
```

## Troubleshooting

- **`No module named note2video`**: not installed, or you used the wrong interpreter — run **User install** above, or explicitly use the venv `python`.
- **PowerPoint errors (Windows)**: open the file in PowerPoint first; requires a Windows desktop with Office installed.
- **`LibreOffice … failed`**: install Impress and Poppler; if needed set `NOTE2VIDEO_LIBREOFFICE` to the `soffice` executable.

## Security

Quote paths, and do not interpolate untrusted user input directly into shell commands.
