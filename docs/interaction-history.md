# Interaction History (Archive)

Date: 2026-04-07  
Workspace: `D:\dev\projects\note2video`  
OS: Windows (`win32 10.0.26200`)  
Shell: PowerShell

## Source transcript

- [GUI preview crash fix](f139c81f-5b9a-4519-9828-aa57d984430c)

## What happened (high-level)

- Synced repository to the latest `origin/main`.
- Installed GUI extras and launched `note2video-gui`.
- Investigated a crash when clicking **Preview/试听**:
  - Confirmed preview WAV files were being generated in `%TEMP%` with prefix `note2video_preview_*.wav`.
  - Observed native crashes with exit codes like `0xC0000409` (stack buffer overrun).
- Stabilized preview generation and playback:
  - Moved preview synthesis to a subprocess (`Popen`) and polled completion via `QTimer` (no Qt thread/signal for the preview path).
  - Added UI actions to play via the system default player and to reveal the file.
  - Added an embedded HTML `<audio>` player using QtWebEngine, loading local WAV files via a file-directory base URL.
  - Removed the “file generated” popup in favor of updating the right-side embedded player area and buttons.
- Translated project documentation from Chinese to English and verified no remaining CJK characters in `README.md`, `docs/`, and `skills/`.

## Key commands used (representative)

- Git sync:
  - `git fetch origin`
  - `git pull --ff-only`
- Install GUI dependencies:
  - `python -m pip install -e ".[gui]"`
- Run GUI:
  - `note2video-gui`
- Locate preview WAV files:
  - `%TEMP%` (e.g. `C:\Users\mayaz\AppData\Local\Temp\note2video_preview_*.wav`)

## Files changed in this session

### Code

- `src/note2video/gui/app.py`
  - Preview generation refactor to subprocess + `QTimer` polling
  - Embedded preview player (QtWebEngine) and player controls
- `src/note2video/tts/voice.py`
  - Explicit UTF-8 decoding for ffmpeg subprocess output (`encoding="utf-8", errors="replace"`)

### Documentation (translated to English)

- `README.md`
- `docs/cli.md`
- `docs/project-history.md`
- `docs/decisions/001-product-positioning.md`
- `docs/decisions/002-cli-first-strategy.md`
- `docs/decisions/003-mvp-scope.md`
- `skills/note2video/README.md`
- `skills/note2video/SKILL.md`

## Outcomes

- GUI preview no longer crashes the application.
- Preview audio can be played via:
  - System default player (manual action)
  - Embedded browser-based `<audio>` player (QtWebEngine), when available
- Documentation is now English-only.

