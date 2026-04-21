# Note2Video

Note2Video turns PowerPoint speaker notes into narrated videos (voice-over + subtitles + MP4). The project is currently developed **CLI-first**, and can later be wrapped as a reusable skill / automation tool.

## Why

Many training/teaching/content workflows start from PPT + notes. Turning them into publishable videos usually means repetitive manual steps:

- Export slide images
- Clean up notes into a script
- Generate voice-over
- Generate subtitles
- Compose images + audio + subtitles into a final video

Note2Video focuses on making this pipeline **scriptable, repeatable, and automation-friendly**.

## Goals

- Export a `.pptx` into per-slide images
- Extract notes per slide
- Transform notes into a narration-friendly script
- Generate voice-over audio
- Generate subtitle files
- Render the final narrated video
- Provide a stable CLI for automation and skill packaging

## Non-goals

- A full video editor
- A complex motion graphics system
- Perfectly preserving all PowerPoint animations
- Competing head-to-head with full-featured editors (e.g. CapCut) on editing features

## Target users

- Enterprise training / enablement teams
- Teachers, instructors, course creators
- Creators whose source content is PPT
- Developers embedding this pipeline into automated systems

## MVP scope

The first usable version aims to support:

- Input: `.pptx`
- Output: slide images, notes JSON, per-slide text, script JSON, audio files, `srt`, and `mp4`
- One primary language workflow
- One or two TTS providers
- `16:9` and `9:16` ratios
- Step-by-step subcommands for debugging

## CLI overview

Main command:

```bash
note2video build input.pptx --out ./dist
```

## Windows GUI (PySide6)

Install GUI dependencies:

```bash
python -m pip install -e ".[gui]"
```

Launch the app:

```bash
note2video-gui
```

Or launch via the main CLI (recommended):

```bash
note2video gui
```

Standalone Electron editor UI (optional; requires Node.js, run from the repo checkout):

```bash
note2video gui --editor
```

TTS 相关设置（MiniMax、火山引擎 / 豆包等）可在菜单 **设置 → TTS Provider…** 中编辑并写入用户配置文件（见下文各小节）。

Package as an exe (PyInstaller; recommended to run in a clean venv):

```bash
python -m pip install -e ".[gui,dev]"
pyinstaller --noconsole --name note2video-gui -m note2video.gui.app
```

The current `build` pipeline runs, in order:

- `extract`
- `voice`
- `subtitle`
- `render`

Subcommands:

- `build`: run the full pipeline
- `extract`: export slide images, notes, and script
- `voice`: generate voice-over from the script
- `voices`: list available voices
- `subtitle`: generate subtitle files
- `render`: render the final video from prepared assets

See `docs/cli.md` for detailed command documentation.

## MiniMax TTS (optional)

This project supports MiniMax **HTTP T2A API** for voice-over generation, and is designed so additional cloud TTS providers can be added via a provider layer.

**China and Global MiniMax accounts use different API hosts.** Your API key must be used with the corresponding host, otherwise you may get errors like `invalid api key (2049)`. Confirm the correct domain in the official console/docs. Common mapping:

| Account / Console | Common API host (origin) |
|------------------|--------------------------|
| Mainland China platform | `https://api.minimax.chat` |
| Global platform | `https://api.minimaxi.chat` (note the extra **i**) |
| Some OpenAPI examples | `https://api.minimax.io` |

Enable:

- Choose provider in CLI/GUI:
  - China: `minimax_cn` (fixed host `https://api.minimax.chat`)
  - Global: `minimax_global` (fixed host `https://api.minimaxi.chat`)
- API key: **`NOTE2VIDEO_MINIMAX_API_KEY`** or **`MINIMAX_API_KEY`** (can also be saved via GUI settings)
- Optional: `NOTE2VIDEO_MINIMAX_MODEL` (default `speech-2.8-hd`)
- Optional: `NOTE2VIDEO_MINIMAX_TIMEOUT_S` (defaults: ~60s for synthesis, ~30s for listing voices)

**User config file** (shared by CLI and GUI settings):

- **Windows**: `%LOCALAPPDATA%\note2video\config.json`
- **Linux / macOS**: `~/.config/note2video/config.json`

Supported fields (JSON): `tts.default_provider`, plus `tts.providers.minimax_cn.api_key` / `model` / `timeout_s`, `tts.providers.minimax_global.api_key` / `model` / `timeout_s`, and `tts.providers.volcengine.appid` / `token` / `cluster` / `base_url` / `timeout_s` (optional integer seconds). **Priority**: environment variables override config file.

CLI can also override per run:

```bash
note2video build input.pptx --out ./dist --tts-provider minimax_cn --voice "Chinese (Mandarin)_News_Anchor" --tts-rate 1.1
```

## 火山引擎 / 豆包语音合成（可选）

在线短文本合成使用 OpenSpeech HTTP 接口（与控制台「豆包语音 / 语音合成」文档一致）。CLI / GUI 中 Provider 名：`volcengine`；别名 **`doubao`**、**`doubao_tts`**（与 `volcengine` 等价）。

- **凭据**：`NOTE2VIDEO_VOLC_APPID`、`NOTE2VIDEO_VOLC_TOKEN`（或 `VOLC_TOKEN`）；也可写入用户配置 `tts.providers.volcengine.appid` / `token`
- **Cluster**：`NOTE2VIDEO_VOLC_CLUSTER`，或配置 `cluster`（常见为 `volcano_tts`，以控制台为准）
- **接口地址**（可选）：`NOTE2VIDEO_VOLC_TTS_URL` 或 `NOTE2VIDEO_DOUBAO_TTS_URL`，或配置 `base_url`；默认 `https://openspeech.bytedance.com/api/v1/tts`
- **超时**：`NOTE2VIDEO_VOLC_TIMEOUT_S` 或配置 `timeout_s`（默认约 60s）
- **音色**：`voice_type` 与控制台一致（GUI 可手动填写或从内置列表选）；默认示例 `BV700_streaming`

```bash
note2video build deck.pptx --out ./dist --tts-provider volcengine --voice BV700_streaming
# 或
note2video build deck.pptx --out ./dist --tts-provider doubao --voice BV700_streaming
```

## Platforms and slide export

- **Windows**: prefer Microsoft PowerPoint COM to export real slide images; fallback to OpenXML + placeholder images.
- **Linux / macOS**: if both `soffice` (or `libreoffice`) and `pdftoppm` (Poppler) are available in `PATH`, export via “LibreOffice headless → PDF → `pdftoppm` → PNG”; otherwise fallback to OpenXML + placeholders.

Debian / Ubuntu example:

```bash
sudo apt install libreoffice-nogui poppler-utils
```

Optional environment variables:

- `NOTE2VIDEO_USE_LIBREOFFICE`: set to `0`, `false`, or `off` to disable the LibreOffice path (e.g. for testing).
- `NOTE2VIDEO_LIBREOFFICE`: absolute path to the `soffice` executable.
- `NOTE2VIDEO_PDF_RENDER_DPI`: `pdftoppm` DPI (default `150`).

On **Windows**, you do not need LibreOffice; the tool prefers PowerPoint COM export.

### Docker (Linux image only)

The container image includes `libreoffice-nogui` and `poppler-utils` for slide export in headless environments. On Windows, run Python locally and do not use this image as a replacement for the PowerPoint COM path.

```bash
docker build -t note2video .
docker run --rm -v "%CD%:/work" -w /work note2video extract ./deck.pptx --out ./dist
```

(In PowerShell, use `-v "${PWD}:/work"`.)

CI (`.github/workflows/ci.yml`) installs system dependencies on **Ubuntu** to match Docker behavior; on **Windows** it does **not** install LibreOffice. Tests set `NOTE2VIDEO_USE_LIBREOFFICE=0` to avoid unstable conversions for tiny test `.pptx` inputs.

## Example output directory

```text
dist/
  manifest.json
  slides/
    001.png
    002.png
  notes/
    notes.json
    all.txt
    raw/
      001.txt
      002.txt
    speaker/
      001.txt
      002.txt
  scripts/
    script.json
    all.txt
    txt/
      001.txt
      002.txt
  audio/
    001.wav
    002.wav
    merged.wav
  subtitles/
    subtitles.srt
    subtitles.json
  video/
    output.mp4
  logs/
    build.log
```

## Notes on outputs

In addition to structured JSON, `extract` also writes helper `.txt` files for easy manual copy/paste:

- `notes/raw/*.txt`: per-slide raw notes
- `notes/speaker/*.txt`: per-slide cleaned notes for narration
- `scripts/txt/*.txt`: per-slide script text
- `notes/all.txt`: all notes combined
- `scripts/all.txt`: all scripts combined

This preserves real newlines when copying into editors, instead of showing literal `\\n`.

Current text layers:

- `raw_notes`: extracted raw text from PPT notes pages
- `speaker_notes`: lightly cleaned, readable notes
- `script`: further segmented into narration/subtitle-friendly sentences

## Planned architecture

```text
src/note2video/
  cli/
  parser/
  script/
  tts/
  subtitle/
  render/
  schemas/
```

Module responsibilities:

- `parser`: parse `.pptx`, export slide images, extract notes
- `script`: clean/normalize notes, generate narration script
- `tts`: provider abstraction layer for voice generation
- `subtitle`: subtitle splitting and timeline generation
- `render`: compose images + audio + subtitles into the final video
- `schemas`: intermediate artifacts and manifest structures
