# CLI design draft

## Command format

```bash
note2video <command> [options]
```

## Design principles

- Non-technical users should be able to run the full pipeline with a single command
- During development, each step should be runnable independently for debugging
- Output paths and directory structure should remain stable for skill packaging
- JSON output should be designed for automation first
- Also export `.txt` files for easy manual copy/paste and editing

## Main command

### `build`

Run the full pipeline from `.pptx` to the final video in one shot.

Current implementation runs (in order):

- `extract`
- `voice`
- `subtitle`
- `render`

```bash
note2video build input.pptx --out ./dist
```

Optional: override narration script (preferred over PPT notes when provided):

```bash
# Use an existing script.json (recommended for automation)
note2video build input.pptx --out ./dist --script-file ./my-script.json

# Or paste scripts/all.txt-style content (blocks start with `--- Slide 001 ---`)
note2video build input.pptx --out ./dist --script-file ./scripts/all.txt

# Or inline text (blank-line separated blocks; if block count == slide count, maps by order)
note2video build input.pptx --out ./dist --script-text "S1\n\nS2\n\nS3"
```

Typical outputs:

- `slides/`
- `notes/notes.json`
- `notes/raw/*.txt`
- `notes/speaker/*.txt`
- `scripts/script.json`
- `scripts/all.txt`
- `scripts/txt/*.txt`
- `audio/*.wav`
- `subtitles/subtitles.srt`
- `video/output.mp4`
- `manifest.json`

## Subcommands

### `extract`

Extract slide images, notes, and script from a PowerPoint file.

```bash
note2video extract input.pptx --out ./work
```

Outputs:

- `slides/*.png`
- `notes/notes.json`
- `notes/all.txt`
- `notes/raw/*.txt`
- `notes/speaker/*.txt`
- `scripts/script.json`
- `scripts/all.txt`
- `scripts/txt/*.txt`
- `manifest.json`

### `voice`

Generate voice-over audio per slide from the script.

Currently supported:

- `edge`: recommended for higher-quality results (requires network)
- `pyttsx3`: local fallback, useful to validate the pipeline offline

```bash
note2video voice ./work/scripts/script.json --out ./work
note2video voice ./work/scripts/script.json --out ./work --tts-provider edge --voice zh-CN-XiaoxiaoNeural
```

Outputs:

- `audio/001.wav`
- `audio/002.wav`
- `audio/merged.wav`
- `audio/timings.json`
- Updates `manifest.json` with audio paths and durations

### `voices`

List available voices, so you can choose a voice ID before running `voice` or `build`.

```bash
note2video voices --tts-provider edge --keyword zh-CN
note2video voices --tts-provider edge --keyword Xiaoxiao --json
```

### `subtitle`

Generate subtitle files from the script and timing information.

Current implementation splits at sentence level and allocates subtitle timings based on per-slide audio duration.\nIf `audio/timings.json` exists, it will prefer the real per-sentence timing generated during TTS.

```bash
note2video subtitle ./work/scripts/script.json --audio ./work/audio --out ./work/subtitles
```

Outputs:

- `subtitles/subtitles.srt`
- `subtitles/subtitles.json`

### `render`

Render the final video from slide images, audio, and subtitles.

The current implementation uses the ffmpeg binary provided by `imageio-ffmpeg`, so you do not need a system-wide `ffmpeg` install.\nAfter rendering succeeds, it automatically cleans up intermediates like `video_only.mp4` and `slides.ffconcat`.

```bash
note2video render ./work --out ./dist/output.mp4
```

Outputs:

- `video/output.mp4`

## Common options

- `--out <dir>`: output directory
- `--temp <dir>`: temporary working directory
- `--config <file>`: config file path
- `--overwrite`: overwrite existing outputs
- `--verbose`: verbose logs
- `--quiet`: minimal logs
- `--json`: print a machine-readable JSON summary

## Input-related options

- `--notes-source <auto|speaker-notes|file>`
- `--script-file <file>`
- `--pages <range>`

Examples:

```bash
note2video build input.pptx --pages 1-5
note2video build input.pptx --script-file scripts.json
```

## TTS-related options

- `--tts-provider <name>` — 例如 `edge`、`minimax_cn`、`minimax_global`、`volcengine`（火山 / 豆包；别名 `doubao`）
- `--voice <id>`
- `--rate <value>`
- `--pitch <value>`
- `--style <value>`
- `--voice-config <file>`

Examples:

```bash
note2video build input.pptx --tts-provider edge --voice zh-CN-XiaoxiaoNeural
note2video build input.pptx --voice narrator_female --rate 1.05
```

## Script-related options

- `--script-mode <raw|clean|spoken>`
- `--max-sentence-length <n>`
- `--pause-ms <n>`
- `--normalize-numbers`
- `--remove-brackets`

Examples:

```bash
note2video build input.pptx --script-mode spoken --normalize-numbers
```

## Video-related options

- `--ratio <16:9|9:16|1:1>`
- `--resolution <720p|1080p|custom>`
- `--fps <n>`
- `--transition <none|fade>`
- `--slide-padding-ms <n>`

Examples:

```bash
note2video build input.pptx --ratio 9:16 --resolution 1080p
```

## Subtitle-related options

- `--subtitle <on|off|soft|burn>`
- `--subtitle-style <name>`
- `--subtitle-max-chars <n>`

Examples:

```bash
note2video build input.pptx --subtitle burn
note2video subtitle ./work/notes/script.json --subtitle-max-chars 18
```

## JSON output contract

When `--json` is enabled, the command should print a summary object to stdout.

Example:

```json
{
  "command": "build",
  "status": "ok",
  "input": "input.pptx",
  "output_dir": "./dist",
  "slide_count": 12,
  "artifacts": {
    "manifest": "./dist/manifest.json",
    "notes": "./dist/notes/notes.json",
    "script": "./dist/scripts/script.json",
    "subtitle": "./dist/subtitles/subtitles.srt",
    "video": "./dist/video/output.mp4"
  }
}
```

## Exit codes

- `0`: success
- `1`: generic runtime error
- `2`: CLI argument error
- `3`: input file not found or unsupported format
- `4`: PowerPoint parsing/export failed
- `5`: TTS generation failed
- `6`: subtitle generation failed
- `7`: video rendering failed

## Notes for skill packaging

Recommended high-level actions to expose:

- `ppt-to-video`
- `ppt-to-assets`

When packaging, prefer:

- Explicitly specify an output directory
- Use `--json` to obtain machine-readable results
- Keep artifact paths stable for follow-up runs and integration
