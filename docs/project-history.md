# Project history archive

## Purpose

This document archives key background, design decisions, implementation progress, and next steps for the `Note2Video` project.

It is not meant to replace detailed design docs. Instead, it serves as a portable “project context entry point” so someone can quickly pick up the project even if local directory names change, the machine changes, or IDE workspace history is lost.

## Project overview

- Project name: `Note2Video`
- Chinese name: `Beizhu Chengpian`
- Current form: open-source, free, **CLI-first**
- Core goal: automatically turn PowerPoint speaker notes into narrated videos (voice-over + subtitles)

The high-level positioning is recorded in:

- `docs/decisions/001-product-positioning.md`
- `docs/decisions/002-cli-first-strategy.md`
- `docs/decisions/003-mvp-scope.md`

## Starting point

The project originally started from the idea of a reusable “skill / agent tool”, not just a local app.

After narrowing down, the following product judgments were made:

- Build an automatable, debuggable, reusable CLI before UI/add-ins/editors.
- Instead of controlling editors like CapCut directly, first make the core `pptx` → `mp4` pipeline solid.
- The competitive advantage is not “effects editing”, but efficiently turning existing PPT assets into narrated videos.
- If packaged as a skill in the future, it should reuse the CLI rather than re-implementing logic.

## Confirmed product direction

### Main value

- Reuse existing PPT + notes content
- Automate slide export, notes extraction, script generation, voice-over, subtitles, and video rendering
- Make the whole workflow scriptable, repeatable, and agent-callable

### Explicit non-goals

- A full video editor
- Competing directly with CapCut on editing capabilities
- Complex motion effects or multi-speaker dubbing as the first priority
- PowerPoint add-ins and Web SaaS in phase 1

## Current CLI scope

Main command and subcommands:

- `note2video build`
- `note2video extract`
- `note2video voice`
- `note2video voices`
- `note2video subtitle`
- `note2video render`

Core principles:

- One command can run the full pipeline
- Each step can also be run independently
- Intermediate artifacts must be written to disk and be inspectable/debuggable

## Current implementation status

The project already provides a minimal end-to-end loop from `.pptx` to `.mp4`.

### Completed capabilities

- Python project scaffold and CLI entry point
- A real `extract` implementation
- Export slide images via PowerPoint COM
- Extract per-slide notes and produce `raw_notes`, `speaker_notes`, and `script`
- Output structured JSON plus per-slide and aggregated `.txt` files
- Voice pipeline (`voice`)
- Two TTS providers: `pyttsx3` and `edge-tts`
- `voices` command to list available voices
- Subtitle generation (`subtitle`) producing `srt` and `json`
- Video rendering (`render`) via ffmpeg
- Full orchestration via `build`
- Unit tests covering the main CLI paths

### Current output structure

Typical artifacts include:

- `manifest.json`
- `slides/*.png`
- `notes/notes.json`
- `notes/raw/*.txt`
- `notes/speaker/*.txt`
- `notes/all.txt`
- `scripts/script.json`
- `scripts/txt/*.txt`
- `scripts/all.txt`
- `audio/*.wav`
- `audio/merged.wav`
- `audio/timings.json`
- `subtitles/subtitles.srt`
- `subtitles/subtitles.json`
- `video/output.mp4`

## Key design evolution

### 1. Text layering

Intermediate text is currently split into three layers:

- `raw_notes`: raw text extracted from PPT notes pages
- `speaker_notes`: lightly cleaned, readable notes
- `script`: further transformed for narration and subtitle splitting

Why:

- Easier to debug extraction issues
- Easier to review/copy manually
- Leaves space for future script cleaning and TTS optimization

### 2. Pluggable TTS providers

TTS is implemented as a provider interface instead of hardcoding logic into the CLI.

Currently integrated:

- `pyttsx3`: local, useful for offline validation
- `edge-tts`: higher quality, better for public demos

### 3. Subtitle timing upgraded from “estimated” to “real per-sentence durations”

Early subtitle timing was allocated by sentence-length ratios.

To improve sync quality, it was upgraded to:

- `voice` generates audio per sentence based on `script`
- Merge into per-slide and full merged audio
- Write `audio/timings.json`
- `subtitle` prefers this real per-sentence timeline

This means subtitle timing is not merely text-length-based, but derived from real TTS durations.

## Important issues already addressed

### 1. pytest could not import project modules

Problem:

- `pytest` could not find `note2video`

Fix:

- Add `src` to pytest `pythonpath` in `pyproject.toml`

### 2. PowerPoint hidden window compatibility

Problem:

- Some PowerPoint versions behave poorly when `app.Visible = 0`

Fix:

- Use visible mode explicitly for better compatibility

### 3. `Slide.Export` relative path instability

Problem:

- PowerPoint COM export is unreliable with relative paths

Fix:

- Always pass absolute paths

### 4. Notes extraction mistakenly captured placeholders/page numbers

Problem:

- When there are no real notes, extraction may still return meaningless text like `1 2 3`

Fix:

- Tighten notes block detection
- Distinguish placeholders from real notes content
- Add cleanup and filtering

### 5. `raw_notes` and `script` were too similar

Problem:

- Early script layer did not differ meaningfully from raw notes

Fix:

- Add `_to_speaker_notes` and `_to_script` conversions
- Make `script` closer to narration/subtitle usage

### 6. Subtitles only showed the first sentence (or only the first sentence per slide)

Problem:

- The SRT had multiple lines, but the burned subtitles did not update correctly

Root cause:

- The initial slide-video track generation method was unfriendly for per-sentence subtitle updates

Fix:

- Generate a fixed-fps base video first, then burn subtitles

### 7. Edge TTS output format inconsistency

Problem:

- `edge-tts` output was not fully compatible with the WAV processing chain

Fix:

- Write temporary audio, then convert into a standard WAV format

## Current verification status

Verified so far:

- `tests/test_cli.py` passes
- The full `build` pipeline runs on a sample `.pptx`
- Real outputs include:
  - `audio/timings.json`
  - `subtitles/subtitles.srt`
  - `video/output.mp4`
- `manifest.json` correctly records:
  - `merged_audio`
  - `timings`
  - `subtitle`
  - `video`
  - `video_subtitles_burned`

## Highest-value next directions

### Direction 1: improve script cleaning

Candidate improvements:

- Number normalization / pronunciation
- Handling of English abbreviations
- Bracket content cleanup
- Further splitting of long sentences
- Better segmentation for spoken Chinese rhythm

### Direction 2: improve subtitles and narration pacing

With per-sentence TTS timing already available, further work can include:

- Pause control between sentences
- Minimum subtitle duration
- Auto-merge very short sentences
- Auto-split overly long sentences
- In-slide timing micro-adjustments

### Direction 3: prepare for public open-sourcing

Before publishing to GitHub, consider adding:

- `.gitignore`
- A first public README
- Example inputs/outputs
- Installation and environment requirements
- Clear statements about Windows / PowerPoint dependencies

## Migration recommendations

If you rename directories, change machines, or move IDE workspaces, keep at least:

- This document: `docs/project-history.md`
- `README.md`
- `docs/cli.md`
- `docs/decisions/*.md`

If you need richer conversation context, capture key discussions as new decision docs rather than relying only on IDE chat history.

## Maintenance recommendations

After each major milestone, update these sections:

- Current implementation status
- Important issues already addressed
- Current verification status
- Highest-value next directions

This helps preserve project history over time instead of scattering it across chat logs.
