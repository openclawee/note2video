# 003 MVP scope

## Date

2026-03-25

## Status

Accepted

## Decision

The first MVP of `Note2Video` focuses on a narrow but end-to-end narrated-video pipeline.

## In scope

- Input: `.pptx`
- Export slide images
- Extract notes per slide
- Transform notes into a narration script
- Use one primary TTS provider to generate voice-over
- Generate `srt` subtitles
- Render the final `mp4`
- Support `16:9` and `9:16`
- Provide both a one-shot command and step-by-step subcommands

## Core commands

- `note2video build`
- `note2video extract`
- `note2video voice`
- `note2video subtitle`
- `note2video render`

## Out of scope

- Full-fidelity preservation of PowerPoint animations
- Advanced multi-speaker voice-over
- Complex music mixing
- Automatic visual enhancement / content generation
- CapCut integrations
- PowerPoint add-in form factor
- Web SaaS form factor
- Multi-language localization workflows

## Rationale

This scope is small enough to build and test quickly, while still validating product value:

- Users can go from a `.pptx` to a final video
- Each step produces inspectable, debuggable intermediate artifacts
- The CLI serves both individual usage and future automation packaging

## MVP quality bar

The MVP can be considered successful if:

- One command can generate a usable video from a PPT with notes
- Intermediate outputs are clear, readable, and debuggable
- Repeated runs are stable
- The project structure is suitable for future skill packaging

## Implementation note

During `extract`, in addition to JSON artifacts, also output per-slide and aggregated `.txt` files for easy manual copy/review/editing.

## Next development priorities

1. `extract`
2. Notes/manifest schema
3. First TTS provider
4. Subtitle generation
5. Video rendering
6. Top-level `build` orchestration
