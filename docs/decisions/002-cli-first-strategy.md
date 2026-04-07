# 002 CLI-first strategy

## Date

2026-03-25

## Status

Accepted

## Decision

The first-stage deliverable of `Note2Video` is a CLI (not a desktop app, PowerPoint add-in, or web product).

The command name is:

```bash
note2video
```

## Rationale

A CLI is the best fit at this stage:

- Fastest way to validate the end-to-end pipeline
- Easier to rerun locally and debug iteratively
- Clearer input/output boundaries
- Easier to wrap as a reusable skill
- Better for early-stage adoption in an open-source project

The main problem to solve in phase 1 is not UI, but a stable core pipeline:

- Parse `.pptx`
- Export slide images and notes
- Generate voice-over
- Generate subtitles
- Render the final video

## Alternatives considered

### Desktop app first

Not chosen: it significantly increases scope and UI cost before the core capabilities are stable.

### PowerPoint add-in first

Not chosen: runtime constraints from the host app make local processing and skill packaging harder.

### Web app first

Not chosen: file privacy, rendering compatibility, and server-side cost are not the most urgent problems right now.

## Implications

- Each major pipeline step should be independently invokable
- Outputs should be stable and persisted to disk
- Logs and JSON summaries should be first-class features
- Future desktop/skill experiences should reuse the same core pipeline
