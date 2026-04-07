# 001 Product positioning

## Date

2026-03-25

## Status

Accepted

## Decision

The project name is `Note2Video` (Chinese name: `Beizhu Chengpian`).

It is positioned as an open-source, free, **CLI-first** tool that converts PowerPoint speaker notes into narrated videos with voice-over and subtitles.

## Rationale

This positioning is clearer than a generic “PPT-to-video tool”:

- The real input is speaker notes
- The output is a narrated/explanatory video
- It fits the open-source distribution model on GitHub
- It works well for a CLI and future skill packaging

The project does not compete directly with full editors. Its core value is automation, repeatability, and reusing existing PPT content assets.

## Target users

- Enterprise training / enablement teams
- Teachers, instructors, course creators
- Creators who already structure content in PPT
- Developers embedding this capability into automated workflows

## Non-goals

- A full video editor
- Advanced editing effects as the primary focus
- Competing on the number of TTS voices as a key differentiator

## Implications

- The first public deliverable should be a CLI
- The core pipeline must support scripting and automation
- The project must persist stable intermediate artifacts, not just the final MP4
- Future skill packaging should wrap the CLI instead of re-implementing the logic
