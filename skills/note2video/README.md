# OpenClaw Skill: `note2video`

This directory contains an **OpenClaw** skill that follows the **[Agent Skills](https://agentskills.io)** spec. It explains how to install and run the **note2video** CLI (PPTX → assets / MP4).

## Load into an OpenClaw workspace

1. Copy this entire directory into the workspace `skills/` root, for example:  
   `<workspace>/skills/note2video/`  
   so that `SKILL.md` is at `<workspace>/skills/note2video/SKILL.md`.

2. Start a new session or restart the Gateway so the skill is reloaded (e.g. `/new` or `openclaw gateway restart`).

3. Verify: run `openclaw skills list` and confirm `note2video` appears.

**OpenClaw does not automatically install Python packages.** After the skill is loaded, run the **User install (one-time)** section in `SKILL.md` in your local terminal (installing into a venv via `pip install git+https://github.com/openclawee/note2video.git`). Only let an agent run the install via `exec` if your environment allows networked `exec` and you explicitly authorize it.

If the skill is published on **ClawHub**, you may also use `openclaw skills install …` (follow the platform’s documentation at that time).

## Relation to this repository

After cloning this repository, the skill lives at `skills/note2video/` under the repo root. The `{baseDir}` placeholder in `SKILL.md` is resolved by OpenClaw to the skill directory, which helps the agent locate relative paths.

## ClawHub licensing note

Skills published on ClawHub are **MIT-0** per registry policy. Do not add license terms in `SKILL.md` that conflict with it.
