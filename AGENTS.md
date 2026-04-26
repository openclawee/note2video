# AGENTS.md

## Cursor Cloud specific instructions

### Overview

Note2Video is a Python CLI tool that converts PowerPoint speaker notes into narrated videos. The source lives in `src/note2video/` and is installed as an editable package via `pip install -e .`.

### Running tests

```bash
python3 -m pytest tests/ -v
```

11 of 13 tests pass on Linux. Two `extract`-related tests (`test_extract_command_writes_expected_files`, `test_extract_command_filters_pages`) fail because `extract_project()` has a hard `sys.platform != "win32"` guard that raises before the monkeypatch can intercept `_extract_with_powerpoint`. This is a pre-existing code issue, not an environment problem.

### Running CLI commands

The `note2video` CLI entry point is installed to `~/.local/bin/`. Ensure `PATH` includes it:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

Supported commands on Linux: `voice`, `subtitle`, `render`, `voices`, `build` (partially — `extract` step requires Windows). The `extract` command requires Windows with Microsoft PowerPoint installed and will error on Linux.

### Hello-world demo (Linux)

To exercise the pipeline on Linux, prepare a `script.json` manually (or use test fixtures), then run:

```bash
note2video voice script.json --out ./dist --tts-provider edge --json
note2video subtitle script.json --out ./dist --json
```

The `render` step additionally requires slide images and merged audio to exist.

### Dependencies

All Python dependencies are declared in `pyproject.toml`. There is no `requirements.txt` or lock file. `pytest` is needed for testing but is not listed as a project dependency.
