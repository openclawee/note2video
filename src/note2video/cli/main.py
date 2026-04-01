from __future__ import annotations

import argparse
import json
from pathlib import Path

from note2video import __version__
from note2video.parser.extract import PowerPointUnavailableError, extract_project
from note2video.render.video import RenderError, render_video
from note2video.subtitle.generate import SubtitleGenerationError, generate_subtitles
from note2video.tts.voice import (
    VoiceGenerationError,
    generate_voice_assets,
    list_available_voices,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="note2video",
        description="Turn PowerPoint speaker notes into narrated videos.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    subparsers = parser.add_subparsers(dest="command", required=True)

    build_cmd = subparsers.add_parser("build", help="Run the full pipeline.")
    _add_common_arguments(build_cmd)
    build_cmd.add_argument("input", help="Path to the input .pptx file.")
    build_cmd.add_argument("--ratio", default="16:9", help="Output ratio.")
    build_cmd.add_argument("--pages", default="all", help="Page selection, e.g. 1-3,5.")
    build_cmd.add_argument("--voice", default="", help="Voice ID placeholder.")
    build_cmd.add_argument("--tts-provider", default="", help="TTS provider placeholder.")
    build_cmd.set_defaults(handler=handle_build)

    extract_cmd = subparsers.add_parser("extract", help="Extract slide assets and notes.")
    _add_common_arguments(extract_cmd)
    extract_cmd.add_argument("input", help="Path to the input .pptx file.")
    extract_cmd.add_argument("--pages", default="all", help="Page selection, e.g. 1-3,5.")
    extract_cmd.set_defaults(handler=handle_extract)

    voice_cmd = subparsers.add_parser("voice", help="Generate voice-over audio.")
    _add_common_arguments(voice_cmd)
    voice_cmd.add_argument("input", help="Path to notes or script JSON.")
    voice_cmd.add_argument("--tts-provider", default="pyttsx3", help="TTS provider name.")
    voice_cmd.add_argument("--voice", default="", help="Voice ID.")
    voice_cmd.set_defaults(handler=handle_voice)

    voices_cmd = subparsers.add_parser("voices", help="List available voices.")
    voices_cmd.add_argument("--tts-provider", default="edge", help="TTS provider name.")
    voices_cmd.add_argument("--keyword", default="", help="Keyword filter.")
    voices_cmd.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="Print a machine-readable JSON summary.",
    )
    voices_cmd.set_defaults(handler=handle_voices)

    subtitle_cmd = subparsers.add_parser("subtitle", help="Generate subtitle files.")
    _add_common_arguments(subtitle_cmd)
    subtitle_cmd.add_argument("input", help="Path to notes or script JSON.")
    subtitle_cmd.set_defaults(handler=handle_subtitle)

    render_cmd = subparsers.add_parser("render", help="Render the final video.")
    _add_common_arguments(render_cmd)
    render_cmd.add_argument("input", help="Path to the prepared work directory.")
    render_cmd.set_defaults(handler=handle_render)

    return parser


def _add_common_arguments(command: argparse.ArgumentParser) -> None:
    command.add_argument("--out", default="./dist", help="Output directory.")
    command.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="Print a machine-readable JSON summary.",
    )


def handle_build(args: argparse.Namespace) -> int:
    out_dir = Path(args.out)
    manifest = extract_project(args.input, str(out_dir), pages=args.pages)
    script_path = out_dir / "scripts" / "script.json"

    voice_result = generate_voice_assets(
        str(script_path),
        str(out_dir),
        provider_name=args.tts_provider or "pyttsx3",
        voice_id=args.voice,
    )
    subtitle_result = generate_subtitles(str(script_path), str(out_dir))
    render_result = render_video(str(out_dir))

    payload = {
        "command": "build",
        "status": "ok",
        "phase": "complete",
        "input": args.input,
        "output_dir": str(out_dir),
        "artifacts": {
            "manifest": "manifest.json",
            "notes": "notes/notes.json",
            "script": "scripts/script.json",
            "audio_dir": "audio",
            "merged_audio": "audio/merged.wav",
            "subtitle": "subtitles/subtitles.srt",
            "subtitle_json": "subtitles/subtitles.json",
            "video": render_result["video"],
        },
        "slide_count": manifest.slide_count,
        "segment_count": subtitle_result["segment_count"],
        "voice_provider": voice_result["provider"],
        "subtitles_burned": render_result["subtitles_burned"],
        "message": "Full pipeline completed.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_extract(args: argparse.Namespace) -> int:
    manifest = extract_project(args.input, args.out, pages=args.pages)
    payload = {
        "command": "extract",
        "status": "ok",
        "phase": "implemented",
        "input": args.input,
        "output_dir": str(Path(args.out)),
        "artifacts": manifest.outputs,
        "slide_count": manifest.slide_count,
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_voice(args: argparse.Namespace) -> int:
    result = generate_voice_assets(
        args.input,
        args.out,
        provider_name=args.tts_provider,
        voice_id=args.voice,
    )
    payload = {
        "command": "voice",
        "status": "ok",
        "input": args.input,
        "output_dir": str(Path(args.out)),
        "slide_count": result["slide_count"],
        "provider": result["provider"],
        "voice": result["voice"],
        "artifacts": {
            "audio_dir": "audio",
            "merged_audio": "audio/merged.wav",
            "timings": "audio/timings.json",
        },
        "message": "Voice assets generated.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_voices(args: argparse.Namespace) -> int:
    voices = list_available_voices(provider_name=args.tts_provider, keyword=args.keyword)
    payload = {
        "command": "voices",
        "status": "ok",
        "provider": args.tts_provider,
        "count": len(voices),
        "voices": voices,
        "message": f"Listed {len(voices)} voices.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_subtitle(args: argparse.Namespace) -> int:
    result = generate_subtitles(args.input, args.out)
    payload = {
        "command": "subtitle",
        "status": "ok",
        "input": args.input,
        "output_dir": str(Path(args.out)),
        "slide_count": result["slide_count"],
        "segment_count": result["segment_count"],
        "artifacts": {
            "subtitle": "subtitles/subtitles.srt",
            "subtitle_json": "subtitles/subtitles.json",
        },
        "message": "Subtitle assets generated.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_render(args: argparse.Namespace) -> int:
    result = render_video(args.input, args.out if args.out != "./dist" else None)
    payload = {
        "command": "render",
        "status": "ok",
        "input": args.input,
        "output_dir": str(Path(args.input)),
        "slide_count": result["slide_count"],
        "artifacts": {
            "video": result["video"],
        },
        "subtitles_burned": result["subtitles_burned"],
        "message": "Video rendered.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_stub_command(args: argparse.Namespace) -> int:
    payload = {
        "command": args.command,
        "status": "todo",
        "message": f"The '{args.command}' command is not implemented yet.",
    }
    return _emit_result(payload, json_output=args.json_output)


def _emit_result(payload: dict, *, json_output: bool) -> int:
    if json_output:
        print(json.dumps(payload, indent=2, ensure_ascii=False))
    else:
        print(f"[{payload['status']}] {payload.get('message', payload['command'])}")
        if payload.get("command") == "voices":
            for voice in payload.get("voices", [])[:20]:
                print(
                    f"{voice.get('name')} | {voice.get('locale')} | "
                    f"{voice.get('gender')} | {voice.get('display_name')}"
                )
        if "output_dir" in payload:
            print(f"output: {payload['output_dir']}")
    return 0 if payload["status"] in {"ok", "todo"} else 1


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        handler = args.handler
        return handler(args)
    except FileNotFoundError as exc:
        print(f"[error] {exc}")
        return 3
    except ValueError as exc:
        print(f"[error] {exc}")
        return 2
    except PowerPointUnavailableError as exc:
        print(f"[error] {exc}")
        return 4
    except VoiceGenerationError as exc:
        print(f"[error] {exc}")
        return 5
    except SubtitleGenerationError as exc:
        print(f"[error] {exc}")
        return 6
    except RenderError as exc:
        print(f"[error] {exc}")
        return 7
    except Exception as exc:  # pragma: no cover - top-level safety net
        print(f"[error] Unexpected failure: {exc}")
        return 1
