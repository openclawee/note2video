from __future__ import annotations

import argparse
import json
from pathlib import Path

from note2video import __version__
from note2video.app.pipeline_service import (
    BuildRequest,
    ExtractRequest,
    RenderRequest,
    SubtitleRequest,
    VoiceRequest,
    VoicesRequest,
    run_build_pipeline,
    run_extract_pipeline,
    run_render_pipeline,
    run_subtitle_pipeline,
    run_voice_pipeline,
    run_voices_pipeline,
)
from note2video.parser.extract import PowerPointUnavailableError, extract_project
from note2video.render.video import RenderError, render_video
from note2video.subtitle.generate import SubtitleGenerationError, generate_subtitles
from note2video.tts.voice import (
    VoiceGenerationError,
    generate_voice_assets,
    list_available_voices,
)
from note2video.compose.pptx import ComposeError, compose_pptx_from_template


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
    build_cmd.add_argument(
        "--script-file",
        default="",
        help="Optional script file to override PPT notes. Supports script.json or scripts/all.txt format.",
    )
    build_cmd.add_argument(
        "--script-text",
        default="",
        help="Optional inline script text to override PPT notes (use quotes).",
    )
    build_cmd.add_argument("--voice", default="", help="Voice ID placeholder.")
    build_cmd.add_argument("--tts-provider", default="", help="TTS provider placeholder.")
    build_cmd.add_argument(
        "--tts-rate",
        type=float,
        default=1.0,
        help="Speech rate multiplier (0.5–2.0); applied during TTS so subtitles stay aligned.",
    )
    build_cmd.add_argument("--bgm", default="", help="Optional background music file to mix into the final video.")
    build_cmd.add_argument("--bgm-volume", type=float, default=0.18, help="Background music volume (default 0.18).")
    build_cmd.add_argument("--bgm-fade-in", type=float, default=0.0, help="BGM fade-in seconds (default 0).")
    build_cmd.add_argument("--bgm-fade-out", type=float, default=0.0, help="BGM fade-out seconds (default 0).")
    build_cmd.add_argument(
        "--narration-volume",
        type=float,
        default=1.0,
        help="Narration volume before mixing (default 1.0).",
    )
    build_cmd.add_argument(
        "--subtitle-color",
        default="",
        help="Subtitle color when burning into video (hex #RRGGBB). Example: #FFFFFF",
    )
    build_cmd.add_argument(
        "--subtitle-font",
        default="",
        help="Subtitle font family name when burning into video (e.g. Microsoft YaHei).",
    )
    build_cmd.add_argument(
        "--subtitle-size",
        type=int,
        default=0,
        help="Subtitle font size in pixels when burning into video (0 = default).",
    )
    build_cmd.add_argument(
        "--subtitle-fade-in-ms",
        type=int,
        default=80,
        help="ASS fade-in duration per sentence (ms).",
    )
    build_cmd.add_argument(
        "--subtitle-fade-out-ms",
        type=int,
        default=120,
        help="ASS fade-out duration per sentence (ms).",
    )
    build_cmd.add_argument(
        "--subtitle-scale-from",
        type=int,
        default=100,
        help="ASS scale at sentence start (percent).",
    )
    build_cmd.add_argument(
        "--subtitle-scale-to",
        type=int,
        default=104,
        help="ASS scale after short ease-in (percent).",
    )
    build_cmd.add_argument(
        "--subtitle-outline",
        type=int,
        default=1,
        help="ASS outline thickness (0+).",
    )
    build_cmd.add_argument(
        "--subtitle-shadow",
        type=int,
        default=0,
        help="ASS shadow depth (0+).",
    )
    build_cmd.add_argument(
        "--subtitle-y-ratio",
        type=float,
        default=None,
        help="Optional subtitle vertical position ratio (0-1). Uses ASS \\pos() with horizontal center.",
    )
    # MiniMax CN/Global are separate providers; host selection is implied by --tts-provider.
    build_cmd.set_defaults(handler=handle_build)

    extract_cmd = subparsers.add_parser("extract", help="Extract slide assets and notes.")
    _add_common_arguments(extract_cmd)
    extract_cmd.add_argument("input", help="Path to the input .pptx or .pdf file.")
    extract_cmd.add_argument("--pages", default="all", help="Page selection, e.g. 1-3,5.")
    extract_cmd.set_defaults(handler=handle_extract)

    compose_cmd = subparsers.add_parser("compose", help="Generate a .pptx from a one-slide template + params.json.")
    compose_cmd.add_argument("template", help="Path to the template .pptx (single slide).")
    compose_cmd.add_argument("params", help="Path to params.json containing a 'pages' array.")
    compose_cmd.add_argument("--out", default="./dist/deck.pptx", help="Output .pptx path.")
    compose_cmd.add_argument(
        "--assets-base-dir",
        default="",
        help="Optional base directory for resolving relative image paths in params.json.",
    )
    compose_cmd.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="Print a machine-readable JSON summary.",
    )
    compose_cmd.set_defaults(handler=handle_compose)

    voice_cmd = subparsers.add_parser("voice", help="Generate voice-over audio.")
    _add_common_arguments(voice_cmd)
    voice_cmd.add_argument("input", help="Path to notes or script JSON.")
    voice_cmd.add_argument("--tts-provider", default="pyttsx3", help="TTS provider name.")
    voice_cmd.add_argument("--voice", default="", help="Voice ID.")
    voice_cmd.add_argument(
        "--tts-rate",
        type=float,
        default=1.0,
        help="Speech rate multiplier (0.5–2.0); applied during TTS so subtitles stay aligned.",
    )
    # MiniMax CN/Global are separate providers; host selection is implied by --tts-provider.
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
    render_cmd.add_argument("--ratio", default="16:9", help="Output ratio (16:9, 9:16, 1:1).")
    render_cmd.add_argument("--bgm", default="", help="Optional background music file to mix into the final video.")
    render_cmd.add_argument("--bgm-volume", type=float, default=0.18, help="Background music volume (default 0.18).")
    render_cmd.add_argument("--bgm-fade-in", type=float, default=0.0, help="BGM fade-in seconds (default 0).")
    render_cmd.add_argument("--bgm-fade-out", type=float, default=0.0, help="BGM fade-out seconds (default 0).")
    render_cmd.add_argument(
        "--narration-volume",
        type=float,
        default=1.0,
        help="Narration volume before mixing (default 1.0).",
    )
    render_cmd.add_argument(
        "--subtitle-color",
        default="",
        help="Subtitle color when burning into video (hex #RRGGBB). Example: #FFFFFF",
    )
    render_cmd.add_argument(
        "--subtitle-font",
        default="",
        help="Subtitle font family name when burning into video (e.g. Microsoft YaHei).",
    )
    render_cmd.add_argument(
        "--subtitle-size",
        type=int,
        default=0,
        help="Subtitle font size in pixels when burning into video (0 = default).",
    )
    render_cmd.add_argument(
        "--subtitle-fade-in-ms",
        type=int,
        default=80,
        help="ASS fade-in duration per sentence (ms).",
    )
    render_cmd.add_argument(
        "--subtitle-fade-out-ms",
        type=int,
        default=120,
        help="ASS fade-out duration per sentence (ms).",
    )
    render_cmd.add_argument(
        "--subtitle-scale-from",
        type=int,
        default=100,
        help="ASS scale at sentence start (percent).",
    )
    render_cmd.add_argument(
        "--subtitle-scale-to",
        type=int,
        default=104,
        help="ASS scale after short ease-in (percent).",
    )
    render_cmd.add_argument(
        "--subtitle-outline",
        type=int,
        default=1,
        help="ASS outline thickness (0+).",
    )
    render_cmd.add_argument(
        "--subtitle-shadow",
        type=int,
        default=0,
        help="ASS shadow depth (0+).",
    )
    render_cmd.add_argument(
        "--subtitle-y-ratio",
        type=float,
        default=None,
        help="Optional subtitle vertical position ratio (0-1). Uses ASS \\pos() with horizontal center.",
    )
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
    result = run_build_pipeline(
        BuildRequest(
            input_file=args.input,
            out_dir=args.out,
            pages=args.pages,
            ratio=args.ratio,
            tts_provider=(args.tts_provider or "pyttsx3"),
            voice_id=args.voice,
            tts_rate=float(args.tts_rate),
            script_file=(args.script_file.strip() or None),
            script_text=(args.script_text if str(args.script_text or "").strip() else None),
            bgm_path=(args.bgm.strip() or None),
            bgm_volume=float(args.bgm_volume),
            bgm_fade_in_s=float(args.bgm_fade_in),
            bgm_fade_out_s=float(args.bgm_fade_out),
            narration_volume=float(args.narration_volume),
            subtitle_color=(args.subtitle_color.strip() or None),
            subtitle_fade_in_ms=int(args.subtitle_fade_in_ms),
            subtitle_fade_out_ms=int(args.subtitle_fade_out_ms),
            subtitle_scale_from=int(args.subtitle_scale_from),
            subtitle_scale_to=int(args.subtitle_scale_to),
            subtitle_outline=int(args.subtitle_outline),
            subtitle_shadow=int(args.subtitle_shadow),
            subtitle_font=(args.subtitle_font.strip() or None),
            subtitle_size=(int(args.subtitle_size) if int(args.subtitle_size or 0) > 0 else None),
            subtitle_y_ratio=(float(args.subtitle_y_ratio) if args.subtitle_y_ratio is not None else None),
        ),
        extract_project_fn=extract_project,
        generate_voice_assets_fn=generate_voice_assets,
        generate_subtitles_fn=generate_subtitles,
        render_video_fn=render_video,
    )
    out_dir = Path(result["output_dir"])

    payload = {
        "command": "build",
        "status": "ok",
        "phase": "complete",
        "input": args.input,
        "output_dir": str(out_dir),
        "artifacts": result["artifacts"],
        "slide_count": result["slide_count"],
        "segment_count": result["segment_count"],
        "voice_provider": result["voice_provider"],
        "subtitles_burned": result["subtitles_burned"],
        "message": "Full pipeline completed.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_extract(args: argparse.Namespace) -> int:
    result = run_extract_pipeline(
        ExtractRequest(input_file=args.input, out_dir=args.out, pages=args.pages),
        extract_project_fn=extract_project,
    )
    payload = {
        "command": "extract",
        "status": "ok",
        "phase": "implemented",
        "input": args.input,
        "output_dir": str(Path(result["output_dir"])),
        "artifacts": result["artifacts"],
        "slide_count": result["slide_count"],
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_compose(args: argparse.Namespace) -> int:
    stats = compose_pptx_from_template(
        template_pptx=args.template,
        params_json=args.params,
        output_pptx=args.out,
        assets_base_dir=(args.assets_base_dir.strip() or None),
    )
    payload = {
        "command": "compose",
        "status": "ok",
        "template": args.template,
        "params": args.params,
        "output_pptx": args.out,
        "slide_count": int(stats.slide_count),
        "applied_text_fields": int(stats.applied_text_fields),
        "ignored_text_fields": int(stats.ignored_text_fields),
        "applied_images": int(stats.applied_images),
        "ignored_images": int(stats.ignored_images),
        "applied_notes": int(stats.applied_notes),
        "message": "PPTX composed.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_voice(args: argparse.Namespace) -> int:
    result = run_voice_pipeline(
        VoiceRequest(
            input_json=args.input,
            out_dir=args.out,
            tts_provider=args.tts_provider,
            voice_id=args.voice,
            tts_rate=float(args.tts_rate),
        ),
        generate_voice_assets_fn=generate_voice_assets,
    )
    payload = {
        "command": "voice",
        "status": "ok",
        "input": args.input,
        "output_dir": str(Path(args.out)),
        "slide_count": result["slide_count"],
        "provider": result["provider"],
        "voice": result["voice"],
        "tts_rate": result["tts_rate"],
        "artifacts": {
            "audio_dir": "audio",
            "merged_audio": "audio/merged.wav",
            "timings": "audio/timings.json",
        },
        "message": "Voice assets generated.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_voices(args: argparse.Namespace) -> int:
    result = run_voices_pipeline(
        VoicesRequest(tts_provider=args.tts_provider, keyword=args.keyword),
        list_available_voices_fn=list_available_voices,
    )
    payload = {
        "command": "voices",
        "status": "ok",
        "provider": args.tts_provider,
        "count": result["count"],
        "voices": result["voices"],
        "message": f"Listed {result['count']} voices.",
    }
    return _emit_result(payload, json_output=args.json_output)


def handle_subtitle(args: argparse.Namespace) -> int:
    result = run_subtitle_pipeline(
        SubtitleRequest(input_json=args.input, out_dir=args.out),
        generate_subtitles_fn=generate_subtitles,
    )
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
    result = run_render_pipeline(
        RenderRequest(
            project_dir=args.input,
            output_path=(args.out if args.out != "./dist" else None),
            ratio=args.ratio,
            bgm_path=(args.bgm.strip() or None),
            bgm_volume=float(args.bgm_volume),
            bgm_fade_in_s=float(args.bgm_fade_in),
            bgm_fade_out_s=float(args.bgm_fade_out),
            narration_volume=float(args.narration_volume),
            subtitle_color=(args.subtitle_color.strip() or None),
            subtitle_fade_in_ms=int(args.subtitle_fade_in_ms),
            subtitle_fade_out_ms=int(args.subtitle_fade_out_ms),
            subtitle_scale_from=int(args.subtitle_scale_from),
            subtitle_scale_to=int(args.subtitle_scale_to),
            subtitle_outline=int(args.subtitle_outline),
            subtitle_shadow=int(args.subtitle_shadow),
            subtitle_font=(args.subtitle_font.strip() or None),
            subtitle_size=(int(args.subtitle_size) if int(args.subtitle_size or 0) > 0 else None),
            subtitle_y_ratio=(float(args.subtitle_y_ratio) if args.subtitle_y_ratio is not None else None),
        ),
        render_video_fn=render_video,
    )
    payload = {
        "command": "render",
        "status": "ok",
        "input": args.input,
        "output_dir": str(Path(args.input)),
        "slide_count": result["slide_count"],
        "artifacts": {
            "video": result["artifacts"]["video"],
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
        try:
            print(json.dumps(payload, indent=2, ensure_ascii=False))
        except UnicodeEncodeError:
            # Some Windows consoles run in legacy encodings (e.g. GBK) and may fail
            # to print certain characters from upstream providers. Fall back to ASCII-escaped JSON.
            print(json.dumps(payload, indent=2, ensure_ascii=True))
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
    except ComposeError as exc:
        print(f"[error] {exc}")
        return 8
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


if __name__ == "__main__":
    raise SystemExit(main())
