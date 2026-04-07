from __future__ import annotations

import argparse
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="note2video-tts-preview")
    parser.add_argument("--provider", required=True, help="TTS provider name.")
    parser.add_argument("--voice", default="", help="Voice ID.")
    parser.add_argument("--tts-rate", type=float, default=1.0, help="Speech rate multiplier (0.5–2.0).")
    parser.add_argument("--text", required=True, help="Preview text.")
    parser.add_argument("--out", required=True, help="Output WAV path.")
    args = parser.parse_args(argv)

    from note2video.tts.voice import synthesize_preview_sample

    out_path = Path(args.out)
    synthesize_preview_sample(
        provider_name=args.provider,
        voice_id=args.voice,
        tts_rate=args.tts_rate,
        text=args.text,
        output_file=out_path,
        minimax_base_url=None,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

