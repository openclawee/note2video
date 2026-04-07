from __future__ import annotations

import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

# Allow `python scripts/minimax_ping.py` from repo root without a prior `pip install -e .`.
_repo_src = Path(__file__).resolve().parent.parent / "src"
if _repo_src.is_dir():
    sys.path.insert(0, str(_repo_src))

from note2video.tts.voice import get_minimax_api_base_url


def main() -> int:
    key = (os.getenv("MINIMAX_API_KEY") or os.getenv("NOTE2VIDEO_MINIMAX_API_KEY") or "").strip()
    if not key:
        print("MINIMAX_API_KEY (or NOTE2VIDEO_MINIMAX_API_KEY) not set.", file=sys.stderr)
        print('PowerShell example:  $env:MINIMAX_API_KEY="your_key_here"', file=sys.stderr)
        return 2

    masked = f"{key[:4]}...{key[-4:]}" if len(key) >= 8 else "***"
    print(f"key length: {len(key)}  masked: {masked}")
    if key.lower().startswith("bearer "):
        print("Warning: key includes 'Bearer ' prefix. Please set the raw API key only.", file=sys.stderr)

    text_override = os.getenv("MINIMAX_TEXT_BASE_URL", "").strip()
    base = text_override.rstrip("/") if text_override else get_minimax_api_base_url()
    model = (os.getenv("MINIMAX_TEXT_MODEL") or "MiniMax-M2.7").strip()

    # Try both:
    # - OpenAI-compatible: POST /v1/chat/completions
    # - Native:            POST /v1/text/chatcompletion_v2
    attempts: list[tuple[str, str, dict[str, Any], dict[str, str]]] = [
        (
            "openai_compatible",
            f"{base}/v1/chat/completions",
            {
                "model": model,
                "messages": [{"role": "user", "content": "回复我：token ok"}],
                "temperature": 0,
            },
            {
                "Authorization": f"Bearer {key}",
                "Content-Type": "application/json",
            },
        ),
        (
            "native_v2",
            f"{base}/v1/text/chatcompletion_v2",
            {
                "model": model,
                "messages": [
                    {"role": "system", "name": "MiniMax AI", "content": "You are a helpful assistant."},
                    {"role": "user", "name": "user", "content": "回复我：token ok"},
                ],
                "stream": False,
            },
            {
                "Authorization": f"Bearer {key}",
                "Content-Type": "application/json",
            },
        ),
        (
            "anthropic_compatible",
            f"{base}/anthropic/v1/messages",
            {
                "model": model,
                "max_tokens": 256,
                "messages": [
                    {
                        "role": "user",
                        "content": [{"type": "text", "text": "回复我：token ok"}],
                    }
                ],
            },
            {
                "x-api-key": key,
                "anthropic-version": "2023-06-01",
                "Content-Type": "application/json",
            },
        ),
    ]

    for name, url, body, extra_headers in attempts:
        print(f"\n== Attempt: {name} ==")
        print("url:", url)
        req = urllib.request.Request(
            url=url,
            data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
            headers=extra_headers,
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read()
                print(f"HTTP {resp.status}")
        except urllib.error.HTTPError as exc:
            try:
                detail = exc.read().decode("utf-8", errors="replace")
            except Exception:
                detail = ""
            print(f"HTTP {exc.code}", file=sys.stderr)
            if detail:
                print(detail, file=sys.stderr)
            continue
        except urllib.error.URLError as exc:
            print(f"Network error: {exc}", file=sys.stderr)
            continue

        try:
            payload = json.loads(raw.decode("utf-8", errors="replace"))
        except Exception:
            print(raw.decode("utf-8", errors="replace"))
            return 0

        # Print a compact view first (easy to eyeball), then the full JSON.
        content = None
        if isinstance(payload, dict):
            try:
                content = payload.get("choices", [{}])[0].get("message", {}).get("content")
            except Exception:
                content = None
        if content:
            print("content:", content)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    return 1


if __name__ == "__main__":
    raise SystemExit(main())

