from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any


def _as_dict(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def normalize_user_config(raw: dict[str, Any]) -> dict[str, Any]:
    """
    Normalize config to a provider-extensible structure.

    Current canonical layout:

    {
      "tts": {
        "default_provider": "edge" | "minimax" | ...,
        "providers": {
          "minimax": {"api_key": "...", "base_url": "...", "region": "cn|global", "model": "...", "timeout_s": 60},
          "edge": {...}
        }
      }
    }
    """
    cfg: dict[str, Any] = dict(raw or {})

    tts = dict(_as_dict(cfg.get("tts")))
    providers = dict(_as_dict(tts.get("providers")))
    providers = dict(providers)
    # Keep a stable nested structure for GUI and future providers.
    tts["providers"] = providers
    cfg["tts"] = tts

    gui = dict(_as_dict(cfg.get("gui")))
    cfg["gui"] = gui
    return cfg


def gui_state(cfg: dict[str, Any]) -> dict[str, Any]:
    cfg = normalize_user_config(cfg)
    return dict(_as_dict(cfg.get("gui")))


def tts_provider_config(cfg: dict[str, Any], provider: str) -> dict[str, Any]:
    cfg = normalize_user_config(cfg)
    tts = _as_dict(cfg.get("tts"))
    providers = _as_dict(tts.get("providers"))
    return dict(_as_dict(providers.get(provider.strip().lower())))


def default_tts_provider(cfg: dict[str, Any]) -> str | None:
    cfg = normalize_user_config(cfg)
    tts = _as_dict(cfg.get("tts"))
    raw = str(tts.get("default_provider") or "").strip().lower()
    return raw or None


def minimax_host_ui_index_from_provider_cfg(provider_cfg: dict[str, Any]) -> tuple[int, str]:
    """
    Map provider config to (settings-dialog API combo index, custom URL).

    Index mapping (same as GUI):
    0 env only, 1 region=cn, 2 region=global, 3 url chat, 4 url intl, 5 url io, 6 custom url
    """
    url = str(provider_cfg.get("base_url") or "").strip().rstrip("/")
    reg = str(provider_cfg.get("region") or "").strip().lower()
    if url == "https://api.minimax.chat":
        return 3, ""
    if url == "https://api.minimaxi.chat":
        return 4, ""
    if url == "https://api.minimax.io":
        return 5, ""
    if url:
        return 6, url
    if reg in ("cn", "china", "domestic", "zh", "国内"):
        return 1, ""
    if reg in ("global", "intl", "international", "int", "国际"):
        return 2, ""
    return 0, ""


def user_config_path() -> Path:
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or str(Path.home())
        return Path(base) / "note2video" / "config.json"
    return Path.home() / ".config" / "note2video" / "config.json"


def load_user_config() -> dict[str, Any]:
    path = user_config_path()
    if not path.is_file():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return {}
    return raw if isinstance(raw, dict) else {}


def save_user_config(data: dict[str, Any]) -> None:
    """Write the full config dict to disk and refresh the in-process TTS cache."""
    path = user_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    try:
        from note2video.tts.voice import invalidate_user_config_cache
    except ImportError:  # pragma: no cover — voice not loaded (e.g. partial test import)
        pass
    else:
        invalidate_user_config_cache()
