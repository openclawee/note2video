from __future__ import annotations

import json
from pathlib import Path
from typing import Any

BUILD_PROFILE_KIND = "note2video.build_profile"
BUILD_PROFILE_VERSION = 1


def _as_dict(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def _clean_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _clean_optional_str(value: Any) -> str | None:
    text = _clean_str(value)
    return text or None


def _clean_optional_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if text.strip() else None


def _clean_int(value: Any, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return int(default)


def _clean_float(value: Any, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


def _clean_optional_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _resolve_relative_path(value: str | None, *, base_dir: Path | None) -> str | None:
    if value is None:
        return None
    raw = str(value).strip()
    if not raw:
        return None
    path = Path(raw)
    if path.is_absolute() or base_dir is None:
        return str(path)
    return str((base_dir / path).resolve())


def default_build_request_kwargs() -> dict[str, Any]:
    return {
        "input_file": "",
        "out_dir": "./dist",
        "pages": "all",
        "ratio": "16:9",
        "resolution": "1080p",
        "fps": 30,
        "quality": "standard",
        "tts_provider": "pyttsx3",
        "voice_id": "",
        "tts_rate": 1.0,
        "script_file": None,
        "script_text": None,
        "bgm_path": None,
        "bgm_volume": 0.18,
        "narration_volume": 1.0,
        "bgm_fade_in_s": 0.0,
        "bgm_fade_out_s": 0.0,
        "subtitle_color": None,
        "subtitle_font": None,
        "subtitle_size": None,
        "subtitle_y_ratio": None,
        "subtitle_fade_in_ms": 80,
        "subtitle_fade_out_ms": 120,
        "subtitle_scale_from": 100,
        "subtitle_scale_to": 104,
        "subtitle_outline": 1,
        "subtitle_shadow": 0,
        "avatar_video": None,
        "avatar_pos": "bl",
        "avatar_scale": 0.25,
        "avatar_key": "auto",
        "avatar_key_color": "#00ff00",
        "avatar_key_similarity": 0.15,
        "avatar_key_blend": 0.02,
        "avatar_x_ratio": None,
        "avatar_y_ratio": None,
    }


def normalize_build_profile(raw: dict[str, Any]) -> dict[str, Any]:
    data = dict(raw or {})
    input_cfg = dict(_as_dict(data.get("input")))
    output_cfg = dict(_as_dict(data.get("output")))
    video_cfg = dict(_as_dict(data.get("video")))
    tts_cfg = dict(_as_dict(data.get("tts")))
    audio_cfg = dict(_as_dict(data.get("audio")))
    subtitle_cfg = dict(_as_dict(data.get("subtitle")))
    subtitle_effects = dict(_as_dict(subtitle_cfg.get("effects")))
    avatar_cfg = dict(_as_dict(data.get("avatar")))

    # Accept flat/manual-edited keys as a convenience.
    if "input_file" in data and "file" not in input_cfg:
        input_cfg["file"] = data.get("input_file")
    if "out_dir" in data and "dir" not in output_cfg:
        output_cfg["dir"] = data.get("out_dir")
    if "pages" in data and "pages" not in input_cfg:
        input_cfg["pages"] = data.get("pages")
    if "script_file" in data and "script_file" not in input_cfg:
        input_cfg["script_file"] = data.get("script_file")
    if "script_text" in data and "script_text" not in input_cfg:
        input_cfg["script_text"] = data.get("script_text")
    if "ratio" in data and "ratio" not in video_cfg:
        video_cfg["ratio"] = data.get("ratio")
    if "resolution" in data and "resolution" not in video_cfg:
        video_cfg["resolution"] = data.get("resolution")
    if "fps" in data and "fps" not in video_cfg:
        video_cfg["fps"] = data.get("fps")
    if "quality" in data and "quality" not in video_cfg:
        video_cfg["quality"] = data.get("quality")
    if "tts_provider" in data and "provider" not in tts_cfg:
        tts_cfg["provider"] = data.get("tts_provider")
    if "voice_id" in data and "voice_id" not in tts_cfg:
        tts_cfg["voice_id"] = data.get("voice_id")
    if "tts_rate" in data and "rate" not in tts_cfg:
        tts_cfg["rate"] = data.get("tts_rate")
    if "bgm_path" in data and "bgm_path" not in audio_cfg:
        audio_cfg["bgm_path"] = data.get("bgm_path")
    if "bgm_volume" in data and "bgm_volume" not in audio_cfg:
        audio_cfg["bgm_volume"] = data.get("bgm_volume")
    if "narration_volume" in data and "narration_volume" not in audio_cfg:
        audio_cfg["narration_volume"] = data.get("narration_volume")
    if "bgm_fade_in_s" in data and "bgm_fade_in_s" not in audio_cfg:
        audio_cfg["bgm_fade_in_s"] = data.get("bgm_fade_in_s")
    if "bgm_fade_out_s" in data and "bgm_fade_out_s" not in audio_cfg:
        audio_cfg["bgm_fade_out_s"] = data.get("bgm_fade_out_s")
    if "subtitle_color" in data and "color" not in subtitle_cfg:
        subtitle_cfg["color"] = data.get("subtitle_color")
    if "subtitle_font" in data and "font" not in subtitle_cfg:
        subtitle_cfg["font"] = data.get("subtitle_font")
    if "subtitle_size" in data and "size" not in subtitle_cfg:
        subtitle_cfg["size"] = data.get("subtitle_size")
    if "subtitle_y_ratio" in data and "y_ratio" not in subtitle_cfg:
        subtitle_cfg["y_ratio"] = data.get("subtitle_y_ratio")
    if "subtitle_fade_in_ms" in data and "fade_in_ms" not in subtitle_effects:
        subtitle_effects["fade_in_ms"] = data.get("subtitle_fade_in_ms")
    if "subtitle_fade_out_ms" in data and "fade_out_ms" not in subtitle_effects:
        subtitle_effects["fade_out_ms"] = data.get("subtitle_fade_out_ms")
    if "subtitle_scale_from" in data and "scale_from" not in subtitle_effects:
        subtitle_effects["scale_from"] = data.get("subtitle_scale_from")
    if "subtitle_scale_to" in data and "scale_to" not in subtitle_effects:
        subtitle_effects["scale_to"] = data.get("subtitle_scale_to")
    if "subtitle_outline" in data and "outline" not in subtitle_effects:
        subtitle_effects["outline"] = data.get("subtitle_outline")
    if "subtitle_shadow" in data and "shadow" not in subtitle_effects:
        subtitle_effects["shadow"] = data.get("subtitle_shadow")
    if "avatar_video" in data and "video" not in avatar_cfg:
        avatar_cfg["video"] = data.get("avatar_video")
    if "avatar_pos" in data and "pos" not in avatar_cfg:
        avatar_cfg["pos"] = data.get("avatar_pos")
    if "avatar_scale" in data and "scale" not in avatar_cfg:
        avatar_cfg["scale"] = data.get("avatar_scale")
    if "avatar_key" in data and "key" not in avatar_cfg:
        avatar_cfg["key"] = data.get("avatar_key")
    if "avatar_key_color" in data and "key_color" not in avatar_cfg:
        avatar_cfg["key_color"] = data.get("avatar_key_color")
    if "avatar_key_similarity" in data and "key_similarity" not in avatar_cfg:
        avatar_cfg["key_similarity"] = data.get("avatar_key_similarity")
    if "avatar_key_blend" in data and "key_blend" not in avatar_cfg:
        avatar_cfg["key_blend"] = data.get("avatar_key_blend")
    if "avatar_x_ratio" in data and "x_ratio" not in avatar_cfg:
        avatar_cfg["x_ratio"] = data.get("avatar_x_ratio")
    if "avatar_y_ratio" in data and "y_ratio" not in avatar_cfg:
        avatar_cfg["y_ratio"] = data.get("avatar_y_ratio")

    return {
        "kind": BUILD_PROFILE_KIND,
        "version": int(data.get("version") or BUILD_PROFILE_VERSION),
        "input": {
            "file": _clean_str(input_cfg.get("file")),
            "pages": _clean_str(input_cfg.get("pages")) or "all",
            "script_file": _clean_str(input_cfg.get("script_file")),
            "script_text": str(input_cfg.get("script_text") or ""),
        },
        "output": {
            "dir": _clean_str(output_cfg.get("dir")) or "./dist",
        },
        "video": {
            "ratio": _clean_str(video_cfg.get("ratio")) or "16:9",
            "resolution": _clean_str(video_cfg.get("resolution")) or "1080p",
            "fps": _clean_int(video_cfg.get("fps"), 30),
            "quality": _clean_str(video_cfg.get("quality")) or "standard",
        },
        "tts": {
            "provider": _clean_str(tts_cfg.get("provider")) or "pyttsx3",
            "voice_id": _clean_str(tts_cfg.get("voice_id")),
            "rate": _clean_float(tts_cfg.get("rate"), 1.0),
        },
        "audio": {
            "bgm_path": _clean_str(audio_cfg.get("bgm_path")),
            "bgm_volume": _clean_float(audio_cfg.get("bgm_volume"), 0.18),
            "narration_volume": _clean_float(audio_cfg.get("narration_volume"), 1.0),
            "bgm_fade_in_s": _clean_float(audio_cfg.get("bgm_fade_in_s"), 0.0),
            "bgm_fade_out_s": _clean_float(audio_cfg.get("bgm_fade_out_s"), 0.0),
        },
        "subtitle": {
            "color": _clean_str(subtitle_cfg.get("color")),
            "font": _clean_str(subtitle_cfg.get("font")),
            "size": _clean_int(subtitle_cfg.get("size"), 0),
            "y_ratio": _clean_optional_float(subtitle_cfg.get("y_ratio")),
            "effects": {
                "fade_in_ms": _clean_int(subtitle_effects.get("fade_in_ms"), 80),
                "fade_out_ms": _clean_int(subtitle_effects.get("fade_out_ms"), 120),
                "scale_from": _clean_int(subtitle_effects.get("scale_from"), 100),
                "scale_to": _clean_int(subtitle_effects.get("scale_to"), 104),
                "outline": _clean_int(subtitle_effects.get("outline"), 1),
                "shadow": _clean_int(subtitle_effects.get("shadow"), 0),
            },
        },
        "avatar": {
            "video": _clean_str(avatar_cfg.get("video")),
            "pos": _clean_str(avatar_cfg.get("pos")) or "bl",
            "scale": _clean_float(avatar_cfg.get("scale"), 0.25),
            "key": _clean_str(avatar_cfg.get("key")) or "auto",
            "key_color": _clean_str(avatar_cfg.get("key_color")) or "#00ff00",
            "key_similarity": _clean_float(avatar_cfg.get("key_similarity"), 0.15),
            "key_blend": _clean_float(avatar_cfg.get("key_blend"), 0.02),
            "x_ratio": _clean_optional_float(avatar_cfg.get("x_ratio")),
            "y_ratio": _clean_optional_float(avatar_cfg.get("y_ratio")),
        },
    }


def load_build_profile(path: str | Path) -> dict[str, Any]:
    profile_path = Path(path)
    payload = json.loads(profile_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("Build profile must be a JSON object.")
    return normalize_build_profile(payload)


def save_build_profile(path: str | Path, profile: dict[str, Any]) -> None:
    profile_path = Path(path)
    profile_path.parent.mkdir(parents=True, exist_ok=True)
    normalized = normalize_build_profile(profile)
    profile_path.write_text(json.dumps(normalized, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def build_profile_to_request_kwargs(
    profile: dict[str, Any],
    *,
    profile_path: str | Path | None = None,
) -> dict[str, Any]:
    normalized = normalize_build_profile(profile)
    base_dir = Path(profile_path).resolve().parent if profile_path is not None else None

    input_cfg = normalized["input"]
    output_cfg = normalized["output"]
    video_cfg = normalized["video"]
    tts_cfg = normalized["tts"]
    audio_cfg = normalized["audio"]
    subtitle_cfg = normalized["subtitle"]
    effects_cfg = subtitle_cfg["effects"]
    avatar_cfg = _as_dict(normalized.get("avatar"))

    size = int(subtitle_cfg["size"])

    return {
        "input_file": _resolve_relative_path(_clean_optional_str(input_cfg.get("file")), base_dir=base_dir) or "",
        "out_dir": _resolve_relative_path(_clean_optional_str(output_cfg.get("dir")), base_dir=base_dir) or "./dist",
        "pages": _clean_str(input_cfg.get("pages")) or "all",
        "ratio": _clean_str(video_cfg.get("ratio")) or "16:9",
        "resolution": _clean_str(video_cfg.get("resolution")) or "1080p",
        "fps": _clean_int(video_cfg.get("fps"), 30),
        "quality": _clean_str(video_cfg.get("quality")) or "standard",
        "tts_provider": _clean_str(tts_cfg.get("provider")) or "pyttsx3",
        "voice_id": _clean_str(tts_cfg.get("voice_id")),
        "tts_rate": _clean_float(tts_cfg.get("rate"), 1.0),
        "script_file": _resolve_relative_path(_clean_optional_str(input_cfg.get("script_file")), base_dir=base_dir),
        "script_text": _clean_optional_text(input_cfg.get("script_text")),
        "bgm_path": _resolve_relative_path(_clean_optional_str(audio_cfg.get("bgm_path")), base_dir=base_dir),
        "bgm_volume": _clean_float(audio_cfg.get("bgm_volume"), 0.18),
        "narration_volume": _clean_float(audio_cfg.get("narration_volume"), 1.0),
        "bgm_fade_in_s": _clean_float(audio_cfg.get("bgm_fade_in_s"), 0.0),
        "bgm_fade_out_s": _clean_float(audio_cfg.get("bgm_fade_out_s"), 0.0),
        "subtitle_color": _clean_optional_str(subtitle_cfg.get("color")),
        "subtitle_font": _clean_optional_str(subtitle_cfg.get("font")),
        "subtitle_size": size if size > 0 else None,
        "subtitle_y_ratio": _clean_optional_float(subtitle_cfg.get("y_ratio")),
        "subtitle_fade_in_ms": _clean_int(effects_cfg.get("fade_in_ms"), 80),
        "subtitle_fade_out_ms": _clean_int(effects_cfg.get("fade_out_ms"), 120),
        "subtitle_scale_from": _clean_int(effects_cfg.get("scale_from"), 100),
        "subtitle_scale_to": _clean_int(effects_cfg.get("scale_to"), 104),
        "subtitle_outline": _clean_int(effects_cfg.get("outline"), 1),
        "subtitle_shadow": _clean_int(effects_cfg.get("shadow"), 0),
        "avatar_video": _resolve_relative_path(_clean_optional_str(avatar_cfg.get("video")), base_dir=base_dir),
        "avatar_pos": _clean_str(avatar_cfg.get("pos")) or "bl",
        "avatar_scale": _clean_float(avatar_cfg.get("scale"), 0.25),
        "avatar_key": _clean_str(avatar_cfg.get("key")) or "auto",
        "avatar_key_color": _clean_str(avatar_cfg.get("key_color")) or "#00ff00",
        "avatar_key_similarity": _clean_float(avatar_cfg.get("key_similarity"), 0.15),
        "avatar_key_blend": _clean_float(avatar_cfg.get("key_blend"), 0.02),
        "avatar_x_ratio": _clean_optional_float(avatar_cfg.get("x_ratio")),
        "avatar_y_ratio": _clean_optional_float(avatar_cfg.get("y_ratio")),
    }


def request_kwargs_to_build_profile(values: dict[str, Any]) -> dict[str, Any]:
    merged = default_build_request_kwargs()
    merged.update(dict(values or {}))

    subtitle_size = merged.get("subtitle_size")
    subtitle_y_ratio = merged.get("subtitle_y_ratio")
    avatar_x_ratio = merged.get("avatar_x_ratio")
    avatar_y_ratio = merged.get("avatar_y_ratio")

    return normalize_build_profile(
        {
            "kind": BUILD_PROFILE_KIND,
            "version": BUILD_PROFILE_VERSION,
            "input": {
                "file": merged.get("input_file") or "",
                "pages": merged.get("pages") or "all",
                "script_file": merged.get("script_file") or "",
                "script_text": merged.get("script_text") or "",
            },
            "output": {
                "dir": merged.get("out_dir") or "./dist",
            },
            "video": {
                "ratio": merged.get("ratio") or "16:9",
                "resolution": merged.get("resolution") or "1080p",
                "fps": int(merged.get("fps") or 30),
                "quality": merged.get("quality") or "standard",
            },
            "tts": {
                "provider": merged.get("tts_provider") or "pyttsx3",
                "voice_id": merged.get("voice_id") or "",
                "rate": float(merged.get("tts_rate") or 1.0),
            },
            "audio": {
                "bgm_path": merged.get("bgm_path") or "",
                "bgm_volume": float(merged.get("bgm_volume") or 0.18),
                "narration_volume": float(merged.get("narration_volume") or 1.0),
                "bgm_fade_in_s": float(merged.get("bgm_fade_in_s") or 0.0),
                "bgm_fade_out_s": float(merged.get("bgm_fade_out_s") or 0.0),
            },
            "subtitle": {
                "color": merged.get("subtitle_color") or "",
                "font": merged.get("subtitle_font") or "",
                "size": int(subtitle_size or 0),
                "y_ratio": float(subtitle_y_ratio) if subtitle_y_ratio is not None else None,
                "effects": {
                    "fade_in_ms": int(merged.get("subtitle_fade_in_ms") or 80),
                    "fade_out_ms": int(merged.get("subtitle_fade_out_ms") or 120),
                    "scale_from": int(merged.get("subtitle_scale_from") or 100),
                    "scale_to": int(merged.get("subtitle_scale_to") or 104),
                    "outline": int(merged.get("subtitle_outline") or 1),
                    "shadow": int(merged.get("subtitle_shadow") or 0),
                },
            },
            "avatar": {
                "video": merged.get("avatar_video") or "",
                "pos": merged.get("avatar_pos") or "bl",
                "scale": float(merged.get("avatar_scale") or 0.25),
                "key": merged.get("avatar_key") or "auto",
                "key_color": merged.get("avatar_key_color") or "#00ff00",
                "key_similarity": float(merged.get("avatar_key_similarity") or 0.15),
                "key_blend": float(merged.get("avatar_key_blend") or 0.02),
                "x_ratio": float(avatar_x_ratio) if avatar_x_ratio is not None else None,
                "y_ratio": float(avatar_y_ratio) if avatar_y_ratio is not None else None,
            },
        }
    )


__all__ = [
    "BUILD_PROFILE_KIND",
    "BUILD_PROFILE_VERSION",
    "build_profile_to_request_kwargs",
    "default_build_request_kwargs",
    "load_build_profile",
    "normalize_build_profile",
    "request_kwargs_to_build_profile",
    "save_build_profile",
]
