from __future__ import annotations

import json
from pathlib import Path

from note2video.build_profile import (
    BUILD_PROFILE_KIND,
    BUILD_PROFILE_VERSION,
    build_profile_to_request_kwargs,
    load_build_profile,
    request_kwargs_to_build_profile,
    save_build_profile,
)


def test_build_profile_roundtrip(tmp_path: Path) -> None:
    profile = request_kwargs_to_build_profile(
        {
            "input_file": "deck/demo.pptx",
            "out_dir": "dist/out",
            "pages": "1-3",
            "ratio": "9:16",
            "resolution": "720p",
            "fps": 24,
            "quality": "high",
            "tts_provider": "edge",
            "voice_id": "zh-CN-XiaoxiaoNeural",
            "tts_rate": 1.15,
            "script_file": "scripts/all.txt",
            "bgm_path": "audio/bgm.mp3",
            "bgm_volume": 0.22,
            "narration_volume": 0.95,
            "bgm_fade_in_s": 1.0,
            "bgm_fade_out_s": 2.0,
            "subtitle_color": "#FFFFFF",
            "subtitle_font": "Microsoft YaHei",
            "subtitle_size": 42,
            "subtitle_y_ratio": 0.88,
            "subtitle_fade_in_ms": 60,
            "subtitle_fade_out_ms": 140,
            "subtitle_scale_from": 96,
            "subtitle_scale_to": 108,
            "subtitle_outline": 2,
            "subtitle_shadow": 1,
        }
    )
    path = tmp_path / "demo.build.json"
    save_build_profile(path, profile)

    loaded = load_build_profile(path)
    kwargs = build_profile_to_request_kwargs(loaded, profile_path=path)

    assert loaded["kind"] == BUILD_PROFILE_KIND
    assert loaded["version"] == BUILD_PROFILE_VERSION
    assert loaded["video"]["resolution"] == "720p"
    assert loaded["video"]["fps"] == 24
    assert loaded["video"]["quality"] == "high"
    assert kwargs["input_file"] == str((tmp_path / "deck" / "demo.pptx").resolve())
    assert kwargs["out_dir"] == str((tmp_path / "dist" / "out").resolve())
    assert kwargs["script_file"] == str((tmp_path / "scripts" / "all.txt").resolve())
    assert kwargs["bgm_path"] == str((tmp_path / "audio" / "bgm.mp3").resolve())
    assert kwargs["fps"] == 24
    assert kwargs["quality"] == "high"
    assert kwargs["subtitle_size"] == 42
    assert abs(float(kwargs["subtitle_y_ratio"]) - 0.88) < 1e-9


def test_load_build_profile_accepts_flat_keys(tmp_path: Path) -> None:
    path = tmp_path / "flat.json"
    path.write_text(
        json.dumps(
            {
                "input_file": "deck.pptx",
                "out_dir": "dist",
                "ratio": "1:1",
                "resolution": "1440p",
                "fps": 60,
                "quality": "high",
                "tts_provider": "edge",
                "voice_id": "demo",
                "subtitle_color": "#ABCDEF",
            }
        ),
        encoding="utf-8",
    )
    loaded = load_build_profile(path)
    assert loaded["input"]["file"] == "deck.pptx"
    assert loaded["video"]["resolution"] == "1440p"
    assert loaded["video"]["fps"] == 60
    assert loaded["tts"]["voice_id"] == "demo"
