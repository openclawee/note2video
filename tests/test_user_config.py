from __future__ import annotations

import json

import pytest

from note2video import user_config
from note2video.tts.voice import invalidate_user_config_cache
from note2video.user_config import minimax_host_ui_index_from_provider_cfg, normalize_user_config, tts_provider_config


def test_user_config_roundtrip(monkeypatch, tmp_path) -> None:
    path = tmp_path / "config.json"
    monkeypatch.setattr(user_config, "user_config_path", lambda: path)

    user_config.save_user_config(
        {
            "tts": {
                "providers": {
                    "minimax": {"region": "cn", "model": "speech-2.8-hd"},
                }
            }
        }
    )
    assert path.is_file()
    loaded = user_config.load_user_config()
    norm = normalize_user_config(loaded)
    mm = tts_provider_config(norm, "minimax")
    assert mm["region"] == "cn"
    assert mm["model"] == "speech-2.8-hd"


def test_get_minimax_api_base_url_reads_config_file(monkeypatch, tmp_path) -> None:
    path = tmp_path / "config.json"
    monkeypatch.setattr(user_config, "user_config_path", lambda: path)
    invalidate_user_config_cache()
    path.write_text(
        json.dumps({"tts": {"providers": {"minimax_cn": {"api_key": "k", "model": "speech-2.8-hd"}}}}),
        encoding="utf-8",
    )
    invalidate_user_config_cache()
    monkeypatch.delenv("NOTE2VIDEO_MINIMAX_API_KEY", raising=False)
    monkeypatch.delenv("MINIMAX_API_KEY", raising=False)
    from note2video.tts.voice import _minimax_fixed_base_url, _provider_api_key

    assert _minimax_fixed_base_url("minimax_cn") == "https://api.minimax.chat"
    assert _provider_api_key("minimax_cn") == "k"


def test_minimax_host_ui_index_mapping() -> None:
    assert minimax_host_ui_index_from_provider_cfg({}) == (0, "")
    assert minimax_host_ui_index_from_provider_cfg({"region": "cn"}) == (1, "")
    assert minimax_host_ui_index_from_provider_cfg({"region": "global"}) == (2, "")
    assert minimax_host_ui_index_from_provider_cfg({"base_url": "https://api.minimax.chat"}) == (3, "")
    assert minimax_host_ui_index_from_provider_cfg({"base_url": "https://api.minimaxi.chat"}) == (4, "")
    assert minimax_host_ui_index_from_provider_cfg({"base_url": "https://api.minimax.io"}) == (5, "")
    assert minimax_host_ui_index_from_provider_cfg({"base_url": "https://x.example"}) == (6, "https://x.example")


def test_minimax_provider_maps_to_fixed_host(monkeypatch, tmp_path) -> None:
    path = tmp_path / "config.json"
    monkeypatch.setattr(user_config, "user_config_path", lambda: path)
    path.write_text(
        json.dumps({"tts": {"providers": {"minimax_cn": {"api_key": "k"}}}}),
        encoding="utf-8",
    )
    invalidate_user_config_cache()
    # Region selection is now done by choosing minimax_cn vs minimax_global, not by region env.
    from note2video.tts.voice import _minimax_fixed_base_url

    assert _minimax_fixed_base_url("minimax_cn") == "https://api.minimax.chat"
    assert _minimax_fixed_base_url("minimax_global") == "https://api.minimaxi.chat"
