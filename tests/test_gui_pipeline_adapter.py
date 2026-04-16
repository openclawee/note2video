from __future__ import annotations

import json
from pathlib import Path

from note2video.app.pipeline_service import BuildRequest
from note2video.gui.app import (
    JobConfig,
    _build_cli_argv_for_config,
    _build_request_from_job_config,
    _job_config_from_build_profile,
    _job_config_to_build_profile,
    _run_extract_or_build,
)


def _make_job_config(tmp_path: Path, mode: str = "build") -> JobConfig:
    return JobConfig(
        mode=mode,
        pptx_path=tmp_path / "demo.pptx",
        out_dir=tmp_path / "dist",
        pages="all",
        ratio="16:9",
        tts_provider="edge",
        voice_id="",
        tts_rate=1.0,
    )


def test_run_extract_mode_delegates_to_extract_service(monkeypatch, tmp_path) -> None:
    logs: list[str] = []
    calls: list[tuple[str, object]] = []

    def fake_extract(request):
        calls.append(("extract", request))
        return {"slide_count": 4}

    def fake_build(_request):
        calls.append(("build", _request))
        return {}

    monkeypatch.setattr("note2video.gui.app.run_extract_pipeline", fake_extract)
    monkeypatch.setattr("note2video.gui.app.run_build_pipeline", fake_build)

    exit_code = _run_extract_or_build(_make_job_config(tmp_path, mode="extract"), logs.append)

    assert exit_code == 0
    assert len(calls) == 1
    assert calls[0][0] == "extract"
    request = calls[0][1]
    assert request.input_file.endswith("demo.pptx")
    assert request.pages == "all"
    assert "细节：slides=4" in logs
    assert "完成：extract" in logs


def test_run_build_mode_delegates_to_build_service(monkeypatch, tmp_path) -> None:
    logs: list[str] = []
    calls: list[tuple[str, object]] = []

    def fake_extract(request):
        calls.append(("extract", request))
        return {"slide_count": 5}

    def fake_build(request):
        calls.append(("build", request))
        return {
            "voice_provider": "pyttsx3",
            "segment_count": 12,
            "slide_count": 5,
            "subtitles_burned": True,
            "mixed_audio": "audio/mixed.wav",
            "artifacts": {"video": "video/output.mp4"},
        }

    monkeypatch.setattr("note2video.gui.app.run_extract_pipeline", fake_extract)
    monkeypatch.setattr("note2video.gui.app.run_build_pipeline", fake_build)

    exit_code = _run_extract_or_build(_make_job_config(tmp_path, mode="build"), logs.append)

    assert exit_code == 0
    assert [c[0] for c in calls] == ["extract", "build"]
    build_request = calls[1][1]
    assert build_request.tts_provider == "edge"
    assert build_request.tts_rate == 1.0
    assert any("voice" in s for s in logs)
    assert any("tts_rate=1.0" in s and "voice=" in s for s in logs)
    assert "细节：segments=12, slides=5" in logs
    assert "细节：subtitles_burned=True, mixed_audio=True" in logs
    assert "输出视频：video/output.mp4" in logs


def test_build_request_mapping_from_job_config(tmp_path) -> None:
    cfg = JobConfig(
        mode="build",
        pptx_path=tmp_path / "demo.pptx",
        out_dir=tmp_path / "dist",
        pages="1-3",
        ratio="9:16",
        resolution="1440p",
        fps=60,
        quality="high",
        tts_provider="edge",
        voice_id="zh-CN-XiaoxiaoNeural",
        tts_rate=1.25,
        subtitle_fade_in_ms=90,
        subtitle_fade_out_ms=130,
        subtitle_scale_from=98,
        subtitle_scale_to=106,
        subtitle_outline=2,
        subtitle_shadow=1,
        subtitle_font="Microsoft YaHei",
        subtitle_size=38,
        bgm_path="bgm.mp3",
        bgm_volume=0.2,
        narration_volume=0.95,
        bgm_fade_in_s=0.5,
        bgm_fade_out_s=1.0,
    )

    req = _build_request_from_job_config(cfg)

    assert isinstance(req, BuildRequest)
    assert req.input_file.endswith("demo.pptx")
    assert req.out_dir.endswith("dist")
    assert req.pages == "1-3"
    assert req.ratio == "9:16"
    assert req.resolution == "1440p"
    assert req.fps == 60
    assert req.quality == "high"
    assert req.tts_provider == "edge"
    assert req.voice_id == "zh-CN-XiaoxiaoNeural"
    assert req.tts_rate == 1.25
    assert req.subtitle_fade_in_ms == 90
    assert req.subtitle_fade_out_ms == 130
    assert req.subtitle_scale_from == 98
    assert req.subtitle_scale_to == 106
    assert req.subtitle_outline == 2
    assert req.subtitle_shadow == 1
    assert req.subtitle_font == "Microsoft YaHei"
    assert req.subtitle_size == 38
    assert req.bgm_path == "bgm.mp3"
    assert req.bgm_volume == 0.2
    assert req.narration_volume == 0.95
    assert req.bgm_fade_in_s == 0.5
    assert req.bgm_fade_out_s == 1.0


def test_build_cli_argv_includes_script_file_when_temp_path_set(tmp_path) -> None:
    script_path = tmp_path / "override.txt"
    script_path.write_text("hello", encoding="utf-8")
    cfg = JobConfig(
        mode="build",
        pptx_path=tmp_path / "demo.pptx",
        out_dir=tmp_path / "dist",
        pages="all",
        ratio="1:1",
        resolution="720p",
        fps=24,
        quality="high",
        tts_provider="edge",
        voice_id="",
        tts_rate=1.0,
        script_text="ignored when temp path set",
        script_temp_path=str(script_path),
        subtitle_y_ratio=0.88,
    )
    argv = _build_cli_argv_for_config(cfg)
    assert argv[1:3] == ["-X", "utf8"]
    assert "--ratio" in argv
    r = argv[argv.index("--ratio") + 1]
    assert r == "1:1"
    assert argv[argv.index("--resolution") + 1] == "720p"
    assert argv[argv.index("--fps") + 1] == "24"
    assert argv[argv.index("--quality") + 1] == "high"
    assert "--subtitle-y-ratio" in argv
    y = float(argv[argv.index("--subtitle-y-ratio") + 1])
    assert abs(y - 0.88) < 1e-9
    assert "--script-file" in argv
    i = argv.index("--script-file")
    assert Path(argv[i + 1]).resolve() == script_path.resolve()


def test_build_request_prefers_script_file_over_script_text(tmp_path) -> None:
    p = tmp_path / "s.txt"
    p.write_text("from file", encoding="utf-8")
    cfg = JobConfig(
        mode="build",
        pptx_path=tmp_path / "demo.pptx",
        out_dir=tmp_path / "dist",
        pages="all",
        tts_provider="edge",
        voice_id="",
        tts_rate=1.0,
        script_text="from ui",
        script_temp_path=str(p),
    )
    req = _build_request_from_job_config(cfg)
    assert req.script_file == str(p)
    assert req.script_text is None


def test_run_pipeline_with_log_returns_error_and_trace_on_failure(monkeypatch, tmp_path) -> None:
    from note2video.gui.app import _run_pipeline_with_log

    logs: list[str] = []

    def fake_extract(_request):
        raise RuntimeError("extract failed for test")

    monkeypatch.setattr("note2video.gui.app.run_extract_pipeline", fake_extract)
    exit_code = _run_pipeline_with_log(_make_job_config(tmp_path, mode="extract"), logs.append)

    assert exit_code == 1
    assert any("RuntimeError: extract failed for test" in line for line in logs)


def test_job_config_profile_roundtrip(tmp_path) -> None:
    cfg = JobConfig(
        mode="build",
        pptx_path=tmp_path / "demo.pptx",
        out_dir=tmp_path / "dist",
        pages="2-4",
        ratio="9:16",
        resolution="720p",
        fps=24,
        quality="high",
        tts_provider="edge",
        voice_id="zh-CN-YunyangNeural",
        tts_rate=1.15,
        script_text="hello",
        subtitle_color="#FFFFFF",
        subtitle_fade_in_ms=90,
        subtitle_fade_out_ms=110,
        subtitle_scale_from=97,
        subtitle_scale_to=105,
        subtitle_outline=2,
        subtitle_shadow=1,
        subtitle_font="Microsoft YaHei",
        subtitle_size=36,
        subtitle_y_ratio=0.87,
        bgm_path="bgm.mp3",
        bgm_volume=0.22,
        narration_volume=0.93,
        bgm_fade_in_s=0.5,
        bgm_fade_out_s=0.8,
    )

    profile = _job_config_to_build_profile(cfg)
    assert profile["video"]["resolution"] == "720p"
    assert profile["video"]["fps"] == 24
    assert profile["video"]["quality"] == "high"

    profile_path = tmp_path / "profile.json"
    profile_path.write_text(json.dumps(profile, ensure_ascii=False), encoding="utf-8")
    loaded = _job_config_from_build_profile(profile, profile_path=profile_path)

    assert loaded.mode == "build"
    assert loaded.pptx_path == tmp_path / "demo.pptx"
    assert loaded.out_dir == tmp_path / "dist"
    assert loaded.pages == "2-4"
    assert loaded.ratio == "9:16"
    assert loaded.resolution == "720p"
    assert loaded.fps == 24
    assert loaded.quality == "high"
    assert loaded.voice_id == "zh-CN-YunyangNeural"
    assert loaded.tts_rate == 1.15
    assert loaded.script_text == "hello"
    assert loaded.bgm_path == str((tmp_path / "bgm.mp3").resolve())
