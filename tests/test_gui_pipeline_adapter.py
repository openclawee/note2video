from __future__ import annotations

from pathlib import Path

from note2video.gui.app import JobConfig, _run_extract_or_build


def _make_job_config(tmp_path: Path, mode: str = "build") -> JobConfig:
    return JobConfig(
        mode=mode,
        pptx_path=tmp_path / "demo.pptx",
        out_dir=tmp_path / "dist",
        pages="all",
        tts_provider="pyttsx3",
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
    assert build_request.tts_provider == "pyttsx3"
    assert build_request.tts_rate == 1.0
    assert "阶段：voice" in logs
    assert "细节：provider=pyttsx3, voice=default, tts_rate=1.0" in logs
    assert "细节：segments=12, slides=5" in logs
    assert "细节：subtitles_burned=True, mixed_audio=True" in logs
    assert "输出视频：video/output.mp4" in logs
