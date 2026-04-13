from __future__ import annotations

from types import SimpleNamespace

from note2video.app.pipeline_service import (
    BuildRequest,
    ExtractRequest,
    RenderRequest,
    SubtitleRequest,
    VoiceRequest,
    run_build_pipeline,
    run_extract_pipeline,
    run_render_pipeline,
    run_subtitle_pipeline,
    run_voice_pipeline,
)


def test_build_runs_pipeline_in_order(monkeypatch, tmp_path) -> None:
    calls: list[tuple] = []

    def fake_extract(input_path, out_dir, pages=None):
        calls.append(("extract", input_path, out_dir, pages))
        return SimpleNamespace(slide_count=3)

    def fake_voice(input_json, out_dir, *, provider_name, voice_id, tts_rate, minimax_base_url):
        calls.append(("voice", input_json, out_dir, provider_name, voice_id, tts_rate, minimax_base_url))
        return {"provider": provider_name}

    def fake_subtitle(input_json, out_dir):
        calls.append(("subtitle", input_json, out_dir))
        return {"segment_count": 5}

    def fake_render(project_dir, output_path=None, **kwargs):
        calls.append(("render", project_dir, output_path, kwargs.get("subtitle_fade_in_ms")))
        return {"video": "video/output.mp4", "subtitles_burned": True, "mixed_audio": "audio/mixed.wav"}

    req = BuildRequest(
        input_file=str(tmp_path / "demo.pptx"),
        out_dir=str(tmp_path / "dist"),
        pages="1-3",
        tts_provider="pyttsx3",
        subtitle_color=None,
    )
    run_build_pipeline(
        req,
        extract_project_fn=fake_extract,
        generate_voice_assets_fn=fake_voice,
        generate_subtitles_fn=fake_subtitle,
        render_video_fn=fake_render,
    )

    assert calls[0][0] == "extract"
    assert calls[1][0] == "voice"
    assert calls[2][0] == "subtitle"
    assert calls[3][0] == "render"


def test_build_includes_mixed_audio_flag(tmp_path) -> None:
    req = BuildRequest(input_file="in.pptx", out_dir=str(tmp_path / "dist"))
    result = run_build_pipeline(
        req,
        extract_project_fn=lambda *args, **kwargs: SimpleNamespace(slide_count=1),
        generate_voice_assets_fn=lambda *args, **kwargs: {"provider": "pyttsx3"},
        generate_subtitles_fn=lambda *args, **kwargs: {"segment_count": 1},
        render_video_fn=lambda *args, **kwargs: {
            "video": "video/output.mp4",
            "subtitles_burned": True,
            "mixed_audio": "audio/mixed.wav",
        },
    )
    assert result["mixed_audio"] is True


def test_extract_delegates(monkeypatch, tmp_path) -> None:
    calls: list[tuple] = []

    def fake_extract(input_path, out_dir, pages=None):
        calls.append((input_path, out_dir, pages))
        return SimpleNamespace(slide_count=1, outputs={})

    req = ExtractRequest(input_file="in.pptx", out_dir=str(tmp_path / "dist"), pages="all")
    result = run_extract_pipeline(req, extract_project_fn=fake_extract)

    assert calls == [("in.pptx", str(tmp_path / "dist"), "all")]
    assert result["slide_count"] == 1


def test_voice_subtitle_render_delegate(monkeypatch, tmp_path) -> None:
    voice_result = run_voice_pipeline(
        VoiceRequest(input_json="a.json", out_dir=str(tmp_path), tts_provider="edge"),
        generate_voice_assets_fn=lambda *args, **kwargs: {
            "provider": kwargs["provider_name"],
            "slide_count": 2,
            "voice": "",
            "tts_rate": 1.0,
        },
    )
    subtitle_result = run_subtitle_pipeline(
        SubtitleRequest(input_json="a.json", out_dir=str(tmp_path)),
        generate_subtitles_fn=lambda *args, **kwargs: {"slide_count": 2, "segment_count": 3},
    )
    render_result = run_render_pipeline(
        RenderRequest(project_dir=str(tmp_path)),
        render_video_fn=lambda *args, **kwargs: {
            "video": "video/output.mp4",
            "subtitles_burned": True,
            "slide_count": 2,
        },
    )

    assert voice_result["provider"] == "edge"
    assert subtitle_result["segment_count"] == 3
    assert render_result["artifacts"]["video"] == "video/output.mp4"
