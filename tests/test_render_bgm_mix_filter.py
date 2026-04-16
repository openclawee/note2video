"""Regression: BGM volume must apply linearly; ffmpeg amix normalize must be off."""

from __future__ import annotations

import json
import wave
from pathlib import Path

import pytest

from note2video.render.video import render_video


def test_bgm_mix_uses_amix_without_normalize(tmp_path, monkeypatch: pytest.MonkeyPatch) -> None:
    root = tmp_path / "proj"
    audio_dir = root / "audio"
    video_dir = root / "video"
    slides_dir = root / "slides"
    subtitles_dir = root / "subtitles"
    audio_dir.mkdir(parents=True)
    video_dir.mkdir(parents=True)
    slides_dir.mkdir(parents=True)
    subtitles_dir.mkdir(parents=True)

    (slides_dir / "001.png").write_bytes(b"\x89PNG\r\n\x1a\n")
    merged = audio_dir / "merged.wav"
    with wave.open(str(merged), "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(44_100)
        wf.writeframes(b"\x00\x00" * 44_100)
    (subtitles_dir / "subtitles.srt").write_text("1\n00:00:00,000 --> 00:00:01,000\nHi\n", encoding="utf-8")
    bgm = root / "bgm.mp3"
    bgm.write_bytes(b"fake-mp3")

    (root / "manifest.json").write_text(
        json.dumps(
            {
                "project_name": "demo",
                "input_file": "demo.pptx",
                "slide_count": 1,
                "outputs": {
                    "merged_audio": "audio/merged.wav",
                    "subtitle": "subtitles/subtitles.srt",
                },
                "slides": [
                    {
                        "page": 1,
                        "image": "slides/001.png",
                        "duration_ms": 1000,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    commands: list[list[str]] = []

    def fake_run_ffmpeg(command: list[str]) -> None:
        commands.append(command)
        out = command[-1]
        if out.endswith("video_only.mp4"):
            Path(out).write_bytes(b"video-only")
        elif out.endswith("output.mp4"):
            Path(out).write_bytes(b"video-final")
        elif out.endswith("mixed.m4a"):
            Path(out).write_bytes(b"mixed-audio")

    monkeypatch.setattr("note2video.render.video._get_ffmpeg_path", lambda: "ffmpeg")
    monkeypatch.setattr("note2video.render.video._run_ffmpeg", fake_run_ffmpeg)

    render_video(str(root), bgm_path=str(bgm), bgm_volume=0.05, narration_volume=1.0)

    mix_cmds = [c for c in commands if "-filter_complex" in c]
    assert mix_cmds, "expected a filter_complex ffmpeg invocation for BGM mix"
    fc_idx = mix_cmds[0].index("-filter_complex") + 1
    graph = mix_cmds[0][fc_idx]
    assert "volume=0.050" in graph
    assert "normalize=0" in graph
    assert "alimiter=limit=0.98:level=0" in graph
