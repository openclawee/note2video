from __future__ import annotations

import json
from pathlib import Path

from note2video.gui.preview_model import load_preview_data


def _write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def test_load_preview_data_prefers_subtitles_and_reads_manifest_outputs(tmp_path: Path) -> None:
    (tmp_path / "custom" / "slides").mkdir(parents=True)
    image_path = tmp_path / "custom" / "slides" / "cover.png"
    image_path.write_bytes(b"png")

    _write_json(
        tmp_path / "manifest.json",
        {
            "slides": [{"page": 1, "title": "封面", "image": "custom/slides/cover.png"}],
            "outputs": {
                "script": "custom/scripts/story.json",
                "subtitle_json": "custom/subtitles/captions.json",
            },
        },
    )
    _write_json(
        tmp_path / "custom" / "scripts" / "story.json",
        {"slides": [{"page": 1, "script": "脚本文本"}]},
    )
    _write_json(
        tmp_path / "custom" / "subtitles" / "captions.json",
        {
            "segments": [
                {"index": 2, "page": 1, "start_ms": 500, "end_ms": 900, "text": "第二句"},
                {"index": 1, "page": 1, "start_ms": 0, "end_ms": 500, "text": "第一句"},
            ]
        },
    )

    data = load_preview_data(
        project_dir=tmp_path,
        page=1,
        ratio="9:16",
        resolution="720p",
        sample_text="示例文本",
    )

    assert data.page == 1
    assert data.available_pages == (1,)
    assert data.page_count == 1
    assert data.canvas_w == 720
    assert data.canvas_h == 1280
    assert data.title == "封面"
    assert data.image_path == image_path
    assert data.text_source == "subtitle"
    assert data.cue_count == 2
    assert [cue.text for cue in data.cues] == ["第一句", "第二句"]
    assert data.active_cue_index == 0
    assert data.active_text == "第一句"
    assert data.status_text == "文本来源：subtitles/subtitles.json · 当前页字幕 1/2"

    second = load_preview_data(
        project_dir=tmp_path,
        page=1,
        ratio="9:16",
        resolution="720p",
        sample_text="示例文本",
        cue_index=1,
    )
    assert second.active_cue_index == 1
    assert second.active_text == "第二句"
    assert second.status_text == "文本来源：subtitles/subtitles.json · 当前页字幕 2/2"


def test_load_preview_data_falls_back_to_script_then_sample(tmp_path: Path) -> None:
    _write_json(
        tmp_path / "scripts" / "script.json",
        {"slides": [{"page": 3, "script": "第三页脚本"}]},
    )

    script_data = load_preview_data(
        project_dir=tmp_path,
        page=3,
        ratio="16:9",
        resolution="1080p",
        sample_text="示例文本",
    )

    assert script_data.page == 3
    assert script_data.text_source == "script"
    assert script_data.cue_count == 1
    assert script_data.active_text == "第三页脚本"
    assert script_data.status_text == "文本来源：scripts/script.json"

    sample_data = load_preview_data(
        project_dir=tmp_path / "empty-project",
        page=1,
        ratio="16:9",
        resolution="1080p",
        sample_text="示例文本",
    )

    assert sample_data.page == 1
    assert sample_data.text_source == "sample"
    assert sample_data.cue_count == 1
    assert sample_data.active_text == "示例文本"
    assert sample_data.status_text == "输出目录不存在，当前显示示例文本。"




def test_load_preview_data_preserves_sparse_real_page_numbers(tmp_path: Path) -> None:
    slides_dir = tmp_path / "slides"
    slides_dir.mkdir(parents=True)
    (slides_dir / "002.png").write_bytes(b"a")
    (slides_dir / "005.png").write_bytes(b"b")
    _write_json(
        tmp_path / "subtitles" / "subtitles.json",
        {
            "segments": [
                {"index": 1, "page": 2, "start_ms": 0, "end_ms": 500, "text": "第二页字幕"},
                {"index": 1, "page": 5, "start_ms": 0, "end_ms": 500, "text": "第五页字幕"},
            ]
        },
    )

    early = load_preview_data(
        project_dir=tmp_path,
        page=1,
        ratio="16:9",
        resolution="1080p",
        sample_text="示例文本",
    )
    assert early.available_pages == (2, 5)
    assert early.page == 2
    assert early.image_path == slides_dir / "002.png"
    assert [cue.page for cue in early.cues] == [2]
    assert early.active_text == "第二页字幕"

    later = load_preview_data(
        project_dir=tmp_path,
        page=4,
        ratio="16:9",
        resolution="1080p",
        sample_text="示例文本",
    )
    assert later.available_pages == (2, 5)
    assert later.page == 5
    assert later.image_path == slides_dir / "005.png"
    assert [cue.page for cue in later.cues] == [5]
    assert later.active_text == "第五页字幕"
