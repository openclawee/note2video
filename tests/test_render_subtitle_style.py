from __future__ import annotations

from pathlib import Path

import pytest

from note2video.render.video import RenderError, _build_subtitle_filter


def test_build_subtitle_filter_default(tmp_path) -> None:
    srt = tmp_path / "a.srt"
    srt.write_text("1\n00:00:00,000 --> 00:00:01,000\nhi\n", encoding="utf-8")
    flt = _build_subtitle_filter(Path(srt))
    assert "subtitles='" in flt
    assert "force_style" not in flt


def test_build_subtitle_filter_with_color(tmp_path) -> None:
    srt = tmp_path / "a.srt"
    srt.write_text("1\n00:00:00,000 --> 00:00:01,000\nhi\n", encoding="utf-8")
    flt = _build_subtitle_filter(Path(srt), subtitle_color="#FF0000")
    assert "force_style" in flt
    # ASS PrimaryColour is &HAABBGGRR (FF0000 -> RR=FF GG=00 BB=00 => &H000000FF)
    assert "PrimaryColour=&H000000FF" in flt


@pytest.mark.parametrize("bad", ["", "fff", "#GGGGGG", "red", "#12345", "#1234567"])
def test_subtitle_color_validation(tmp_path, bad) -> None:
    srt = tmp_path / "a.srt"
    srt.write_text("1\n00:00:00,000 --> 00:00:01,000\nhi\n", encoding="utf-8")
    if not bad:
        # empty means no override
        flt = _build_subtitle_filter(Path(srt), subtitle_color=None)
        assert "force_style" not in flt
        return
    with pytest.raises(RenderError):
        _build_subtitle_filter(Path(srt), subtitle_color=bad)

