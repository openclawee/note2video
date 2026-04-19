from note2video.video_canvas import canvas_size, normalize_ratio, normalize_resolution


def test_canvas_size_16_9_1080p() -> None:
    w, h = canvas_size(ratio="16:9", resolution="1080p")
    assert w == 1920 and h == 1080


def test_canvas_size_9_16_720p() -> None:
    w, h = canvas_size(ratio="9:16", resolution="720p")
    assert w == 720 and h == 1280


def test_normalize_helpers() -> None:
    assert normalize_ratio("16：9") == "16:9"
    assert normalize_resolution("1080P") == "1080p"
