from note2video.subtitle.wrap import (
    SubtitleWrapLayout,
    estimate_max_chars_per_line,
    subtitle_wrap_layout_from_canvas,
    wrap_subtitle_text,
)


def test_short_text_single_line() -> None:
    assert wrap_subtitle_text("大家好", max_chars_per_line=18, max_lines=4) == "大家好"


def test_balanced_two_lines_not_split_at_first_comma() -> None:
    # Early comma must not force a 4-char first line when a balanced wrap exists.
    t = "其实，这是一个中等长度的句子用来测试换行"
    out = wrap_subtitle_text(t, max_chars_per_line=18, max_lines=4)
    lines = out.split("\n")
    assert len(lines) == 2
    assert min(len(line) for line in lines) >= 6


def test_three_lines_not_middle_heavy() -> None:
    # 40 CJK chars → needs 3 lines at 18/line; lines should stay similar length.
    t = "第一第二第三第四第五第六第七第八第九第十第十一第十二第十三第十四第十五第十六第十七第十八第十九第二十"
    out = wrap_subtitle_text(t, max_chars_per_line=18, max_lines=4)
    lines = out.split("\n")
    assert len(lines) == 3
    lengths = [len(x) for x in lines]
    assert max(lengths) - min(lengths) <= 4
    assert all(ln <= 18 for ln in lengths)


def test_preserves_explicit_newlines() -> None:
    assert wrap_subtitle_text("a\nb", max_chars_per_line=18, max_lines=4) == "a\nb"


def test_very_long_respects_max_lines() -> None:
    t = "字" * 80
    out = wrap_subtitle_text(t, max_chars_per_line=18, max_lines=4)
    lines = out.split("\n")
    assert len(lines) == 4
    assert all(len(line) <= 18 for line in lines)


def test_layout_from_canvas_reserves_safe_width_for_cjk() -> None:
    text = "前面我们讨论的大多是AI如何完成给定的任务"
    layout = subtitle_wrap_layout_from_canvas(
        canvas_w=1080,
        canvas_h=1920,
        font_size=48,
        margin_l=45,
        margin_r=45,
        outline=1,
        max_lines=4,
    )

    budget = estimate_max_chars_per_line(text=text, font_size=layout.font_size, max_width_px=layout.max_width_px)
    assert budget <= 20
    wrapped = wrap_subtitle_text(text, layout=layout)
    assert "\n" in wrapped

