from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class AssStyle:
    font: str = ""
    font_size: int = 48
    primary_color: str = "#FFFFFF"  # base color
    outline: int = 1
    shadow: int = 0


def _ass_color_from_rgb_hex(value: str) -> str:
    raw = (value or "").strip()
    if raw.startswith("#"):
        raw = raw[1:]
    if len(raw) != 6:
        raise ValueError(f"Invalid color: {value!r}")
    rr = raw[0:2]
    gg = raw[2:4]
    bb = raw[4:6]
    return f"&H00{bb}{gg}{rr}"


def _fmt_time(ms: int) -> str:
    # ASS: H:MM:SS.cc (centiseconds)
    if ms < 0:
        ms = 0
    cs = int(round(ms / 10))
    s, cc = divmod(cs, 100)
    m, ss = divmod(s, 60)
    h, mm = divmod(m, 60)
    return f"{h:d}:{mm:02d}:{ss:02d}.{cc:02d}"


def _escape_text(text: str) -> str:
    # ASS uses {\} for override tags; escape braces and newlines.
    t = (text or "").replace("\r", "")
    t = t.replace("{", r"\{").replace("}", r"\}")
    t = t.replace("\n", r"\N")
    return t


def build_ass(
    *,
    segments: list[dict[str, Any]],
    highlight_mode: str = "none",  # none | line | word
    base_color: str = "#FFFFFF",
    highlight_color: str = "#FFD400",
    fade_in_ms: int = 80,
    fade_out_ms: int = 120,
    scale_from: int = 100,
    scale_to: int = 104,
    outline: int = 1,
    shadow: int = 0,
    font: str = "",
    font_size: int = 48,
) -> str:
    """
    Build an ASS subtitle document.

    segments: list of sentence-level segments with {start_ms, end_ms, text}
    If highlight_mode == 'word', each segment may include `words` with
    {text, offset_ms, duration_ms} in ms relative to sentence audio.
    """
    highlight_mode = (highlight_mode or "none").strip().lower()
    if highlight_mode not in {"none", "line", "word"}:
        highlight_mode = "none"

    fade_in_ms = max(0, int(fade_in_ms))
    fade_out_ms = max(0, int(fade_out_ms))
    scale_from = int(scale_from or 100)
    scale_to = int(scale_to or scale_from)
    outline = max(0, int(outline))
    shadow = max(0, int(shadow))
    font_size = max(8, int(font_size))

    primary = _ass_color_from_rgb_hex(base_color)
    hi = _ass_color_from_rgb_hex(highlight_color)

    style_parts = [
        "[Script Info]",
        "ScriptType: v4.00+",
        "PlayResX: 1920",
        "PlayResY: 1080",
        "ScaledBorderAndShadow: yes",
        "",
        "[V4+ Styles]",
        "Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, "
        "Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, "
        "Alignment, MarginL, MarginR, MarginV, Encoding",
    ]

    fontname = font.strip() or "Arial"
    # OutlineColour/BackColour use black; SecondaryColour used for karaoke but we use override tags.
    style_parts.append(
        "Style: Default,"
        f"{fontname},{font_size},{primary},&H00000000,&H00000000,&H64000000,"
        "0,0,0,0,100,100,0,0,1,"
        f"{outline},{shadow},2,80,80,60,1"
    )

    style_parts += ["", "[Events]", "Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text"]

    events: list[str] = []

    for seg in segments:
        text = _escape_text(str(seg.get("text", "") or ""))
        if not text.strip():
            continue
        start_ms = int(seg.get("start_ms", 0))
        end_ms = int(seg.get("end_ms", 0))
        if end_ms <= start_ms:
            continue

        dur_ms = end_ms - start_ms
        fi = min(fade_in_ms, max(0, dur_ms // 2))
        fo = min(fade_out_ms, max(0, dur_ms // 2))

        # Base animation: fade + slight scale-in using \t.
        # \fscx/\fscy are percentages.
        anim = rf"{{\fad({fi},{fo})\fscx{scale_from}\fscy{scale_from}\t(0,{min(200, dur_ms)},\fscx{scale_to}\fscy{scale_to})}}"

        if highlight_mode == "none":
            line = f"Dialogue: 0,{_fmt_time(start_ms)},{_fmt_time(end_ms)},Default,,0,0,0,,{anim}{text}"
            events.append(line)
            continue

        if highlight_mode == "line":
            # Whole-line highlight color at appearance, then transition back to base color.
            # Transition ends at 35% duration, capped to 900ms.
            t_end = min(int(dur_ms * 0.35), 900)
            line = (
                f"Dialogue: 0,{_fmt_time(start_ms)},{_fmt_time(end_ms)},Default,,0,0,0,,"
                rf"{{\1c{hi}}}{anim}"
                rf"{{\t(0,{t_end},\1c{primary})}}"
                f"{text}"
            )
            events.append(line)
            continue

        # highlight_mode == "word": use per-word color highlight on the subtitle line itself.
        # Approach: draw a highlight-colored rectangle behind each word using \clip with
        # a vector drawing. Character positions are proportional to word timing:
        #   word_char_end = word_offset_ms / sentence_duration_ms * total_chars
        words = seg.get("words") or []
        if not isinstance(words, list) or not words:
            # Fallback: treat as line highlight.
            t_end = min(int(dur_ms * 0.35), 900)
            line = (
                f"Dialogue: 0,{_fmt_time(start_ms)},{_fmt_time(end_ms)},Default,,0,0,0,,"
                rf"{{\1c{hi}}}{anim}"
                rf"{{\t(0,{t_end},\1c{primary})}}"
                f"{text}"
            )
            events.append(line)
            continue

        # Build vector path for the full sentence: each character as a rect of equal width.
        # Path coordinate system: 1 DU = 1/20th of a pixel at PlayResX=1920.
        # Subtitle area: x=80..x=1840 (MarginL=80, MarginR=80), total width = 1760 px = 35200 DU.
        SUB_AREA_LEFT = 80 * 20   # = 1600 DU
        SUB_AREA_RIGHT = 1840 * 20  # = 36800 DU
        SUB_AREA_WIDTH = SUB_AREA_RIGHT - SUB_AREA_LEFT  # = 35200 DU
        SUB_H = font_size * 20 * 2   # height: 2x font size in DU

        raw_text = str(seg.get("text", "") or "").strip()
        n_chars = len(raw_text) or 1
        CHAR_W = SUB_AREA_WIDTH // n_chars  # drawing units per character slot

        def _word_clip_rect(word_text: str, word_offset_ms: int, word_dur_ms: int) -> str:
            """Build a vector clip rect for a word's character range in drawing units."""
            if not raw_text:
                return ""
            # Find word's character range: proportional to timing within the sentence.
            word_start_ratio = max(0.0, min(1.0, word_offset_ms / dur_ms)) if dur_ms > 0 else 0.0
            word_end_ratio = max(0.0, min(1.0, (word_offset_ms + word_dur_ms) / dur_ms)) if dur_ms > 0 else 1.0
            word_start_char = int(round(word_start_ratio * n_chars))
            word_end_char = int(round(word_end_ratio * n_chars))
            word_end_char = max(word_end_char, word_start_char + 1)  # at least 1 char
            x0 = SUB_AREA_LEFT + word_start_char * CHAR_W
            x1 = SUB_AREA_LEFT + word_end_char * CHAR_W
            y0 = 0
            y1 = SUB_H
            return f"m {x0} {y0} l {x1} {y0} l {x1} {y1} l {x0} {y1} 0 0"

        base_line = f"Dialogue: 0,{_fmt_time(start_ms)},{_fmt_time(end_ms)},Default,,0,0,0,,{anim}{text}"
        events.append(base_line)

        for w in words:
            try:
                off = int(w.get("offset_ms", 0))
                d = int(w.get("duration_ms", 0))
            except Exception:
                continue
            if d <= 0:
                continue
            w_start = start_ms + max(0, off)
            w_end = min(end_ms, w_start + d)
            if w_end <= w_start:
                continue
            clip_rect = _word_clip_rect(str(w.get("text", "") or "").strip(), off, d)
            if not clip_rect:
                continue
            # Clip the full subtitle text to just this word's character range, color hi.
            hi_line = (
                f"Dialogue: 5,{_fmt_time(w_start)},{_fmt_time(w_end)},Default,,0,0,0,,"
                rf"{{\pos(80,900)\an4\clip(\p1{clip_rect}\p0)\1c{hi}\3c{hi}}}{text}"
            )
            events.append(hi_line)

    return "\n".join(style_parts + events) + "\n"

