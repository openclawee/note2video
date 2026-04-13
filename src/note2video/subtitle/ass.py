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
    base_color: str = "#FFFFFF",
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
    """
    fade_in_ms = max(0, int(fade_in_ms))
    fade_out_ms = max(0, int(fade_out_ms))
    scale_from = int(scale_from or 100)
    scale_to = int(scale_to or scale_from)
    outline = max(0, int(outline))
    shadow = max(0, int(shadow))
    font_size = max(8, int(font_size))

    primary = _ass_color_from_rgb_hex(base_color)

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
        line = f"Dialogue: 0,{_fmt_time(start_ms)},{_fmt_time(end_ms)},Default,,0,0,0,,{anim}{text}"
        events.append(line)

    return "\n".join(style_parts + events) + "\n"

