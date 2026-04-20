from __future__ import annotations

"""
Subtitle line wrapping: width-aware (font + canvas) with balanced lines.

Sentence segmentation already removes most sentence-ending punctuation from cues,
so soft breaks prefer list separators / quotes / spaces — not 。！？ as primary anchors.
"""

from dataclasses import dataclass
from typing import Protocol

_PUNCT_SOFT = frozenset("，,、；;：:")  # list-like / mid-sentence (cues rarely end with 。！？)
_QUOTE_LIKE = frozenset("""'"「」『』（）()[]【】《》<>""")
_WHITESPACE = frozenset(" \t")


class _FontMeasure(Protocol):
    def line_width(self, text: str) -> int: ...


@dataclass
class SubtitleWrapLayout:
    """Pixel budget for one subtitle cue (matches ASS margins + outline)."""

    max_width_px: int
    font_size: int = 48
    max_lines: int = 4


def subtitle_wrap_layout_from_canvas(
    *,
    canvas_w: int,
    canvas_h: int,
    font_size: int = 48,
    margin_l: int = 80,
    margin_r: int = 80,
    outline: int = 1,
    max_lines: int = 4,
) -> SubtitleWrapLayout:
    cw = max(64, int(canvas_w))
    fs = max(8, int(font_size))
    ml = max(0, int(margin_l))
    mr = max(0, int(margin_r))
    ol = max(0, int(outline))
    # Reserve horizontal space for outline/shadow (ScaledBorderAndShadow).
    usable = cw - ml - mr - 2 * ol - 8
    safety_px = max(16, int(round(fs * 0.65)))
    max_w = max(120, usable - safety_px)
    return SubtitleWrapLayout(max_width_px=max_w, font_size=fs, max_lines=max(1, int(max_lines)))


def wrap_subtitle_text(
    text: str,
    *,
    max_chars_per_line: int = 18,
    max_lines: int = 4,
    layout: SubtitleWrapLayout | None = None,
    font_name: str = "",
) -> str:
    """
    Wrap subtitle text. If ``layout`` is set, break by measured pixel width (Pillow);
    otherwise fall back to character-count heuristics (``max_chars_per_line``).
    """
    t = (text or "").replace("\r", "\n").strip()
    if not t:
        return ""
    if "\n" in t:
        return "\n".join(line.strip() for line in t.splitlines() if line.strip())
    if max_lines <= 1:
        return t

    if layout is not None and layout.max_width_px > 0:
        meas = _try_create_pil_measure(font_name=font_name, font_size=layout.font_size)
        if meas is not None:
            return _wrap_balanced_pixels(t, meas=meas, max_px=layout.max_width_px, max_lines=layout.max_lines)
        m = estimate_max_chars_per_line(text=t, font_size=layout.font_size, max_width_px=layout.max_width_px)
        return _wrap_balanced_chars(t, max_chars_per_line=m, max_lines=layout.max_lines)

    return _wrap_balanced_chars(t, max_chars_per_line=max_chars_per_line, max_lines=max_lines)


def _try_create_pil_measure(*, font_name: str, font_size: int) -> _FontMeasure | None:
    try:
        from PIL.ImageFont import FreeTypeFont
    except Exception:
        FreeTypeFont = None  # type: ignore[misc,assignment]

    font = _load_pil_font(font_name, font_size)
    if font is None:
        return None
    # Bitmap default font gives misleading widths for CJK; use char fallback instead.
    if FreeTypeFont is not None and not isinstance(font, FreeTypeFont):
        return None
    return _PillowFontMeasure(font=font)


class _PillowFontMeasure:
    def __init__(self, *, font) -> None:
        self._font = font

    def line_width(self, text: str) -> int:
        if not text:
            return 0
        if hasattr(self._font, "getlength"):
            return int(round(float(self._font.getlength(text))))  # type: ignore[union-attr]
        bbox = self._font.getbbox(text)
        return int(bbox[2] - bbox[0])


def _load_pil_font(font_name: str, font_size: int):
    from pathlib import Path

    from PIL import ImageFont

    size = max(8, int(font_size))
    raw = (font_name or "").strip()
    candidates: list[str] = []
    if raw:
        p = Path(raw)
        if p.is_file():
            candidates.append(str(p))
        candidates.extend(_named_font_candidates(raw))
        candidates.extend([raw, f"{raw}.ttf", f"{raw}.ttc", f"{raw}.otf"])
    # Common Linux paths (Pillow may not resolve family names without fc-list).
    candidates.extend(
        [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\msyhbd.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            r"C:\Windows\Fonts\simsun.ttc",
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\segoeui.ttf",
        ]
    )
    for path in candidates:
        if not path:
            continue
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return None


def _named_font_candidates(font_name: str) -> list[str]:
    key = font_name.strip().casefold()
    if not key:
        return []
    windows = {
        "microsoft yahei": [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyhbd.ttc"],
        "微软雅黑": [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyhbd.ttc"],
        "simhei": [r"C:\Windows\Fonts\simhei.ttf"],
        "黑体": [r"C:\Windows\Fonts\simhei.ttf"],
        "simsun": [r"C:\Windows\Fonts\simsun.ttc"],
        "宋体": [r"C:\Windows\Fonts\simsun.ttc"],
        "arial": [r"C:\Windows\Fonts\arial.ttf"],
        "segoe ui": [r"C:\Windows\Fonts\segoeui.ttf"],
    }
    return windows.get(key, [])


def _span_width(meas: _FontMeasure, t: str, start: int, end: int) -> int:
    if start >= end:
        return 0
    return meas.line_width(t[start:end])


def _max_fit_end(meas: _FontMeasure, t: str, start: int, max_px: int) -> int:
    """Largest end index such that width(t[start:end]) <= max_px (end > start)."""
    if start >= len(t):
        return start
    lo, hi = start + 1, len(t)
    best = start + 1
    while lo <= hi:
        mid = (lo + hi) // 2
        w = _span_width(meas, t, start, mid)
        if w <= max_px:
            best = mid
            lo = mid + 1
        else:
            hi = mid - 1
    return best


def _min_lines_greedy(meas: _FontMeasure, t: str, start: int, end: int, max_px: int) -> int:
    """Lower bound on line count for t[start:end] with greedy max-fit wrapping."""
    n = 0
    i = start
    while i < end:
        j = _max_fit_end(meas, t, i, max_px)
        if j <= i:
            j = i + 1
        i = j
        n += 1
    return n


def _end_for_target_width(meas: _FontMeasure, t: str, prev: int, target_w: float, lo: int, hi: int) -> int:
    """Pick ``cut`` in [lo, hi] inclusive so width(prev:cut) is closest to ``target_w``."""
    lo = max(prev + 1, lo)
    hi = max(lo, min(len(t), hi))
    if lo > hi:
        return lo
    return min(range(lo, hi + 1), key=lambda e: abs(_span_width(meas, t, prev, e) - target_w))


def _wrap_balanced_pixels(t: str, *, meas: _FontMeasure, max_px: int, max_lines: int) -> str:
    if max_px <= 0:
        return t
    total_w = _span_width(meas, t, 0, len(t))
    if total_w <= max_px:
        return t

    n_lines = min(max_lines, max(2, _min_lines_greedy(meas, t, 0, len(t), max_px)))
    if n_lines <= 1:
        return t

    cuts: list[int] = [0]
    for j in range(1, n_lines):
        prev = cuts[-1]
        rem = len(t) - prev
        k = n_lines - (j - 1)
        max_end = _max_fit_end(meas, t, prev, max_px)
        max_chars = max(1, max_end - prev)
        lo_seg = max(1, rem - max_chars * (k - 1))
        hi_seg = min(max_chars, rem - (k - 1))
        if lo_seg > hi_seg:
            lo_seg = max(1, rem - (k - 1))
            hi_seg = min(max_chars, rem - (k - 1))
            if lo_seg > hi_seg:
                hi_seg = max(lo_seg, min(max_chars, rem - (k - 1)))

        rem_w = float(_span_width(meas, t, prev, len(t)))
        ideal_seg_w = max(1.0, rem_w / float(k))
        lo_i = prev + lo_seg
        hi_i = prev + hi_seg
        ideal_cut = _end_for_target_width(meas, t, prev, ideal_seg_w, lo_i, hi_i)

        window = max(12, (hi_seg - lo_seg) // 2 + 2)
        lo2 = max(lo_i, ideal_cut - window)
        hi2 = min(hi_i, ideal_cut + window)
        if lo2 > hi2:
            lo2, hi2 = lo_i, hi_i

        cut = _choose_cut_px(t, meas=meas, prev=prev, lo=lo2, hi=hi2, ideal_cut=ideal_cut, ideal_seg_w=ideal_seg_w)
        cut = max(lo_i, min(hi_i, cut))
        max_end2 = _max_fit_end(meas, t, prev, max_px)
        if cut > max_end2:
            cut = max_end2
        if cut <= prev:
            cut = min(prev + hi_seg, len(t) - (n_lines - j))
            cut = max(prev + lo_seg, cut)
        cuts.append(cut)

    cuts.append(len(t))
    lines = [t[cuts[i] : cuts[i + 1]].strip() for i in range(len(cuts) - 1)]
    lines = [ln for ln in lines if ln]
    return "\n".join(lines)


def _choose_cut_px(
    t: str,
    *,
    meas: _FontMeasure,
    prev: int,
    lo: int,
    hi: int,
    ideal_cut: int,
    ideal_seg_w: float,
) -> int:
    ideal_cut = max(lo, min(hi, ideal_cut))

    def sort_key(cut: int) -> tuple[float, int, int, int]:
        w = _span_width(meas, t, prev, cut)
        balance = abs(float(w) - ideal_seg_w)
        dist = abs(cut - ideal_cut)
        last = t[cut - 1]
        on_soft = last in _PUNCT_SOFT or last in _WHITESPACE or last.isspace()
        on_quote = last in _QUOTE_LIKE
        punct_bonus = 0 if on_soft else (1 if on_quote else 2)
        primary = 1.4 * balance + float(dist)
        return (primary, punct_bonus, _break_rank(last) if on_soft else 50, cut)

    return min(range(lo, hi + 1), key=sort_key)


def _wrap_balanced_chars(t: str, *, max_chars_per_line: int, max_lines: int) -> str:
    m = max_chars_per_line
    l_total = len(t)
    if m <= 0:
        return t
    if l_total <= m:
        return t

    cap = m * max_lines
    if l_total > cap:
        lines: list[str] = []
        rest = t
        while rest and len(lines) < max_lines:
            if len(rest) <= m:
                lines.append(rest)
                break
            lines.append(rest[:m].rstrip())
            rest = rest[m:].lstrip()
        return "\n".join(lines)

    n_lines = min(max_lines, (l_total + m - 1) // m)
    if n_lines <= 1:
        return t

    window = max(m // 2, 8)
    cuts: list[int] = [0]
    for j in range(1, n_lines):
        prev = cuts[-1]
        rem = l_total - prev
        k = n_lines - (j - 1)
        lo_seg = max(1, rem - m * (k - 1))
        hi_seg = min(m, rem - (k - 1))
        if lo_seg > hi_seg:
            lo_seg = max(1, min(m, rem - (k - 1)))
            hi_seg = max(lo_seg, min(m, rem - (k - 1)))

        ideal_len = max(lo_seg, min(hi_seg, round(rem / k)))
        ideal_cut = prev + ideal_len

        lo = prev + lo_seg
        hi = prev + hi_seg
        lo2 = max(lo, ideal_cut - window)
        hi2 = min(hi, ideal_cut + window)
        if lo2 > hi2:
            lo2, hi2 = lo, hi

        cut = _choose_cut_chars(t, prev=prev, lo=lo2, hi=hi2, ideal_cut=ideal_cut, ideal_len=ideal_len)
        cut = max(lo, min(hi, cut))
        if cut <= prev:
            cut = min(prev + hi_seg, l_total - (n_lines - j))
            cut = max(prev + lo_seg, cut)
        cuts.append(cut)

    cuts.append(l_total)
    lines = [t[cuts[i] : cuts[i + 1]].strip() for i in range(len(cuts) - 1)]
    lines = [ln for ln in lines if ln]
    return "\n".join(lines)


def _break_rank(ch: str) -> int:
    if ch in "，,":
        return 0
    if ch in "、":
        return 1
    if ch in "；;":
        return 2
    if ch in "：:":
        return 3
    if ch.isspace() or ch in _WHITESPACE:
        return 4
    if ch in _QUOTE_LIKE:
        return 5
    return 50


def _choose_cut_chars(
    t: str,
    *,
    prev: int,
    lo: int,
    hi: int,
    ideal_cut: int,
    ideal_len: int,
) -> int:
    l_total = len(t)
    lo = max(prev + 1, min(lo, l_total - 1))
    hi = max(lo, min(hi, l_total - 1))
    ideal_cut = max(lo, min(hi, ideal_cut))

    def sort_key(cut: int) -> tuple[int, int, int, int]:
        seg_len = cut - prev
        balance = abs(seg_len - ideal_len)
        dist = abs(cut - ideal_cut)
        last = t[cut - 1]
        on_soft = last in _PUNCT_SOFT or last in _WHITESPACE or last.isspace()
        on_quote = last in _QUOTE_LIKE
        punct_bonus = 0 if on_soft else (1 if on_quote else 2)
        primary = 14 * balance + dist
        return (primary, punct_bonus, _break_rank(last) if on_soft else 50, cut)

    return min(range(lo, hi + 1), key=sort_key)


def estimate_max_chars_per_line(*, text: str = "", font_size: int, max_width_px: int) -> int:
    """Rough char budget when Pillow font loading fails."""
    fs = max(8, int(font_size))
    mw = max(80, int(max_width_px))
    stripped = str(text or "").strip()
    cjk_ratio = _cjk_ratio(stripped)
    per_char = 0.95 if cjk_ratio >= 0.5 else 0.58
    est = int(mw / (per_char * float(fs)))
    upper = 22 if cjk_ratio >= 0.5 else 40
    lower = 6 if cjk_ratio >= 0.5 else 8
    return max(lower, min(upper, est))


def _cjk_ratio(text: str) -> float:
    if not text:
        return 0.0
    visible = [ch for ch in text if not ch.isspace()]
    if not visible:
        return 0.0
    cjk = sum(1 for ch in visible if _is_cjk(ch))
    return float(cjk) / float(len(visible))


def _is_cjk(ch: str) -> bool:
    code = ord(ch)
    return (
        0x3400 <= code <= 0x4DBF
        or 0x4E00 <= code <= 0x9FFF
        or 0xF900 <= code <= 0xFAFF
        or 0x3040 <= code <= 0x30FF
        or 0xAC00 <= code <= 0xD7AF
    )
