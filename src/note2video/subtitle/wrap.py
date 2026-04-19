from __future__ import annotations

"""Balanced subtitle line wrapping (shared by subtitle generation and video render)."""

_PUNCT_BREAK = frozenset("，,、：:；;。！？!?")
_WHITESPACE = frozenset(" \t")


def wrap_subtitle_text(
    text: str,
    *,
    max_chars_per_line: int = 18,
    max_lines: int = 4,
) -> str:
    """
    Wrap subtitle text into up to ``max_lines`` lines without exceeding ``max_chars_per_line``.

    Uses proportional targets so line lengths stay similar; prefers breaking after
    punctuation or space when close to the target. If the text is longer than
    ``max_lines * max_chars_per_line``, falls back to fixed-width chunks (same as before).
    """
    t = (text or "").replace("\r", "\n").strip()
    if not t:
        return ""
    if "\n" in t:
        return "\n".join(line.strip() for line in t.splitlines() if line.strip())
    if max_chars_per_line <= 0 or max_lines <= 1:
        return t
    m = max_chars_per_line
    l_total = len(t)
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

        cut = _choose_cut(t, prev=prev, lo=lo2, hi=hi2, ideal_cut=ideal_cut, ideal_len=ideal_len)
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
    """Lower is better for a mid-sentence wrap (prefer comma over period)."""
    if ch in "，,":
        return 0
    if ch in "、":
        return 1
    if ch in "；;":
        return 2
    if ch in "：:":
        return 3
    if ch.isspace():
        return 4
    if ch in "。！？!?":
        return 20
    if ch == ".":
        return 20
    return 50


def _choose_cut(
    t: str,
    *,
    prev: int,
    lo: int,
    hi: int,
    ideal_cut: int,
    ideal_len: int,
) -> int:
    """Pick split index ``cut`` so first segment is ``t[prev:cut]``."""
    l_total = len(t)
    lo = max(prev + 1, min(lo, l_total - 1))
    hi = max(lo, min(hi, l_total - 1))
    ideal_cut = max(lo, min(hi, ideal_cut))

    def sort_key(cut: int) -> tuple[int, int, int, int]:
        seg_len = cut - prev
        balance = abs(seg_len - ideal_len)
        dist = abs(cut - ideal_cut)
        last = t[cut - 1]
        on_punct = last in _PUNCT_BREAK or last in _WHITESPACE or last.isspace()
        primary = 14 * balance + dist
        punct_bonus = _break_rank(last) if on_punct else 99
        # Prefer punctuation only when it does not destroy balance; then finer punct rank.
        punct_pref = 0 if on_punct else 1
        return (primary, punct_pref, punct_bonus, cut)

    return min(range(lo, hi + 1), key=sort_key)
