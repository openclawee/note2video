from __future__ import annotations

from dataclasses import dataclass


_SENTENCE_END_PUNCT = {"。", "！", "？", "!", "?", "；", ";", "."}
_PAUSE_AFTER_PUNCT: dict[str, int] = {
    "。": 420,
    "！": 420,
    "？": 420,
    ".": 360,
    "!": 380,
    "?": 380,
    "；": 320,
    ";": 320,
}


@dataclass(frozen=True)
class _SplitUnit:
    text: str
    boundary: str  # punct | newline | paragraph | eof
    punct: str | None = None


def split_sentences(text: str) -> list[str]:
    return [unit.text for unit in _iter_split_units(text)]


def split_sentences_with_pauses(
    text: str,
    *,
    newline_pause_ms: int = 380,
    paragraph_pause_ms: int = 620,
) -> list[tuple[str, int]]:
    out: list[tuple[str, int]] = []
    for unit in _iter_split_units(text):
        pause = 0
        if unit.boundary == "punct" and unit.punct is not None:
            pause = _PAUSE_AFTER_PUNCT.get(unit.punct, 0)
        elif unit.boundary == "newline":
            pause = int(newline_pause_ms)
        elif unit.boundary == "paragraph":
            pause = int(paragraph_pause_ms)
        out.append((unit.text, max(0, pause)))
    return out


def _iter_split_units(text: str) -> list[_SplitUnit]:
    normalized = _normalize_text(text)
    if not normalized:
        return []

    units: list[_SplitUnit] = []
    buf: list[str] = []
    i = 0
    n = len(normalized)

    while i < n:
        ch = normalized[i]
        if ch in _SENTENCE_END_PUNCT:
            buf.append(ch)
            j = i + 1
            while j < n and normalized[j] == "\n":
                j += 1
            if j > i + 1:
                boundary = "paragraph" if (j - i - 1) >= 2 else "newline"
                _flush(units, buf, boundary=boundary, punct=ch)
                i = j
                continue
            _flush(units, buf, boundary="punct", punct=ch)
            i += 1
            continue

        if ch == "\n":
            j = i
            while j < n and normalized[j] == "\n":
                j += 1
            boundary = "paragraph" if (j - i) >= 2 else "newline"
            _flush(units, buf, boundary=boundary)
            i = j
            continue

        buf.append(ch)
        i += 1

    _flush(units, buf, boundary="eof")
    return units


def _flush(units: list[_SplitUnit], buf: list[str], *, boundary: str, punct: str | None = None) -> None:
    text = "".join(buf).strip()
    buf.clear()
    if not text:
        return
    units.append(_SplitUnit(text=text, boundary=boundary, punct=punct))


def _normalize_text(text: str) -> str:
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
