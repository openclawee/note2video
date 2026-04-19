from __future__ import annotations

"""Output canvas size from ratio + resolution preset (shared by render and subtitle layout)."""


def normalize_ratio(value: str | None) -> str:
    raw = str(value or "").strip().lower()
    if not raw:
        return "16:9"
    raw = raw.replace("：", ":").replace("x", ":")
    raw = raw.replace(" ", "")
    if raw in {"16:9", "9:16", "1:1"}:
        return raw
    raise ValueError(f"Unsupported ratio: {value!r}. Use 16:9, 9:16, or 1:1.")


def normalize_resolution(value: str | None) -> str:
    raw = str(value or "").strip().lower()
    if not raw:
        return "1080p"
    if raw in {"720p", "1080p", "1440p"}:
        return raw
    raise ValueError(f"Unsupported resolution: {value!r}. Use 720p, 1080p, or 1440p.")


def ratio_base_size(ratio: str) -> tuple[int, int]:
    r = normalize_ratio(ratio)
    if r == "16:9":
        return 1920, 1080
    if r == "9:16":
        return 1080, 1920
    return 1080, 1080


def canvas_size(*, ratio: str | None, resolution: str | None) -> tuple[int, int]:
    base_w, base_h = ratio_base_size(ratio or "16:9")
    normalized = normalize_resolution(resolution or "1080p")
    scale_map = {"720p": 2.0 / 3.0, "1080p": 1.0, "1440p": 4.0 / 3.0}
    scale = scale_map[normalized]
    return int(round(base_w * scale)), int(round(base_h * scale))
