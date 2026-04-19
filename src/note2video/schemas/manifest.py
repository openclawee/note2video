from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class SlideRecord:
    page: int
    title: str = ""
    image: str = ""
    raw_notes: str = ""
    script: str = ""
    audio: str = ""
    duration_ms: int = 0


@dataclass
class Manifest:
    project_name: str
    input_file: str
    slide_count: int
    ratio: str = "16:9"
    resolution: str = "1080p"
    fps: int = 30
    quality: str = "standard"
    tts_provider: str = ""
    voice: str = ""
    outputs: dict[str, str] = field(default_factory=dict)
    slides: list[SlideRecord] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
