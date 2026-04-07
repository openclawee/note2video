from __future__ import annotations

import json
from pathlib import Path

import pytest

from note2video.app.publish_service import (
    PublishService,
    PublishServiceError,
    PublishServiceRequest,
)


def _request(**overrides) -> PublishServiceRequest:
    data = {
        "platform": "douyin",
        "method": "web",
        "out_dir": "/tmp/out",
        "video_source": "manual",
        "manual_video_path": "/tmp/demo.mp4",
        "title": "测试发布",
        "topics": "AI,效率",
        "description": "这是一段描述",
        "cover_path": "",
        "visibility": "public",
        "schedule_enabled": False,
        "schedule_time": "",
        "dry_run": True,
        "auto_confirm": False,
    }
    data.update(overrides)
    return PublishServiceRequest(**data)


def test_build_payload_manual_video_missing(tmp_path: Path) -> None:
    service = PublishService(profile_root=tmp_path / "profiles")
    req = _request(
        out_dir=str(tmp_path / "out"),
        manual_video_path=str(tmp_path / "missing.mp4"),
    )
    with pytest.raises(PublishServiceError):
        service.build_payload(req)


def test_build_payload_from_manifest(tmp_path: Path) -> None:
    out_dir = tmp_path / "out"
    (out_dir / "video").mkdir(parents=True)
    video = out_dir / "video" / "output.mp4"
    video.write_bytes(b"demo")
    (out_dir / "manifest.json").write_text(
        json.dumps({"outputs": {"video": "video/output.mp4"}}, ensure_ascii=False),
        encoding="utf-8",
    )

    service = PublishService(profile_root=tmp_path / "profiles")
    req = _request(
        out_dir=str(out_dir),
        video_source="last_build",
        manual_video_path="",
        topics="AI,教育,AI",
    )
    payload = service.build_payload(req)
    assert payload.video_path == str(video)
    assert payload.title == "测试发布"
    assert payload.topics == "AI,教育,AI"


def test_build_payload_requires_title(tmp_path: Path) -> None:
    video = tmp_path / "demo.mp4"
    video.write_bytes(b"v")
    service = PublishService(profile_root=tmp_path / "profiles")
    req = _request(
        out_dir=str(tmp_path / "out"),
        manual_video_path=str(video),
        title="",
    )
    with pytest.raises(PublishServiceError):
        service.build_payload(req)
