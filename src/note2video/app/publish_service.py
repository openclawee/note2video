from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from note2video.publish import (
    WebPublishPayload,
    check_web_auth_status,
    dump_publish_log,
    login_with_browser,
    platform_upload_url,
    profile_root_from_config,
    publish_with_browser,
)
from note2video.user_config import load_user_config, normalize_user_config


class PublishServiceError(RuntimeError):
    """Raised when publish service validation or execution fails."""


@dataclass(frozen=True)
class PublishFormInput:
    platform: str
    method: str
    video_source: str
    out_dir: str
    manual_video_path: str
    title: str
    topics: str
    description: str
    cover_path: str
    visibility: str
    schedule_enabled: bool
    schedule_time: str
    dry_run: bool
    auto_confirm: bool


@dataclass(frozen=True)
class PublishExecution:
    ui_payload: dict[str, Any]
    web_payload: WebPublishPayload
    profile_root: Path
    log_path: Path


def resolve_profile_root() -> Path:
    cfg = normalize_user_config(load_user_config())
    return profile_root_from_config(cfg)


def build_publish_execution(form: PublishFormInput, *, profile_root: Path | None = None) -> PublishExecution:
    video_path = _resolve_video_path(
        video_source=form.video_source,
        out_dir=form.out_dir,
        manual_video_path=form.manual_video_path,
    )
    title = (form.title or "").strip()
    if not title:
        raise PublishServiceError("标题不能为空。")

    cover_text = (form.cover_path or "").strip()
    if cover_text:
        cover_path = Path(cover_text.strip('"'))
        if not cover_path.exists():
            raise PublishServiceError("封面图路径不存在。")

    out_dir_path = Path((form.out_dir or "dist").strip().strip('"') or "dist")
    ui_payload = {
        "platform": (form.platform or "douyin").strip(),
        "method": (form.method or "web").strip(),
        "video_path": str(video_path),
        "title": title,
        "topics": (form.topics or "").strip(),
        "description": (form.description or "").strip(),
        "cover_path": cover_text,
        "visibility": (form.visibility or "public").strip(),
        "schedule_enabled": bool(form.schedule_enabled),
        "schedule_time": (form.schedule_time or "").strip(),
        "dry_run": bool(form.dry_run),
        "auto_confirm": bool(form.auto_confirm),
    }
    web_payload = WebPublishPayload(
        platform=ui_payload["platform"],
        method=ui_payload["method"],
        video_path=ui_payload["video_path"],
        title=ui_payload["title"],
        topics=ui_payload["topics"],
        description=ui_payload["description"],
        cover_path=ui_payload["cover_path"],
        visibility=ui_payload["visibility"],
        schedule_enabled=bool(ui_payload["schedule_enabled"]),
        schedule_time=ui_payload["schedule_time"],
        dry_run=bool(ui_payload["dry_run"]),
        auto_confirm=bool(ui_payload["auto_confirm"]),
    )
    return PublishExecution(
        ui_payload=ui_payload,
        web_payload=web_payload,
        profile_root=Path(profile_root) if profile_root is not None else resolve_profile_root(),
        log_path=out_dir_path / "logs" / "publish.log",
    )


def login_via_web(*, platform: str, method: str, wait_seconds: int = 60) -> dict[str, Any]:
    if (method or "").strip().lower() != "web":
        raise PublishServiceError("当前仅支持 web 方式登录。")
    return login_with_browser(
        platform=(platform or "douyin").strip(),
        profile_root=resolve_profile_root(),
        wait_seconds=wait_seconds,
    )


def check_auth_via_web(*, platform: str, method: str, headless: bool = True) -> dict[str, Any]:
    if (method or "").strip().lower() != "web":
        return {"ok": True, "platform": platform, "logged_in": False, "status": "未实现", "url": ""}
    return check_web_auth_status(
        platform=(platform or "douyin").strip(),
        profile_root=resolve_profile_root(),
        headless=headless,
    )


def submit_publish(
    execution: PublishExecution,
    *,
    stage_callback: Callable[[str], None] | None = None,
    headless: bool = False,
) -> dict[str, Any]:
    return publish_with_browser(
        execution.web_payload,
        profile_root=execution.profile_root,
        headless=headless,
        stage_callback=stage_callback,
    )


def write_publish_record(
    *,
    log_path: Path,
    payload: dict[str, Any],
    result: dict[str, Any] | None = None,
    error: str | None = None,
) -> None:
    record = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "payload": payload,
        "result": result or {},
        "error": error or "",
    }
    dump_publish_log(log_path, record)


def last_publish_status_text(payload: dict[str, Any]) -> tuple[str, str]:
    platform = str(payload.get("platform") or "douyin")
    method = str(payload.get("method") or "web")
    line = "[status] 最近一次任务：" + f"platform={platform}, title={payload.get('title')}, method={method}"
    return line, f"[status] 入口地址：{platform_upload_url(platform)}"


def _resolve_video_path(*, video_source: str, out_dir: str, manual_video_path: str) -> Path:
    source = (video_source or "last_build").strip()
    if source == "manual":
        manual = Path((manual_video_path or "").strip().strip('"'))
        if not manual.exists():
            raise PublishServiceError("手动选择的视频文件不存在。")
        return manual

    out_dir_path = Path((out_dir or "dist").strip().strip('"') or "dist")
    manifest_path = out_dir_path / "manifest.json"
    if not manifest_path.exists():
        raise PublishServiceError("找不到 manifest.json：Build 输出目录中没有 manifest.json 文件。请先执行 Build 或选择手动指定视频文件。")
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise PublishServiceError(f"读取 manifest.json 失败：{exc}") from exc
    rel_video = str((manifest.get("outputs") or {}).get("video") or "").strip()
    if not rel_video:
        raise PublishServiceError("manifest.json 中没有 video 输出，请先运行 Build。")
    candidate = Path(rel_video)
    video_path = candidate if candidate.is_absolute() else (out_dir_path / candidate)
    if not video_path.exists():
        raise PublishServiceError(f"视频文件不存在：{video_path}")
    return video_path


@dataclass(frozen=True)
class PublishRequest(PublishFormInput):
    """UI-agnostic publish request for adapters like GUI/CLI/Electron."""


PublishServiceRequest = PublishRequest


class PublishService:
    def __init__(self, *, profile_root: Path | None = None) -> None:
        self.profile_root = Path(profile_root) if profile_root is not None else resolve_profile_root()

    def build_payload(self, request: PublishRequest) -> WebPublishPayload:
        return self.build_execution(request).web_payload

    def build_execution(self, request: PublishRequest) -> PublishExecution:
        return build_publish_execution(request, profile_root=self.profile_root)

    def login(self, *, platform: str, method: str, wait_seconds: int = 60) -> dict[str, Any]:
        if (method or "").strip().lower() != "web":
            raise PublishServiceError("当前仅支持 web 方式登录。")
        return login_with_browser(
            platform=(platform or "douyin").strip(),
            profile_root=self.profile_root,
            wait_seconds=wait_seconds,
        )

    def check_auth_status(self, *, platform: str, method: str, headless: bool = True) -> dict[str, Any]:
        if (method or "").strip().lower() != "web":
            return {"ok": True, "platform": platform, "logged_in": False, "status": "未实现", "url": ""}
        return check_web_auth_status(
            platform=(platform or "douyin").strip(),
            profile_root=self.profile_root,
            headless=headless,
        )

    def publish(
        self,
        request: PublishRequest,
        *,
        stage_callback: Callable[[str], None] | None = None,
        headless: bool = False,
    ) -> tuple[PublishExecution, dict[str, Any]]:
        execution = self.build_execution(request)
        result = submit_publish(execution, stage_callback=stage_callback, headless=headless)
        return execution, result

    def publish_execution(
        self,
        execution: PublishExecution,
        *,
        stage_callback: Callable[[str], None] | None = None,
        headless: bool = False,
    ) -> dict[str, Any]:
        return submit_publish(execution, stage_callback=stage_callback, headless=headless)


def build_publish_request(request: PublishRequest) -> PublishExecution:
    return PublishService().build_execution(request)


def perform_publish_login(*, platform: str, method: str, wait_seconds: int = 60) -> dict[str, Any]:
    return PublishService().login(platform=platform, method=method, wait_seconds=wait_seconds)


def check_publish_auth_status(*, platform: str, method: str, headless: bool = True) -> dict[str, Any]:
    return PublishService().check_auth_status(platform=platform, method=method, headless=headless)


def perform_publish(
    request: PublishRequest,
    *,
    stage_callback: Callable[[str], None] | None = None,
    headless: bool = False,
) -> tuple[PublishExecution, dict[str, Any]]:
    return PublishService().publish(request, stage_callback=stage_callback, headless=headless)


def perform_publish_execution(
    execution: PublishExecution,
    *,
    stage_callback: Callable[[str], None] | None = None,
    headless: bool = False,
) -> dict[str, Any]:
    return PublishService().publish_execution(execution, stage_callback=stage_callback, headless=headless)


def last_publish_status_lines(payload: dict[str, Any]) -> tuple[str, str]:
    return last_publish_status_text(payload)

