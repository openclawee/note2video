from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable


class WebPublishError(RuntimeError):
    """Raised when browser-driven publish flow fails."""


@dataclass(frozen=True)
class WebPublishPayload:
    platform: str
    method: str
    video_path: str
    title: str
    topics: str
    description: str
    cover_path: str
    visibility: str
    schedule_enabled: bool
    schedule_time: str
    dry_run: bool
    auto_confirm: bool


def platform_upload_url(platform: str) -> str:
    normalized = _normalize_platform(platform)
    if normalized == "douyin":
        return "https://creator.douyin.com/creator-micro/content/upload"
    return "https://channels.weixin.qq.com/platform/post/create"


def normalize_topics(value: str | Iterable[str]) -> list[str]:
    if isinstance(value, str):
        raw = value.replace("\n", ",").replace("，", ",").split(",")
    else:
        raw = [str(x) for x in value]
    dedup: list[str] = []
    seen: set[str] = set()
    for item in raw:
        topic = str(item).strip().lstrip("#").strip()
        if not topic:
            continue
        key = topic.casefold()
        if key in seen:
            continue
        seen.add(key)
        dedup.append(topic)
    return dedup


def merge_description_and_topics(description: str, topics: str | Iterable[str]) -> str:
    base = (description or "").strip()
    tags = normalize_topics(topics)
    if tags:
        tag_text = " ".join(f"#{t}" for t in tags)
        return f"{base}\n{tag_text}".strip()
    return base


def check_web_auth_status(
    *,
    platform: str,
    profile_root: Path,
    headless: bool = True,
) -> dict[str, Any]:
    publisher = _WebPublisher(profile_root=profile_root, headless=headless)
    return publisher.check_auth_status(platform=platform)


def login_with_browser(
    *,
    platform: str,
    profile_root: Path,
    wait_seconds: int = 60,
) -> dict[str, Any]:
    publisher = _WebPublisher(profile_root=profile_root, headless=False)
    return publisher.login_interactive(platform=platform, wait_seconds=max(15, int(wait_seconds)))


def publish_with_browser(
    payload: WebPublishPayload,
    *,
    profile_root: Path,
    headless: bool = False,
    stage_callback: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    if (payload.method or "").strip().lower() != "web":
        raise WebPublishError("Only web publish method is supported right now.")
    publisher = _WebPublisher(profile_root=profile_root, headless=headless)
    return publisher.publish(payload=payload, stage_callback=stage_callback)


def profile_root_from_config(cfg: dict[str, Any] | None = None) -> Path:
    default = Path.home() / ".config" / "note2video" / "publish-web"
    if not cfg:
        return default
    gui = cfg.get("gui") if isinstance(cfg, dict) else {}
    if not isinstance(gui, dict):
        return default
    override = str(gui.get("publish_profile_root") or "").strip()
    if not override:
        return default
    return Path(override)


def dump_publish_log(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


class _WebPublisher:
    def __init__(self, *, profile_root: Path, headless: bool) -> None:
        self.profile_root = Path(profile_root)
        self.profile_root.mkdir(parents=True, exist_ok=True)
        self.headless = bool(headless)

    def check_auth_status(self, *, platform: str) -> dict[str, Any]:
        normalized = _normalize_platform(platform)
        with self._context(platform=normalized, headless=self.headless) as context:
            page = context.new_page()
            page.goto(platform_upload_url(normalized), wait_until="domcontentloaded", timeout=90_000)
            page.wait_for_timeout(1_000)
            logged_in = _is_logged_in(page)
            return {
                "ok": True,
                "platform": normalized,
                "logged_in": logged_in,
                "url": page.url,
            }

    def login_interactive(self, *, platform: str, wait_seconds: int) -> dict[str, Any]:
        normalized = _normalize_platform(platform)
        with self._context(platform=normalized, headless=False) as context:
            page = context.new_page()
            page.goto(platform_upload_url(normalized), wait_until="domcontentloaded", timeout=90_000)
            page.wait_for_timeout(max(15, wait_seconds) * 1_000)
            logged_in = _is_logged_in(page)
            return {
                "ok": True,
                "platform": normalized,
                "logged_in": logged_in,
                "url": page.url,
                "message": "Interactive browser session finished.",
            }

    def publish(
        self,
        *,
        payload: WebPublishPayload,
        stage_callback: Callable[[str], None] | None = None,
    ) -> dict[str, Any]:
        normalized_platform = _normalize_platform(payload.platform)
        video = Path(payload.video_path)
        if not video.exists():
            raise WebPublishError(f"Video file not found: {video}")
        if not payload.title.strip():
            raise WebPublishError("Title is required.")

        def stage(name: str) -> None:
            if stage_callback is not None:
                stage_callback(name)

        stage("open_page")
        with self._context(platform=normalized_platform, headless=self.headless) as context:
            page = context.new_page()
            page.goto(platform_upload_url(normalized_platform), wait_until="domcontentloaded", timeout=90_000)
            page.wait_for_timeout(1_000)

            stage("check_auth")
            if not _is_logged_in(page):
                raise WebPublishError("未检测到登录态，请先执行“登录/刷新登录”。")

            stage("upload_video")
            self._upload_video(page, video)

            stage("fill_form")
            self._fill_publish_form(page=page, payload=payload)

            stage("confirm_publish")
            status = "ready"
            note = "已完成自动填充，未点击发布按钮。"
            if payload.dry_run:
                status = "dry_run"
                note = "dry-run 模式：已上传并填写，未发布。"
            elif payload.auto_confirm:
                if not _click_first_available(
                    page,
                    [
                        "button:has-text('发布')",
                        "button:has-text('立即发布')",
                        "button:has-text('发表')",
                        "div[role='button']:has-text('发布')",
                        "div[role='button']:has-text('发表')",
                    ],
                ):
                    raise WebPublishError("未找到可点击的发布按钮，请检查页面是否变化。")
                status = "submitted"
                note = "已自动点击发布按钮。"

            stage("done")
            return {
                "ok": True,
                "platform": normalized_platform,
                "status": status,
                "note": note,
                "url": page.url,
                "publish_id": _generate_publish_id(normalized_platform),
            }

    def _upload_video(self, page: Any, video: Path) -> None:
        try:
            page.set_input_files("input[type='file']", str(video.resolve()), timeout=60_000)
            page.wait_for_timeout(2_000)
            return
        except Exception:
            pass

        locator = page.locator("input[type='file']")
        if locator.count() <= 0:
            raise WebPublishError("未找到视频上传控件（input[type=file]）。")
        locator.first.set_input_files(str(video.resolve()))
        page.wait_for_timeout(2_000)

    def _fill_publish_form(self, *, page: Any, payload: WebPublishPayload) -> None:
        title_done = _fill_first_available(
            page,
            [
                "textarea[placeholder*='标题']",
                "input[placeholder*='标题']",
                "div[contenteditable='true'][data-placeholder*='标题']",
                "div[contenteditable='true'][placeholder*='标题']",
            ],
            payload.title.strip(),
        )
        if not title_done:
            raise WebPublishError("未找到标题输入框，请检查页面布局或登录状态。")

        merged_desc = merge_description_and_topics(payload.description, payload.topics)
        if merged_desc:
            _fill_first_available(
                page,
                [
                    "textarea[placeholder*='简介']",
                    "textarea[placeholder*='描述']",
                    "div[contenteditable='true'][data-placeholder*='简介']",
                    "div[contenteditable='true'][data-placeholder*='描述']",
                ],
                merged_desc,
            )

    def _context(self, *, platform: str, headless: bool) -> "_ContextScope":
        sync_playwright = _require_playwright()
        profile_dir = self.profile_root / platform
        profile_dir.mkdir(parents=True, exist_ok=True)
        manager = sync_playwright()
        playwright = manager.start()
        try:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(profile_dir.resolve()),
                headless=headless,
                viewport={"width": 1360, "height": 900},
            )
        except Exception as exc:
            try:
                playwright.stop()
            except Exception:
                pass
            raise WebPublishError(f"Failed to start Chromium context: {exc}") from exc
        return _ContextScope(manager=manager, playwright=playwright, context=context)


class _ContextScope:
    def __init__(self, *, manager: Any, playwright: Any, context: Any) -> None:
        self._manager = manager
        self._playwright = playwright
        self._context = context

    def __enter__(self) -> Any:
        return self._context

    def __exit__(self, exc_type, exc, tb) -> None:
        try:
            self._context.close()
        finally:
            try:
                self._playwright.stop()
            finally:
                try:
                    self._manager.__exit__(exc_type, exc, tb)
                except Exception:
                    pass


def _normalize_platform(platform: str) -> str:
    normalized = (platform or "").strip().lower()
    if normalized not in {"douyin", "channels"}:
        raise WebPublishError(f"Unsupported platform: {platform}")
    return normalized


def _generate_publish_id(platform: str) -> str:
    return f"{platform}-{datetime.now().strftime('%Y%m%d%H%M%S')}"


def _is_logged_in(page: Any) -> bool:
    url = (page.url or "").lower()
    blocked = ("login", "passport", "qrcode", "scan")
    if any(word in url for word in blocked):
        return False
    text = ""
    try:
        text = (page.inner_text("body", timeout=2_000) or "").lower()
    except Exception:
        return True
    login_markers = ("扫码登录", "登录", "请登录")
    return not any(marker in text for marker in login_markers)


def _fill_first_available(page: Any, selectors: list[str], value: str) -> bool:
    text = (value or "").strip()
    if not text:
        return False
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if locator.count() <= 0:
                continue
            locator.click(timeout=1_500)
            tag = (locator.evaluate("el => (el.tagName || '').toLowerCase()") or "").strip().lower()
            if tag in {"input", "textarea"}:
                locator.fill(text, timeout=3_000)
            else:
                locator.evaluate(
                    """(el, val) => {
                        if (el.isContentEditable) {
                            el.focus();
                            el.innerText = val;
                            const event = new Event('input', { bubbles: true });
                            el.dispatchEvent(event);
                        }
                    }""",
                    text,
                )
            return True
        except Exception:
            continue
    return False


def _click_first_available(page: Any, selectors: list[str]) -> bool:
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if locator.count() <= 0:
                continue
            locator.click(timeout=3_000)
            return True
        except Exception:
            continue
    return False


def _require_playwright():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise WebPublishError(
            "Playwright is required for web publish. "
            "Install with `python -m pip install -e .[gui]` and run "
            "`python -m playwright install chromium`."
        ) from exc
    return sync_playwright
