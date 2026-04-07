from __future__ import annotations

from note2video.publish.web import (
    WebPublishError,
    WebPublishPayload,
    profile_root_from_config,
    check_web_auth_status,
    dump_publish_log,
    login_with_browser,
    merge_description_and_topics,
    normalize_topics,
    platform_upload_url,
    publish_with_browser,
)

# Backward-compatible aliases for in-flight GUI imports.
PublishError = WebPublishError
PublishPayload = WebPublishPayload

__all__ = [
    "WebPublishError",
    "WebPublishPayload",
    "check_web_auth_status",
    "dump_publish_log",
    "login_with_browser",
    "merge_description_and_topics",
    "normalize_topics",
    "platform_upload_url",
    "profile_root_from_config",
    "publish_with_browser",
    # compatibility exports
    "PublishError",
    "PublishPayload",
]
