from __future__ import annotations

import asyncio
import base64
import json
import os
import re
import subprocess
import time
import urllib.error
import urllib.request
import wave
from contextlib import closing
from pathlib import Path
from typing import Any

from note2video.text_segmentation import split_sentences, split_sentences_with_pauses
from note2video.user_config import load_user_config, normalize_user_config, tts_provider_config


class VoiceGenerationError(RuntimeError):
    """Raised when TTS generation fails."""


_user_config_snapshot: dict[str, Any] | None = None


def invalidate_user_config_cache() -> None:
    """Call after saving user config so the next TTS request reloads from disk."""
    global _user_config_snapshot
    _user_config_snapshot = None


def _user_cfg() -> dict[str, Any]:
    global _user_config_snapshot
    if _user_config_snapshot is None:
        _user_config_snapshot = normalize_user_config(load_user_config())
    return _user_config_snapshot


def _minimax_fixed_base_url(provider_name: str) -> str:
    normalized = (provider_name or "").strip().lower()
    if normalized in {"minimax_cn", "mimax_cn", "minimax-china", "minimax-cn"}:
        return "https://api.minimax.chat"
    if normalized in {"minimax_global", "mimax_global", "minimax-intl", "minimax-global"}:
        return "https://api.minimaxi.chat"
    raise ValueError(f"Unsupported MiniMax provider: {provider_name!r}")


def _provider_api_key(provider_name: str) -> str:
    key = (os.getenv("NOTE2VIDEO_MINIMAX_API_KEY") or os.getenv("MINIMAX_API_KEY") or "").strip()
    if key:
        return key
    return str(tts_provider_config(_user_cfg(), provider_name).get("api_key") or "").strip()


def _provider_model(provider_name: str, explicit: str | None) -> str:
    for candidate in (
        (explicit or "").strip(),
        (os.getenv("NOTE2VIDEO_MINIMAX_MODEL") or "").strip(),
        str(tts_provider_config(_user_cfg(), provider_name).get("model") or "").strip(),
    ):
        if candidate:
            return candidate
    return "speech-2.8-hd"


def _provider_timeout_s(provider_name: str, *, for_list: bool) -> float:
    env = (os.getenv("NOTE2VIDEO_MINIMAX_TIMEOUT_S") or "").strip()
    if env:
        return float(env)
    raw = tts_provider_config(_user_cfg(), provider_name).get("timeout_s")
    if raw is not None and str(raw).strip() != "":
        return float(raw)
    return 30.0 if for_list else 60.0


def list_available_voices(
    *,
    provider_name: str = "edge",
    keyword: str = "",
    minimax_base_url: str | None = None,
) -> list[dict[str, Any]]:
    normalized = provider_name.strip().lower() or "edge"
    if normalized == "edge":
        return _list_edge_voices(keyword=keyword)
    if normalized in {"volcengine", "volc", "huoshan", "doubao", "doubao_tts"}:
        return _list_volcengine_voices(keyword=keyword)
    raise ValueError(f"Unsupported TTS provider: {provider_name}")


def _clamp_tts_rate(rate: float) -> float:
    """Speech rate multiplier; applied during synthesis so WAV durations match subtitles."""
    try:
        value = float(rate)
    except (TypeError, ValueError) as exc:
        raise VoiceGenerationError(f"Invalid tts_rate: {rate!r}") from exc
    if value < 0.5 or value > 2.0:
        raise VoiceGenerationError("tts_rate must be between 0.5 and 2.0.")
    return value


def generate_voice_assets(
    input_json: str,
    output_dir: str,
    *,
    provider_name: str = "edge",
    voice_id: str = "",
    tts_rate: float = 1.0,
    minimax_base_url: str | None = None,
) -> dict[str, Any]:
    input_path = Path(input_json)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    project_dir = _resolve_project_dir(input_path, output_dir)
    scripts = _load_scripts(input_path)
    audio_dir = project_dir / "audio"
    logs_dir = project_dir / "logs"
    manifest_path = project_dir / "manifest.json"
    timings_path = audio_dir / "timings.json"

    audio_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)

    rate_mult = _clamp_tts_rate(tts_rate)
    provider = _create_provider(
        provider_name=provider_name,
        voice_id=voice_id,
        tts_rate=rate_mult,
        minimax_base_url=minimax_base_url,
    )
    generated: list[dict[str, Any]] = []
    timing_segments: list[dict[str, Any]] = []
    global_cursor_ms = 0

    for item in scripts:
        page = int(item["page"])
        text = item.get("script", "")
        output_file = audio_dir / f"{page:03d}.wav"
        sentence_files: list[Path] = []

        if text.strip():
            sentence_cursor_ms = global_cursor_ms
            parts = _split_tts_chunks_with_pauses(text)
            if not parts:
                parts = [(text.strip(), 0)]

            for idx, (sentence, _pause_ms) in enumerate(parts, start=1):
                sentence_file = audio_dir / f"{page:03d}.{idx:02d}.wav"
                try:
                    provider.synthesize_to_file(text=sentence, output_file=sentence_file)
                except Exception as exc:
                    snippet = _safe_snippet(sentence, limit=120)
                    raise VoiceGenerationError(
                        "TTS synthesis failed.\n"
                        f"- page: {page}\n"
                        f"- sentence_index: {idx}\n"
                        f"- chars: {len(sentence)}\n"
                        f"- text: {snippet}\n"
                        f"- provider: {provider_name}\n"
                        f"- voice: {voice_id or 'default'}\n"
                        f"- tts_rate: {rate_mult}\n"
                        f"- cause: {type(exc).__name__}: {exc}"
                    ) from exc
                sentence_duration_ms = _read_wav_duration_ms(sentence_file)
                sentence_files.append(sentence_file)
                timing_segments.append(
                    {
                        "page": page,
                        "index": len(timing_segments) + 1,
                        "sentence_index": idx,
                        "text": sentence,
                        "start_ms": sentence_cursor_ms,
                        "end_ms": sentence_cursor_ms + sentence_duration_ms,
                        "duration_ms": sentence_duration_ms,
                    }
                )
                sentence_cursor_ms += sentence_duration_ms

            _merge_wav_files(input_files=sentence_files, output_file=output_file)
            duration_ms = _read_wav_duration_ms(output_file)
            global_cursor_ms += duration_ms
        else:
            _write_silence_wav(output_file, duration_ms=300)
            duration_ms = 300
            timing_segments.append(
                {
                    "page": page,
                    "index": len(timing_segments) + 1,
                    "sentence_index": 1,
                    "text": "",
                    "start_ms": global_cursor_ms,
                    "end_ms": global_cursor_ms + duration_ms,
                    "duration_ms": duration_ms,
                }
            )
            global_cursor_ms += duration_ms

        generated.append(
            {
                "page": page,
                "audio": f"audio/{output_file.name}",
                "duration_ms": duration_ms,
            }
        )

    merged_path = audio_dir / "merged.wav"
    _merge_wav_files(
        input_files=[audio_dir / f"{item['page']:03d}.wav" for item in scripts],
        output_file=merged_path,
    )
    timings_path.write_text(
        json.dumps({"segments": timing_segments}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    _update_manifest(
        manifest_path=manifest_path,
        generated=generated,
        provider_name=provider_name,
        voice_id=voice_id,
        timings_path="audio/timings.json",
        tts_rate=rate_mult,
    )

    (logs_dir / "voice.log").write_text(
        f"provider: {provider_name}\nvoice: {voice_id or 'default'}\ntts_rate: {rate_mult}\nslides: {len(generated)}\n",
        encoding="utf-8",
    )

    return {
        "audio_dir": str(audio_dir),
        "merged_audio": str(merged_path),
        "timings": str(timings_path),
        "slide_count": len(generated),
        "provider": provider_name,
        "voice": voice_id,
        "tts_rate": rate_mult,
    }


def _resolve_project_dir(input_path: Path, output_dir: str) -> Path:
    if output_dir and output_dir != "./dist":
        return Path(output_dir)
    return input_path.parent.parent


def _load_scripts(input_path: Path) -> list[dict[str, Any]]:
    payload = json.loads(input_path.read_text(encoding="utf-8"))
    slides = payload.get("slides", [])
    if not isinstance(slides, list):
        raise ValueError("Script file is missing a valid 'slides' array.")
    return slides


def _create_provider(
    *,
    provider_name: str,
    voice_id: str,
    tts_rate: float = 1.0,
    minimax_base_url: str | None = None,
):
    normalized = provider_name.strip().lower() or "pyttsx3"
    rate_mult = _clamp_tts_rate(tts_rate)
    if normalized == "edge":
        return EdgeTTSProvider(
            voice_id=voice_id or "zh-CN-XiaoxiaoNeural",
            tts_rate=rate_mult,
        )
    if normalized in {"volcengine", "volc", "huoshan", "doubao", "doubao_tts"}:
        return VolcengineTTSProvider(
            voice_id=voice_id or "BV700_streaming",
            tts_rate=rate_mult,
        )
    raise ValueError(f"Unsupported TTS provider: {provider_name}")


def synthesize_preview_sample(
    *,
    provider_name: str,
    voice_id: str,
    tts_rate: float,
    text: str,
    output_file: Path,
    minimax_base_url: str | None = None,
) -> None:
    """Synthesize a short clip for GUI preview (same engines as the build pipeline)."""
    normalized = (provider_name or "edge").strip().lower() or "edge"
    vid = (voice_id or "").strip()
    if not vid:
        if normalized == "edge":
            vid = "zh-CN-XiaoxiaoNeural"
        elif normalized in {"volcengine", "volc", "huoshan", "doubao", "doubao_tts"}:
            vid = "BV700_streaming"
    provider = _create_provider(
        provider_name=provider_name,
        voice_id=vid,
        tts_rate=tts_rate,
        minimax_base_url=minimax_base_url,
    )
    output_file.parent.mkdir(parents=True, exist_ok=True)
    provider.synthesize_to_file(text=text, output_file=output_file)


class Pyttsx3Provider:
    def __init__(self, *, voice_id: str = "", tts_rate: float = 1.0) -> None:
        self.voice_id = voice_id
        self.tts_rate = _clamp_tts_rate(tts_rate)

    def synthesize_to_file(self, *, text: str, output_file: Path) -> None:
        try:
            import pyttsx3
        except ImportError as exc:  # pragma: no cover
            raise VoiceGenerationError("pyttsx3 is required for local voice generation.") from exc

        try:
            engine = pyttsx3.init()
            if self.voice_id:
                engine.setProperty("voice", self.voice_id)
            base = engine.getProperty("rate")
            try:
                base_wpm = float(base) if base is not None else 200.0
            except (TypeError, ValueError):
                base_wpm = 200.0
            new_wpm = int(round(base_wpm * self.tts_rate))
            new_wpm = max(50, min(400, new_wpm))
            engine.setProperty("rate", new_wpm)
            engine.save_to_file(text, str(output_file))
            engine.runAndWait()
            engine.stop()
        except Exception as exc:  # pragma: no cover - runtime specific
            raise VoiceGenerationError(f"pyttsx3 synthesis failed: {exc}") from exc


class EdgeTTSProvider:
    def __init__(self, *, voice_id: str, tts_rate: float = 1.0) -> None:
        self.voice_id = voice_id
        self.tts_rate = _clamp_tts_rate(tts_rate)

    def synthesize_to_file(self, *, text: str, output_file: Path) -> None:
        try:
            import edge_tts
        except ImportError as exc:  # pragma: no cover
            raise VoiceGenerationError("edge-tts is required for Edge voice generation.") from exc

        text = _sanitize_tts_text(text)
        temp_audio = output_file.with_suffix(".edge.mp3")
        pct = int(round((self.tts_rate - 1.0) * 100))
        pct = max(-50, min(100, pct))
        rate_str = f"{pct:+d}%"

        async def _run(rate: str) -> None:
            communicate = edge_tts.Communicate(text=text, voice=self.voice_id, rate=rate)
            await communicate.save(str(temp_audio))

        last_exc: Exception | None = None
        # Edge TTS can intermittently return empty audio on some networks. We retry a few times with
        # exponential backoff, and also fall back to a neutral rate to reduce parameter sensitivity.
        rate_candidates = [rate_str, "+0%"] if rate_str != "+0%" else ["+0%"]
        max_attempts = int(os.getenv("NOTE2VIDEO_EDGE_TTS_RETRIES", "4") or "4")
        max_attempts = max(1, min(max_attempts, 8))
        attempt = 0
        for rate in rate_candidates:
            for _ in range(max_attempts):
                attempt += 1
                sleep_s = min(3.0, 0.35 * (2 ** max(0, attempt - 1)))
            try:
                    asyncio.run(_run(rate))
                    _convert_audio_to_wav(temp_audio=temp_audio, output_file=output_file)
                    return
            except Exception as exc:  # pragma: no cover - runtime/network specific
                last_exc = exc
                # Some networks intermittently fail to return audio; best-effort cleanup then retry.
                try:
                    if temp_audio.exists():
                        temp_audio.unlink()
                except OSError:
                    pass
                # Backoff before retrying. Keep it short so build doesn't feel "hung".
                time.sleep(sleep_s)
                continue
        msg = (
            f"edge-tts synthesis failed: {last_exc}\n"
            f"- voice: {self.voice_id}\n"
            f"- rate: {rate_str}\n"
            f"- chars: {len(text)}\n"
            "Hints:\n"
            "- If preview works but build fails, it may be intermittent network/proxy stability or long text.\n"
            "- Try reducing tts_rate to 1.0, or try a different voice ID.\n"
            "- If you're behind a proxy/VPN/firewall, ensure the Python subprocess inherits proxy env vars.\n"
        )
        raise VoiceGenerationError(msg) from last_exc
        # Should never reach here.
        if last_exc is not None:
            raise VoiceGenerationError(f"edge-tts synthesis failed: {last_exc}") from last_exc


def _sanitize_tts_text(text: str) -> str:
    """
    Best-effort sanitization for cloud TTS engines.

    - Remove NUL/control characters that can break WS payloads.
    - Normalize CRLF to LF.
    - Collapse excessive whitespace.
    """
    s = str(text or "")
    s = s.replace("\r", "\n")
    # Drop ASCII control chars except \n and \t.
    s = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", s)
    # Avoid pathological whitespace; keep newlines as they may indicate pauses.
    s = re.sub(r"[ \t]{2,}", " ", s)
    return s.strip()


def _safe_snippet(text: str, *, limit: int = 120) -> str:
    s = (text or "").replace("\r", " ").replace("\n", " ").strip()
    if len(s) <= limit:
        return s
    return s[: max(0, limit - 1)] + "…"


def _split_tts_chunks_with_pauses(text: str) -> list[tuple[str, int]]:
    """
    Split into reasonably sized chunks for TTS stability.

    Some providers (notably edge-tts on certain networks) are more likely to fail on long sentences.
    We first split by punctuation/newlines, then further split overlong chunks by commas/whitespace.
"""
    normalized = _sanitize_tts_text(text)
    parts = _split_sentences(normalized)
    if not parts:
        return []

    # Conservative default; can be overridden for power users.
    try:
        max_chars = int(os.getenv("NOTE2VIDEO_TTS_MAX_CHARS", "240") or "240")
    except ValueError:
        max_chars = 240
    max_chars = max(80, min(max_chars, 800))

    out: list[tuple[str, int]] = []
    for chunk in parts:
        c = chunk.strip()
        if not c:
            continue
        if len(c) <= max_chars:
            out.append((c, 0))
            continue

        # First try to split by commas/顿号; keep punctuation.
        subparts = re.split(r"(?<=[，,、])\s*", c)
        buf = ""
        for sp in subparts:
            sp = sp.strip()
            if not sp:
                continue
            candidate = (buf + (" " if buf and not buf.endswith(("\n", " ")) else "") + sp).strip()
            if not buf:
                buf = sp
                continue
            if len(candidate) <= max_chars:
                buf = candidate
                continue
            out.append((buf.strip(), 0))
            buf = sp
        if buf.strip():
            out.append((buf.strip(), 0))

    # Final pass: hard split any remaining very long chunks (no punctuation/commas).
    final: list[tuple[str, int]] = []
    for chunk, pause_ms in out:
        if len(chunk) <= max_chars:
            final.append((chunk, pause_ms))
            continue
        start = 0
        while start < len(chunk):
            piece = chunk[start : start + max_chars].strip()
            start += max_chars
            if not piece:
                continue
            final.append((piece, 0))
    return final


class MiniMaxTTSProvider:
    """
    MiniMax (mimax) Text-to-Speech HD via HTTP API.

    We keep this provider stateless and API-driven so it can be swapped for other vendors later.
    """

    def __init__(
        self,
        *,
        provider_name: str,
        voice_id: str,
        tts_rate: float = 1.0,
        model: str | None = None,
        api_base_url: str | None = None,
    ) -> None:
        self.provider_name = provider_name.strip().lower()
        self.voice_id = voice_id
        self.tts_rate = _clamp_tts_rate(tts_rate)
        self.model = _provider_model(self.provider_name, model)
        self._api_base_url = (api_base_url or "").strip().rstrip("/") or _minimax_fixed_base_url(self.provider_name)

    def synthesize_to_file(self, *, text: str, output_file: Path) -> None:
        api_key = _provider_api_key(self.provider_name)
        if not api_key:
            raise VoiceGenerationError(
                "MiniMax API key missing. Set NOTE2VIDEO_MINIMAX_API_KEY or MINIMAX_API_KEY, "
                "or save it in the GUI settings (user config file)."
            )

        base_url = self._api_base_url
        url = f"{base_url}/v1/t2a_v2"

        # For simplicity we request mp3 (hex encoded) and convert to our project WAV format.
        req = {
            "model": self.model,
            "text": text,
            "stream": False,
            "output_format": "hex",
            "language_boost": "auto",
            "voice_setting": {
                "voice_id": self.voice_id,
                "speed": float(self.tts_rate),
                "vol": 1,
                "pitch": 0,
            },
            "audio_setting": {
                "sample_rate": 22050,
                "bitrate": 128000,
                "format": "mp3",
                "channel": 1,
            },
        }

        payload = _http_post_json(
            url,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            body=req,
            timeout_s=_provider_timeout_s(self.provider_name, for_list=False),
        )

        base_resp = (payload or {}).get("base_resp") or {}
        status_code = base_resp.get("status_code", 0)
        if status_code not in (0, "0"):
            status_msg = base_resp.get("status_msg") or "MiniMax TTS request failed."
            raise VoiceGenerationError(f"MiniMax TTS failed: {status_code} {status_msg}".strip())

        data = (payload or {}).get("data") or {}
        hex_audio = data.get("audio")
        if not isinstance(hex_audio, str) or not hex_audio.strip():
            raise VoiceGenerationError("MiniMax TTS returned empty audio.")

        temp_audio = output_file.with_suffix(".minimax.mp3")
        try:
            temp_audio.write_bytes(bytes.fromhex(hex_audio))
            _convert_audio_to_wav(temp_audio=temp_audio, output_file=output_file)
        except ValueError as exc:
            raise VoiceGenerationError("MiniMax TTS returned invalid hex audio.") from exc
        finally:
            if temp_audio.exists():
                temp_audio.unlink()


def _volc_appid() -> str:
    env = (os.getenv("NOTE2VIDEO_VOLC_APPID") or "").strip()
    if env:
        return env
    return str(tts_provider_config(_user_cfg(), "volcengine").get("appid") or "").strip()


def _volc_token() -> str:
    env = (os.getenv("NOTE2VIDEO_VOLC_TOKEN") or os.getenv("VOLC_TOKEN") or "").strip()
    if env:
        return env
    return str(tts_provider_config(_user_cfg(), "volcengine").get("token") or "").strip()


def _volc_cluster() -> str:
    env = (os.getenv("NOTE2VIDEO_VOLC_CLUSTER") or "").strip()
    if env:
        return env
    raw = str(tts_provider_config(_user_cfg(), "volcengine").get("cluster") or "").strip()
    return raw or "volcano_tts"


# Default to Doubao TTS 2.0 (V3) unidirectional endpoint.
# Users can override via NOTE2VIDEO_VOLC_TTS_URL / NOTE2VIDEO_DOUBAO_TTS_URL or config tts.providers.volcengine.base_url.
_DEFAULT_VOLC_TTS_URL = "https://openspeech.bytedance.com/api/v3/tts/unidirectional"


def _volc_base_url() -> str:
    for key in ("NOTE2VIDEO_VOLC_TTS_URL", "NOTE2VIDEO_DOUBAO_TTS_URL"):
        raw = (os.getenv(key) or "").strip().rstrip("/")
        if raw:
            return raw
    cfg = str(tts_provider_config(_user_cfg(), "volcengine").get("base_url") or "").strip().rstrip("/")
    return cfg or _DEFAULT_VOLC_TTS_URL


def _volc_timeout_s() -> float:
    env = (os.getenv("NOTE2VIDEO_VOLC_TIMEOUT_S") or "").strip()
    if env:
        return float(env)
    raw = tts_provider_config(_user_cfg(), "volcengine").get("timeout_s")
    if raw is not None and str(raw).strip() != "":
        return float(raw)
    return 60.0


def _volc_resource_id() -> str:
    """
    Some Doubao/Volcengine V3 endpoints require a Resource-Id header.

    Example (varies by console/service): seed-tts-2.0
    """
    for key in ("NOTE2VIDEO_VOLC_RESOURCE_ID", "NOTE2VIDEO_DOUBAO_RESOURCE_ID"):
        raw = (os.getenv(key) or "").strip()
        if raw:
            return raw
    return str(tts_provider_config(_user_cfg(), "volcengine").get("resource_id") or "").strip()


def _volc_tts_success(code: Any) -> bool:
    if code is None:
        return True
    if code in (0, "0", 3000, "3000"):
        return True
    try:
        return int(code) in (0, 3000)
    except (TypeError, ValueError):
        return str(code).strip().lower() in ("success", "ok")


def _extract_volcengine_audio_b64(payload: dict[str, Any]) -> str:
    data = payload.get("data")
    if isinstance(data, str) and data.strip():
        return data.strip()
    if isinstance(data, dict):
        for key in ("audio", "content", "binary", "b64"):
            val = data.get(key)
            if isinstance(val, str) and val.strip():
                return val.strip()
    return ""


class VolcengineTTSProvider:
    """
    火山引擎「豆包」在线语音合成（OpenSpeech HTTP TTS）。

    默认地址：https://openspeech.bytedance.com/api/v1/tts
    鉴权：控制台 AppID + Token；Token 同时放在 JSON 与 ``Authorization: Bearer`` 头中。
    """

    def __init__(self, *, voice_id: str, tts_rate: float = 1.0) -> None:
        self.voice_id = voice_id
        self.tts_rate = _clamp_tts_rate(tts_rate)

    def synthesize_to_file(self, *, text: str, output_file: Path) -> None:
        appid = _volc_appid()
        token = _volc_token()
        if not appid or not token:
            raise VoiceGenerationError(
                "Volcengine / 豆包凭据未配置。请设置环境变量 NOTE2VIDEO_VOLC_APPID、NOTE2VIDEO_VOLC_TOKEN（或 VOLC_TOKEN），"
                "或在菜单「设置 → TTS Provider…」中选择火山引擎/豆包并保存 App ID 与 Access Token。"
            )

        import uuid

        temp_audio = output_file.with_suffix(".volc.mp3")
        url = _volc_base_url()
        is_v3_streaming = "/api/v3/tts/unidirectional" in url
        # Volc speed ratio is typically 0.8-1.2; map our multiplier directly but clamp.
        speed_ratio = max(0.5, min(2.0, float(self.tts_rate)))
        body = {
            "app": {"appid": appid, "token": token, "cluster": _volc_cluster()},
            "user": {"uid": "note2video"},
            "audio": {
                "voice_type": self.voice_id,
                "encoding": "mp3",
                "speed_ratio": speed_ratio,
                "volume_ratio": 1.0,
                "pitch_ratio": 1.0,
            },
            "request": {
                "reqid": str(uuid.uuid4()),
                "text": text,
                "text_type": "plain",
                "operation": "query",
                "with_frontend": 1,
                "frontend_type": "unitTson",
            },
        }
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            # Some V3 docs/examples use X-Api-* style headers.
            "X-Api-App-Id": appid,
            "X-Api-Access-Key": token,
        }

        if is_v3_streaming:
            rid = _volc_resource_id()
            if not rid:
                raise VoiceGenerationError(
                    "豆包/火山 V3 TTS 需要 Resource-Id，但当前未配置。"
                    "请在「设置 → TTS Provider…」里填写 Resource-Id（或设置 NOTE2VIDEO_VOLC_RESOURCE_ID）。"
                )
            headers["Resource-Id"] = rid
            headers["X-Api-Resource-Id"] = rid
            audio_bytes = _http_post_sse_collect_audio_bytes(
                url,
                headers=headers,
                body=body,
                timeout_s=_volc_timeout_s(),
            )
            if not audio_bytes:
                raise VoiceGenerationError("Volcengine V3 TTS returned empty audio.")
            temp_audio.write_bytes(audio_bytes)
        else:
            payload = _http_post_json(
                url,
                headers=headers,
                body=body,
                timeout_s=_volc_timeout_s(),
            )
            code = payload.get("code")
            if not _volc_tts_success(code):
                raise VoiceGenerationError(
                    f"Volcengine TTS failed: code={code!r} message={payload.get('message')!r}"
                )
            b64 = _extract_volcengine_audio_b64(payload)
            if not b64:
                raise VoiceGenerationError("Volcengine TTS returned empty audio.")
            try:
                temp_audio.write_bytes(base64.b64decode(b64))
            except Exception as exc:
                raise VoiceGenerationError("Volcengine TTS returned invalid base64 audio.") from exc
        try:
            _convert_audio_to_wav(temp_audio=temp_audio, output_file=output_file)
        finally:
            if temp_audio.exists():
                try:
                    temp_audio.unlink()
                except OSError:
                    pass


# Curated voice_type samples for 豆包/火山「在线语音合成」. There is no stable list-voices HTTP API
# wired here; the console / 文档「发音人参数列表」 is authoritative. See also volcengine.com/docs/6561/97465 .
_VOLC_VOICE_TYPE_SAMPLES: list[tuple[str, str, str, str]] = [
    # (voice_type, locale, gender, display_name)
    ("BV700_streaming", "zh-CN", "", "BV700_streaming（文档常见）"),
    ("BV701_streaming", "zh-CN", "", "BV701_streaming（文档常见）"),
    ("BV700_V2_streaming", "zh-CN", "", "灿灿 2.0 · BV700_V2_streaming"),
    ("BV705_streaming", "zh-CN", "", "阳阳 · BV705_streaming"),
    ("BV123_streaming", "zh-CN", "", "阳光青年 · BV123_streaming"),
    ("BV120_streaming", "zh-CN", "", "反卷青年 · BV120_streaming"),
    ("BV119_streaming", "zh-CN", "", "通用女婿 · BV119_streaming"),
    ("BV115_streaming", "zh-CN", "", "古风少御 · BV115_streaming"),
    ("BV107_streaming", "zh-CN", "", "霸道少爷 · BV107_streaming"),
    ("BV100_streaming", "zh-CN", "", "质朴青年 · BV100_streaming"),
    ("BV104_streaming", "zh-CN", "", "温柔淑女 · BV104_streaming"),
    ("BV004_streaming", "zh-CN", "", "开朗青年 · BV004_streaming"),
    ("BV113_streaming", "zh-CN", "", "甜宠少御 · BV113_streaming"),
    ("BV102_streaming", "zh-CN", "", "儒雅青年 · BV102_streaming"),
    ("BV405_streaming", "zh-CN", "", "甜美小源 · BV405_streaming"),
    ("BV007_streaming", "zh-CN", "Female", "亲切女声 · BV007_streaming"),
    ("BV009_streaming", "zh-CN", "Female", "知性女声 · BV009_streaming"),
    ("BV419_streaming", "zh-CN", "", "丞丞 · BV419_streaming"),
    ("BV415_streaming", "zh-CN", "", "彤彤 · BV415_streaming"),
    ("BV008_streaming", "zh-CN", "Male", "亲切男声 · BV008_streaming"),
    ("BV408_streaming", "zh-CN", "Male", "译制片男声 · BV408_streaming"),
    ("BV426_streaming", "zh-CN", "", "懒小羊 · BV426_streaming"),
    ("BV428_streaming", "zh-CN", "Female", "清新文艺女声 · BV428_streaming"),
    ("BV403_streaming", "zh-CN", "Female", "励志女声 · BV403_streaming"),
    ("BV158_streaming", "zh-CN", "", "睿智老者 · BV158_streaming"),
    ("BV157_streaming", "zh-CN", "Female", "慈爱奶奶 · BV157_streaming"),
    ("BR001_streaming", "zh-CN", "", "说唱小哥 · BR001_streaming"),
    ("BV410_streaming", "zh-CN", "Male", "活力解说男声 · BV410_streaming"),
    ("BV411_streaming", "zh-CN", "Male", "影视解说小帅 · BV411_streaming"),
    ("BV437_streaming", "zh-CN", "Male", "解说小帅多情感 · BV437_streaming"),
    ("BV412_streaming", "zh-CN", "Female", "影视解说小美 · BV412_streaming"),
    ("BV159_streaming", "zh-CN", "", "花花公子 · BV159_streaming"),
    ("BV418_streaming", "zh-CN", "Female", "直播一姐 · BV418_streaming"),
    ("BV142_streaming", "zh-CN", "Male", "沉稳解说男声 · BV142_streaming"),
    ("BV143_streaming", "zh-CN", "", "潇洒青年 · BV143_streaming"),
    (
        "zh_female_shuangkuaisisi_moon_bigtts",
        "zh-CN",
        "Female",
        "双快思思 · zh_female_shuangkuaisisi_moon_bigtts",
    ),
]


def _list_volcengine_voices(*, keyword: str = "") -> list[dict[str, Any]]:
    samples = [
        {
            "provider": "volcengine",
            "name": vid,
            "locale": loc,
            "gender": gender,
            "display_name": label,
        }
        for vid, loc, gender, label in _VOLC_VOICE_TYPE_SAMPLES
    ]
    kw = keyword.strip().lower()
    if not kw:
        return samples
    out: list[dict[str, Any]] = []
    for it in samples:
        hay = " ".join(str(v) for v in it.values()).lower()
        if kw in hay:
            out.append(it)
    return out

def _list_edge_voices(*, keyword: str = "") -> list[dict[str, Any]]:
    try:
        import edge_tts
    except ImportError as exc:  # pragma: no cover
        raise VoiceGenerationError("edge-tts is required for Edge voice generation.") from exc

    async def _run() -> list[dict[str, Any]]:
        voices = await edge_tts.list_voices()
        return voices

    try:
        voices = asyncio.run(_run())
    except Exception as exc:  # pragma: no cover
        raise VoiceGenerationError(f"Failed to list Edge voices: {exc}") from exc

    results = []
    keyword_lower = keyword.strip().lower()
    for voice in voices:
        item = {
            "provider": "edge",
            "name": voice.get("ShortName", ""),
            "locale": voice.get("Locale", ""),
            "gender": voice.get("Gender", ""),
            "display_name": voice.get("FriendlyName", ""),
        }
        haystack = " ".join(str(value) for value in item.values()).lower()
        if keyword_lower and keyword_lower not in haystack:
            continue
        results.append(item)
    return results


def _list_pyttsx3_voices(*, keyword: str = "") -> list[dict[str, Any]]:
    try:
        import pyttsx3
    except ImportError as exc:  # pragma: no cover
        raise VoiceGenerationError("pyttsx3 is required for local voice generation.") from exc

    engine = pyttsx3.init()
    voices = engine.getProperty("voices") or []
    keyword_lower = keyword.strip().lower()
    results = []
    for voice in voices:
        item = {
            "provider": "pyttsx3",
            "name": getattr(voice, "id", ""),
            "locale": ",".join(
                value.decode("utf-8", errors="ignore") if isinstance(value, bytes) else str(value)
                for value in (getattr(voice, "languages", []) or [])
            ),
            "gender": getattr(voice, "gender", ""),
            "display_name": getattr(voice, "name", ""),
        }
        haystack = " ".join(str(value) for value in item.values()).lower()
        if keyword_lower and keyword_lower not in haystack:
            continue
        results.append(item)
    engine.stop()
    return results


def _list_minimax_voices(*, provider_name: str, keyword: str = "") -> list[dict[str, Any]]:
    api_key = _provider_api_key(provider_name)
    if not api_key:
        raise VoiceGenerationError(
            "MiniMax API key missing. Set NOTE2VIDEO_MINIMAX_API_KEY or MINIMAX_API_KEY, "
            "or save it in the GUI settings (user config file)."
        )
    origin = _minimax_fixed_base_url(provider_name)
    url = f"{origin}/v1/get_voice"

    payload = _http_post_json(
        url,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        body={"voice_type": "system"},
        timeout_s=_provider_timeout_s(provider_name, for_list=True),
    )
    base_resp = (payload or {}).get("base_resp") or {}
    status_code = base_resp.get("status_code", 0)
    if status_code not in (0, "0"):
        status_msg = base_resp.get("status_msg") or "MiniMax get_voice failed."
        raise VoiceGenerationError(f"MiniMax get_voice failed: {status_code} {status_msg}".strip())

    voices = (payload or {}).get("system_voice") or []
    keyword_lower = keyword.strip().lower()
    results: list[dict[str, Any]] = []
    for v in voices:
        voice_id = str(v.get("voice_id", "") or "")
        voice_name = str(v.get("voice_name", "") or "")
        desc = v.get("description") or []
        item = {
            "provider": provider_name,
            "name": voice_id,
            "locale": "",
            "gender": "",
            "display_name": voice_name,
            "description": desc,
        }
        haystack = " ".join(
            [voice_id, voice_name, " ".join(str(x) for x in (desc if isinstance(desc, list) else [desc]))]
        ).lower()
        if keyword_lower and keyword_lower not in haystack:
            continue
        results.append(item)
    return results


def _http_post_json(url: str, *, headers: dict[str, str], body: dict[str, Any], timeout_s: float) -> dict[str, Any]:
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url=url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as exc:
        try:
            detail = exc.read().decode("utf-8", errors="replace")
        except Exception:
            detail = ""
        raise VoiceGenerationError(f"HTTP {exc.code} from {url}: {detail}".strip()) from exc
    except urllib.error.URLError as exc:
        raise VoiceGenerationError(f"Failed to reach {url}: {exc}") from exc

    try:
        payload = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise VoiceGenerationError("Failed to decode MiniMax JSON response.") from exc
    if not isinstance(payload, dict):
        raise VoiceGenerationError("MiniMax response is not a JSON object.")
    return payload


def _http_post_sse_collect_audio_bytes(
    url: str, *, headers: dict[str, str], body: dict[str, Any], timeout_s: float
) -> bytes:
    """
    Handle V3 unidirectional TTS which may stream JSON events (SSE / chunked).

    We parse ``data: {...}`` lines, base64-decode the ``data`` field, and append bytes.
    """
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url=url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            buf = bytearray()
            for raw_line in resp:
                line = raw_line.strip()
                if not line:
                    continue
                if line.startswith(b"data:"):
                    line = line[5:].strip()
                # Some servers may send plain JSON per line.
                try:
                    obj = json.loads(line.decode("utf-8", errors="replace"))
                except Exception:
                    continue
                if not isinstance(obj, dict):
                    continue
                header = obj.get("header")
                if isinstance(header, dict):
                    code = header.get("code")
                    # 0 typically means OK; treat others as errors.
                    if code not in (0, "0", None):
                        msg = header.get("message") or ""
                        raise VoiceGenerationError(f"Volcengine V3 TTS failed: {code} {msg}".strip())
                chunk_b64 = obj.get("data")
                if isinstance(chunk_b64, str) and chunk_b64.strip():
                    try:
                        buf.extend(base64.b64decode(chunk_b64))
                    except Exception as exc:
                        raise VoiceGenerationError("Volcengine V3 returned invalid base64 audio chunk.") from exc
            return bytes(buf)
    except urllib.error.HTTPError as exc:
        try:
            detail = exc.read().decode("utf-8", errors="replace")
        except Exception:
            detail = ""
        raise VoiceGenerationError(f"HTTP {exc.code} from {url}: {detail}".strip()) from exc
    except urllib.error.URLError as exc:
        raise VoiceGenerationError(f"Failed to reach {url}: {exc}") from exc


def _split_sentences(text: str) -> list[str]:
    return split_sentences(text)


def _split_sentences_with_pauses(text: str) -> list[tuple[str, int]]:
    return split_sentences_with_pauses(text)


def _read_wav_duration_ms(path: Path) -> int:
    with closing(wave.open(str(path), "rb")) as wav_file:
        frames = wav_file.getnframes()
        frame_rate = wav_file.getframerate()
    if frame_rate <= 0:
        return 0
    return int((frames / frame_rate) * 1000)


def _write_silence_wav(path: Path, *, duration_ms: int) -> None:
    frame_rate = 22050
    sample_width = 2
    channels = 1
    frame_count = int(frame_rate * (duration_ms / 1000))
    silence = b"\x00\x00" * frame_count

    with closing(wave.open(str(path), "wb")) as wav_file:
        wav_file.setnchannels(channels)
        wav_file.setsampwidth(sample_width)
        wav_file.setframerate(frame_rate)
        wav_file.writeframes(silence)


def _merge_wav_files(*, input_files: list[Path], output_file: Path) -> None:
    params = None
    frames: list[bytes] = []

    for file_path in input_files:
        with closing(wave.open(str(file_path), "rb")) as wav_file:
            current_params = (
                wav_file.getnchannels(),
                wav_file.getsampwidth(),
                wav_file.getframerate(),
            )
            if params is None:
                params = current_params
            elif params != current_params:
                raise VoiceGenerationError("All WAV files must share the same audio format.")
            frames.append(wav_file.readframes(wav_file.getnframes()))

    if params is None:
        _write_silence_wav(output_file, duration_ms=300)
        return

    with closing(wave.open(str(output_file), "wb")) as wav_file:
        wav_file.setnchannels(params[0])
        wav_file.setsampwidth(params[1])
        wav_file.setframerate(params[2])
        for chunk in frames:
            wav_file.writeframes(chunk)


def _update_manifest(
    *,
    manifest_path: Path,
    generated: list[dict[str, Any]],
    provider_name: str,
    voice_id: str,
    timings_path: str,
    tts_rate: float,
) -> None:
    if not manifest_path.exists():
        return

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    by_page = {item["page"]: item for item in generated}

    manifest["tts_provider"] = provider_name
    manifest["voice"] = voice_id
    manifest["tts_rate"] = tts_rate
    outputs = manifest.setdefault("outputs", {})
    outputs["audio_dir"] = "audio"
    outputs["merged_audio"] = "audio/merged.wav"
    outputs["timings"] = timings_path

    for slide in manifest.get("slides", []):
        page = slide.get("page")
        if page in by_page:
            slide["audio"] = by_page[page]["audio"]
            slide["duration_ms"] = by_page[page]["duration_ms"]

    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")


def _convert_audio_to_wav(*, temp_audio: Path, output_file: Path) -> None:
    ffmpeg = _get_ffmpeg_path()
    result = subprocess.run(
        [
            ffmpeg,
            "-y",
            "-i",
            str(temp_audio),
            "-ac",
            "1",
            "-ar",
            "22050",
            str(output_file),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode != 0:
        raise VoiceGenerationError(result.stderr.strip() or "Failed to convert Edge TTS audio to WAV.")


def _get_ffmpeg_path() -> str:
    try:
        import imageio_ffmpeg
    except ImportError as exc:  # pragma: no cover
        raise VoiceGenerationError("imageio-ffmpeg is required for audio conversion.") from exc
    return imageio_ffmpeg.get_ffmpeg_exe()
