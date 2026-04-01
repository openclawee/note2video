from __future__ import annotations

import asyncio
import json
import re
import subprocess
import wave
from contextlib import closing
from pathlib import Path
from typing import Any


class VoiceGenerationError(RuntimeError):
    """Raised when TTS generation fails."""


def list_available_voices(*, provider_name: str = "edge", keyword: str = "") -> list[dict[str, Any]]:
    normalized = provider_name.strip().lower() or "edge"
    if normalized == "edge":
        return _list_edge_voices(keyword=keyword)
    if normalized == "pyttsx3":
        return _list_pyttsx3_voices(keyword=keyword)
    raise ValueError(f"Unsupported TTS provider: {provider_name}")


def generate_voice_assets(
    input_json: str,
    output_dir: str,
    *,
    provider_name: str = "pyttsx3",
    voice_id: str = "",
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

    provider = _create_provider(provider_name=provider_name, voice_id=voice_id)
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
            sentences = _split_sentences(text)
            if not sentences:
                sentences = [text.strip()]

            for idx, sentence in enumerate(sentences, start=1):
                sentence_file = audio_dir / f"{page:03d}.{idx:02d}.wav"
                provider.synthesize_to_file(text=sentence, output_file=sentence_file)
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
    )

    (logs_dir / "voice.log").write_text(
        f"provider: {provider_name}\nvoice: {voice_id or 'default'}\nslides: {len(generated)}\n",
        encoding="utf-8",
    )

    return {
        "audio_dir": str(audio_dir),
        "merged_audio": str(merged_path),
        "timings": str(timings_path),
        "slide_count": len(generated),
        "provider": provider_name,
        "voice": voice_id,
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


def _create_provider(*, provider_name: str, voice_id: str):
    normalized = provider_name.strip().lower() or "pyttsx3"
    if normalized == "pyttsx3":
        return Pyttsx3Provider(voice_id=voice_id)
    if normalized == "edge":
        return EdgeTTSProvider(voice_id=voice_id or "zh-CN-XiaoxiaoNeural")
    raise ValueError(f"Unsupported TTS provider: {provider_name}")


class Pyttsx3Provider:
    def __init__(self, *, voice_id: str = "") -> None:
        self.voice_id = voice_id

    def synthesize_to_file(self, *, text: str, output_file: Path) -> None:
        try:
            import pyttsx3
        except ImportError as exc:  # pragma: no cover
            raise VoiceGenerationError("pyttsx3 is required for local voice generation.") from exc

        try:
            engine = pyttsx3.init()
            if self.voice_id:
                engine.setProperty("voice", self.voice_id)
            engine.save_to_file(text, str(output_file))
            engine.runAndWait()
            engine.stop()
        except Exception as exc:  # pragma: no cover - runtime specific
            raise VoiceGenerationError(f"pyttsx3 synthesis failed: {exc}") from exc


class EdgeTTSProvider:
    def __init__(self, *, voice_id: str) -> None:
        self.voice_id = voice_id

    def synthesize_to_file(self, *, text: str, output_file: Path) -> None:
        try:
            import edge_tts
        except ImportError as exc:  # pragma: no cover
            raise VoiceGenerationError("edge-tts is required for Edge voice generation.") from exc

        temp_audio = output_file.with_suffix(".edge.mp3")

        async def _run() -> None:
            communicate = edge_tts.Communicate(text=text, voice=self.voice_id)
            await communicate.save(str(temp_audio))

        try:
            asyncio.run(_run())
            _convert_audio_to_wav(temp_audio=temp_audio, output_file=output_file)
        except Exception as exc:  # pragma: no cover - runtime/network specific
            raise VoiceGenerationError(f"edge-tts synthesis failed: {exc}") from exc
        finally:
            if temp_audio.exists():
                temp_audio.unlink()


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


def _split_sentences(text: str) -> list[str]:
    normalized = text.replace("\r", "\n").strip()
    if not normalized:
        return []
    raw_parts = re.split(r"(?<=[。！？!?；;])|\n+", normalized)
    return [part.strip() for part in raw_parts if part and part.strip()]


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
) -> None:
    if not manifest_path.exists():
        return

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    by_page = {item["page"]: item for item in generated}

    manifest["tts_provider"] = provider_name
    manifest["voice"] = voice_id
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
    )
    if result.returncode != 0:
        raise VoiceGenerationError(result.stderr.strip() or "Failed to convert Edge TTS audio to WAV.")


def _get_ffmpeg_path() -> str:
    try:
        import imageio_ffmpeg
    except ImportError as exc:  # pragma: no cover
        raise VoiceGenerationError("imageio-ffmpeg is required for audio conversion.") from exc
    return imageio_ffmpeg.get_ffmpeg_exe()
