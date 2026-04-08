import json
import wave
from contextlib import closing
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZIP_DEFLATED, ZipFile

from note2video.cli.main import main
from note2video.parser.extract import (
    _apply_png_sequence_to_slides,
    _is_notes_body_placeholder,
    _to_script,
    _to_speaker_notes,
)
from note2video.tts.voice import (
    EdgeTTSProvider,
    VolcengineTTSProvider,
    _convert_audio_to_wav,
    _create_provider,
)


def test_voice_command_generates_audio_files(tmp_path, monkeypatch) -> None:
    project_dir = tmp_path / "dist"
    scripts_dir = project_dir / "scripts"
    scripts_dir.mkdir(parents=True)
    script_path = scripts_dir / "script.json"
    script_path.write_text(
        json.dumps(
            {
                "slides": [
                    {"page": 1, "title": "A", "script": "hello"},
                    {"page": 2, "title": "B", "script": ""},
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    (project_dir / "manifest.json").write_text(
        json.dumps(
            {
                "project_name": "demo",
                "input_file": "demo.pptx",
                "slide_count": 2,
                "outputs": {},
                "slides": [
                    {"page": 1, "audio": "", "duration_ms": 0},
                    {"page": 2, "audio": "", "duration_ms": 0},
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    class FakeProvider:
        def synthesize_to_file(self, *, text, output_file):
            with closing(wave.open(str(output_file), "wb")) as wav_file:
                wav_file.setnchannels(1)
                wav_file.setsampwidth(2)
                wav_file.setframerate(22050)
                wav_file.writeframes(b"\x01\x02" * 22050)

    def fake_create_provider(*, provider_name, voice_id, tts_rate=1.0, minimax_base_url=None):
        return FakeProvider()

    monkeypatch.setattr("note2video.tts.voice._create_provider", fake_create_provider)

    exit_code = main(["voice", str(script_path), "--out", str(project_dir), "--json"])

    assert exit_code == 0
    assert (project_dir / "audio" / "001.wav").exists()
    assert (project_dir / "audio" / "002.wav").exists()
    assert (project_dir / "audio" / "merged.wav").exists()
    assert (project_dir / "audio" / "timings.json").exists()

    manifest = json.loads((project_dir / "manifest.json").read_text(encoding="utf-8"))
    timings = json.loads((project_dir / "audio" / "timings.json").read_text(encoding="utf-8"))
    assert manifest["outputs"]["merged_audio"] == "audio/merged.wav"
    assert manifest["outputs"]["timings"] == "audio/timings.json"
    assert manifest["slides"][0]["audio"] == "audio/001.wav"
    assert manifest["slides"][0]["duration_ms"] > 0
    assert manifest["slides"][1]["audio"] == "audio/002.wav"
    assert timings["segments"][0]["page"] == 1
    assert (project_dir / "logs" / "voice.log").exists()


def test_subtitle_command_generates_srt_and_json(tmp_path) -> None:
    project_dir = tmp_path / "dist"
    scripts_dir = project_dir / "scripts"
    subtitles_dir = project_dir / "subtitles"
    scripts_dir.mkdir(parents=True)
    subtitles_dir.mkdir(parents=True)

    script_path = scripts_dir / "script.json"
    script_path.write_text(
        json.dumps(
            {
                "slides": [
                    {"page": 1, "title": "A", "script": "第一句。第二句！"},
                    {"page": 2, "title": "B", "script": ""},
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    (project_dir / "manifest.json").write_text(
        json.dumps(
            {
                "project_name": "demo",
                "input_file": "demo.pptx",
                "slide_count": 2,
                "outputs": {},
                "slides": [
                    {"page": 1, "duration_ms": 2000},
                    {"page": 2, "duration_ms": 500},
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    exit_code = main(["subtitle", str(script_path), "--out", str(project_dir), "--json"])

    assert exit_code == 0
    srt_text = (project_dir / "subtitles" / "subtitles.srt").read_text(encoding="utf-8")
    subtitle_json = json.loads(
        (project_dir / "subtitles" / "subtitles.json").read_text(encoding="utf-8")
    )
    manifest = json.loads((project_dir / "manifest.json").read_text(encoding="utf-8"))

    assert "第一句。" in srt_text
    assert "第二句！" in srt_text
    assert subtitle_json["segments"][0]["page"] == 1
    assert manifest["outputs"]["subtitle"] == "subtitles/subtitles.srt"
    assert manifest["outputs"]["subtitle_json"] == "subtitles/subtitles.json"


def test_subtitle_prefers_voice_timings(tmp_path) -> None:
    project_dir = tmp_path / "dist"
    scripts_dir = project_dir / "scripts"
    audio_dir = project_dir / "audio"
    subtitles_dir = project_dir / "subtitles"
    scripts_dir.mkdir(parents=True)
    audio_dir.mkdir(parents=True)
    subtitles_dir.mkdir(parents=True)

    script_path = scripts_dir / "script.json"
    script_path.write_text(
        json.dumps({"slides": [{"page": 1, "title": "A", "script": "第一句。第二句！"}]}, ensure_ascii=False),
        encoding="utf-8",
    )
    (audio_dir / "timings.json").write_text(
        json.dumps(
            {
                "segments": [
                    {"index": 1, "page": 1, "text": "第一句。", "start_ms": 0, "end_ms": 700},
                    {"index": 2, "page": 1, "text": "第二句！", "start_ms": 700, "end_ms": 1500},
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    (project_dir / "manifest.json").write_text(
        json.dumps(
            {
                "project_name": "demo",
                "input_file": "demo.pptx",
                "slide_count": 1,
                "outputs": {"timings": "audio/timings.json"},
                "slides": [{"page": 1, "duration_ms": 1500}],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    exit_code = main(["subtitle", str(script_path), "--out", str(project_dir), "--json"])
    assert exit_code == 0
    subtitle_json = json.loads((project_dir / "subtitles" / "subtitles.json").read_text(encoding="utf-8"))
    assert subtitle_json["segments"][0]["start_ms"] == 0
    assert subtitle_json["segments"][1]["start_ms"] == 700


def test_render_command_updates_manifest(tmp_path, monkeypatch) -> None:
    project_dir = tmp_path / "dist"
    slides_dir = project_dir / "slides"
    audio_dir = project_dir / "audio"
    subtitles_dir = project_dir / "subtitles"
    video_dir = project_dir / "video"
    slides_dir.mkdir(parents=True)
    audio_dir.mkdir(parents=True)
    subtitles_dir.mkdir(parents=True)
    video_dir.mkdir(parents=True)

    (slides_dir / "001.png").write_bytes(b"fake")
    (audio_dir / "merged.wav").write_bytes(b"fake")
    (subtitles_dir / "subtitles.srt").write_text("1\n00:00:00,000 --> 00:00:01,000\nHi\n", encoding="utf-8")
    (project_dir / "manifest.json").write_text(
        json.dumps(
            {
                "project_name": "demo",
                "input_file": "demo.pptx",
                "slide_count": 1,
                "outputs": {
                    "merged_audio": "audio/merged.wav",
                    "subtitle": "subtitles/subtitles.srt",
                },
                "slides": [
                    {
                        "page": 1,
                        "image": "slides/001.png",
                        "duration_ms": 1000,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    monkeypatch.setattr("note2video.render.video._get_ffmpeg_path", lambda: "ffmpeg")
    commands = []

    def fake_run_ffmpeg(command):
        commands.append(command)
        if command[-1].endswith("video_only.mp4"):
            (video_dir / "video_only.mp4").write_bytes(b"video-only")
        elif command[-1].endswith("output.mp4"):
            (video_dir / "output.mp4").write_bytes(b"video-final")

    monkeypatch.setattr("note2video.render.video._run_ffmpeg", fake_run_ffmpeg)

    exit_code = main(["render", str(project_dir), "--json"])

    assert exit_code == 0
    manifest = json.loads((project_dir / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["outputs"]["video"] == "video/output.mp4"
    assert (project_dir / "video" / "output.mp4").exists()
    first_command = commands[0]
    assert "fps=30,format=yuv420p" in first_command
    assert "vfr" not in first_command
    assert not (project_dir / "video" / "video_only.mp4").exists()
    assert not (project_dir / "video" / "slides.ffconcat").exists()


def test_build_command_runs_full_pipeline(tmp_path, monkeypatch) -> None:
    input_file = tmp_path / "demo.pptx"
    output_dir = tmp_path / "dist"
    input_file.write_bytes(b"placeholder")

    calls = []

    def fake_extract(input_path, out_dir, pages=None):
        calls.append(("extract", input_path, out_dir, pages))
        scripts_dir = output_dir / "scripts"
        scripts_dir.mkdir(parents=True, exist_ok=True)
        (scripts_dir / "script.json").write_text('{"slides":[]}', encoding="utf-8")
        return SimpleNamespace(slide_count=3)

    def fake_voice(input_json, out_dir, *, provider_name="pyttsx3", voice_id="", tts_rate=1.0, minimax_base_url=None):
        calls.append(("voice", input_json, out_dir, provider_name, voice_id, tts_rate, minimax_base_url))
        return {"provider": provider_name, "slide_count": 3, "tts_rate": tts_rate}

    def fake_subtitle(input_json, out_dir):
        calls.append(("subtitle", input_json, out_dir))
        return {"segment_count": 5}

    def fake_render(
        project_dir,
        output_path=None,
        *,
        bgm_path=None,
        bgm_volume=0.18,
        bgm_fade_in_s=0.0,
        bgm_fade_out_s=0.0,
        narration_volume=1.0,
        subtitle_color=None,
    ):
        calls.append(
            ("render", project_dir, bgm_path, bgm_volume, bgm_fade_in_s, bgm_fade_out_s, narration_volume, subtitle_color)
        )
        return {"video": "video/output.mp4", "subtitles_burned": True}

    monkeypatch.setattr("note2video.cli.main.extract_project", fake_extract)
    monkeypatch.setattr("note2video.cli.main.generate_voice_assets", fake_voice)
    monkeypatch.setattr("note2video.cli.main.generate_subtitles", fake_subtitle)
    monkeypatch.setattr("note2video.cli.main.render_video", fake_render)

    exit_code = main(
        [
            "build",
            str(input_file),
            "--out",
            str(output_dir),
            "--pages",
            "1-3",
            "--tts-provider",
            "pyttsx3",
            "--json",
        ]
    )

    assert exit_code == 0
    assert calls[0] == ("extract", str(input_file), str(output_dir), "1-3")
    assert calls[1][0] == "voice"
    assert calls[2][0] == "subtitle"
    assert calls[3] == ("render", str(output_dir), None, 0.18, 0.0, 0.0, 1.0, None)


def test_extract_command_writes_expected_files(tmp_path, monkeypatch) -> None:
    input_file = tmp_path / "demo.pptx"
    input_file.write_bytes(b"placeholder")

    def fake_extract_slide_data(_input_path, _slides_dir, *, selected_pages=None):
        return [
            {
                "page": 1,
                "title": "Intro",
                "image": "slides/001.png",
                "raw_notes": "Hello world",
            }
        ], "mock-extractor"

    monkeypatch.setattr(
        "note2video.parser.extract._extract_slide_data",
        fake_extract_slide_data,
    )

    output_dir = tmp_path / "dist"
    exit_code = main(["extract", str(input_file), "--out", str(output_dir)])

    assert exit_code == 0
    manifest = json.loads((output_dir / "manifest.json").read_text(encoding="utf-8"))
    notes = json.loads((output_dir / "notes" / "notes.json").read_text(encoding="utf-8"))
    script = json.loads((output_dir / "scripts" / "script.json").read_text(encoding="utf-8"))
    log_text = (output_dir / "logs" / "build.log").read_text(encoding="utf-8")
    raw_txt = (output_dir / "notes" / "raw" / "001.txt").read_text(encoding="utf-8")
    speaker_txt = (output_dir / "notes" / "speaker" / "001.txt").read_text(encoding="utf-8")
    script_txt = (output_dir / "scripts" / "txt" / "001.txt").read_text(encoding="utf-8")
    all_script_txt = (output_dir / "scripts" / "all.txt").read_text(encoding="utf-8")

    assert manifest["slide_count"] == 1
    assert notes["slide_count"] == 1
    assert script["slides"][0]["script"] == "Hello world"
    assert "exported_slides: 1" in log_text
    assert raw_txt == "Hello world"
    assert speaker_txt == "Hello world"
    assert script_txt == "Hello world"
    assert "Hello world" in all_script_txt
    assert "extractor: mock-extractor" in log_text


def test_extract_command_filters_pages(tmp_path, monkeypatch) -> None:
    input_file = tmp_path / "demo.pptx"
    input_file.write_bytes(b"placeholder")

    def fake_extract_slide_data(_input_path, _slides_dir, *, selected_pages=None):
        items = [
            {"page": 1, "title": "A", "image": "slides/001.png", "raw_notes": "one"},
            {"page": 2, "title": "B", "image": "slides/002.png", "raw_notes": "two"},
            {"page": 3, "title": "C", "image": "slides/003.png", "raw_notes": "three"},
        ]
        if selected_pages is not None:
            items = [x for x in items if x["page"] in selected_pages]
        return items, "mock-extractor"

    monkeypatch.setattr(
        "note2video.parser.extract._extract_slide_data",
        fake_extract_slide_data,
    )
    monkeypatch.setattr("note2video.parser.extract._count_slides_openxml", lambda _p: 3)

    output_dir = tmp_path / "dist"
    exit_code = main(
        ["extract", str(input_file), "--out", str(output_dir), "--pages", "1,3"]
    )

    assert exit_code == 0
    notes = json.loads((output_dir / "notes" / "notes.json").read_text(encoding="utf-8"))
    assert notes["slide_count"] == 2
    assert [slide["page"] for slide in notes["slides"]] == [1, 3]

    # Ensure we didn't generate slide images for excluded pages.
    slides_dir = output_dir / "slides"
    assert (slides_dir / "002.png").exists() is False


def test_extract_command_openxml_fallback_on_linux(tmp_path, monkeypatch) -> None:
    input_file = tmp_path / "demo.pptx"
    output_dir = tmp_path / "dist"
    _write_minimal_openxml_pptx(input_file)
    monkeypatch.setenv("NOTE2VIDEO_USE_LIBREOFFICE", "0")

    exit_code = main(["extract", str(input_file), "--out", str(output_dir)])

    assert exit_code == 0
    notes = json.loads((output_dir / "notes" / "notes.json").read_text(encoding="utf-8"))
    script = json.loads((output_dir / "scripts" / "script.json").read_text(encoding="utf-8"))
    log_text = (output_dir / "logs" / "build.log").read_text(encoding="utf-8")

    assert notes["slide_count"] == 1
    assert notes["slides"][0]["title"] == "Demo Slide"
    assert notes["slides"][0]["raw_notes"] == "第一句。\n第二句！"
    assert script["slides"][0]["script"] == "第一句。\n第二句！"
    assert (output_dir / "slides" / "001.png").exists()
    assert "extractor: openxml" in log_text


def test_apply_png_sequence_to_slides_copies_and_pads(tmp_path) -> None:
    slides_dir = tmp_path / "slides"
    slides_dir.mkdir()
    png_a = tmp_path / "a.png"
    png_b = tmp_path / "b.png"
    png_a.write_bytes(b"img1")
    png_b.write_bytes(b"img2")
    meta = [
        {"page": 1, "title": "T1", "raw_notes": "n1"},
        {"page": 2, "title": "T2", "raw_notes": "n2"},
        {"page": 3, "title": "T3", "raw_notes": "n3"},
    ]
    out = _apply_png_sequence_to_slides(meta, [png_a, png_b], slides_dir)
    assert len(out) == 3
    assert (slides_dir / "001.png").read_bytes() == b"img1"
    assert (slides_dir / "002.png").read_bytes() == b"img2"
    assert (slides_dir / "003.png").exists()
    assert out[2]["raw_notes"] == "n3"


def test_apply_png_sequence_to_slides_respects_selected_pages(tmp_path) -> None:
    slides_dir = tmp_path / "slides"
    slides_dir.mkdir()
    png_a = tmp_path / "a.png"
    png_b = tmp_path / "b.png"
    png_c = tmp_path / "c.png"
    png_a.write_bytes(b"img1")
    png_b.write_bytes(b"img2")
    png_c.write_bytes(b"img3")
    meta = [
        {"page": 1, "title": "T1", "raw_notes": "n1"},
        {"page": 2, "title": "T2", "raw_notes": "n2"},
        {"page": 3, "title": "T3", "raw_notes": "n3"},
    ]
    out = _apply_png_sequence_to_slides(meta, [png_a, png_b, png_c], slides_dir, selected_pages={1, 3})
    assert [row["page"] for row in out] == [1, 3]
    assert (slides_dir / "001.png").exists()
    assert not (slides_dir / "002.png").exists()
    assert (slides_dir / "003.png").exists()


def test_notes_body_placeholder_filter() -> None:
    class PlaceholderFormat:
        def __init__(self, placeholder_type):
            self.Type = placeholder_type

    class Shape:
        def __init__(self, shape_type, placeholder_type):
            self.Type = shape_type
            self.PlaceholderFormat = PlaceholderFormat(placeholder_type)

    assert _is_notes_body_placeholder(Shape(14, 2)) is True
    assert _is_notes_body_placeholder(Shape(14, 13)) is False
    assert _is_notes_body_placeholder(Shape(14, 15)) is False
    assert _is_notes_body_placeholder(Shape(1, 2)) is False


def test_speaker_notes_and_script_cleanup() -> None:
    raw = "备注：仅供演讲者参考\n大家好，欢迎来到课程。这是第一段。接着讲第二段！"
    speaker_notes = _to_speaker_notes(raw)
    script = _to_script(raw)

    assert "仅供演讲者参考" not in speaker_notes
    assert speaker_notes == "大家好，欢迎来到课程。这是第一段。接着讲第二段！"
    assert script == "大家好，欢迎来到课程。\n这是第一段。\n接着讲第二段！"


def test_tts_provider_selection() -> None:
    edge_provider = _create_provider(provider_name="edge", voice_id="")
    assert isinstance(edge_provider, EdgeTTSProvider)
    assert edge_provider.voice_id == "zh-CN-XiaoxiaoNeural"
    assert edge_provider.tts_rate == 1.0
    volc = _create_provider(provider_name="volcengine", voice_id="")
    assert isinstance(volc, VolcengineTTSProvider)
    assert volc.voice_id == "BV700_streaming"
    doubao = _create_provider(provider_name="doubao", voice_id="")
    assert isinstance(doubao, VolcengineTTSProvider)


def test_voices_command_returns_success(monkeypatch) -> None:
    monkeypatch.setattr(
        "note2video.cli.main.list_available_voices",
        lambda provider_name, keyword, minimax_base_url=None: [
            {
                "provider": provider_name,
                "name": "zh-CN-XiaoxiaoNeural",
                "locale": "zh-CN",
                "gender": "Female",
                "display_name": "Xiaoxiao",
            }
        ],
    )
    exit_code = main(["voices", "--tts-provider", "edge", "--keyword", "zh-CN", "--json"])
    assert exit_code == 0


def test_convert_audio_to_wav_invokes_ffmpeg(tmp_path, monkeypatch) -> None:
    temp_audio = tmp_path / "temp.mp3"
    output_file = tmp_path / "out.wav"
    temp_audio.write_bytes(b"mp3")

    monkeypatch.setattr("note2video.tts.voice._get_ffmpeg_path", lambda: "ffmpeg")

    class Result:
        returncode = 0
        stderr = ""

    def fake_run(command, capture_output, text, **kwargs):
        assert kwargs.get("encoding") == "utf-8"
        assert kwargs.get("errors") == "replace"
        output_file.write_bytes(b"RIFF")
        return Result()

    monkeypatch.setattr("note2video.tts.voice.subprocess.run", fake_run)

    _convert_audio_to_wav(temp_audio=temp_audio, output_file=output_file)
    assert output_file.exists()


def test_powerpoint_export_uses_absolute_image_path(tmp_path, monkeypatch) -> None:
    import note2video.parser.extract as extract_module

    exported_paths = []

    class FakeSlide:
        def __init__(self, index):
            self.index = index
            self.Shapes = SimpleNamespace(HasTitle=False)
            self.NotesPage = SimpleNamespace(Shapes=[])

        def Export(self, path, image_format):
            exported_paths.append((path, image_format))
            Path(path).write_bytes(b"png")

    class FakePresentation:
        def __init__(self):
            self.Slides = [FakeSlide(1)]

        def Close(self):
            return None

    class FakePresentations:
        def Open(self, *_args, **_kwargs):
            return FakePresentation()

    class FakeApp:
        def __init__(self):
            self.Visible = 0
            self.Presentations = FakePresentations()

        def Quit(self):
            return None

    fake_pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    fake_client = SimpleNamespace(DispatchEx=lambda _name: FakeApp())
    monkeypatch.setitem(__import__("sys").modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(
        __import__("sys").modules,
        "win32com",
        SimpleNamespace(client=fake_client),
    )
    monkeypatch.setitem(__import__("sys").modules, "win32com.client", fake_client)

    input_file = tmp_path / "demo.pptx"
    input_file.write_bytes(b"placeholder")
    slides_dir = tmp_path / "dist" / "slides"
    slides_dir.mkdir(parents=True)

    result = extract_module._extract_with_powerpoint(input_file, slides_dir)

    assert result[0]["image"] == "slides/001.png"
    assert exported_paths
    export_path, export_format = exported_paths[0]
    assert Path(export_path).is_absolute()
    assert export_format == "PNG"


def _write_minimal_openxml_pptx(output_path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

    presentation = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
 <p:sldIdLst>
  <p:sldId id="256" r:id="rId1"/>
 </p:sldIdLst>
 <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
 <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""

    presentation_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"""

    slide1 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
 <p:cSld>
  <p:spTree>
   <p:nvGrpSpPr>
    <p:cNvPr id="1" name=""/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
   </p:nvGrpSpPr>
   <p:grpSpPr>
    <a:xfrm>
     <a:off x="0" y="0"/>
     <a:ext cx="0" cy="0"/>
     <a:chOff x="0" y="0"/>
     <a:chExt cx="0" cy="0"/>
    </a:xfrm>
   </p:grpSpPr>
   <p:sp>
    <p:nvSpPr>
     <p:cNvPr id="2" name="Title 1"/>
     <p:cNvSpPr/>
     <p:nvPr><p:ph type="title"/></p:nvPr>
    </p:nvSpPr>
    <p:spPr/>
    <p:txBody>
     <a:bodyPr/>
     <a:lstStyle/>
     <a:p><a:r><a:t>Demo Slide</a:t></a:r></a:p>
    </p:txBody>
   </p:sp>
  </p:spTree>
 </p:cSld>
 <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

    slide1_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>"""

    notes_slide1 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
 <p:cSld>
  <p:spTree>
   <p:nvGrpSpPr>
    <p:cNvPr id="1" name=""/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
   </p:nvGrpSpPr>
   <p:grpSpPr><a:xfrm/></p:grpSpPr>
   <p:sp>
    <p:nvSpPr>
     <p:cNvPr id="2" name="Notes Placeholder 1"/>
     <p:cNvSpPr/>
     <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
    </p:nvSpPr>
    <p:spPr/>
    <p:txBody>
     <a:bodyPr/>
     <a:lstStyle/>
     <a:p><a:r><a:t>第一句。</a:t></a:r></a:p>
     <a:p><a:r><a:t>第二句！</a:t></a:r></a:p>
    </p:txBody>
   </p:sp>
  </p:spTree>
 </p:cSld>
 <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>"""

    with ZipFile(output_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("ppt/presentation.xml", presentation)
        archive.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
        archive.writestr("ppt/slides/slide1.xml", slide1)
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", slide1_rels)
        archive.writestr("ppt/notesSlides/notesSlide1.xml", notes_slide1)
