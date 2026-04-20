from note2video.subtitle.generate import _split_sentences as subtitle_split_sentences
from note2video.text_segmentation import split_sentences
from note2video.tts.voice import _split_sentences as voice_split_sentences
from note2video.tts.voice import _split_tts_chunks_with_pauses


def test_split_sentences_uses_punctuation_and_paragraph_breaks() -> None:
    text = "Alpha. Beta!\n\nGamma?\nDelta;"
    assert split_sentences(text) == [
        "Alpha.",
        "Beta!",
        "Gamma?",
        "Delta;",
    ]


def test_split_sentences_uses_comma_and_colon_in_both_languages() -> None:
    text = "中文，下一句：继续。 English, next: continue."
    assert split_sentences(text) == [
        "中文，",
        "下一句：",
        "继续。",
        "English,",
        "next:",
        "continue.",
    ]


def test_tts_chunk_splitter_does_not_add_punctuation_pauses() -> None:
    chunks = _split_tts_chunks_with_pauses("One, Two: Three")
    assert chunks == [("One,", 0), ("Two:", 0), ("Three", 0)]


def test_subtitle_and_voice_wrappers_use_shared_segmentation() -> None:
    text = "A.\n\nB?\nC"
    assert subtitle_split_sentences(text) == split_sentences(text)
    assert voice_split_sentences(text) == split_sentences(text)
