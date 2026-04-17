from note2video.subtitle.generate import _split_sentences as subtitle_split_sentences
from note2video.text_segmentation import split_sentences, split_sentences_with_pauses
from note2video.tts.voice import _split_sentences as voice_split_sentences
from note2video.tts.voice import _split_sentences_with_pauses as voice_split_sentences_with_pauses


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


def test_split_sentences_with_pauses_uses_paragraph_pause() -> None:
    text = "Alpha.\n\nBeta"
    chunks = split_sentences_with_pauses(text)
    assert chunks[0][0] == "Alpha."
    assert chunks[0][1] == 620
    assert chunks[1] == ("Beta", 0)


def test_split_sentences_with_pauses_uses_dot_pause() -> None:
    chunks = split_sentences_with_pauses("One. Two")
    assert chunks[0] == ("One.", 360)
    assert chunks[1] == ("Two", 0)


def test_split_sentences_with_pauses_uses_comma_and_colon_pause() -> None:
    chunks = split_sentences_with_pauses("One, Two: Three")
    assert chunks[0] == ("One,", 240)
    assert chunks[1] == ("Two:", 280)
    assert chunks[2] == ("Three", 0)


def test_subtitle_and_voice_wrappers_use_shared_segmentation() -> None:
    text = "A.\n\nB?\nC"
    assert subtitle_split_sentences(text) == split_sentences(text)
    assert voice_split_sentences(text) == split_sentences(text)
    assert voice_split_sentences_with_pauses(text) == split_sentences_with_pauses(text)
