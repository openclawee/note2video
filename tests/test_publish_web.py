from __future__ import annotations

from note2video.publish.web import merge_description_and_topics, normalize_topics


def test_normalize_topics_empty() -> None:
    assert normalize_topics("") == []
    assert normalize_topics(" , , ") == []


def test_normalize_topics_adds_hash_and_dedups() -> None:
    normalized = normalize_topics("AI, #教育,AI,  #效率 ")
    assert normalized == ["AI", "教育", "效率"]


def test_normalize_topics_iterable() -> None:
    normalized = normalize_topics(["AI", " #教育 ", "ai"])
    assert normalized == ["AI", "教育"]


def test_merge_description_and_topics() -> None:
    merged = merge_description_and_topics("这是一段描述", "AI,效率")
    assert merged == "这是一段描述\n#AI #效率"


def test_merge_description_and_topics_topics_only() -> None:
    merged = merge_description_and_topics("", "AI,效率")
    assert merged == "#AI #效率"
