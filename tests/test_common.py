from pathlib import Path

from common import Warnings, read_json, write_json


def test_json_roundtrip_keeps_korean(tmp_path: Path):
    p = tmp_path / "a.json"
    write_json(p, {"name": "이용기관 목록"})
    assert read_json(p) == {"name": "이용기관 목록"}
    assert "이용기관" in p.read_text(encoding="utf-8")


def test_warnings_collects_and_formats():
    w = Warnings()
    w.add("B2BISMT1001", "no-image", "이미지를 찾지 못했습니다")
    w.add(None, "shape-not-found", "제목 도형 없음")
    assert len(w) == 2
    assert w.to_list()[0] == {
        "screen_id": "B2BISMT1001",
        "code": "no-image",
        "message": "이미지를 찾지 못했습니다",
    }
    text = w.format()
    assert "B2BISMT1001" in text
    assert "no-image" in text


def test_warnings_format_is_empty_when_none():
    assert Warnings().format() == ""
