from pathlib import Path

from build import build
from common import Warnings
from fixtures import make_png, make_template_pptx
from verify import verify_output


def _mapping(template: Path) -> dict:
    return {
        "template": {
            "file": str(template),
            "mode": "clone",
            "source_slide": 0,
            "shapes": {
                "title": "제목 13",
                "screen_id": "텍스트 개체 틀 14",
                "image": "그림 18",
                "detail_tables": ["표 7", "표 8", "표 9", "표 10", "표 11"],
            },
            "table_columns": {"no": 0, "text": 1},
        },
        "options": {"detail_text_source": "desc", "clear_unused_slots": True},
    }


def _data(image: str | None) -> dict:
    return {
        "meta": {"title": "화면설계서"},
        "screens": [
            {"id": "SCR001", "name": "이용기관 목록",
             "images": [image] if image else [], "fields": {},
             "details": [{"no": "1", "desc": "등록한다"}]}
        ],
    }


def test_verify_passes_for_good_output(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    make_png(tmp_path / "images" / "SCR001.png")
    data = _data("images/SCR001.png")
    out = tmp_path / "out.pptx"
    report = build(data, _mapping(tpl), tmp_path, out, Warnings())

    result = verify_output(out, data, _mapping(tpl), report["slides"])
    assert result["ok"] is True
    assert all(c["ok"] for c in result["checks"])


def test_verify_detects_slide_count_mismatch(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _data(None)
    out = tmp_path / "out.pptx"
    build(data, _mapping(tpl), tmp_path, out, Warnings())

    result = verify_output(out, data, _mapping(tpl), 5)
    assert result["ok"] is False
    failed = [c["name"] for c in result["checks"] if not c["ok"]]
    assert "슬라이드 수" in failed
