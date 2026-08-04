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


def test_verify_fails_size_check_when_template_missing(tmp_path: Path):
    """비교 대상 템플릿이 없으면 크기 검사를 건너뛰지 말고 실패로 보고해야 한다."""
    tpl = make_template_pptx(tmp_path / "t.pptx")
    make_png(tmp_path / "images" / "SCR001.png")
    data = _data("images/SCR001.png")
    out = tmp_path / "out.pptx"
    report = build(data, _mapping(tpl), tmp_path, out, Warnings())

    bad_mapping = _mapping(tpl)
    bad_mapping["template"]["file"] = str(tmp_path / "does-not-exist.pptx")

    result = verify_output(out, data, bad_mapping, report["slides"])
    assert result["ok"] is False
    failed = {c["name"]: c for c in result["checks"] if not c["ok"]}
    assert "슬라이드 크기" in failed
    assert "does-not-exist.pptx" in failed["슬라이드 크기"]["detail"]
    # 다른 검사는 여전히 정상이어야 한다 — 크기 검사만 실패로 좁혀졌는지 확인
    assert failed.keys() == {"슬라이드 크기"}


def test_verify_detects_missing_image_for_screen(tmp_path: Path):
    """상세에 이미지가 있다고 되어 있지만 실제 파일이 없어 배치가 실패하는 상황.

    build()는 화면 단위로 예외를 격리하므로 슬라이드는 만들어지지만(제목은
    '[생성 실패] ...'로 바뀐다) 그림 도형은 존재하지 않는다. 이미지 배치 검사가
    이 화면을 정확히 짚어내야 한다 — 전체 그림 개수만 세면 놓치는 사례다.
    """
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _data("images/missing.png")  # 파일을 만들지 않아 place_image가 실패한다
    out = tmp_path / "out.pptx"
    report = build(data, _mapping(tpl), tmp_path, out, Warnings())

    result = verify_output(out, data, _mapping(tpl), report["slides"])
    assert result["ok"] is False
    failed = {c["name"]: c for c in result["checks"] if not c["ok"]}
    assert "이미지 배치" in failed
    assert "SCR001" in failed["이미지 배치"]["detail"]
