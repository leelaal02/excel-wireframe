from pathlib import Path

from build import build, chunk_details, page_title
from common import Warnings, write_json
from fixtures import make_png, make_template_pptx
from pptx import Presentation


def _mapping(template: Path) -> dict:
    return {
        "version": 1,
        "excel": {"layout": "sheet-per-screen"},
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
        "options": {
            "detail_text_source": "desc",
            "overflow": "split",
            "clear_unused_slots": True,
        },
    }


def _screens(n_details: int, image: str | None = None) -> dict:
    return {
        "meta": {"title": "화면설계서", "source": "s.xlsx", "template": "t.pptx"},
        "screens": [
            {
                "id": "SCR001",
                "name": "이용기관 목록",
                "images": [image] if image else [],
                "fields": {},
                "details": [
                    {"no": str(i + 1), "desc": "설명 %d" % (i + 1)}
                    for i in range(n_details)
                ],
            }
        ],
    }


def test_chunk_details_splits_by_slot_count():
    d = [{"no": str(i)} for i in range(25)]
    chunks = chunk_details(d, 20)
    assert [len(c) for c in chunks] == [20, 5]
    assert chunk_details([], 20) == [[]]
    assert len(chunk_details(d[:20], 20)) == 1


def test_page_title_marks_split():
    assert page_title("목록", 0, 1) == "목록"
    assert page_title("목록", 0, 2) == "목록 (1/2)"
    assert page_title("목록", 1, 2) == "목록 (2/2)"


def test_build_creates_one_slide_per_screen(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    report = build(_screens(6), _mapping(tpl), tmp_path, out, Warnings())

    assert report["slides"] == 1
    prs = Presentation(str(out))
    assert len(prs.slides) == 1
    slide = prs.slides[0]
    title = next(s for s in slide.shapes if s.name == "제목 13")
    assert title.text_frame.text == "이용기관 목록"
    sid = next(s for s in slide.shapes if s.name == "텍스트 개체 틀 14")
    assert sid.text_frame.text == "SCR001"


def test_build_keeps_slide_size(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    build(_screens(3), _mapping(tpl), tmp_path, out, Warnings())
    assert Presentation(str(out)).slide_width == 9906000


def test_build_splits_on_overflow(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    warns = Warnings()
    report = build(_screens(25), _mapping(tpl), tmp_path, out, warns)

    assert report["slides"] == 2
    assert report["split"] == ["SCR001"]
    prs = Presentation(str(out))
    titles = [
        next(s for s in sl.shapes if s.name == "제목 13").text_frame.text
        for sl in prs.slides
    ]
    assert titles == ["이용기관 목록 (1/2)", "이용기관 목록 (2/2)"]
    assert "slide-split" in [w["code"] for w in warns.to_list()]


def test_build_places_image(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    make_png(tmp_path / "images" / "SCR001.png", size=(800, 400))
    out = tmp_path / "out.pptx"
    build(_screens(4, "images/SCR001.png"), _mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    pics = [s for s in slide.shapes if s.shape_type == 13]
    assert len(pics) == 1
    assert not any(s.name == "그림 18" for s in slide.shapes)


def test_build_isolates_screen_failure(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["screens"].append(
        {"id": "BROKEN", "name": "깨진 화면", "images": ["images/없는파일.png"],
         "fields": {}, "details": [{"no": "1", "desc": "x"}]}
    )
    out = tmp_path / "out.pptx"
    warns = Warnings()
    report = build(data, _mapping(tpl), tmp_path, out, warns)

    assert report["screens"] == 2
    assert report["failed"] == ["BROKEN"]
    assert out.exists()
    assert "screen-failed" in [w["code"] for w in warns.to_list()]

    prs = Presentation(str(out))
    assert len(prs.slides) == report["slides"]
    titles = [
        next(s for s in sl.shapes if s.name == "제목 13").text_frame.text
        for sl in prs.slides
    ]
    assert titles == ["이용기관 목록", "[생성 실패] 깨진 화면"]


def test_build_does_not_import_openpyxl():
    import build as build_mod
    import inspect

    src = inspect.getsource(build_mod)
    assert "openpyxl" not in src
