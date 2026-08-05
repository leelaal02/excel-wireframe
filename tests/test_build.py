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


def test_build_isolates_screen_missing_id(tmp_path: Path):
    """블로킹 발견 3: failed_ids.append(scr["id"])가 except 블록 안에 있으면,
    id가 아예 빠진 screen dict 하나가 scr["id"] 참조에서 또 KeyError를 내며
    빌드 전체를 무너뜨린다 — 결과물이 통째로 안 생긴다. 중간 산출물인
    screens.json에서 id가 빠지는 것(추출이 어긋났거나 사람이 손댄 경우)은
    화면 단위 예외 격리가 정확히 보호해야 하는 시나리오다. 다른 화면들은
    끝까지 완성되고 파일이 반드시 저장돼야 한다."""
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["screens"].append(
        {"name": "id 없는 화면", "images": [], "fields": {},
         "details": [{"no": "1", "desc": "x"}]}
    )
    out = tmp_path / "out.pptx"
    warns = Warnings()

    report = build(data, _mapping(tpl), tmp_path, out, warns)

    assert out.exists()
    assert report["screens"] == 2
    codes = [w["code"] for w in warns.to_list()]
    assert "screen-failed" in codes

    prs = Presentation(str(out))
    titles = [
        next(s for s in sl.shapes if s.name == "제목 13").text_frame.text
        for sl in prs.slides
    ]
    assert "이용기관 목록" in titles


def test_build_mode_layout_gives_readable_error_not_typeerror(tmp_path: Path):
    """블로킹 발견 4: source_slide: null인 채로 clone 모드로 빌드를 시도하면
    int(None)이 raw TypeError를 던졌다. 원인과 대안을 알려주는 ValueError로
    바뀌어야 한다. (mode: layout + source_slide: null 조합은 이제 실제로
    지원되는 정상 경로다 — test_build_layout_mode_* 테스트들이 그 경로를
    검증한다. 이 테스트는 clone 모드에서 source_slide 자체가 없는 경우만
    다룬다.)"""
    tpl = make_template_pptx(tmp_path / "t.pptx")
    mapping = _mapping(tpl)
    mapping["template"]["source_slide"] = None
    out = tmp_path / "out.pptx"

    try:
        build(_screens(2), mapping, tmp_path, out, Warnings())
        assert False, "ValueError가 나야 한다"
    except TypeError:
        raise AssertionError("원시 TypeError가 그대로 새고 있습니다")
    except ValueError as exc:
        msg = str(exc)
        assert "source_slide" in msg
        assert "--template" in msg or "기본 템플릿" in msg


def test_build_does_not_import_openpyxl():
    import build as build_mod
    import inspect

    src = inspect.getsource(build_mod)
    assert "openpyxl" not in src


def _mapping_with_meta_shape(template: Path) -> dict:
    m = _mapping(template)
    m["template"]["shapes"]["프로젝트명"] = "텍스트 개체 틀 14"
    return m


def test_build_fills_meta_into_named_shape(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["meta"]["프로젝트명"] = "통합관리시스템"
    out = tmp_path / "out.pptx"
    build(data, _mapping_with_meta_shape(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    shp = next(s for s in slide.shapes if s.name == "텍스트 개체 틀 14")
    assert shp.text_frame.text == "통합관리시스템"


def test_build_screen_fields_beat_document_meta(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["meta"]["프로젝트명"] = "문서 전체 값"
    data["screens"][0]["fields"]["프로젝트명"] = "화면별 값"
    out = tmp_path / "out.pptx"
    build(data, _mapping_with_meta_shape(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    shp = next(s for s in slide.shapes if s.name == "텍스트 개체 틀 14")
    assert shp.text_frame.text == "화면별 값"


def test_build_ignores_meta_without_matching_shape(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["meta"]["아무도_안_쓰는_키"] = "값"
    out = tmp_path / "out.pptx"
    warns = Warnings()
    build(data, _mapping(tpl), tmp_path, out, warns)
    assert "shape-not-found" not in [w["code"] for w in warns.to_list()]


def test_build_meta_does_not_override_title_shape(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["meta"]["title"] = "문서 제목이 화면명을 덮으면 안 된다"
    out = tmp_path / "out.pptx"
    build(data, _mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    shp = next(s for s in slide.shapes if s.name == "제목 13")
    assert shp.text_frame.text == "이용기관 목록"


from pptx.util import Emu


def _layout_template(path: Path) -> Path:
    """placeholder만 있는 레이아웃형 템플릿 (예시 슬라이드 없음)."""
    prs = Presentation()
    prs.slide_width = Emu(9906000)
    prs.slide_height = Emu(6858000)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path


def _layout_mapping(template: Path) -> dict:
    return {
        "version": 1,
        "excel": {"layout": "sheet-per-screen"},
        "template": {
            "file": str(template),
            "mode": "layout",
            "layout": "Title and Content",
            "placeholders": {"title": 0, "screen_id": 1, "작성일": 10,
                             "문서제목": 11},
            "shapes": {
                "title": "제목",
                "screen_id": "화면ID",
                "image": "화면이미지",
                "작성일": "작성일",
                "문서제목": "문서제목",
                "detail_tables": ["상세표1", "상세표2", "상세표3",
                                  "상세표4", "상세표5"],
            },
            "content_area": [-12319, 337940, 9957099, 6331421],
            "detail_tables": {"count": 5, "rows": 4},
            "table_columns": {"no": 0, "text": 1},
        },
        "options": {
            "detail_text_source": "desc",
            "overflow": "split",
            "clear_unused_slots": True,
        },
    }


def test_build_layout_mode_creates_slides(tmp_path: Path):
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    report = build(_screens(6), _layout_mapping(tpl), tmp_path, out, Warnings())

    assert report["slides"] == 1
    slide = Presentation(str(out)).slides[0]
    by_name = {s.name: s for s in slide.shapes}
    assert by_name["제목"].text_frame.text == "이용기관 목록"
    assert by_name["화면ID"].text_frame.text == "SCR001"


def test_build_layout_mode_fills_detail_tables(tmp_path: Path):
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    build(_screens(6), _layout_mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) == 5
    first = next(t for t in tables if t.name == "상세표1").table
    assert first.cell(0, 0).text == "1"
    assert first.cell(0, 1).text == "설명 1"


def test_build_layout_mode_splits_when_details_exceed_slots(tmp_path: Path):
    """슬롯은 5표 × 4행 = 20개다."""
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    warns = Warnings()
    report = build(_screens(25), _layout_mapping(tpl), tmp_path, out, warns)

    assert report["slides"] == 2
    assert report["split"] == ["SCR001"]
    assert any(w["code"] == "slide-split" for w in warns.to_list())


def test_build_layout_mode_places_image(tmp_path: Path):
    img = make_png(tmp_path / "images" / "SCR001.png")
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    build(_screens(3, image="images/SCR001.png"), _layout_mapping(tpl),
          tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    assert any("PICTURE" in str(s.shape_type) for s in slide.shapes)
    assert not any(s.name == "화면이미지" for s in slide.shapes)


def test_build_layout_mode_without_image_shape_leaves_no_orphan_anchor(tmp_path: Path):
    """최종 리뷰 지적 1: 이미지 자리를 그리는 쪽(_new_layout_slide)은 기본 이름
    '화면이미지'를 쓰고, 그것을 그림으로 바꾸는 쪽(_fill_page)은 기본값 없이
    shapes.image를 읽었다. 그래서 매핑에서 shapes.image를 빼면 본문 영역만 한
    회색 사각형이 슬라이드마다 그대로 남았다 — 아무 경고도 없이. 레이아웃에는
    이미지 placeholder가 없으니 이름을 지을 게 없다고 판단하는 매핑 작성자가
    실제로 밟을 수 있는 길이다."""
    tpl = _layout_template(tmp_path / "t.pptx")
    mapping = _layout_mapping(tpl)
    del mapping["template"]["shapes"]["image"]
    out = tmp_path / "out.pptx"
    warns = Warnings()

    report = build(_screens(3, image="images/SCR001.png"), mapping, tmp_path,
                   out, warns)

    assert report["slides"] == 1
    slide = Presentation(str(out)).slides[0]
    names = [s.name for s in slide.shapes]
    assert "화면이미지" not in names
    # 이미지 자리로 쓰이던 큰 사각형이 다른 이름으로 남지도 않아야 한다.
    # (표 5개와 placeholder만 남는다)
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    autoshapes = [s for s in slide.shapes
                  if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    assert autoshapes == [], [s.name for s in autoshapes]
    assert len([s for s in slide.shapes if s.has_table]) == 5
    assert warns.to_list() == []


def test_build_layout_mode_drops_empty_placeholders(tmp_path: Path):
    """값이 없는 문서제목 placeholder가 산출물에 남으면 안 된다."""
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    screens = _screens(3)
    screens["meta"] = {"source": "s.xlsx", "template": "t.pptx"}  # 문서제목 없음
    build(screens, _layout_mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    assert "문서제목" not in [s.name for s in slide.shapes]


def test_build_layout_mode_fills_doc_title_from_meta(tmp_path: Path):
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    screens = _screens(3)
    screens["meta"]["문서제목"] = "발행기관관리"
    build(screens, _layout_mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    by_name = {s.name: s for s in slide.shapes}
    assert by_name["문서제목"].text_frame.text == "발행기관관리"


def test_build_layout_mode_raises_for_unknown_layout(tmp_path: Path):
    import pytest
    tpl = _layout_template(tmp_path / "t.pptx")
    mapping = _layout_mapping(tpl)
    mapping["template"]["layout"] = "없는레이아웃"
    with pytest.raises(ValueError) as exc:
        build(_screens(3), mapping, tmp_path, tmp_path / "out.pptx", Warnings())
    assert "없는레이아웃" in str(exc.value)


def test_build_layout_mode_raises_for_global_row_misconfig_not_per_screen(
    tmp_path: Path,
):
    """rows_per_table이 과도해 이미지 자리가 안 나오는 것은 화면 하나의 사고가
    아니라 mapping 전체에 걸린 설정 오류다. 화면 단위 격리(screen-failed)로
    흩어지면 build()가 조용히 끝나면서 모든 슬라이드가 '[생성 실패]'가 되어
    버린다 — 여기서는 그 대신 ValueError가 루프 밖으로 곧장 터져야 한다."""
    import pytest
    tpl = _layout_template(tmp_path / "t.pptx")
    mapping = _layout_mapping(tpl)
    mapping["template"]["detail_tables"]["rows"] = 40
    warns = Warnings()
    with pytest.raises(ValueError):
        build(_screens(6), mapping, tmp_path, tmp_path / "out.pptx", warns)
    assert "screen-failed" not in [w["code"] for w in warns.to_list()]


def test_build_clone_mode_still_works(tmp_path: Path):
    """layout 모드를 더해도 clone 경로는 그대로여야 한다."""
    tpl = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    report = build(_screens(6), _mapping(tpl), tmp_path, out, Warnings())
    assert report["slides"] == 1
