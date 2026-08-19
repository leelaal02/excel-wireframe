from pathlib import Path

import pytest
from fixtures import make_empty_layout_pptx, make_template_pptx
from pptx_scan import scan_layouts, scan_presentation, suggest_content_area, suggest_mode

ROOT = Path(__file__).resolve().parent.parent
SAMPLE_PPTX = ROOT / "화면설계서_저작권_발행기관.정산처_발행기관관리_v1.0_20260427.pptx"


def test_scan_reports_slide_size_and_shapes(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    report = scan_presentation(pptx)
    assert report["slide_width"] == 9144000
    assert round(report["slide_size_in"][0], 2) == 10.0
    slide = report["slides"][0]
    names = [s["name"] for s in slide["shapes"]]
    assert "제목 13" in names
    assert "그림 18" in names
    assert names.count("표 7") == 1


def test_scan_reports_table_dimensions(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    report = scan_presentation(pptx)
    tables = [s for s in report["slides"][0]["shapes"] if s["table"]]
    assert len(tables) == 5
    assert tables[0]["table"] == {"rows": 4, "cols": 2}


def test_suggest_mode_picks_clone_for_example_slide(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    got = suggest_mode(scan_presentation(pptx))
    assert got["mode"] == "clone"
    assert got["source_slide"] == 0


def test_suggest_mode_falls_back_to_layout(tmp_path: Path):
    pptx = make_empty_layout_pptx(tmp_path / "e.pptx")
    got = suggest_mode(scan_presentation(pptx))
    assert got["mode"] == "layout"
    assert got["source_slide"] is None


def test_scan_layouts_lists_placeholders(tmp_path: Path):
    from default_template import DEFAULT_LAYOUT_NAME, build_default_template
    path = build_default_template(tmp_path / "d.pptx")
    layouts = scan_layouts(path)
    names = [lay["name"] for lay in layouts]
    assert DEFAULT_LAYOUT_NAME in names
    target = next(lay for lay in layouts if lay["name"] == DEFAULT_LAYOUT_NAME)
    idxs = [ph["idx"] for ph in target["placeholders"]]
    # 화면명·작성일은 placeholder가 아니라 상단 메타 표가 담당한다
    assert set(idxs) == {1, 11, 12}
    shape_names = {s["name"] for s in target["shapes"]}
    assert {"메타표1", "메타표2", "본문박스", "하단바"} <= shape_names


def test_suggest_content_area_avoids_header_and_footer(tmp_path: Path):
    from default_template import DEFAULT_LAYOUT_NAME, build_default_template
    path = build_default_template(tmp_path / "d.pptx")
    layout_info = next(lay for lay in scan_layouts(path)
                       if lay["name"] == DEFAULT_LAYOUT_NAME)

    area = suggest_content_area(layout_info, 9144000, 6858000)
    left, top, width, height = area

    # 메타표2(top 195617 + height 116632 = 312249) 아래에서 시작한다
    assert top >= 312249
    # 하단바(top 6716266) 위에서 끝난다
    assert top + height <= 6716266
    assert width > 0 and height > 0


def test_suggest_content_area_ignores_shape_crossing_vertical_center():
    """세로 중앙을 걸치는 넓은 도형은 장애물로 보지 않고 무시해야 한다.

    실제 샘플(화면설계서_저작권_발행기관.정산처_발행기관관리_v1.0_20260427.pptx)의
    마스터 0 레이아웃 0에 있는 '그룹 22'가 실례다 — 좌표
    (0, 188657, 9892977, 6480704)로 슬라이드의 top(0)~mid(3429000)~
    bottom(6858000)을 모두 걸치는 전폭 장식 배경이다. 이런 도형까지
    상단/하단 장애물로 잡으면 top이 중앙까지 밀리거나 bottom이 중앙까지
    당겨져 본문 영역이 사라진다. 여기서는 진짜 상단띠/하단바만 경계로
    반영되어야 한다.
    """
    layout_info = {
        "shapes": [
            {"name": "상단띠", "left": 0, "top": 0, "width": 9896172, "height": 137234},
            {"name": "그룹 22", "left": 0, "top": 188657,
             "width": 9892977, "height": 6480704},
            {"name": "하단바", "left": 0, "top": 6716266, "width": 9906000, "height": 144000},
        ],
        "placeholders": [],
    }
    area = suggest_content_area(layout_info, 9906000, 6858000)
    left, top, width, height = area

    assert top == 137234            # 상단띠 하단에서 시작 (그룹 22는 무시됨)
    assert top + height == 6716266  # 하단바 상단에서 끝남
    assert width > 0 and height > 0


@pytest.mark.skipif(not SAMPLE_PPTX.exists(), reason="실제 샘플 파일이 없습니다")
def test_scan_layouts_covers_all_masters_in_real_sample():
    """실제 샘플은 마스터가 둘(17개 + 6개 레이아웃)이다.

    scan_layouts가 첫 마스터만 훑고 끝나는 회귀를 잡기 위한 테스트다.
    """
    layouts = scan_layouts(SAMPLE_PPTX)
    masters = {lay["master"] for lay in layouts}
    assert masters == {0, 1}
    assert sum(1 for lay in layouts if lay["master"] == 0) == 17
    assert sum(1 for lay in layouts if lay["master"] == 1) == 6
