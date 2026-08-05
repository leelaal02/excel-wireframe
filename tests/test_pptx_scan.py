from pathlib import Path

from fixtures import make_empty_layout_pptx, make_template_pptx
from pptx_scan import scan_layouts, scan_presentation, suggest_content_area, suggest_mode


def test_scan_reports_slide_size_and_shapes(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    report = scan_presentation(pptx)
    assert report["slide_width"] == 9906000
    assert round(report["slide_size_in"][0], 2) == 10.83
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
    assert {0, 1, 10, 11, 12} <= set(idxs)
    assert any(s["name"] == "상단띠" for s in target["shapes"])


def test_suggest_content_area_avoids_header_and_footer(tmp_path: Path):
    from default_template import DEFAULT_LAYOUT_NAME, build_default_template
    path = build_default_template(tmp_path / "d.pptx")
    layout_info = next(lay for lay in scan_layouts(path)
                       if lay["name"] == DEFAULT_LAYOUT_NAME)

    area = suggest_content_area(layout_info, 9906000, 6858000)
    left, top, width, height = area

    # 구분선(top 404664 + height 216024 = 620688) 아래에서 시작한다
    assert top >= 620688
    # 하단바(top 6716266) 위에서 끝난다
    assert top + height <= 6716266
    assert width > 0 and height > 0
