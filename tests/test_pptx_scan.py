from pathlib import Path

from fixtures import make_empty_layout_pptx, make_template_pptx
from pptx_scan import scan_presentation, suggest_mode


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
