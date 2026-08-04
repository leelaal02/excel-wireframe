from pathlib import Path

from fixtures import make_template_pptx
from pptx import Presentation
from pptx.util import Pt
from slide_fill import estimate_overflow, find_shape, set_cell_text, set_text


def test_find_shape_by_name(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    slide = prs.slides[0]
    assert find_shape(slide, "제목 13") is not None
    assert find_shape(slide, "없는도형") is None


def test_set_text_keeps_font_size(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    slide = prs.slides[0]
    shape = find_shape(slide, "제목 13")
    shape.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    set_text(shape, "이용기관 목록")

    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.text == "이용기관 목록"
    assert run.font.size == Pt(18)
    assert len(shape.text_frame.paragraphs) == 1


def test_set_cell_text_splits_newlines(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    table = next(s for s in prs.slides[0].shapes if s.has_table).table
    cell = table.cell(0, 1)
    cell.text_frame.paragraphs[0].runs[0].font.size = Pt(9)

    set_cell_text(cell, "- 구분값 : 전체, Y, N\n- 디폴트 : 전체")

    paras = cell.text_frame.paragraphs
    assert len(paras) == 2
    assert paras[0].runs[0].text == "- 구분값 : 전체, Y, N"
    assert paras[1].runs[0].text == "- 디폴트 : 전체"
    assert paras[0].runs[0].font.size == Pt(9)
    assert paras[1].runs[0].font.size == Pt(9)


def test_set_cell_text_clears_previous_content(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    table = next(s for s in prs.slides[0].shapes if s.has_table).table
    cell = table.cell(0, 1)
    set_cell_text(cell, "첫 줄\n둘째 줄")
    set_cell_text(cell, "짧게")
    assert cell.text == "짧게"


def test_set_cell_text_accepts_empty(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    table = next(s for s in prs.slides[0].shapes if s.has_table).table
    cell = table.cell(0, 1)
    set_cell_text(cell, "")
    assert cell.text == ""


def test_estimate_overflow():
    assert estimate_overflow("짧은 글", 1974850) is False
    assert estimate_overflow("가" * 400, 1974850) is True
