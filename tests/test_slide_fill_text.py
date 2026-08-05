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


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _add_date_field(shape, shown: str) -> None:
    """자동 날짜 필드를 심는다. 사용자 템플릿의 DATE placeholder가 이 꼴이다."""
    from lxml import etree

    p = shape.text_frame.paragraphs[0]._p
    for r in list(p.findall("{%s}r" % A_NS)):
        p.remove(r)
    fld = etree.SubElement(p, "{%s}fld" % A_NS)
    fld.set("id", "{7DE4667A-1AB8-4DCD-A2F5-D16009686B1A}")
    fld.set("type", "datetime1")
    t = etree.SubElement(fld, "{%s}t" % A_NS)
    t.text = shown


def test_set_text_replaces_auto_field(tmp_path: Path):
    """자동 날짜 필드가 있는 자리에 값을 쓰면 필드 값이 남아 이어 붙으면 안 된다.

    a:fld는 런이 아니라서 paragraph.runs에 잡히지 않는다. 그 사실을 모르고
    런만 갈아끼우면 '2026-05-292026-06-25'처럼 두 값이 붙어 나온다.
    """
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    shape = find_shape(prs.slides[0], "제목 13")
    _add_date_field(shape, "2026-05-29")
    assert shape.text_frame.text == "2026-05-29"

    set_text(shape, "2026-06-25")

    assert shape.text_frame.text == "2026-06-25"
    assert shape._element.find(".//{%s}fld" % A_NS) is None
