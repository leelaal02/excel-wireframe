from pathlib import Path

from fixtures import make_template_pptx
from pptx import Presentation
from slide_clone import clone_slide


def test_clone_copies_all_shapes(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    src = prs.slides[0]
    new = clone_slide(prs, src)
    assert len(new.shapes) == len(src.shapes)
    assert [s.name for s in new.shapes] == [s.name for s in src.shapes]


def test_clone_survives_save_and_reload(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    clone_slide(prs, prs.slides[0])
    out = tmp_path / "out.pptx"
    prs.save(out)

    chk = Presentation(str(out))
    assert len(chk.slides) == 2
    last = chk.slides[-1]
    tables = [s for s in last.shapes if s.has_table]
    assert len(tables) == 5
    assert tables[0].table.cell(0, 0).text == "1"


def test_clone_twice_produces_three_slides(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    src = prs.slides[0]
    clone_slide(prs, src)
    clone_slide(prs, src)
    out = tmp_path / "out.pptx"
    prs.save(out)
    assert len(Presentation(str(out)).slides) == 3
