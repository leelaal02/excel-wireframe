from pathlib import Path

from default_template import (
    DEFAULT_SHAPE_NAMES,
    build_default_template,
    default_template_mapping,
)
from pptx import Presentation


def test_default_template_is_16_9(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    prs = Presentation(str(p))
    assert prs.slide_width == 12192000
    assert prs.slide_height == 6858000
    assert len(prs.slides) == 1


def test_default_template_has_named_shapes(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    names = [s.name for s in Presentation(str(p)).slides[0].shapes]
    assert DEFAULT_SHAPE_NAMES["title"] in names
    assert DEFAULT_SHAPE_NAMES["screen_id"] in names
    assert DEFAULT_SHAPE_NAMES["image"] in names
    assert "상세표1" in names and "상세표5" in names


def test_default_template_tables_have_20_slots(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    tables = [s for s in Presentation(str(p)).slides[0].shapes if s.has_table]
    assert len(tables) == 5
    assert sum(len(t.table.rows) for t in tables) == 20
    assert tables[0].table.cell(0, 0).text == "1"


def test_default_template_tables_do_not_overlap(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    prs = Presentation(str(p))
    tables = sorted(
        (s for s in prs.slides[0].shapes if s.has_table), key=lambda s: s.left
    )
    for a, b in zip(tables, tables[1:]):
        assert a.left + a.width <= b.left + 1
    assert tables[-1].left + tables[-1].width <= prs.slide_width


def test_default_template_size_is_configurable(tmp_path: Path):
    p = build_default_template(
        tmp_path / "d.pptx", slide_width_emu=9906000, table_count=3
    )
    prs = Presentation(str(p))
    assert prs.slide_width == 9906000
    assert sum(1 for s in prs.slides[0].shapes if s.has_table) == 3


def test_default_template_image_slot_does_not_overlap_tables(tmp_path: Path):
    """가로축(표끼리 겹치지 않음)은 이미 검사한다. 세로축은 아무도 다시 안 본다 —
    표 높이가 rows_per_table에 비례해 커지면 이미지 자리가 슬라이드 하단
    표 위치까지 침범할 수 있다."""
    p = build_default_template(tmp_path / "d.pptx", rows_per_table=8)
    prs = Presentation(str(p))
    shapes = prs.slides[0].shapes
    img_slot = next(s for s in shapes if s.name == DEFAULT_SHAPE_NAMES["image"])
    tables = [s for s in shapes if s.has_table]
    assert img_slot.top + img_slot.height <= tables[0].top
    for t in tables:
        assert t.top + t.height <= prs.slide_height


def test_default_template_mapping_matches_shapes(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    mp = default_template_mapping(p)
    names = [s.name for s in Presentation(str(p)).slides[0].shapes]
    assert mp["mode"] == "clone"
    assert mp["source_slide"] == 0
    assert mp["file"] == str(p)
    assert mp["shapes"]["title"] in names
    assert mp["shapes"]["detail_tables"] == [
        "상세표1", "상세표2", "상세표3", "상세표4", "상세표5"
    ]
    assert mp["table_columns"] == {"no": 0, "text": 1}
