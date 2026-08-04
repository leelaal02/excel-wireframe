from pathlib import Path

import pytest
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


def test_default_template_table_height_scales_with_rows_per_table(tmp_path: Path):
    """이전 세로-간격 테스트는 rows_per_table과 무관하게 고정된 표 높이
    공식에서도 통과했다 — 표 높이가 선언된 기하 안에서 이미지 자리와
    안 겹치기만 하면 됐기 때문에, 옛 코드(표 높이 1.55in 고정)를 그대로
    둬도 이 assertion만으로는 안 잡혔다. 실제로 달라져야 하는 성질(표
    높이가 행 수에 비례해 커진다) 그 자체를 직접 비교한다."""
    p4 = build_default_template(tmp_path / "d4.pptx", rows_per_table=4)
    p8 = build_default_template(tmp_path / "d8.pptx", rows_per_table=8)

    t4 = next(s for s in Presentation(str(p4)).slides[0].shapes if s.has_table)
    t8 = next(s for s in Presentation(str(p8)).slides[0].shapes if s.has_table)

    assert t8.height > t4.height
    assert t8.height == pytest.approx(t4.height * 2, rel=0.01)


def test_default_template_image_slot_does_not_overlap_tables(tmp_path: Path):
    """가로축(표끼리 겹치지 않음)은 이미 검사한다. 세로축은 아무도 다시 안 본다 —
    표 높이가 rows_per_table에 비례해 커지면 이미지 자리가 슬라이드 하단
    표 위치까지 침범할 수 있다. 위 스케일 테스트와 짝을 이룬다: 그 스케일이
    실제로 이미지 자리를 침범하지 않는지는 여기서 확인한다."""
    p = build_default_template(tmp_path / "d.pptx", rows_per_table=8)
    prs = Presentation(str(p))
    shapes = prs.slides[0].shapes
    img_slot = next(s for s in shapes if s.name == DEFAULT_SHAPE_NAMES["image"])
    tables = [s for s in shapes if s.has_table]
    assert img_slot.top + img_slot.height <= tables[0].top
    for t in tables:
        assert t.top + t.height <= prs.slide_height


def test_default_template_raises_when_rows_per_table_too_large(tmp_path: Path):
    """회귀 4: table_h가 rows_per_table에 비례하므로 무한정 키우면 이미지
    자리 높이가 음수가 될 수 있다(측정: rows=17 → -0.14in). 조용히 겹치는
    도형을 만드는 대신, 기본 슬라이드 크기(7.5in 높이)에서 실제로 쓸 수
    있는 최대값을 알려주는 ValueError로 막아야 한다."""
    with pytest.raises(ValueError) as excinfo:
        build_default_template(tmp_path / "d.pptx", rows_per_table=17)
    msg = str(excinfo.value)
    assert "rows_per_table" in msg
    assert "14" in msg  # 기본 슬라이드 크기에서 이미지 자리 최소 1in을 지키는 상한


def test_default_template_allows_max_workable_rows_per_table(tmp_path: Path):
    """위 테스트가 말하는 상한(14)에서는 여전히 만들어져야 한다 — 경계값이
    실수로 한 칸 안쪽으로 당겨지지 않았는지 확인한다."""
    build_default_template(tmp_path / "d.pptx", rows_per_table=14)


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


def test_default_template_has_meta_shapes(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    names = [s.name for s in Presentation(str(p)).slides[0].shapes]
    assert "문서제목" in names
    assert "작성일" in names


def test_default_template_meta_shapes_are_empty(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    slide = Presentation(str(p)).slides[0]
    for name in ("문서제목", "작성일"):
        shp = next(s for s in slide.shapes if s.name == name)
        assert shp.text_frame.text == ""


def test_default_template_meta_shapes_stay_inside_slide(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    prs = Presentation(str(p))
    slide = prs.slides[0]
    for name in ("문서제목", "작성일"):
        shp = next(s for s in slide.shapes if s.name == name)
        assert shp.left >= 0
        assert shp.top >= 0
        assert shp.left + shp.width <= prs.slide_width
        assert shp.top + shp.height <= prs.slide_height


def test_default_template_mapping_includes_meta_shapes(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    shapes = default_template_mapping(p)["shapes"]
    assert shapes["문서제목"] == "문서제목"
    assert shapes["작성일"] == "작성일"
