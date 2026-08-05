# -*- coding: utf-8 -*-
from pathlib import Path

import pytest
from common import Warnings
from pptx import Presentation
from pptx.util import Emu, Pt
from slide_layout import (
    DEFAULT_CONTENT_AREA,
    P_NS,
    add_detail_table,
    add_image_anchor,
    drop_empty_placeholders,
    find_layout,
    inherit_placeholders,
    name_placeholders,
    split_content_area,
)


def _prs():
    """Title and Content 레이아웃(TITLE/OBJECT/DATE/FOOTER/SLIDE_NUMBER)을 가진 프레젠테이션."""
    return Presentation()


def test_find_layout_by_name():
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    assert lay.name == "Title and Content"


def test_find_layout_by_index():
    prs = _prs()
    assert find_layout(prs, 1).name == prs.slide_layouts[1].name


def test_find_layout_raises_for_unknown_name():
    prs = _prs()
    with pytest.raises(ValueError) as exc:
        find_layout(prs, "없는레이아웃")
    assert "없는레이아웃" in str(exc.value)


def test_inherit_placeholders_adds_date_and_footer():
    """python-pptx는 date/footer/slidenumber를 복제하지 않는다 — 우리가 채운다."""
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    before = {ph.placeholder_format.idx for ph in slide.placeholders}
    assert 10 not in before

    added = inherit_placeholders(slide, lay)

    after = {ph.placeholder_format.idx for ph in slide.placeholders}
    assert {10, 11, 12} <= after
    assert set(added) == after - before


def test_inherit_placeholders_assigns_unique_shape_ids():
    """레이아웃 id를 그대로 쓰면 사용자 템플릿에서 중복 id가 생겨 PPT가 복구를 요구한다."""
    prs = _prs()
    lay = find_layout(prs, "Title and Content")

    # 레이아웃의 날짜 placeholder(idx=10)에, 슬라이드가 실제로 쓸 id(제목 placeholder의
    # id)를 일부러 심어 충돌을 재현한다. 이렇게 강제하지 않으면 python-pptx 기본
    # 템플릿에서는 번호가 우연히 겹치지 않아 회귀를 못 잡는다.
    date_ph = next(ph for ph in lay.placeholders if ph.placeholder_format.idx == 10)
    date_ph._element.find(".//{%s}cNvPr" % P_NS).set("id", "2")

    slide = prs.slides.add_slide(lay)
    assert slide.placeholders[0].shape_id == 2  # 제목이 이미 id=2를 쓰고 있음을 확인

    inherit_placeholders(slide, lay)

    ids = [shape.shape_id for shape in slide.shapes]
    assert len(ids) == len(set(ids)), ids


def test_inherit_placeholders_is_idempotent():
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    inherit_placeholders(slide, lay)
    n = len(list(slide.placeholders))
    assert inherit_placeholders(slide, lay) == []
    assert len(list(slide.placeholders)) == n


def test_name_placeholders_renames_by_idx():
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    inherit_placeholders(slide, lay)

    warns = Warnings()
    name_placeholders(
        slide,
        {"title": 0, "screen_id": 1, "작성일": 10},
        {"title": "제목", "screen_id": "화면ID"},
        warns,
        "SCR001",
    )

    names = [s.name for s in slide.shapes]
    assert "제목" in names
    assert "화면ID" in names
    assert "작성일" in names  # shapes에 없으면 키를 그대로 이름으로 쓴다
    assert len(warns) == 0


def test_name_placeholders_warns_for_missing_idx():
    prs = _prs()
    slide = prs.slides.add_slide(find_layout(prs, "Title and Content"))
    warns = Warnings()
    name_placeholders(slide, {"title": 0, "없는것": 99}, {}, warns, "SCR001")
    items = warns.to_list()
    assert len(items) == 1
    assert items[0]["code"] == "shape-not-found"
    assert "99" in items[0]["message"]


def test_drop_empty_placeholders_removes_blank_ones():
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    slide.placeholders[0].text_frame.text = "제목 있음"

    removed = drop_empty_placeholders(slide)

    assert removed >= 1
    remaining = [ph.placeholder_format.idx for ph in slide.placeholders]
    assert 0 in remaining
    assert 1 not in remaining


def test_drop_empty_placeholders_keeps_field_placeholders():
    """쪽번호는 자동 번호 필드라 텍스트가 비어 보여도 지우면 안 된다."""
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    inherit_placeholders(slide, lay)
    drop_empty_placeholders(slide)
    assert 12 in [ph.placeholder_format.idx for ph in slide.placeholders]


def test_split_content_area_puts_tables_at_the_bottom():
    area = DEFAULT_CONTENT_AREA
    image_box, table_boxes = split_content_area(area, 5, 4)

    assert len(table_boxes) == 5
    # 표는 본문 영역 아래쪽에 붙는다
    table_top = table_boxes[0][1]
    table_height = table_boxes[0][3]
    assert table_top + table_height == area[1] + area[3]
    # 이미지는 위쪽 나머지를 전부 쓴다
    assert image_box[1] == area[1]
    assert image_box[1] + image_box[3] == table_top


def test_split_content_area_matches_measured_geometry():
    """실측 조합(표 5개 × 4행)에서 표 상단이 원본과 같은 자리에 온다."""
    _, table_boxes = split_content_area(DEFAULT_CONTENT_AREA, 5, 4)
    assert table_boxes[0][1] == 5253244
    assert table_boxes[0][3] == 382457 + 268746 + 496168 + 268746


def test_split_content_area_divides_width_evenly():
    area = (0, 0, 10000, 4000000)
    _, boxes = split_content_area(area, 5, 4)
    assert [b[0] for b in boxes] == [0, 2000, 4000, 6000, 8000]
    assert all(b[2] == 2000 for b in boxes)


def test_split_content_area_grows_tables_with_rows():
    _, four = split_content_area(DEFAULT_CONTENT_AREA, 5, 4)
    _, six = split_content_area(DEFAULT_CONTENT_AREA, 5, 6)
    assert six[0][3] > four[0][3]
    # 마지막 행높이를 반복한다
    assert six[0][3] == four[0][3] + 268746 * 2


def test_split_content_area_raises_when_image_slot_too_small():
    with pytest.raises(ValueError) as exc:
        split_content_area(DEFAULT_CONTENT_AREA, 5, 40)
    assert "이미지" in str(exc.value)


def test_add_detail_table_applies_measured_formatting():
    prs = _prs()
    slide = prs.slides.add_slide(find_layout(prs, "Title and Content"))
    frame = add_detail_table(slide, (0, 0, 1971135, 1416117), 4, "상세표1")

    assert frame.name == "상세표1"
    table = frame.table
    assert len(table.rows) == 4
    assert len(table.columns) == 2
    assert [r.height for r in table.rows] == [382457, 268746, 496168, 268746]
    # 열폭은 실측 비율을 표 폭에 맞춰 나눈다
    assert sum(c.width for c in table.columns) == 1971135
    assert table.columns[0].width < table.columns[1].width

    num, txt = table.cell(0, 0), table.cell(0, 1)
    assert num.margin_left == 18000
    assert txt.margin_left == 9525
    assert txt.margin_bottom == 0
    assert num.text_frame.paragraphs[0].runs[0].font.size == Pt(6.5)
    assert num.text_frame.paragraphs[0].runs[0].font.bold is True
    assert txt.text_frame.paragraphs[0].runs[0].font.size == Pt(7)
    assert txt.text_frame.paragraphs[0].runs[0].font.name == "맑은 고딕"


def test_add_detail_table_starts_empty():
    """빈 Excel에서 예시 문구가 산출물에 찍히면 안 된다."""
    prs = _prs()
    slide = prs.slides.add_slide(find_layout(prs, "Title and Content"))
    table = add_detail_table(slide, (0, 0, 1971135, 1416117), 4, "상세표1").table
    for r in range(4):
        assert table.cell(r, 0).text == ""
        assert table.cell(r, 1).text == ""


def test_add_image_anchor_uses_the_box():
    prs = _prs()
    slide = prs.slides.add_slide(find_layout(prs, "Title and Content"))
    shp = add_image_anchor(slide, (100, 200, 3000, 4000), "화면이미지")
    assert (shp.left, shp.top, shp.width, shp.height) == (100, 200, 3000, 4000)
    assert shp.name == "화면이미지"
