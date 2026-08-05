# -*- coding: utf-8 -*-
from pathlib import Path

import pytest
from common import Warnings
from pptx import Presentation
from pptx.util import Emu
from slide_layout import (
    drop_empty_placeholders,
    find_layout,
    inherit_placeholders,
    name_placeholders,
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
