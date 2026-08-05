# -*- coding: utf-8 -*-
"""레이아웃을 골라 슬라이드를 만들고, 그 레이아웃이 물려주는 자리에 값을 채운다.

디자인은 레이아웃이 담당한다. 이 모듈은 자리를 찾아 이름을 붙이고, 레이아웃에
자리가 없는 두 가지(화면 이미지, 상세표)만 본문 영역 안에 만든다.
"""
from __future__ import annotations

import copy

from common import EMU_PER_INCH
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def find_layout(prs, spec):
    """이름 또는 인덱스로 레이아웃을 찾는다.

    마스터가 여러 개인 템플릿이 흔하므로 전부 훑는다. python-pptx의
    prs.slide_layouts는 첫 번째 마스터만 보여줘서 그것만으로는 부족하다.
    """
    layouts = [lay for master in prs.slide_masters for lay in master.slide_layouts]

    if isinstance(spec, int):
        if 0 <= spec < len(layouts):
            return layouts[spec]
        raise ValueError(
            "레이아웃 인덱스 %d가 범위를 벗어납니다 (레이아웃 %d개)"
            % (spec, len(layouts))
        )

    matches = [lay for lay in layouts if lay.name == spec]
    if not matches:
        raise ValueError(
            "레이아웃 '%s'을(를) 찾지 못했습니다. 있는 레이아웃: %s"
            % (spec, ", ".join(lay.name for lay in layouts))
        )
    if len(matches) > 1:
        # 마스터가 여러 개면 같은 이름이 겹칠 수 있다. 멈출 일은 아니지만
        # 어느 것을 골랐는지는 알려야 한다.
        import sys
        print("경고: 레이아웃 '%s'이(가) %d개 있어 첫 번째를 씁니다"
              % (spec, len(matches)), file=sys.stderr)
    return matches[0]


def inherit_placeholders(slide, layout) -> list[int]:
    """레이아웃에 있으나 슬라이드에 없는 placeholder를 복제한다.

    python-pptx의 add_slide는 date/footer/slidenumber 계열을 복제하지 않는다.
    PowerPoint 관례상 그 셋은 마스터 설정으로 표시되기 때문인데, 우리는 거기에
    값을 써야 하므로 직접 옮긴다.

    레이아웃 XML의 cNvPr/@id는 그대로 들고 오면 안 된다. id는 슬라이드 XML
    전체에서 유일해야 하는데(OOXML 규칙), 레이아웃 쪽 id와 슬라이드가 이미 쓴
    id가 우연히 겹칠 수 있다 — 특히 사용자가 준 임의의 템플릿에서는 흔하다.
    겹치면 PowerPoint가 파일을 열 때 복구를 요구한다. python-pptx 자신도
    clone_placeholder에서 새 id를 매길 때 같은 방식(문서 전체 최댓값+1)을 쓴다.
    """
    have = {ph.placeholder_format.idx for ph in slide.placeholders}
    spTree = slide.shapes._spTree
    added: list[int] = []
    for ph in layout.placeholders:
        idx = ph.placeholder_format.idx
        if idx in have:
            continue
        new_sp = copy.deepcopy(ph._element)
        cNvPr = new_sp.find(".//{%s}cNvPr" % P_NS)
        cNvPr.set("id", str(spTree.max_shape_id + 1))
        spTree.append(new_sp)
        added.append(idx)
    return added


def name_placeholders(slide, placeholders_cfg, shapes_cfg, warns, screen_id) -> None:
    """placeholder에 mapping이 정한 이름을 붙인다.

    add_slide가 주는 이름은 'Title 1', 'Content Placeholder 2'처럼 그때그때
    달라진다. verify.py는 template.shapes의 *이름*으로 도형을 찾으므로,
    이름을 고정해 두지 않으면 검증이 느슨한 경로로 떨어진다.
    """
    by_idx = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
    for key, idx in (placeholders_cfg or {}).items():
        ph = by_idx.get(int(idx))
        if ph is None:
            warns.add(screen_id, "shape-not-found",
                      "레이아웃에 placeholder idx=%s('%s')가 없습니다" % (idx, key))
            continue
        name = shapes_cfg.get(key, key)
        if isinstance(name, str):
            ph.name = name


def _has_field(shape) -> bool:
    """쪽번호처럼 자동 필드를 담은 도형인가."""
    return shape._element.find(".//{%s}fld" % A_NS) is not None


def drop_empty_placeholders(slide) -> int:
    """값이 없는 placeholder를 지운다.

    남겨 두면 PowerPoint가 '제목을 입력하십시오' 프롬프트를 그려서 산출물에
    빈 안내 문구가 보인다. 자동 필드(쪽번호)는 텍스트가 비어 보여도 남긴다.
    """
    removed = 0
    for ph in list(slide.placeholders):
        if _has_field(ph):
            continue
        if ph.has_text_frame and ph.text_frame.text.strip():
            continue
        ph._element.getparent().remove(ph._element)
        removed += 1
    return removed


# --- 실측 상수 (원본 화면설계서 화면 페이지) ---
DEFAULT_CONTENT_AREA = (-12319, 337940, 9957099, 6331421)
ROW_HEIGHTS = [382457, 268746, 496168, 268746]
COL_WIDTH_RATIO = (160215, 1810920)
MIN_IMAGE_HEIGHT_EMU = EMU_PER_INCH  # 1인치

SLOT_BORDER = RGBColor(0xBB, 0xBB, 0xBB)
SLOT_FILL = RGBColor(0xF5, 0xF6, 0xF8)


def _row_heights(rows_per_table: int) -> list[int]:
    """실측 행높이를 쓰되, 4행을 넘으면 마지막 값을 반복한다."""
    if rows_per_table <= len(ROW_HEIGHTS):
        return ROW_HEIGHTS[:rows_per_table]
    tail = [ROW_HEIGHTS[-1]] * (rows_per_table - len(ROW_HEIGHTS))
    return ROW_HEIGHTS + tail


def split_content_area(area, table_count: int, rows_per_table: int):
    """본문 영역을 이미지 자리와 상세표 자리로 나눈다.

    표는 아래쪽에 붙이고 폭을 균등 분할한다. 원본의 표 간격은 0.01인치라
    사실상 붙어 있으므로 간격을 두지 않는다. 이미지는 위쪽 나머지를 전부 쓴다.
    """
    left, top, width, height = area
    table_h = sum(_row_heights(rows_per_table))
    table_top = top + height - table_h
    image_h = table_top - top

    if image_h < MIN_IMAGE_HEIGHT_EMU:
        max_rows = 0
        while sum(_row_heights(max_rows + 1)) <= height - MIN_IMAGE_HEIGHT_EMU:
            max_rows += 1
        raise ValueError(
            "rows_per_table=%d면 이미지 자리 높이가 %.2fin로 너무 작아집니다"
            "(최소 %.2fin 필요). 이 content_area 높이(%.2fin)에서는 "
            "rows_per_table을 %d 이하로 쓰세요."
            % (rows_per_table, image_h / EMU_PER_INCH,
               MIN_IMAGE_HEIGHT_EMU / EMU_PER_INCH, height / EMU_PER_INCH,
               max_rows)
        )

    table_w = width // table_count
    table_boxes = [
        (left + table_w * i, table_top, table_w, table_h)
        for i in range(table_count - 1)
    ]
    # 나머지(정수 나눗셈 몫에서 버려지는 폭)는 마지막 표에 몰아준다 — 그래야
    # 마지막 표의 오른쪽 끝이 본문 영역의 오른쪽 끝과 정확히 맞는다.
    last_left = left + table_w * (table_count - 1)
    last_width = width - table_w * (table_count - 1)
    table_boxes.append((last_left, table_top, last_width, table_h))
    return (left, top, width, image_h), table_boxes


def add_image_anchor(slide, box, name: str):
    """이미지가 들어갈 자리 사각형. place_image가 이것을 지우고 그림으로 바꾼다.

    이미지가 없는 화면에서는 이 사각형이 그대로 남아 '여기에 스크린샷' 자리로
    보인다 — clone 모드에서 템플릿의 이미지 자리 도형이 남는 것과 같은 동작이다.
    """
    left, top, width, height = box
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(left), Emu(top),
                                 Emu(width), Emu(height))
    shp.name = name
    shp.fill.solid()
    shp.fill.fore_color.rgb = SLOT_FILL
    shp.line.color.rgb = SLOT_BORDER
    shp.text_frame.text = ""
    return shp


def _format_cell(cell, size_pt, bold, align, anchor, margin, margin_bottom, font=None):
    cell.margin_left = Emu(margin)
    cell.margin_right = Emu(margin)
    cell.margin_top = Emu(margin)
    cell.margin_bottom = Emu(margin_bottom)
    if anchor is not None:
        cell.vertical_anchor = anchor
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = ""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if font:
        run.font.name = font


def add_detail_table(slide, box, rows_per_table: int, name: str):
    """상세표 하나를 만들고 실측 서식을 적용한다. 셀은 비운 채로 둔다."""
    left, top, width, height = box
    frame = slide.shapes.add_table(rows_per_table, 2, Emu(left), Emu(top),
                                   Emu(width), Emu(height))
    frame.name = name
    table = frame.table

    ratio_total = sum(COL_WIDTH_RATIO)
    no_w = width * COL_WIDTH_RATIO[0] // ratio_total
    table.columns[0].width = Emu(no_w)
    table.columns[1].width = Emu(width - no_w)
    for ri, rh in enumerate(_row_heights(rows_per_table)):
        table.rows[ri].height = Emu(rh)

    for r in range(rows_per_table):
        _format_cell(table.cell(r, 0), 6.5, True, PP_ALIGN.CENTER,
                     MSO_ANCHOR.MIDDLE, 18000, 18000)
        _format_cell(table.cell(r, 1), 7.0, False, PP_ALIGN.LEFT, None,
                     9525, 0, font="맑은 고딕")
    return frame
