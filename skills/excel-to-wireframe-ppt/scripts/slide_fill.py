# -*- coding: utf-8 -*-
"""슬라이드 도형에 값을 채운다.

핵심은 서식 보존이다. 런을 전부 지우고 새로 만들면 폰트·크기·색이 초기화되어
템플릿 디자인이 무너지므로, 첫 런의 텍스트만 갈아끼우는 방식을 쓴다.
"""
from __future__ import annotations

import copy
from pathlib import Path

from PIL import Image

EMU_PER_INCH = 914400


def find_shape(slide, name: str):
    for shp in slide.shapes:
        if shp.name == name:
            return shp
    return None


def _fill_text_frame(tf, text: str) -> None:
    lines = str(text).split("\n") if text else [""]
    p0 = tf.paragraphs[0]

    if p0.runs:
        base_run = p0.runs[0]
    else:
        base_run = p0.add_run()
    rPr = base_run._r.find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"
    )
    base_rPr = copy.deepcopy(rPr) if rPr is not None else None

    base_run.text = lines[0]
    for extra in p0.runs[1:]:
        extra._r.getparent().remove(extra._r)
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)

    for line in lines[1:]:
        p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        if base_rPr is not None:
            run._r.insert(0, copy.deepcopy(base_rPr))


def set_text(shape, text: str) -> None:
    if not shape.has_text_frame:
        return
    _fill_text_frame(shape.text_frame, text)


def set_cell_text(cell, text: str) -> None:
    _fill_text_frame(cell.text_frame, text)


def estimate_overflow(text: str, cell_width_emu: int, limit_chars: int = 60) -> bool:
    """셀 폭 대비 글자 수로 잘림 가능성을 추정한다.

    정확한 텍스트 측정은 폰트 메트릭이 필요해 과하다. 자동 축소로 서식을 무너뜨리는
    것보다 사람이 확인하도록 경고만 올리는 편이 낫다.
    """
    if not text:
        return False
    inches = max(cell_width_emu / EMU_PER_INCH, 0.1)
    return len(str(text)) > limit_chars * inches


def collect_tables(slide, names: list[str] | None):
    """상세 표를 슬롯 순서대로 모은다.

    이름이 지정되면 그 순서를 그대로 따른다. 없으면 좌→우로 정렬한다 —
    화면상 왼쪽 표가 앞 번호를 담는 것이 사람의 읽기 순서와 맞기 때문이다.
    """
    tables = [s for s in slide.shapes if s.has_table]
    if names:
        by_name = {t.name: t for t in tables}
        return [by_name[n] for n in names if n in by_name]
    return sorted(tables, key=lambda t: (t.left or 0, t.top or 0))


def count_slots(tables) -> int:
    return sum(len(t.table.rows) for t in tables)


def place_image(slide, anchor_shape, image_path: Path):
    """앵커 도형 자리에 종횡비를 유지해 이미지를 넣고 앵커는 지운다."""
    left = anchor_shape.left
    top = anchor_shape.top
    box_w = anchor_shape.width
    box_h = anchor_shape.height

    with Image.open(image_path) as im:
        iw, ih = im.size

    scale = min(box_w / iw, box_h / ih)
    new_w = int(iw * scale)
    new_h = int(ih * scale)
    new_left = left + (box_w - new_w) // 2
    new_top = top + (box_h - new_h) // 2

    pic = slide.shapes.add_picture(str(image_path), new_left, new_top, new_w, new_h)
    anchor_shape._element.getparent().remove(anchor_shape._element)
    return pic


def fill_slots(
    tables,
    details: list[dict],
    cols: dict,
    text_key: str,
    clear_unused: bool,
    warns,
    screen_id: str,
) -> int:
    """상세 항목을 표1 r0…rN → 표2 r0…rN 순으로 흘려 넣는다."""
    no_col = int(cols.get("no", 0))
    text_col = int(cols.get("text", 1))

    slots = []
    for t in tables:
        table = t.table
        width = 0
        if text_col < len(table.columns):
            width = int(table.columns[text_col].width or 0)
        for r in range(len(table.rows)):
            slots.append((table, r, width))

    filled = 0
    for i, (table, row, width) in enumerate(slots):
        if i < len(details):
            d = details[i]
            text = str(d.get(text_key, "") or "")
            if no_col < len(table.columns):
                set_cell_text(table.cell(row, no_col), str(d.get("no", "") or ""))
            if text_col < len(table.columns):
                set_cell_text(table.cell(row, text_col), text)
            if estimate_overflow(text, width):
                warns.add(screen_id, "text-overflow",
                          "%s번 항목의 설명이 셀 폭을 넘길 수 있습니다" % d.get("no", "?"))
            filled += 1
        elif clear_unused:
            if no_col < len(table.columns):
                set_cell_text(table.cell(row, no_col), "")
            if text_col < len(table.columns):
                set_cell_text(table.cell(row, text_col), "")

    if len(details) > len(slots):
        warns.add(screen_id, "slot-shortage",
                  "상세 %d건 중 %d건만 이 슬라이드에 들어갔습니다"
                  % (len(details), len(slots)))
    return filled
