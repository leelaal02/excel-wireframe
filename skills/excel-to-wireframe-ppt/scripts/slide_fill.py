# -*- coding: utf-8 -*-
"""슬라이드 도형에 값을 채운다.

핵심은 서식 보존이다. 런을 전부 지우고 새로 만들면 폰트·크기·색이 초기화되어
템플릿 디자인이 무너지므로, 첫 런의 텍스트만 갈아끼우는 방식을 쓴다.
"""
from __future__ import annotations

import copy

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
