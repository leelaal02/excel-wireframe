# -*- coding: utf-8 -*-
"""레이아웃을 골라 슬라이드를 만들고, 그 레이아웃이 물려주는 자리에 값을 채운다.

디자인은 레이아웃이 담당한다. 이 모듈은 자리를 찾아 이름을 붙이고, 레이아웃에
자리가 없는 두 가지(화면 이미지, 상세표)만 본문 영역 안에 만든다.
"""
from __future__ import annotations

import copy

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


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
    """
    have = {ph.placeholder_format.idx for ph in slide.placeholders}
    added: list[int] = []
    for ph in layout.placeholders:
        idx = ph.placeholder_format.idx
        if idx in have:
            continue
        slide.shapes._spTree.append(copy.deepcopy(ph._element))
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
