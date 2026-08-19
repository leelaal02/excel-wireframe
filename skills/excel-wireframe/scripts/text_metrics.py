# -*- coding: utf-8 -*-
"""표 셀에 들어갈 텍스트가 몇 줄이 되고 그 행에 몇 줄이 들어가는지 센다.

표 자리는 모든 화면에서 같으므로 행을 늘리지 않는다. 대신 정해진 행에
글자가 들어가는지를 재서 넘치면 글자를 줄이거나 `text-overflow`로 알린다.

폰트 메트릭 없이 전각/반각만 구분해 폭을 누적한다. 화면설계서 상세는 대부분
한글이라 이 근사로 충분하다 — 정확한 측정은 폰트 파일을 읽어야 하는데, 얻는
정확도에 비해 의존성이 과하다.

순수 함수만 둔다. python-pptx도 openpyxl도 모른다.
"""
from __future__ import annotations

import unicodedata

EMU_PER_PT = 12700
BASE_LINE_SPACING = 1.2      # 맑은 고딕이 기본으로 잡는 줄높이 (폰트 크기 대비)
LINE_SPACING_RATIO = 0.95    # 셀 문단에 실제로 넣는 줄간격 배수
# 계산에 쓰는 줄높이. slide_layout이 LINE_SPACING_RATIO를 셀 서식에 그대로
# 넣으므로 이 값과 산출물이 일치한다 — 한쪽만 바꾸면 표가 자리에 안 맞는다.
LINE_SPACING = BASE_LINE_SPACING * LINE_SPACING_RATIO
CELL_MARGIN = 4762           # 설명 칸 여백. slide_layout이 셀 서식에 그대로 쓴다
DEFAULT_MARGIN_TOP = CELL_MARGIN


def _char_width(ch: str, size_emu: int) -> int:
    """전각은 폰트 크기와 같은 폭, 반각은 절반으로 친다."""
    if unicodedata.east_asian_width(ch) in ("W", "F"):
        return size_emu
    return size_emu // 2


def text_lines(text, width_emu: int, size_pt: float) -> int:
    """폭 안에서 텍스트가 몇 줄이 되는지 센다.

    빈 텍스트도 1을 돌려준다 — 빈 셀이 행 높이를 0으로 만들면 안 된다.
    폭보다 넓은 한 글자도 한 줄로 센다.
    """
    if not text:
        return 1

    size_emu = int(size_pt * EMU_PER_PT)
    total = 0
    # set_cell_text가 \n을 문단으로 나누므로 줄 수에도 그대로 반영한다.
    for para in str(text).split("\n"):
        count = 1
        used = 0
        for ch in para:
            w = _char_width(ch, size_emu)
            if used and used + w > width_emu:
                count += 1
                used = w
            else:
                used += w
        total += count
    return total


def fits_lines(height_emu: int, size_pt: float,
               margin_top: int = DEFAULT_MARGIN_TOP,
               margin_bottom: int = 0) -> int:
    """그 높이의 행에 몇 줄이 들어가는지. text_lines의 반대 방향이다.

    행이 아무리 낮아도 1을 돌려준다 — 0이면 빈 셀까지 넘침으로 잡힌다.
    """
    line_h = int(size_pt * LINE_SPACING * EMU_PER_PT)
    return max(1, (height_emu - margin_top - margin_bottom) // line_h)
