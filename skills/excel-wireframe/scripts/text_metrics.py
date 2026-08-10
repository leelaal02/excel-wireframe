# -*- coding: utf-8 -*-
"""표 셀에 들어갈 텍스트가 몇 줄이 되고 행이 얼마나 높아야 하는지 센다.

PowerPoint 표의 행 높이는 하한이다. 지정한 값보다 텍스트가 길면 렌더링 시 행이
스스로 자라고 표가 슬라이드 밖으로 밀린다. 미리 필요한 높이를 계산해 두면 표를
위로 늘려 담을 수 있다.

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


def row_height(lines: int, size_pt: float, margin_top: int = DEFAULT_MARGIN_TOP,
               margin_bottom: int = 0) -> int:
    """줄 수에 필요한 행 높이(EMU)."""
    line_h = int(size_pt * LINE_SPACING * EMU_PER_PT)
    return max(1, lines) * line_h + margin_top + margin_bottom


def fits_lines(height_emu: int, size_pt: float,
               margin_top: int = DEFAULT_MARGIN_TOP,
               margin_bottom: int = 0) -> int:
    """그 높이의 행에 몇 줄이 들어가는지. row_height의 반대 방향이다.

    행이 아무리 낮아도 1을 돌려준다 — 0이면 빈 셀까지 넘침으로 잡힌다.
    """
    line_h = int(size_pt * LINE_SPACING * EMU_PER_PT)
    return max(1, (height_emu - margin_top - margin_bottom) // line_h)


def plan_row_heights(pages, rows_per_table: int, width_emu: int,
                     size_pt: float, floors) -> list[int]:
    """한 화면의 모든 장을 통틀어 행 인덱스별 높이를 정한다.

    표가 여러 개 나란히 놓이므로 행 높이는 표를 가로질러 하나로 통일해야 한다 —
    표마다 다르면 행이 어긋나 보인다. 장끼리도 통일한다. 장마다 표 높이가 다르면
    이미지 자리가 장마다 달라지고, 이미지 분할이 표 높이에 의존하는 순환에 빠진다.

    슬롯 i는 표 i // rows_per_table, 행 i % rows_per_table에 대응한다.
    floors(실측 행 높이)를 하한으로 삼는다.
    """
    heights = [floors[r] if r < len(floors) else floors[-1]
               for r in range(rows_per_table)]
    for page in pages:
        for i, text in enumerate(page):
            r = i % rows_per_table
            need = row_height(text_lines(text, width_emu, size_pt), size_pt)
            if need > heights[r]:
                heights[r] = need
    return heights
