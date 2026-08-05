# -*- coding: utf-8 -*-
"""PPT 템플릿을 사용자가 주지 않았을 때 쓸 기본 템플릿을 만든다.

디자인은 레이아웃이 담당한다. 상단 띠·하단 바 같은 껍데기는 레이아웃 위 도형이고,
글자가 들어가는 자리는 전부 placeholder다 — 빌드는 레이아웃을 골라 슬라이드를
추가하고 그 자리에 값을 채운다. 슬라이드는 한 장도 넣지 않는다.

python-pptx 기본 템플릿의 'Title and Content' 레이아웃이 가진 placeholder 다섯 개
(TITLE/OBJECT/DATE/FOOTER/SLIDE_NUMBER)를 실측 자리로 옮겨 쓴다. 다섯 개가 필요한
자리와 정확히 맞아떨어져서 placeholder XML을 새로 만들 필요가 없다.

원본 레이아웃의 저작권 문구는 복제하지 않는다. 이 파일은 스킬에 코드로 담겨
배포되므로, 넣으면 제3자의 저작권 표기가 모든 사용자 산출물에 박힌다.
그 자리(FOOTER)에는 Excel 표지에서 읽은 문서제목이 들어간다.
"""
from __future__ import annotations

import copy
from pathlib import Path

from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu

DEFAULT_LAYOUT_NAME = "화면"

DEFAULT_SHAPE_NAMES = {
    "title": "제목",
    "screen_id": "화면ID",
    "image": "화면이미지",
    "doc_title": "문서제목",
    "date": "작성일",
    "page_no": "쪽번호",
}

# python-pptx 기본 템플릿 'Title and Content'의 placeholder idx
PLACEHOLDER_IDX = {
    "title": 0,
    "screen_id": 1,
    "작성일": 10,
    "문서제목": 11,
    "쪽번호": 12,
}

BASE_LAYOUT_INDEX = 1  # 'Title and Content'

# --- 실측값 (실제 화면설계서 화면 페이지) ---
MEASURED_SLIDE = (9906000, 6858000)
MEASURED_PLACEHOLDERS = {
    0: (3722514, 0, 1260000, 144000),        # 제목
    1: (8121353, 188640, 1766860, 138032),   # 화면ID
    10: (8146752, 0, 504000, 144000),        # 작성일
    11: (0, 6738252, 2648744, 100027),       # 문서제목 (원본의 저작권 문구 자리)
    12: (4734198, 6716266, 437604, 144000),  # 쪽번호
}
MEASURED_SHELL = {
    "상단띠": (0, 0, 9896172, 137234),
    "상단띠2": (0, 195617, 8049346, 137234),
    "구분선": (-1, 404664, 9892977, 216024),
    "화면ID배경": (8121352, 188657, 1771625, 144000),
    "하단바": (0, 6716266, 9906000, 144000),
}
# 껍데기를 그리는 순서. 화면ID배경은 화면ID placeholder보다 먼저 와야 뒤에 깔린다.
SHELL_ORDER = ("상단띠", "상단띠2", "구분선", "화면ID배경", "하단바")


def _scaled(geom, sx: float, sy: float):
    left, top, width, height = geom
    return (int(left * sx), int(top * sy), int(width * sx), int(height * sy))


def _accent_rect(shapes, name, geom):
    """테마 accent1로 채운 장식 사각형. 테두리는 없앤다."""
    left, top, width, height = geom
    shp = shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(left), Emu(top),
                           Emu(width), Emu(height))
    shp.name = name
    shp.fill.solid()
    shp.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    shp.line.fill.background()
    shp.text_frame.text = ""
    return shp


def _drop_slide(prs, slide) -> None:
    """프레젠테이션에서 슬라이드를 제거한다. python-pptx에 삭제 API가 없다."""
    xml_slides = prs.slides._sldIdLst
    for sld_id in list(xml_slides):
        if prs.part.rels[sld_id.rId].target_part is slide.part:
            prs.part.drop_rel(sld_id.rId)
            xml_slides.remove(sld_id)
            return


def build_default_template(
    path: Path,
    slide_width_emu: int = MEASURED_SLIDE[0],
    slide_height_emu: int = MEASURED_SLIDE[1],
) -> Path:
    """레이아웃 하나짜리 기본 템플릿을 만든다. 슬라이드는 넣지 않는다."""
    prs = Presentation()
    prs.slide_width = Emu(slide_width_emu)
    prs.slide_height = Emu(slide_height_emu)

    sx = slide_width_emu / MEASURED_SLIDE[0]
    sy = slide_height_emu / MEASURED_SLIDE[1]

    layout = prs.slide_layouts[BASE_LAYOUT_INDEX]
    layout.name = DEFAULT_LAYOUT_NAME

    # 1) 껍데기 — LayoutShapes에는 add_shape가 없어 임시 슬라이드를 거친다.
    #    빈 화면 레이아웃을 써서 그 레이아웃의 placeholder가 섞이지 않게 한다.
    tmp = prs.slides.add_slide(prs.slide_layouts[6])
    for name in SHELL_ORDER:
        _accent_rect(tmp.shapes, name, _scaled(MEASURED_SHELL[name], sx, sy))
    for shp in list(tmp.shapes):
        layout.shapes._spTree.append(copy.deepcopy(shp._element))
    _drop_slide(prs, tmp)

    # 2) placeholder를 실측 자리로 옮긴다. 껍데기 뒤에 붙어야 위에 그려진다.
    #    list()로 감싸는 이유: 루프 안에서 _spTree를 재정렬하므로 살아 있는
    #    반복자를 그대로 쓰면 요소를 건너뛴다.
    for ph in list(layout.placeholders):
        idx = ph.placeholder_format.idx
        geom = MEASURED_PLACEHOLDERS.get(idx)
        if geom is None:
            continue
        ph.left, ph.top, ph.width, ph.height = (
            Emu(v) for v in _scaled(geom, sx, sy)
        )
        # python-pptx 기본 레이아웃의 placeholder는 안내 문구('Click to edit
        # Master title style')와 날짜 필드('1/27/13')를 갖고 있다. 날짜 자리는
        # 오늘 날짜가 아니라 Excel 표지에서 읽은 작성일이 들어갈 자리이고,
        # inherit_placeholders는 레이아웃 XML을 그대로 복제하므로 비워 두지
        # 않으면 그 문구가 슬라이드에 그대로 실린다. 쪽번호(12)만 자동 필드를
        # 살려 둔다.
        if idx != PLACEHOLDER_IDX["쪽번호"]:
            ph.text_frame.text = ""
        layout.shapes._spTree.append(ph._element)  # z-order를 맨 위로

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path


def default_template_mapping(path: Path, table_count: int = 5,
                             rows_per_table: int = 4) -> dict:
    """생성한 기본 템플릿에 대응하는 mapping.json의 template 섹션."""
    return {
        "file": str(path),
        "mode": "layout",
        "layout": DEFAULT_LAYOUT_NAME,
        "placeholders": dict(PLACEHOLDER_IDX),
        "shapes": {
            "title": DEFAULT_SHAPE_NAMES["title"],
            "screen_id": DEFAULT_SHAPE_NAMES["screen_id"],
            "image": DEFAULT_SHAPE_NAMES["image"],
            "문서제목": DEFAULT_SHAPE_NAMES["doc_title"],
            "작성일": DEFAULT_SHAPE_NAMES["date"],
            "쪽번호": DEFAULT_SHAPE_NAMES["page_no"],
            "detail_tables": ["상세표%d" % (i + 1) for i in range(table_count)],
        },
        "detail_tables": {"count": table_count, "rows": rows_per_table},
        "table_columns": {"no": 0, "text": 1},
    }
