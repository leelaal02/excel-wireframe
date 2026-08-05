# -*- coding: utf-8 -*-
"""PPT 템플릿을 사용자가 주지 않았을 때 쓸 기본 템플릿을 만든다.

실제 화면설계서의 화면 페이지에서 잰 좌표를 그대로 쓴다. 원본은 상단 띠·하단 바·
쪽번호 같은 껍데기를 슬라이드 레이아웃에 그려서 슬라이드에서 클릭조차 되지 않는데,
여기서는 전부 슬라이드 위 도형으로 올려 모든 글자를 수정할 수 있게 한다.

원본 레이아웃의 저작권 문구는 복제하지 않는다. 이 파일은 스킬에 코드로 담겨
배포되므로, 넣으면 제3자의 저작권 표기가 모든 사용자 산출물에 박힌다.
그 자리에는 Excel 표지에서 읽은 문서제목이 들어간다.
"""
from __future__ import annotations

from pathlib import Path

from common import EMU_PER_INCH
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

DEFAULT_SHAPE_NAMES = {
    "title": "제목",
    "screen_id": "화면ID",
    "image": "화면이미지",
    "doc_title": "문서제목",
    "date": "작성일",
}

# --- 실측값 (실제 화면설계서 화면 페이지) ---
MEASURED_SLIDE = (9906000, 6858000)
MEASURED_TABLE_COUNT = 5
MEASURED_ROWS_PER_TABLE = 4

MEASURED_CONTENT = {
    "제목": (3722514, 0, 1260000, 144000),
    "화면ID": (8121353, 188640, 1766860, 138032),
    "화면이미지": (-12319, 337940, 9957099, 4675235),
}
MEASURED_SHELL = {
    "상단띠": (0, 0, 9896172, 137234),
    "상단띠2": (0, 195617, 8049346, 137234),
    "구분선": (-1, 404664, 9892977, 216024),
    "화면ID배경": (8121352, 188657, 1771625, 144000),
    "하단바": (0, 6716266, 9906000, 144000),
    "문서제목": (0, 6738252, 2648744, 100027),
    "쪽번호": (4734198, 6716266, 437604, 144000),
    "작성일": (8146752, 0, 504000, 144000),
}
MEASURED_TABLE_LEFTS = [-6849, 1974133, 3955115, 5936097, 7917077]
MEASURED_TABLE_TOP = 5253244
MEASURED_TABLE_SIZE = (1971135, 1416117)
MEASURED_COL_WIDTHS = [160215, 1810920]
MEASURED_ROW_HEIGHTS = [382457, 268746, 496168, 268746]

TEXT_ON_BAR = RGBColor(0xFF, 0xFF, 0xFF)
SLOT_BORDER = RGBColor(0xBB, 0xBB, 0xBB)
SLOT_FILL = RGBColor(0xF5, 0xF6, 0xF8)
MIN_IMAGE_HEIGHT_EMU = EMU_PER_INCH  # 1인치


def _is_measured(slide_w: int, slide_h: int, table_count: int, rows_per_table: int) -> bool:
    """실측 좌표를 그대로 쓸 수 있는 조합인가."""
    return (
        (slide_w, slide_h) == MEASURED_SLIDE
        and table_count == MEASURED_TABLE_COUNT
        and rows_per_table == MEASURED_ROWS_PER_TABLE
    )


def _accent_rect(slide, name, geom):
    """테마 accent1로 채운 장식 사각형. 테두리는 없앤다."""
    left, top, width, height = geom
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(left), Emu(top),
                                 Emu(width), Emu(height))
    shp.name = name
    shp.fill.solid()
    shp.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    shp.line.fill.background()
    shp.text_frame.text = ""
    return shp


def _textbox(slide, name, geom, size_pt, color=None, bold=False, align=None):
    left, top, width, height = geom
    box = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run()
    run.text = ""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    return box


def _fill_cell(cell, text, size_pt, bold, align, anchor, margin, font=None):
    cell.margin_left = Emu(margin)
    cell.margin_right = Emu(margin)
    cell.margin_top = Emu(margin)
    cell.margin_bottom = Emu(0 if font else margin)
    if anchor is not None:
        cell.vertical_anchor = anchor
    tf = cell.text_frame
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if font:
        run.font.name = font


def _computed_layout(slide_width_emu, slide_height_emu, table_count, rows_per_table):
    """실측 조합을 벗어났을 때의 폴백. 기존 코드의 비율 계산을 그대로 옮기되,
    반환 형태를 실측 경로(content/shell/표 배치)와 맞춘다.

    상단 두 개 띠 + 구분선이 헤더를, 하단 바 + 문서제목 + 쪽번호가 푸터를
    이룬다. 표 높이는 rows_per_table에 비례해 커지고, 그만큼 이미지 자리가
    줄어든다 — 기존 코드와 같은 트레이드오프다.
    """
    w, h = slide_width_emu, slide_height_emu
    margin = int(0.25 * EMU_PER_INCH)
    inner_w = w - margin * 2

    band_h = int(0.15 * EMU_PER_INCH)
    band_gap = int(0.02 * EMU_PER_INCH)
    row1_top = 0
    row2_top = band_h + band_gap
    divider_top = row2_top + band_h
    divider_h = int(0.03 * EMU_PER_INCH)
    header_bottom = divider_top + divider_h

    footer_h = int(0.16 * EMU_PER_INCH)
    footer_top = h - footer_h

    id_w = int(inner_w * 0.18)
    id_left = w - margin - id_w

    # 표 높이는 rows_per_table에 비례한다(4행일 때 1.55in — 실측값과 근사).
    # 표를 늘릴수록 그만큼 이미지 자리가 줄어들어야 슬라이드 하단을 넘치지 않는다.
    row_h = int((1.55 / 4) * EMU_PER_INCH)
    table_h = row_h * rows_per_table
    bottom_gap = int(0.06 * EMU_PER_INCH)
    tables_top = footer_top - bottom_gap - table_h

    img_gap = int(0.05 * EMU_PER_INCH)
    img_top = header_bottom + img_gap
    img_h = tables_top - img_top - img_gap

    if img_h < MIN_IMAGE_HEIGHT_EMU:
        # table_h가 rows_per_table에 비례하므로 무한정 키우면 이미지 자리
        # 높이가 음수가 될 수 있다. 조용히 겹치는 도형을 만드는 대신, 이
        # 슬라이드 크기에서 실제로 쓸 수 있는 최대 rows_per_table을 계산해
        # 알려준다.
        max_rows = int(
            (footer_top - bottom_gap - img_top - img_gap - MIN_IMAGE_HEIGHT_EMU) // row_h
        )
        raise ValueError(
            "rows_per_table=%d면 이미지 자리 높이가 %.2fin로 너무 작아집니다"
            "(최소 %.2fin 필요). 이 슬라이드 크기(%.2f x %.2fin)에서는 "
            "rows_per_table을 %d 이하로 쓰세요."
            % (
                rows_per_table, img_h / EMU_PER_INCH,
                MIN_IMAGE_HEIGHT_EMU / EMU_PER_INCH,
                slide_width_emu / EMU_PER_INCH, slide_height_emu / EMU_PER_INCH,
                max_rows,
            )
        )

    gap = int(0.06 * EMU_PER_INCH)
    table_w = (inner_w - gap * (table_count - 1)) // table_count
    table_lefts = [margin + (table_w + gap) * i for i in range(table_count)]
    col_widths = [int(table_w * 0.18), table_w - int(table_w * 0.18)]
    row_heights = [row_h] * rows_per_table

    content = {
        "제목": (margin, row1_top, int(inner_w * 0.3), band_h),
        "화면ID": (id_left, row2_top, id_w, band_h),
        "화면이미지": (margin, img_top, inner_w, img_h),
    }
    shell = {
        "상단띠": (0, row1_top, w - margin, band_h),
        "상단띠2": (0, row2_top, int(w * 0.7), band_h),
        "구분선": (0, divider_top, w - margin, divider_h),
        "화면ID배경": (id_left, row2_top, id_w, band_h),
        "하단바": (0, footer_top, w, footer_h),
        "문서제목": (margin, footer_top + int(footer_h * 0.1),
                   int(inner_w * 0.3), int(footer_h * 0.8)),
        "쪽번호": (int(w * 0.47), footer_top, int(w * 0.045), footer_h),
        "작성일": (id_left, row1_top, id_w, band_h),
    }
    return content, shell, table_lefts, tables_top, table_w, table_h, col_widths, row_heights


def build_default_template(
    path: Path,
    slide_width_emu: int = MEASURED_SLIDE[0],
    slide_height_emu: int = MEASURED_SLIDE[1],
    table_count: int = MEASURED_TABLE_COUNT,
    rows_per_table: int = MEASURED_ROWS_PER_TABLE,
) -> Path:
    measured = _is_measured(slide_width_emu, slide_height_emu, table_count, rows_per_table)

    prs = Presentation()
    prs.slide_width = Emu(slide_width_emu)
    prs.slide_height = Emu(slide_height_emu)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 빈 화면

    # python-pptx 기본 템플릿의 "빈 화면" 레이아웃도 master에서 물려받은
    # 날짜/바닥글/쪽번호 placeholder를 갖고 있다. 슬라이드에는 새지 않지만
    # 레이아웃 자체에는 남아 있어 "레이아웃에는 아무것도 그리지 않는다"는
    # 요구를 어긴다 — 지운다.
    layout = slide.slide_layout
    for shp in list(layout.shapes):
        shp._element.getparent().remove(shp._element)

    if measured:
        content = dict(MEASURED_CONTENT)
        shell = dict(MEASURED_SHELL)
        table_lefts = list(MEASURED_TABLE_LEFTS)
        table_top = MEASURED_TABLE_TOP
        table_w, table_h = MEASURED_TABLE_SIZE
        col_widths = list(MEASURED_COL_WIDTHS)
        row_heights = list(MEASURED_ROW_HEIGHTS)
    else:
        content, shell, table_lefts, table_top, table_w, table_h, col_widths, row_heights = (
            _computed_layout(slide_width_emu, slide_height_emu, table_count, rows_per_table)
        )

    # 1) 껍데기 — 배경부터
    for name in ("상단띠", "상단띠2", "구분선", "화면ID배경", "하단바"):
        _accent_rect(slide, name, shell[name])

    # 2) 이미지 자리
    left, top, width, height = content["화면이미지"]
    img = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(left), Emu(top),
                                 Emu(width), Emu(height))
    img.name = DEFAULT_SHAPE_NAMES["image"]
    img.fill.solid()
    img.fill.fore_color.rgb = SLOT_FILL
    img.line.color.rgb = SLOT_BORDER
    img.text_frame.text = ""

    # 3) 표
    for i in range(table_count):
        shp = slide.shapes.add_table(rows_per_table, 2, Emu(table_lefts[i]),
                                     Emu(table_top), Emu(table_w), Emu(table_h))
        shp.name = "상세표%d" % (i + 1)
        table = shp.table
        for ci, cw in enumerate(col_widths):
            table.columns[ci].width = Emu(cw)
        for ri, rh in enumerate(row_heights[:rows_per_table]):
            table.rows[ri].height = Emu(rh)
        for r in range(rows_per_table):
            n = i * rows_per_table + r + 1
            _fill_cell(table.cell(r, 0), str(n), 6.5, bold=True,
                       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=18000)
            _fill_cell(table.cell(r, 1), "예시 설명 %d" % n, 7.0, bold=False,
                       align=PP_ALIGN.LEFT, anchor=None, margin=9525, font="맑은 고딕")

    # 4) 텍스트 — 배경 위에
    _textbox(slide, DEFAULT_SHAPE_NAMES["title"], content["제목"], 6.5,
             align=PP_ALIGN.CENTER)
    _textbox(slide, DEFAULT_SHAPE_NAMES["screen_id"], content["화면ID"], 6.5,
             color=TEXT_ON_BAR)
    _textbox(slide, DEFAULT_SHAPE_NAMES["doc_title"], shell["문서제목"], 6.5,
             color=TEXT_ON_BAR)
    _textbox(slide, "쪽번호", shell["쪽번호"], 6.5, color=TEXT_ON_BAR,
             align=PP_ALIGN.CENTER)
    _textbox(slide, DEFAULT_SHAPE_NAMES["date"], shell["작성일"], 6.5)

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path


def default_template_mapping(path: Path, table_count: int = 5) -> dict:
    """생성한 기본 템플릿에 대응하는 mapping.json의 template 섹션."""
    return {
        "file": str(path),
        "mode": "clone",
        "source_slide": 0,
        "shapes": {
            "title": DEFAULT_SHAPE_NAMES["title"],
            "screen_id": DEFAULT_SHAPE_NAMES["screen_id"],
            "image": DEFAULT_SHAPE_NAMES["image"],
            "문서제목": DEFAULT_SHAPE_NAMES["doc_title"],
            "작성일": DEFAULT_SHAPE_NAMES["date"],
            "detail_tables": ["상세표%d" % (i + 1) for i in range(table_count)],
        },
        "table_columns": {"no": 0, "text": 1},
    }
