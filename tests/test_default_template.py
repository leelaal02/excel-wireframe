# -*- coding: utf-8 -*-
from pathlib import Path

from default_template import (
    DEFAULT_LAYOUT_NAME,
    DEFAULT_SHAPE_NAMES,
    PLACEHOLDER_IDX,
    build_default_template,
    default_template_mapping,
)
from pptx import Presentation

MEASURED_PLACEHOLDERS = {
    0: (3722514, 0, 1260000, 144000),
    1: (8121353, 188640, 1766860, 138032),
    10: (8146752, 0, 504000, 144000),
    11: (0, 6738252, 2648744, 100027),
    12: (4734198, 6716266, 437604, 144000),
}
MEASURED_SHELL = {
    "상단띠": (0, 0, 9896172, 137234),
    "상단띠2": (0, 195617, 8049346, 137234),
    "구분선": (-1, 404664, 9892977, 216024),
    "화면ID배경": (8121352, 188657, 1771625, 144000),
    "하단바": (0, 6716266, 9906000, 144000),
}


def _layout(path: Path):
    prs = Presentation(str(path))
    return next(lay for master in prs.slide_masters
                for lay in master.slide_layouts
                if lay.name == DEFAULT_LAYOUT_NAME)


def test_default_template_uses_measured_size(tmp_path: Path):
    prs = Presentation(str(build_default_template(tmp_path / "d.pptx")))
    assert prs.slide_width == 9906000
    assert prs.slide_height == 6858000


def test_default_template_size_is_configurable(tmp_path: Path):
    prs = Presentation(str(build_default_template(
        tmp_path / "d.pptx", slide_width_emu=12192000, slide_height_emu=6858000)))
    assert prs.slide_width == 12192000


def test_default_template_has_no_slides(tmp_path: Path):
    """슬라이드는 빌드가 레이아웃으로 만든다. 템플릿에 예시가 필요 없다."""
    prs = Presentation(str(build_default_template(tmp_path / "d.pptx")))
    assert len(list(prs.slides)) == 0


def test_default_template_layout_exists(tmp_path: Path):
    lay = _layout(build_default_template(tmp_path / "d.pptx"))
    assert lay.name == DEFAULT_LAYOUT_NAME


def test_default_template_placeholders_match_measurements(tmp_path: Path):
    lay = _layout(build_default_template(tmp_path / "d.pptx"))
    got = {ph.placeholder_format.idx: (ph.left, ph.top, ph.width, ph.height)
           for ph in lay.placeholders}
    for idx, geom in MEASURED_PLACEHOLDERS.items():
        assert got[idx] == geom, "idx=%d" % idx


def test_default_template_shell_is_on_the_layout(tmp_path: Path):
    """껍데기는 레이아웃이 담당한다. 슬라이드마다 다시 그리지 않는다."""
    lay = _layout(build_default_template(tmp_path / "d.pptx"))
    by_name = {s.name: s for s in lay.shapes if not s.is_placeholder}
    for name, geom in MEASURED_SHELL.items():
        assert name in by_name, name
        s = by_name[name]
        assert (s.left, s.top, s.width, s.height) == geom, name


def test_default_template_placeholders_start_empty(tmp_path: Path):
    """표지가 없는 Excel에서 자리표시 문구가 산출물에 찍히면 안 된다."""
    lay = _layout(build_default_template(tmp_path / "d.pptx"))
    for ph in lay.placeholders:
        if ph.placeholder_format.idx == 12:
            continue  # 쪽번호는 자동 필드다
        assert ph.text_frame.text == "", ph.placeholder_format.idx


def test_default_template_id_background_sits_behind_id_text(tmp_path: Path):
    """화면ID배경이 화면ID placeholder보다 먼저 와야 뒤에 깔린다."""
    lay = _layout(build_default_template(tmp_path / "d.pptx"))
    order = [s.name for s in lay.shapes]
    ph_name = next(s.name for s in lay.shapes
                   if s.is_placeholder and s.placeholder_format.idx == 1)
    assert order.index("화면ID배경") < order.index(ph_name)


def test_default_template_carries_no_third_party_copyright(tmp_path: Path):
    """기본 템플릿은 코드로 배포되므로 남의 저작권 표기가 들어가면 안 된다."""
    prs = Presentation(str(build_default_template(tmp_path / "d.pptx")))
    texts = []
    for master in prs.slide_masters:
        for lay in master.slide_layouts:
            for s in lay.shapes:
                if s.has_text_frame:
                    texts.append(s.text_frame.text)
    blob = "\n".join(texts)
    for banned in ("Copyright", "ⓒ", "Nurimedia", "All rights reserved"):
        assert banned not in blob, banned


def test_default_template_mapping_is_layout_mode(tmp_path: Path):
    path = build_default_template(tmp_path / "d.pptx")
    m = default_template_mapping(path)
    assert m["mode"] == "layout"
    assert m["layout"] == DEFAULT_LAYOUT_NAME
    assert "source_slide" not in m
    assert m["placeholders"] == PLACEHOLDER_IDX
    assert m["detail_tables"] == {"count": 5, "rows": 4}
    assert m["shapes"]["title"] == DEFAULT_SHAPE_NAMES["title"]
    assert m["shapes"]["detail_tables"] == [
        "상세표1", "상세표2", "상세표3", "상세표4", "상세표5"]
    assert m["table_columns"] == {"no": 0, "text": 1}


def test_default_template_mapping_is_buildable(tmp_path: Path):
    """제안 매핑을 그대로 build에 넘겨 슬라이드가 나와야 한다."""
    from build import build
    from common import Warnings

    path = build_default_template(tmp_path / "d.pptx")
    mapping = {
        "version": 1,
        "excel": {"layout": "sheet-per-screen"},
        "template": default_template_mapping(path),
        "options": {"detail_text_source": "desc", "clear_unused_slots": True},
    }
    screens = {
        "meta": {"문서제목": "화면설계서"},
        "screens": [{"id": "SCR001", "name": "목록", "images": [], "fields": {},
                     "details": [{"no": "1", "desc": "설명"}]}],
    }
    out = tmp_path / "out.pptx"
    report = build(screens, mapping, tmp_path, out, Warnings())
    assert report["slides"] == 1
    slide = Presentation(str(out)).slides[0]
    by_name = {s.name: s for s in slide.shapes}
    assert by_name["제목"].text_frame.text == "목록"
    assert by_name["문서제목"].text_frame.text == "화면설계서"
