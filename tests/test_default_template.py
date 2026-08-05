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
from slide_layout import P_NS

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


def test_default_template_shell_sits_behind_every_placeholder(tmp_path: Path):
    """껍데기는 전부 placeholder보다 뒤에 깔려야 한다 — 화면ID배경 위에 화면ID
    글자가 보이는 것이 대표적인 경우다.

    예전 테스트는 '화면ID배경'과 idx=1 placeholder 한 쌍의 순서만 봤다. 그런데
    build_default_template은 껍데기를 전부 붙인 뒤 *모든* placeholder를 spTree
    맨 뒤로 다시 붙이므로, 그 한 쌍의 비교는 SHELL_ORDER를 어떻게 바꾸든 항상
    참이라 깨질 수가 없었다. 실제 불변식은 '껍데기 전부가 placeholder 전부보다
    앞'이고, 껍데기를 placeholder 재배치 뒤에 붙이는 순간 이 검사는 실패한다.
    """
    lay = _layout(build_default_template(tmp_path / "d.pptx"))
    shells = [i for i, s in enumerate(lay.shapes) if not s.is_placeholder]
    phs = [i for i, s in enumerate(lay.shapes) if s.is_placeholder]

    # 다섯 껍데기가 전부 레이아웃에 있어야 아래 비교가 의미를 갖는다
    assert {s.name for s in lay.shapes if not s.is_placeholder} == set(MEASURED_SHELL)
    assert len(phs) == len(MEASURED_PLACEHOLDERS)
    assert max(shells) < min(phs), [s.name for s in lay.shapes]


def test_default_template_layout_shape_ids_are_unique(tmp_path: Path):
    """cNvPr/@id는 한 파트 안에서 유일해야 한다 — 겹치면 PowerPoint가 파일을
    열 때 복구를 요구한다. 껍데기 도형을 임시 슬라이드에서 deepcopy로 이식할 때
    슬라이드에서 매겨진 id가 따라와 레이아웃 placeholder의 id와 겹쳤다."""
    lay = _layout(build_default_template(tmp_path / "d.pptx"))
    ids = [el.get("id") for el
           in lay.shapes._spTree.iter("{%s}cNvPr" % P_NS)]
    dupes = sorted({i for i in ids if ids.count(i) > 1})
    assert dupes == [], "중복된 cNvPr/@id: %s" % dupes


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
