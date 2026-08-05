from pathlib import Path

import pytest
from default_template import (
    DEFAULT_SHAPE_NAMES,
    build_default_template,
    default_template_mapping,
)
from pptx import Presentation


def test_default_template_uses_measured_size(tmp_path: Path):
    """실측 화면설계서의 슬라이드 크기(10.83 x 7.50in)를 기본값으로 쓴다.
    예전에는 16:9(12192000)가 기본값이라 이름이 `is_16_9`였지만, 이제 기본
    조합은 실측 크기를 그대로 쓰므로 이름과 기대값 모두 실측 기준으로 고친다."""
    p = build_default_template(tmp_path / "d.pptx")
    prs = Presentation(str(p))
    assert prs.slide_width == 9906000
    assert prs.slide_height == 6858000
    assert len(prs.slides) == 1


def test_default_template_has_named_shapes(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    names = [s.name for s in Presentation(str(p)).slides[0].shapes]
    assert DEFAULT_SHAPE_NAMES["title"] in names
    assert DEFAULT_SHAPE_NAMES["screen_id"] in names
    assert DEFAULT_SHAPE_NAMES["image"] in names
    assert "상세표1" in names and "상세표5" in names


def test_default_template_tables_have_20_slots(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    tables = [s for s in Presentation(str(p)).slides[0].shapes if s.has_table]
    assert len(tables) == 5
    assert sum(len(t.table.rows) for t in tables) == 20
    assert tables[0].table.cell(0, 0).text == "1"


def test_default_template_tables_do_not_overlap(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    prs = Presentation(str(p))
    tables = sorted(
        (s for s in prs.slides[0].shapes if s.has_table), key=lambda s: s.left
    )
    for a, b in zip(tables, tables[1:]):
        assert a.left + a.width <= b.left + 1
    assert tables[-1].left + tables[-1].width <= prs.slide_width


def test_default_template_size_is_configurable(tmp_path: Path):
    """9906000이 기본값이 된 뒤로는 그 값을 넘겨도 아무것도 검증하지 못한다
    (계산 경로로도, 실측 경로로도 안 떨어질 수 있음). 기본값과 다른 폭
    (12192000, 옛 16:9 폭)을 넘겨 "기본값과 다른 값도 된다"를 검증한다."""
    p = build_default_template(
        tmp_path / "d.pptx", slide_width_emu=12192000, table_count=3
    )
    prs = Presentation(str(p))
    assert prs.slide_width == 12192000
    assert sum(1 for s in prs.slides[0].shapes if s.has_table) == 3


def test_default_template_table_height_scales_with_rows_per_table(tmp_path: Path):
    """이전 세로-간격 테스트는 rows_per_table과 무관하게 고정된 표 높이
    공식에서도 통과했다 — 표 높이가 선언된 기하 안에서 이미지 자리와
    안 겹치기만 하면 됐기 때문에, 옛 코드(표 높이 1.55in 고정)를 그대로
    둬도 이 assertion만으로는 안 잡혔다. 실제로 달라져야 하는 성질(표
    높이가 행 수에 비례해 커진다) 그 자체를 직접 비교한다.

    실측 조합(rows_per_table=4)은 실측 경로로 떨어져 고정 표 높이(1416117)를
    쓰므로 rows_per_table을 바꿔도 표 높이가 달라지지 않는다. 스케일 성질은
    계산 경로에서만 참이다 — 6행과 12행 둘 다 실측 조합(4행)을 벗어나므로
    계산 경로로 떨어진다."""
    p6 = build_default_template(tmp_path / "d6.pptx", rows_per_table=6)
    p12 = build_default_template(tmp_path / "d12.pptx", rows_per_table=12)

    t6 = next(s for s in Presentation(str(p6)).slides[0].shapes if s.has_table)
    t12 = next(s for s in Presentation(str(p12)).slides[0].shapes if s.has_table)

    assert t12.height > t6.height
    assert t12.height == pytest.approx(t6.height * 2, rel=0.01)


def test_default_template_image_slot_does_not_overlap_tables(tmp_path: Path):
    """가로축(표끼리 겹치지 않음)은 이미 검사한다. 세로축은 아무도 다시 안 본다 —
    표 높이가 rows_per_table에 비례해 커지면 이미지 자리가 슬라이드 하단
    표 위치까지 침범할 수 있다. 위 스케일 테스트와 짝을 이룬다: 그 스케일이
    실제로 이미지 자리를 침범하지 않는지는 여기서 확인한다."""
    p = build_default_template(tmp_path / "d.pptx", rows_per_table=8)
    prs = Presentation(str(p))
    shapes = prs.slides[0].shapes
    img_slot = next(s for s in shapes if s.name == DEFAULT_SHAPE_NAMES["image"])
    tables = [s for s in shapes if s.has_table]
    assert img_slot.top + img_slot.height <= tables[0].top
    for t in tables:
        assert t.top + t.height <= prs.slide_height


def test_default_template_raises_when_rows_per_table_too_large(tmp_path: Path):
    """회귀 4: table_h가 rows_per_table에 비례하므로 무한정 키우면 이미지
    자리 높이가 음수가 될 수 있다. 조용히 겹치는 도형을 만드는 대신, 기본
    슬라이드 크기(7.5in 높이)에서 실제로 쓸 수 있는 최대값을 알려주는
    ValueError로 막아야 한다. 새 계산 경로 기하에서 상한은 15다(옛 기하의
    14가 아니다 — 헤더/푸터 껍데기가 늘어난 만큼 여유가 달라졌다)."""
    with pytest.raises(ValueError) as excinfo:
        build_default_template(tmp_path / "d.pptx", rows_per_table=17)
    msg = str(excinfo.value)
    assert "rows_per_table" in msg
    assert "15" in msg  # 기본 슬라이드 크기에서 이미지 자리 최소 1in을 지키는 상한


def test_default_template_allows_max_workable_rows_per_table(tmp_path: Path):
    """위 테스트가 말하는 상한(15)에서는 여전히 만들어져야 한다 — 경계값이
    실수로 한 칸 안쪽으로 당겨지지 않았는지 확인한다."""
    build_default_template(tmp_path / "d.pptx", rows_per_table=15)


def test_default_template_mapping_matches_shapes(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    mp = default_template_mapping(p)
    names = [s.name for s in Presentation(str(p)).slides[0].shapes]
    assert mp["mode"] == "clone"
    assert mp["source_slide"] == 0
    assert mp["file"] == str(p)
    assert mp["shapes"]["title"] in names
    assert mp["shapes"]["detail_tables"] == [
        "상세표1", "상세표2", "상세표3", "상세표4", "상세표5"
    ]
    assert mp["table_columns"] == {"no": 0, "text": 1}


def test_default_template_has_meta_shapes(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    names = [s.name for s in Presentation(str(p)).slides[0].shapes]
    assert "문서제목" in names
    assert "작성일" in names


def test_default_template_meta_shapes_are_empty(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    slide = Presentation(str(p)).slides[0]
    for name in ("문서제목", "작성일"):
        shp = next(s for s in slide.shapes if s.name == name)
        assert shp.text_frame.text == ""


def test_default_template_meta_shapes_stay_inside_slide(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    prs = Presentation(str(p))
    slide = prs.slides[0]
    for name in ("문서제목", "작성일"):
        shp = next(s for s in slide.shapes if s.name == name)
        assert shp.left >= 0
        assert shp.top >= 0
        assert shp.left + shp.width <= prs.slide_width
        assert shp.top + shp.height <= prs.slide_height


def test_default_template_mapping_includes_meta_shapes(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    shapes = default_template_mapping(p)["shapes"]
    assert shapes["문서제목"] == "문서제목"
    assert shapes["작성일"] == "작성일"


def test_default_template_matches_measured_slide_size(tmp_path: Path):
    prs = Presentation(str(build_default_template(tmp_path / "d.pptx")))
    assert prs.slide_width == 9906000
    assert prs.slide_height == 6858000


def test_default_template_content_shapes_match_measurements(tmp_path: Path):
    slide = Presentation(str(build_default_template(tmp_path / "d.pptx"))).slides[0]
    by_name = {s.name: s for s in slide.shapes}
    assert (by_name["제목"].left, by_name["제목"].top,
            by_name["제목"].width, by_name["제목"].height) == (3722514, 0, 1260000, 144000)
    assert (by_name["화면ID"].left, by_name["화면ID"].top,
            by_name["화면ID"].width, by_name["화면ID"].height) == (8121353, 188640, 1766860, 138032)
    assert (by_name["화면이미지"].left, by_name["화면이미지"].top,
            by_name["화면이미지"].width, by_name["화면이미지"].height) == (-12319, 337940, 9957099, 4675235)


def test_default_template_tables_match_measurements(tmp_path: Path):
    slide = Presentation(str(build_default_template(tmp_path / "d.pptx"))).slides[0]
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) == 5
    assert [t.left for t in tables] == [-6849, 1974133, 3955115, 5936097, 7917077]
    for t in tables:
        assert t.top == 5253244
        assert t.width == 1971135
        assert [c.width for c in t.table.columns] == [160215, 1810920]
        assert [r.height for r in t.table.rows] == [382457, 268746, 496168, 268746]


def test_default_template_cell_formatting_matches(tmp_path: Path):
    from pptx.util import Pt
    slide = Presentation(str(build_default_template(tmp_path / "d.pptx"))).slides[0]
    table = next(s for s in slide.shapes if s.has_table).table
    num = table.cell(0, 0)
    txt = table.cell(0, 1)
    assert num.text_frame.paragraphs[0].runs[0].font.size == Pt(6.5)
    assert num.text_frame.paragraphs[0].runs[0].font.bold is True
    assert num.margin_left == 18000
    assert txt.text_frame.paragraphs[0].runs[0].font.size == Pt(7)
    assert txt.text_frame.paragraphs[0].runs[0].font.name == "맑은 고딕"
    assert txt.margin_left == 9525


def test_default_template_shell_shapes_are_on_the_slide(tmp_path: Path):
    """껍데기가 레이아웃이 아니라 슬라이드에 있어야 편집할 수 있다."""
    prs = Presentation(str(build_default_template(tmp_path / "d.pptx")))
    slide = prs.slides[0]
    names = [s.name for s in slide.shapes]
    for want in ("상단띠", "상단띠2", "구분선", "화면ID배경", "하단바", "쪽번호"):
        assert want in names, "%s 가 슬라이드에 없다" % want
    # 레이아웃에는 아무것도 그리지 않는다
    assert len(slide.slide_layout.shapes) == 0


def test_default_template_shell_geometry_matches(tmp_path: Path):
    slide = Presentation(str(build_default_template(tmp_path / "d.pptx"))).slides[0]
    by_name = {s.name: s for s in slide.shapes}
    expected = {
        "상단띠": (0, 0, 9896172, 137234),
        "상단띠2": (0, 195617, 8049346, 137234),
        "구분선": (-1, 404664, 9892977, 216024),
        "화면ID배경": (8121352, 188657, 1771625, 144000),
        "하단바": (0, 6716266, 9906000, 144000),
        "문서제목": (0, 6738252, 2648744, 100027),
        "쪽번호": (4734198, 6716266, 437604, 144000),
        "작성일": (8146752, 0, 504000, 144000),
    }
    for name, geom in expected.items():
        s = by_name[name]
        assert (s.left, s.top, s.width, s.height) == geom, name


def test_default_template_carries_no_third_party_copyright(tmp_path: Path):
    """기본 템플릿은 코드로 배포되므로 남의 저작권 표기가 들어가면 안 된다."""
    prs = Presentation(str(build_default_template(tmp_path / "d.pptx")))
    texts = []
    for slide in prs.slides:
        for s in slide.shapes:
            if s.has_text_frame:
                texts.append(s.text_frame.text)
            if s.has_table:
                for r in s.table.rows:
                    for c in r.cells:
                        texts.append(c.text)
    blob = "\n".join(texts)
    for banned in ("Copyright", "ⓒ", "Nurimedia", "All rights reserved"):
        assert banned not in blob, banned


def test_default_template_id_background_sits_behind_id_text(tmp_path: Path):
    slide = Presentation(str(build_default_template(tmp_path / "d.pptx"))).slides[0]
    order = [s.name for s in slide.shapes]
    assert order.index("화면ID배경") < order.index("화면ID")


def test_default_template_shell_text_starts_empty(tmp_path: Path):
    slide = Presentation(str(build_default_template(tmp_path / "d.pptx"))).slides[0]
    by_name = {s.name: s for s in slide.shapes}
    for name in ("문서제목", "작성일", "제목", "화면ID"):
        assert by_name[name].text_frame.text == "", name


def test_default_template_falls_back_to_computed_layout(tmp_path: Path):
    """실측 조합을 벗어나면 계산 배치로 떨어지되 슬라이드 안에 머문다."""
    prs = Presentation(str(build_default_template(tmp_path / "d.pptx", table_count=3)))
    slide = prs.slides[0]
    tables = sorted((s for s in slide.shapes if s.has_table), key=lambda s: s.left)
    assert len(tables) == 3
    assert tables[0].left >= 0
    assert tables[-1].left + tables[-1].width <= prs.slide_width
