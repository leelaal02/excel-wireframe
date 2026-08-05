# 기본 템플릿 실측 동형화 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 템플릿을 주지 않았을 때 생성되는 기본 템플릿을 실제 화면설계서와 같은 모양으로 만들고, 원본에서 레이아웃에 갇혀 편집 불가였던 껍데기 요소를 전부 슬라이드 도형으로 올려 수정 가능하게 한다.

**Architecture:** `default_template.py`를 실측 좌표 기반으로 재작성한다. 지금은 여백 0.25in에서 비율을 계산해 배치하지만, 앞으로는 실제 화면설계서에서 잰 EMU 값을 그대로 쓴다. 껍데기(상단 띠 2개, 구분선, 화면ID 배경, 하단 바, 쪽번호)는 슬라이드 위 도형으로 만들어 이름을 붙인다.

**Tech Stack:** Python 3.13, python-pptx 1.0.2, pytest

## 왜 하는가

현재 기본 템플릿은 16:9(13.33 × 7.50 in)에 비율 계산 배치라 실제 화면설계서(10.83 × 7.50 in)와 모양이 다르다. 그리고 사용자가 준 템플릿을 쓰는 경우, 상단 띠·하단 저작권 바·날짜·쪽번호가 **레이아웃**에 있어 슬라이드에서 클릭조차 되지 않는다. "모든 글자를 수정할 수 있는 형태"라는 요구는 이 두 가지를 함께 풀어야 충족된다.

## 저작권 처리 — 타협 없음

원본 레이아웃 하단에는 `Copyright ⓒ Nurimedia co., Ltd. All rights reserved.`가 있다. **이 문구를 기본 템플릿에 넣지 않는다.** 기본 템플릿은 스킬에 코드로 담겨 배포되고 템플릿 미제공 시마다 생성되므로, 넣으면 제3자의 저작권 표기가 모든 사용자 산출물에 박힌다.

그 자리(`L0 T6738252 W2648744 H100027`)에는 Excel 표지에서 읽은 `문서제목`이 들어간다. 회사 로고·브랜딩 그룹도 복제하지 않는다 — 색과 배치만 따른다.

## Global Constraints

- 스킬은 `skills/excel-wireframe/`. 스크립트는 `scripts/`에 평면 배치, 패키지를 만들지 않는다. import는 `from common import ...`.
- **경고 코드는 아홉 개뿐이다:** `no-image`, `no-detail`, `text-overflow`, `shape-not-found`, `slide-split`, `slot-shortage`, `screen-failed`, `image-convert-failed`, `orphan-row`. `tests/test_warning_codes.py`가 소스를 검사해 강제한다.
- **`build.py`는 openpyxl을 import하지 않는다.** 이 계획은 `build.py`를 건드리지 않는다.
- **껍데기 요소는 전부 슬라이드 도형이어야 한다.** 레이아웃이나 마스터에 그리면 이 작업의 목적이 사라진다. `prs.slide_layouts[6]`(빈 화면)을 그대로 쓰고 그 위에 도형을 올린다.
- **새 도형의 텍스트는 비워 둔다.** 자리표시 문구를 넣으면 표지가 없는 Excel에서 그 문구가 산출물에 인쇄된다. 예시 표 셀(`예시 설명 N`)은 지금처럼 유지한다 — `clear_unused_slots`가 지운다.
- `tests/fixtures.py`의 `make_template_pptx`는 `build_default_template`를 재사용한 뒤 도형 이름만 실제 샘플 이름으로 바꾼다. 생성기가 바뀌면 그 픽스처 산출물도 바뀌므로, 도형 수·이름·크기를 세는 기존 테스트가 깨진다. **깨지는 테스트는 새 기하에 맞게 고친다 — 기하를 되돌리지 않는다.**
- 테스트는 `pytest -q` 한 줄로 전부 돌아야 한다. `pytest.ini`가 `addopts = -q`를 설정하므로 명령줄 `-v`가 무시된다 — 개별 PASSED 줄이 필요하면 `--override-ini="addopts="`.
- 스킬 파일을 수정하면 `C:\Users\user\.claude\skills\excel-wireframe\`에 재설치하고 `__pycache__`를 제외한다.
- Windows·cp949. 한글이 깨지면 `PYTHONIOENCODING=utf-8`을 붙인다.

## 실측값

`work/demo/화면설계서_짧은버전.pptx`(원본 템플릿으로 생성한 결과물)와 그 레이아웃 `내용설명연결`에서 잰 값이다. 단위는 EMU.

**슬라이드:** `9906000 × 6858000` (10.83 × 7.50 in)

### 내용 요소 (원본에서도 슬라이드에 있던 것)

| 이름 | left | top | width | height | 서식 |
|---|---|---|---|---|---|
| `제목` | 3722514 | 0 | 1260000 | 144000 | 6.5pt, 가운데 정렬 |
| `화면ID` | 8121353 | 188640 | 1766860 | 138032 | 6.5pt, 좌측 정렬 |
| `화면이미지` | -12319 | 337940 | 9957099 | 4675235 | 자리 표시 사각형 |

### 상세 표 5개

공통: `W1971135 H1416117`, 열폭 `160215 / 1810920`, 행높이 `382457 / 268746 / 496168 / 268746`

| 표 | left |
|---|---|
| 상세표1 | -6849 |
| 상세표2 | 1974133 |
| 상세표3 | 3955115 |
| 상세표4 | 5936097 |
| 상세표5 | 7917077 |

`top`은 전부 `5253244`로 통일한다. 원본은 5253244와 5256806이 섞여 있는데 3562 EMU(0.004 in) 차이라 수작업 흔적이다. 간격은 균등하게 `1980982`.

셀 서식:

| | 크기 | 굵게 | 정렬 | 수직 | 여백 |
|---|---|---|---|---|---|
| 번호 칸(0열) | 6.5pt (82550) | 예 | 가운데 | 중앙 | 상하좌우 18000 |
| 내용 칸(1열) | 7pt (88900) | 아니오 | 좌측 | 기본 | 좌우 9525, 상 9525, 하 0 |

내용 칸 글꼴은 `맑은 고딕`.

### 껍데기 (레이아웃 → 슬라이드로 올릴 것)

| 이름 | left | top | width | height | 비고 |
|---|---|---|---|---|---|
| `상단띠` | 0 | 0 | 9896172 | 137234 | accent1 |
| `상단띠2` | 0 | 195617 | 8049346 | 137234 | accent1 |
| `구분선` | -1 | 404664 | 9892977 | 216024 | accent1 |
| `화면ID배경` | 8121352 | 188657 | 1771625 | 144000 | accent1, `화면ID` 뒤에 깔린다 |
| `하단바` | 0 | 6716266 | 9906000 | 144000 | accent1 |
| `문서제목` | 0 | 6738252 | 2648744 | 100027 | 6.5pt 흰색. 원본의 저작권 문구 자리 |
| `쪽번호` | 4734198 | 6716266 | 437604 | 144000 | 가운데 정렬, 흰색 |
| `작성일` | 8146752 | 0 | 504000 | 144000 | 우측 상단 |

테마 색은 Office 표준(accent1 `4F81BD`)이라 python-pptx 기본 테마와 일치한다. `MSO_THEME_COLOR.ACCENT_1`을 쓰면 원본과 같은 색이 나온다.

`화면ID배경`은 `화면ID` 텍스트보다 **먼저** 만들어야 뒤에 깔린다. python-pptx는 추가 순서가 z-order다.

## 파일 구조

| 파일 | 변경 |
|---|---|
| `scripts/default_template.py` | 실측 기하로 재작성, 껍데기 도형 추가 |
| `tests/test_default_template.py` | 깨지는 테스트 수정 + 새 기하 검증 |
| `tests/fixtures.py` | `make_template_pptx`의 명시 크기 인자 제거(기본값이 됨) |
| `references/mapping-schema.md` | 기본 템플릿 설명 갱신 |
| `SKILL.md` | 기본 템플릿 설명 갱신 |

---

### Task 1: default_template.py 재작성

**Files:**
- Modify: `skills/excel-wireframe/scripts/default_template.py`
- Modify: `tests/test_default_template.py`
- Modify: `tests/fixtures.py`
- Test: `tests/test_default_template.py`

**Interfaces:**
- `DEFAULT_SHAPE_NAMES` — 기존 5개(`title`/`screen_id`/`image`/`doc_title`/`date`)에 껍데기 이름은 넣지 않는다. 껍데기는 매핑 대상이 아니라 장식이다.
- `build_default_template(path, slide_width_emu=9906000, slide_height_emu=6858000, table_count=5, rows_per_table=4) -> Path` — 기본값이 바뀐다.
- `default_template_mapping(path, table_count=5) -> dict` — 시그니처·반환 구조 불변. `shapes`의 값도 그대로다.

**핵심 설계 결정:** 실측값은 `table_count=5`, `rows_per_table=4`, 슬라이드 9906000 × 6858000일 때의 값이다. 그 조합이면 실측 좌표를 그대로 쓰고, 벗어나면 지금처럼 계산으로 떨어진다. 실측 상수를 모듈 상단에 이름 붙여 모아두고, 기본 조합인지 판정하는 헬퍼를 하나 둔다. 이렇게 하면 기본 경로는 원본과 픽셀 단위로 같고, 파라미터를 바꾸면 합리적으로 늘어난다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_default_template.py`에 추가한다. 기존 테스트는 아직 손대지 않는다 — 어떤 것이 왜 깨지는지 Step 2에서 확인한 뒤 고친다.

```python
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
```

- [ ] **Step 2: 테스트 실패 확인과 기존 테스트 영향 조사**

Run: `python -m pytest tests/ --override-ini="addopts=" -q 2>&1 | tail -30`

새 테스트는 실패한다. **기존 테스트 중 어떤 것이 함께 깨지는지 목록을 만든다.** 예상되는 것:

- `test_default_template_is_16_9` — `12192000`을 기대한다. 이름부터 거짓이 되므로 함수명을 `test_default_template_uses_measured_size`로 바꾸고 `9906000`을 기대하게 고친다.
- `test_default_template_size_is_configurable` — `slide_width_emu=9906000`을 넘겨 "기본값과 다른 값도 된다"를 검증한다. 그 값이 기본값이 되면 아무것도 검증하지 못한다. 다른 폭(예: `12192000`)으로 바꾼다.
- `test_default_template_raises_when_rows_per_table_too_large` / `test_default_template_allows_max_workable_rows_per_table` — 상한값이 새 기하에서 달라진다. 실제 상한을 계산해 기대값을 고친다.
- `test_default_template_table_height_scales_with_rows_per_table` — 실측 경로는 고정 행높이를 쓰므로, 스케일 검증은 계산 경로(실측 조합을 벗어난 `rows_per_table`)에서 해야 한다.
- `test_default_template_meta_shapes_stay_inside_slide` / `..._are_empty` — 좌표가 바뀌므로 재확인.
- `tests/fixtures.py`의 `make_template_pptx`가 `slide_width_emu=9906000`을 명시적으로 넘기고 있다면 이제 불필요하다. 제거해도 동작이 같아야 한다.

깨진 것을 하나씩 새 기하 기준으로 고친다. **기대값을 옛 기하로 되돌려 통과시키지 않는다.**

- [ ] **Step 3: default_template.py 재작성**

실측 상수를 모듈 상단에 모으고, 기본 조합일 때만 쓴다.

```python
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
```

이어서 본체를 쓴다. 순서가 중요하다 — python-pptx는 추가 순서가 z-order이므로 **배경 → 내용** 순으로 만든다.

```python
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
        for ci, w in enumerate(col_widths):
            table.columns[ci].width = Emu(w)
        for ri, h in enumerate(row_heights[:rows_per_table]):
            table.rows[ri].height = Emu(h)
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
```

`_fill_cell`은 셀 서식을 한 곳에 모은다.

```python
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
```

`_computed_layout`은 실측 조합을 벗어났을 때의 폴백이다. 기존 코드의 비율 계산을 그대로 옮겨 오되, 반환 형태를 실측 경로와 맞춘다. `MIN_IMAGE_HEIGHT_EMU` 검증도 이 경로에 유지한다 — 실측 경로는 고정값이라 검증이 필요 없다.

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_default_template.py --override-ini="addopts=" -v`
Expected: 새 테스트 전부 통과, 고친 기존 테스트도 통과

- [ ] **Step 5: 전체 스위트**

Run: `python -m pytest --override-ini="addopts=" -q`

`tests/fixtures.py`의 `make_template_pptx`가 만드는 산출물이 바뀌므로 `test_pptx_scan.py`, `test_slide_fill_*.py`, `test_build.py`, `test_slide_clone.py`가 영향을 받을 수 있다. 깨지면 새 기하 기준으로 고친다. 특히 표 이름 재명명(`상세표N` → `표 N`)과 도형 수를 세는 곳을 확인한다.

- [ ] **Step 6: 눈으로 확인**

```bash
export PYTHONIOENCODING=utf-8
python -c "import sys; sys.path.insert(0,'skills/excel-wireframe/scripts'); from default_template import build_default_template; build_default_template('work/parity-check.pptx')"
python - <<'PY'
import sys; sys.path.insert(0, "skills/excel-wireframe/scripts")
from pptx import Presentation
a = Presentation("work/parity-check.pptx").slides[0]
b = Presentation("work/demo/화면설계서_짧은버전.pptx").slides[0]
ba = {s.name: (s.left, s.top, s.width, s.height) for s in a.shapes}
print("생성된 기본 템플릿 도형 %d개:" % len(a.shapes))
for n, g in ba.items(): print("   %-12s %s" % (n, g))
print("\n목표 파일의 슬라이드 도형 %d개:" % len(b.shapes))
for s in b.shapes: print("   %-20s %s" % (s.name, (s.left, s.top, s.width, s.height)))
PY
```

목표 파일과 좌표가 일치하는지 대조한다. 목표 파일은 껍데기가 레이아웃에 있어 슬라이드 도형이 8개뿐이고, 새 기본 템플릿은 껍데기를 올렸으므로 더 많다 — 그 차이는 의도된 것이다. **내용 요소(제목·화면ID·화면이미지·표 5개)의 좌표가 같아야 한다.**

- [ ] **Step 7: 커밋**

```bash
git add skills/excel-wireframe/scripts/default_template.py tests/test_default_template.py tests/fixtures.py
git commit -m "feat: 기본 템플릿을 실제 화면설계서 실측 기하로 재작성"
```

---

### Task 2: 문서 갱신과 설치

**Files:**
- Modify: `skills/excel-wireframe/references/mapping-schema.md`
- Modify: `skills/excel-wireframe/SKILL.md`
- Modify: `README.md`
- Test: `tests/test_skill_docs.py`

- [ ] **Step 1: 문서에서 옛 설명 찾기**

```bash
grep -rn "16:9\|13.33\|12192000" skills/excel-wireframe/ README.md docs/ | grep -v "^docs/superpowers/plans/2026-08-05"
```

기본 템플릿을 16:9로 설명하는 문장이 여러 곳에 있다. 전부 새 크기로 고친다.

- [ ] **Step 2: `references/mapping-schema.md` 갱신**

기본 템플릿 절을 새 구조로 고친다. 담아야 할 것:

- 슬라이드 10.83 × 7.50 in (9906000 × 6858000 EMU), 실제 화면설계서와 같은 크기
- 도형: `제목`, `화면ID`, `화면이미지`, `상세표1`~`상세표5`(각 4행 = 20슬롯), `문서제목`, `작성일`
- 장식: `상단띠`, `상단띠2`, `구분선`, `화면ID배경`, `하단바`, `쪽번호` — **전부 슬라이드 도형이라 PowerPoint에서 직접 수정된다**
- `build_default_template`의 파라미터를 기본값에서 벗어나게 주면 실측 좌표 대신 계산 배치로 떨어진다는 것
- 제3자 저작권 문구는 복제하지 않으며, 원본에서 그 문구가 있던 자리에 `문서제목`이 들어간다는 것

- [ ] **Step 3: `SKILL.md` 갱신**

1단계의 기본 템플릿 설명을 고친다. 지금은 "기본 템플릿(16:9, 제목 바(`작성일` 포함) + 이미지 자리 + 상세 표 5개 × 4행 + 표 아래 `문서제목`)"으로 되어 있다. 새 설명은 크기와 껍데기가 편집 가능하다는 사실을 담는다.

"하지 않는 것" 절에 한 줄 더한다 — 기본 템플릿에 제3자 저작권 표기를 넣지 않는다는 것.

- [ ] **Step 4: `README.md` 갱신**

"현재 한계"의 다음 두 줄을 고친다.

- "기본 템플릿이 16:9 비율 계산 배치라 실제 화면설계서(10.83in)와 모양이 다르다" — 해결됐으므로 삭제
- "사용자 템플릿의 레이아웃에 박힌 요소는 건드리지 못한다" — 유지한다. 이건 사용자가 준 템플릿에 여전히 해당한다. 다만 기본 템플릿은 그렇지 않다는 것을 덧붙인다.

- [ ] **Step 5: 테스트와 재설치**

```bash
python -m pytest --override-ini="addopts=" -q
rm -rf /c/Users/user/.claude/skills/excel-wireframe
cp -r skills/excel-wireframe /c/Users/user/.claude/skills/
find /c/Users/user/.claude/skills/excel-wireframe -name __pycache__ -type d -exec rm -rf {} +
diff -r skills/excel-wireframe /c/Users/user/.claude/skills/excel-wireframe && echo "설치 일치"
```

- [ ] **Step 6: 실제 샘플로 끝까지 확인**

```bash
export PYTHONIOENCODING=utf-8
S=~/.claude/skills/excel-wireframe/scripts
rm -rf work/parity && mkdir -p work/parity
python "$S/analyze.py" --excel "짧은 버전.xlsx" --out work/parity/structure-report.json
```

리포트의 `template.slide_size_in`이 `10.83 x 7.50`인지 확인한다. 그다음 매핑을 조립해 `extract.py`와 `build.py`를 돌리고, 생성물의 검증 6항목이 모두 통과하는지, 슬라이드 크기가 9906000인지 확인한다. 명령과 출력을 리포트에 붙인다.

- [ ] **Step 7: 커밋**

```bash
git add skills/excel-wireframe/ README.md tests/
git commit -m "docs: 기본 템플릿 실측 동형화 반영"
```

---

## 검증 매트릭스

| 요구 | 커버 위치 |
|---|---|
| 슬라이드 크기 9906000 × 6858000 | Task 1 Step 1 |
| 제목·화면ID·이미지 자리 좌표 실측 일치 | Task 1 Step 1 |
| 표 5개 좌표·열폭·행높이 실측 일치 | Task 1 Step 1 |
| 셀 서식(6.5pt 굵게 가운데 / 7pt 맑은 고딕 좌측) | Task 1 Step 1 |
| 껍데기가 레이아웃이 아닌 슬라이드에 있음 | Task 1 Step 1 |
| 껍데기 좌표 실측 일치 | Task 1 Step 1 |
| 제3자 저작권 문구 부재 | Task 1 Step 1 |
| `화면ID배경`이 `화면ID` 뒤에 깔림 | Task 1 Step 1 |
| 새 도형 텍스트가 비어 있음 | Task 1 Step 1 |
| 실측 조합을 벗어나면 계산 배치로 폴백 | Task 1 Step 1 |
| 문서가 코드와 일치 | Task 2 |
| 실제 샘플 파이프라인 통과 | Task 2 Step 6 |

## 범위 밖

- 사용자가 `--template`으로 준 파일의 레이아웃 요소는 여전히 건드리지 않는다. 그건 그 파일 소유다.
- 표지 슬라이드·목차는 만들지 않는다.
- `layout` 모드 생성은 여전히 지원하지 않는다.
