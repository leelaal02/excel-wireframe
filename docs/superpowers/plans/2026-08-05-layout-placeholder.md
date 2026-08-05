# 레이아웃 선택 + placeholder 채우기 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 슬라이드를 예시 슬라이드 복제가 아니라 **레이아웃 선택 + placeholder 채우기**로 만들고, `screens.json`을 SSOT에서 중간 산출물로 강등한다.

**Architecture:** 새 모듈 `slide_layout.py`가 레이아웃 탐색·placeholder 보충 상속·본문 영역 분할·상세표 생성을 담당한다. `build.py`는 `mode`에 따라 슬라이드를 만드는 방법만 갈라 쓰고, 값을 채우는 `_fill_page`는 두 모드가 공유한다 — layout 모드에서 채운 placeholder에 mapping이 정한 **이름을 붙이기** 때문이다. 기본 템플릿은 슬라이드 0장 + 레이아웃 1개가 된다.

**Tech Stack:** Python 3.13, python-pptx 1.0.2, openpyxl, Pillow, pytest

설계 문서: `docs/superpowers/specs/2026-08-05-layout-placeholder-design.md`

## Global Constraints

- 스크립트는 `skills/excel-wireframe/scripts/`에 **평면 배치**. 패키지로 만들지 않는다. 모듈 간 import는 `from common import ...` 형태.
- **`build.py`는 openpyxl을 import하지 않는다.** `tests/test_build.py`가 소스를 검사해 강제한다.
- **경고 코드는 아홉 개뿐이다:** `no-image`, `no-detail`, `text-overflow`, `shape-not-found`, `slide-split`, `slot-shortage`, `screen-failed`, `image-convert-failed`, `orphan-row`. `tests/test_warning_codes.py`가 소스를 검사해 강제한다. **새 코드를 추가하지 않는다.**
- 모든 CLI 진입점은 `setup_stdio()`를 먼저 호출한다.
- 표 셀에 값을 쓸 때는 `slide_fill.set_cell_text`를 쓴다. 런을 새로 만들면 서식이 초기화된다.
- 슬라이드 복제는 `slide_clone.clone_slide`만 쓴다.
- **슬라이드 크기를 바꾸지 않는다.** 템플릿 크기를 승계한다.
- **화면 단위로 예외를 격리한다.** 한 화면이 실패해도 나머지는 완성하고 `screen-failed` 경고를 남긴다.
- **상세 표 번호는 Excel 값을 그대로 쓴다.** 재부여 금지.
- **기본 템플릿에 제3자 저작권 문구를 넣지 않는다.** 원본에서 그 문구가 있던 자리(`L0 T6738252 W2648744 H100027`)에는 표지에서 읽은 `문서제목`이 들어간다.
- 테스트는 `python -m pytest -q` 한 줄로 전부 돌아야 한다. `pytest.ini`가 `addopts = -q`를 설정하므로 명령줄 `-v`가 무시된다 — 개별 PASSED 줄이 필요하면 `--override-ini="addopts="`.
- Windows·cp949. 한글이 깨지면 `PYTHONIOENCODING=utf-8`을 붙인다.
- 스킬 파일을 수정하면 `C:\Users\user\.claude\skills\excel-wireframe\`에 재설치한다(`__pycache__` 제외).

## 실측 상수

원본 화면설계서에서 잰 값이다. 단위는 EMU. 슬라이드는 `9906000 × 6858000`.

**placeholder 자리** — python-pptx 기본 템플릿의 `Title and Content` 레이아웃(`prs.slide_layouts[1]`)이 가진 다섯 개를 옮겨 쓴다.

| idx | 종류 | 쓸 자리 | left | top | width | height |
|---|---|---|---|---|---|---|
| 0 | TITLE | 제목 | 3722514 | 0 | 1260000 | 144000 |
| 1 | OBJECT | 화면ID | 8121353 | 188640 | 1766860 | 138032 |
| 10 | DATE | 작성일 | 8146752 | 0 | 504000 | 144000 |
| 11 | FOOTER | 문서제목 | 0 | 6738252 | 2648744 | 100027 |
| 12 | SLIDE_NUMBER | 쪽번호 | 4734198 | 6716266 | 437604 | 144000 |

**껍데기 도형** — 레이아웃 위의 일반 도형. 전부 테마 accent1, 테두리 없음.

| 이름 | left | top | width | height |
|---|---|---|---|---|
| 상단띠 | 0 | 0 | 9896172 | 137234 |
| 상단띠2 | 0 | 195617 | 8049346 | 137234 |
| 구분선 | -1 | 404664 | 9892977 | 216024 |
| 화면ID배경 | 8121352 | 188657 | 1771625 | 144000 |
| 하단바 | 0 | 6716266 | 9906000 | 144000 |

**본문 영역과 표**

| 항목 | 값 |
|---|---|
| `content_area` 기본값 | `[-12319, 337940, 9957099, 6331421]` |
| 행높이 | `[382457, 268746, 496168, 268746]` — 4행을 넘으면 마지막 값 반복 |
| 열폭 비율 | `160215 : 1810920` (번호 칸 : 내용 칸) |
| 이미지 자리 최소 높이 | 1인치 (`EMU_PER_INCH`) |

**셀 서식**

| | 크기 | 굵게 | 정렬 | 수직 | 여백 |
|---|---|---|---|---|---|
| 번호 칸(0열) | 6.5pt | 예 | 가운데 | 중앙 | 상하좌우 18000 |
| 내용 칸(1열) | 7pt, 맑은 고딕 | 아니오 | 좌측 | 기본 | 좌·우·상 9525, 하 0 |

## 파일 구조

| 파일 | 책임 | 변경 |
|---|---|---|
| `scripts/slide_layout.py` | 레이아웃 탐색, placeholder 보충 상속·이름 부여·정리, 본문 영역 분할, 상세표 생성 | **신규** |
| `scripts/build.py` | `mode`에 따라 슬라이드 생성 방법을 고른다. 값 채우기는 공유 | 수정 |
| `scripts/default_template.py` | 레이아웃 1개 + 슬라이드 0장짜리 기본 템플릿 | 재작성 |
| `scripts/analyze.py` | 기본 템플릿용 layout 모드 매핑 제안 | 수정 |
| `scripts/extract.py` | `screens.json`을 그냥 덮어쓴다 | 수정 |
| `scripts/slide_clone.py` | clone 모드 전용 | 변경 없음 |
| `scripts/verify.py` | 도형 **이름**으로 검증한다 — layout 모드가 이름을 붙이므로 | 변경 없음 |
| `tests/fixtures.py` | clone 모드용 예시 슬라이드 템플릿을 layout 경로로 만든다 | 수정 |

---

### Task 1: slide_layout.py — 레이아웃 탐색과 placeholder 처리

**Files:**
- Create: `skills/excel-wireframe/scripts/slide_layout.py`
- Test: `tests/test_slide_layout.py` (신규)

**Interfaces:**
- Produces:
  - `find_layout(prs, spec) -> SlideLayout` — `spec`은 레이아웃 이름(str) 또는 인덱스(int). 모든 마스터를 훑는다. 못 찾으면 `ValueError`
  - `inherit_placeholders(slide, layout) -> list[int]` — 레이아웃에 있으나 슬라이드에 없는 placeholder를 XML 복제하고, 복제한 idx 목록을 반환
  - `name_placeholders(slide, placeholders_cfg, shapes_cfg, warns, screen_id) -> None` — `placeholders_cfg`의 `{키: idx}`를 돌며 그 idx의 placeholder 이름을 `shapes_cfg.get(키, 키)`로 바꾼다. idx가 없으면 `shape-not-found` 경고
  - `drop_empty_placeholders(slide) -> int` — 텍스트가 비어 있고 필드(`a:fld`)도 없는 placeholder를 제거하고 제거 수를 반환

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_slide_layout.py`를 새로 만든다.

```python
# -*- coding: utf-8 -*-
from pathlib import Path

import pytest
from common import Warnings
from pptx import Presentation
from pptx.util import Emu
from slide_layout import (
    drop_empty_placeholders,
    find_layout,
    inherit_placeholders,
    name_placeholders,
)


def _prs():
    """Title and Content 레이아웃(TITLE/OBJECT/DATE/FOOTER/SLIDE_NUMBER)을 가진 프레젠테이션."""
    return Presentation()


def test_find_layout_by_name():
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    assert lay.name == "Title and Content"


def test_find_layout_by_index():
    prs = _prs()
    assert find_layout(prs, 1).name == prs.slide_layouts[1].name


def test_find_layout_raises_for_unknown_name():
    prs = _prs()
    with pytest.raises(ValueError) as exc:
        find_layout(prs, "없는레이아웃")
    assert "없는레이아웃" in str(exc.value)


def test_inherit_placeholders_adds_date_and_footer():
    """python-pptx는 date/footer/slidenumber를 복제하지 않는다 — 우리가 채운다."""
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    before = {ph.placeholder_format.idx for ph in slide.placeholders}
    assert 10 not in before

    added = inherit_placeholders(slide, lay)

    after = {ph.placeholder_format.idx for ph in slide.placeholders}
    assert {10, 11, 12} <= after
    assert set(added) == after - before


def test_inherit_placeholders_is_idempotent():
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    inherit_placeholders(slide, lay)
    n = len(list(slide.placeholders))
    assert inherit_placeholders(slide, lay) == []
    assert len(list(slide.placeholders)) == n


def test_name_placeholders_renames_by_idx():
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    inherit_placeholders(slide, lay)

    warns = Warnings()
    name_placeholders(
        slide,
        {"title": 0, "screen_id": 1, "작성일": 10},
        {"title": "제목", "screen_id": "화면ID"},
        warns,
        "SCR001",
    )

    names = [s.name for s in slide.shapes]
    assert "제목" in names
    assert "화면ID" in names
    assert "작성일" in names  # shapes에 없으면 키를 그대로 이름으로 쓴다
    assert len(warns) == 0


def test_name_placeholders_warns_for_missing_idx():
    prs = _prs()
    slide = prs.slides.add_slide(find_layout(prs, "Title and Content"))
    warns = Warnings()
    name_placeholders(slide, {"title": 0, "없는것": 99}, {}, warns, "SCR001")
    items = warns.to_list()
    assert len(items) == 1
    assert items[0]["code"] == "shape-not-found"
    assert "99" in items[0]["message"]


def test_drop_empty_placeholders_removes_blank_ones():
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    slide.placeholders[0].text_frame.text = "제목 있음"

    removed = drop_empty_placeholders(slide)

    assert removed >= 1
    remaining = [ph.placeholder_format.idx for ph in slide.placeholders]
    assert 0 in remaining
    assert 1 not in remaining


def test_drop_empty_placeholders_keeps_field_placeholders():
    """쪽번호는 자동 번호 필드라 텍스트가 비어 보여도 지우면 안 된다."""
    prs = _prs()
    lay = find_layout(prs, "Title and Content")
    slide = prs.slides.add_slide(lay)
    inherit_placeholders(slide, lay)
    drop_empty_placeholders(slide)
    assert 12 in [ph.placeholder_format.idx for ph in slide.placeholders]
```

- [ ] **Step 2: 테스트가 실패하는 것을 확인**

Run: `python -m pytest tests/test_slide_layout.py --override-ini="addopts=" -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'slide_layout'`

- [ ] **Step 3: slide_layout.py 작성 (이 태스크 분량만)**

```python
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
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_slide_layout.py --override-ini="addopts=" -v`
Expected: 8개 전부 PASS

- [ ] **Step 5: 경고 코드 검사와 전체 스위트**

Run: `python -m pytest --override-ini="addopts=" -q`
Expected: 전부 PASS (기존 145개 + 새 8개)

- [ ] **Step 6: 커밋**

```bash
git add skills/excel-wireframe/scripts/slide_layout.py tests/test_slide_layout.py
git commit -m "feat: 레이아웃 탐색과 placeholder 상속·명명·정리"
```

---

### Task 2: slide_layout.py — 본문 영역 분할과 상세표 생성

**Files:**
- Modify: `skills/excel-wireframe/scripts/slide_layout.py`
- Modify: `tests/test_slide_layout.py`

**Interfaces:**
- Consumes: Task 1의 `slide_layout` 모듈
- Produces:
  - `split_content_area(area, table_count, rows_per_table) -> tuple[tuple, list[tuple]]` — `area`는 `(left, top, width, height)`. `(image_box, [table_box, ...])`를 반환. 이미지 자리가 1인치 미만이면 `ValueError`
  - `add_image_anchor(slide, box, name) -> Shape` — 이미지가 들어갈 자리 사각형
  - `add_detail_table(slide, box, rows, name) -> GraphicFrame` — 상세표 하나를 만들고 셀 서식을 적용
  - 모듈 상수: `ROW_HEIGHTS`, `COL_WIDTH_RATIO`, `MIN_IMAGE_HEIGHT_EMU`, `DEFAULT_CONTENT_AREA`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_slide_layout.py` 끝에 덧붙인다.

```python
from pptx.util import Pt
from slide_layout import (
    DEFAULT_CONTENT_AREA,
    add_detail_table,
    add_image_anchor,
    split_content_area,
)


def test_split_content_area_puts_tables_at_the_bottom():
    area = DEFAULT_CONTENT_AREA
    image_box, table_boxes = split_content_area(area, 5, 4)

    assert len(table_boxes) == 5
    # 표는 본문 영역 아래쪽에 붙는다
    table_top = table_boxes[0][1]
    table_height = table_boxes[0][3]
    assert table_top + table_height == area[1] + area[3]
    # 이미지는 위쪽 나머지를 전부 쓴다
    assert image_box[1] == area[1]
    assert image_box[1] + image_box[3] == table_top


def test_split_content_area_matches_measured_geometry():
    """실측 조합(표 5개 × 4행)에서 표 상단이 원본과 같은 자리에 온다."""
    _, table_boxes = split_content_area(DEFAULT_CONTENT_AREA, 5, 4)
    assert table_boxes[0][1] == 5253244
    assert table_boxes[0][3] == 382457 + 268746 + 496168 + 268746


def test_split_content_area_divides_width_evenly():
    area = (0, 0, 10000, 4000000)
    _, boxes = split_content_area(area, 5, 4)
    assert [b[0] for b in boxes] == [0, 2000, 4000, 6000, 8000]
    assert all(b[2] == 2000 for b in boxes)


def test_split_content_area_grows_tables_with_rows():
    _, four = split_content_area(DEFAULT_CONTENT_AREA, 5, 4)
    _, six = split_content_area(DEFAULT_CONTENT_AREA, 5, 6)
    assert six[0][3] > four[0][3]
    # 마지막 행높이를 반복한다
    assert six[0][3] == four[0][3] + 268746 * 2


def test_split_content_area_raises_when_image_slot_too_small():
    with pytest.raises(ValueError) as exc:
        split_content_area(DEFAULT_CONTENT_AREA, 5, 40)
    assert "이미지" in str(exc.value)


def test_add_detail_table_applies_measured_formatting():
    prs = _prs()
    slide = prs.slides.add_slide(find_layout(prs, "Title and Content"))
    frame = add_detail_table(slide, (0, 0, 1971135, 1416117), 4, "상세표1")

    assert frame.name == "상세표1"
    table = frame.table
    assert len(table.rows) == 4
    assert len(table.columns) == 2
    assert [r.height for r in table.rows] == [382457, 268746, 496168, 268746]
    # 열폭은 실측 비율을 표 폭에 맞춰 나눈다
    assert sum(c.width for c in table.columns) == 1971135
    assert table.columns[0].width < table.columns[1].width

    num, txt = table.cell(0, 0), table.cell(0, 1)
    assert num.margin_left == 18000
    assert txt.margin_left == 9525
    assert txt.margin_bottom == 0
    assert num.text_frame.paragraphs[0].runs[0].font.size == Pt(6.5)
    assert num.text_frame.paragraphs[0].runs[0].font.bold is True
    assert txt.text_frame.paragraphs[0].runs[0].font.size == Pt(7)
    assert txt.text_frame.paragraphs[0].runs[0].font.name == "맑은 고딕"


def test_add_detail_table_starts_empty():
    """빈 Excel에서 예시 문구가 산출물에 찍히면 안 된다."""
    prs = _prs()
    slide = prs.slides.add_slide(find_layout(prs, "Title and Content"))
    table = add_detail_table(slide, (0, 0, 1971135, 1416117), 4, "상세표1").table
    for r in range(4):
        assert table.cell(r, 0).text == ""
        assert table.cell(r, 1).text == ""


def test_add_image_anchor_uses_the_box():
    prs = _prs()
    slide = prs.slides.add_slide(find_layout(prs, "Title and Content"))
    shp = add_image_anchor(slide, (100, 200, 3000, 4000), "화면이미지")
    assert (shp.left, shp.top, shp.width, shp.height) == (100, 200, 3000, 4000)
    assert shp.name == "화면이미지"
```

- [ ] **Step 2: 테스트가 실패하는 것을 확인**

Run: `python -m pytest tests/test_slide_layout.py --override-ini="addopts=" -v`
Expected: FAIL — `ImportError: cannot import name 'split_content_area'`

- [ ] **Step 3: slide_layout.py에 구현 추가**

파일 상단 import에 다음을 더한다.

```python
from common import EMU_PER_INCH
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt
```

상수와 함수를 파일 끝에 더한다.

```python
# --- 실측 상수 (원본 화면설계서 화면 페이지) ---
DEFAULT_CONTENT_AREA = (-12319, 337940, 9957099, 6331421)
ROW_HEIGHTS = [382457, 268746, 496168, 268746]
COL_WIDTH_RATIO = (160215, 1810920)
MIN_IMAGE_HEIGHT_EMU = EMU_PER_INCH  # 1인치

SLOT_BORDER = RGBColor(0xBB, 0xBB, 0xBB)
SLOT_FILL = RGBColor(0xF5, 0xF6, 0xF8)


def _row_heights(rows_per_table: int) -> list[int]:
    """실측 행높이를 쓰되, 4행을 넘으면 마지막 값을 반복한다."""
    if rows_per_table <= len(ROW_HEIGHTS):
        return ROW_HEIGHTS[:rows_per_table]
    tail = [ROW_HEIGHTS[-1]] * (rows_per_table - len(ROW_HEIGHTS))
    return ROW_HEIGHTS + tail


def split_content_area(area, table_count: int, rows_per_table: int):
    """본문 영역을 이미지 자리와 상세표 자리로 나눈다.

    표는 아래쪽에 붙이고 폭을 균등 분할한다. 원본의 표 간격은 0.01인치라
    사실상 붙어 있으므로 간격을 두지 않는다. 이미지는 위쪽 나머지를 전부 쓴다.
    """
    left, top, width, height = area
    table_h = sum(_row_heights(rows_per_table))
    table_top = top + height - table_h
    image_h = table_top - top

    if image_h < MIN_IMAGE_HEIGHT_EMU:
        max_rows = 0
        while sum(_row_heights(max_rows + 1)) <= height - MIN_IMAGE_HEIGHT_EMU:
            max_rows += 1
        raise ValueError(
            "rows_per_table=%d면 이미지 자리 높이가 %.2fin로 너무 작아집니다"
            "(최소 %.2fin 필요). 이 content_area 높이(%.2fin)에서는 "
            "rows_per_table을 %d 이하로 쓰세요."
            % (rows_per_table, image_h / EMU_PER_INCH,
               MIN_IMAGE_HEIGHT_EMU / EMU_PER_INCH, height / EMU_PER_INCH,
               max_rows)
        )

    table_w = width // table_count
    table_boxes = [
        (left + table_w * i, table_top, table_w, table_h)
        for i in range(table_count)
    ]
    return (left, top, width, image_h), table_boxes


def add_image_anchor(slide, box, name: str):
    """이미지가 들어갈 자리 사각형. place_image가 이것을 지우고 그림으로 바꾼다.

    이미지가 없는 화면에서는 이 사각형이 그대로 남아 '여기에 스크린샷' 자리로
    보인다 — clone 모드에서 템플릿의 이미지 자리 도형이 남는 것과 같은 동작이다.
    """
    left, top, width, height = box
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(left), Emu(top),
                                 Emu(width), Emu(height))
    shp.name = name
    shp.fill.solid()
    shp.fill.fore_color.rgb = SLOT_FILL
    shp.line.color.rgb = SLOT_BORDER
    shp.text_frame.text = ""
    return shp


def _format_cell(cell, size_pt, bold, align, anchor, margin, font=None):
    cell.margin_left = Emu(margin)
    cell.margin_right = Emu(margin)
    cell.margin_top = Emu(margin)
    cell.margin_bottom = Emu(0 if font else margin)
    if anchor is not None:
        cell.vertical_anchor = anchor
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = ""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if font:
        run.font.name = font


def add_detail_table(slide, box, rows_per_table: int, name: str):
    """상세표 하나를 만들고 실측 서식을 적용한다. 셀은 비운 채로 둔다."""
    left, top, width, height = box
    frame = slide.shapes.add_table(rows_per_table, 2, Emu(left), Emu(top),
                                   Emu(width), Emu(height))
    frame.name = name
    table = frame.table

    ratio_total = sum(COL_WIDTH_RATIO)
    no_w = width * COL_WIDTH_RATIO[0] // ratio_total
    table.columns[0].width = Emu(no_w)
    table.columns[1].width = Emu(width - no_w)
    for ri, rh in enumerate(_row_heights(rows_per_table)):
        table.rows[ri].height = Emu(rh)

    for r in range(rows_per_table):
        _format_cell(table.cell(r, 0), 6.5, True, PP_ALIGN.CENTER,
                     MSO_ANCHOR.MIDDLE, 18000)
        _format_cell(table.cell(r, 1), 7.0, False, PP_ALIGN.LEFT, None,
                     9525, font="맑은 고딕")
    return frame
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_slide_layout.py --override-ini="addopts=" -v`
Expected: 전부 PASS

- [ ] **Step 5: 전체 스위트**

Run: `python -m pytest --override-ini="addopts=" -q`

- [ ] **Step 6: 커밋**

```bash
git add skills/excel-wireframe/scripts/slide_layout.py tests/test_slide_layout.py
git commit -m "feat: 본문 영역 분할과 상세표 생성"
```

---

### Task 3: build.py에 layout 모드 연결

**Files:**
- Modify: `skills/excel-wireframe/scripts/build.py:117-190` (`build` 함수)
- Modify: `tests/test_build.py`

**Interfaces:**
- Consumes: `slide_layout.find_layout`, `inherit_placeholders`, `name_placeholders`, `drop_empty_placeholders`, `split_content_area`, `add_image_anchor`, `add_detail_table`, `DEFAULT_CONTENT_AREA`
- Produces: `build(screens_data, mapping, work_dir, out_path, warns) -> dict` — 시그니처·반환 구조 불변. `template.mode`가 `"layout"`이면 새 경로를 탄다

**핵심 설계 결정:** `_fill_page`는 두 모드가 그대로 공유한다. layout 모드가 placeholder와 표에 mapping이 정한 **이름**을 붙이기 때문에, 이름으로 도형을 찾는 기존 코드가 손대지 않고 동작한다. `verify.py`도 같은 이유로 변경이 없다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_build.py` 끝에 덧붙인다.

```python
from pptx.util import Emu


def _layout_template(path: Path) -> Path:
    """placeholder만 있는 레이아웃형 템플릿 (예시 슬라이드 없음)."""
    prs = Presentation()
    prs.slide_width = Emu(9906000)
    prs.slide_height = Emu(6858000)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path


def _layout_mapping(template: Path) -> dict:
    return {
        "version": 1,
        "excel": {"layout": "sheet-per-screen"},
        "template": {
            "file": str(template),
            "mode": "layout",
            "layout": "Title and Content",
            "placeholders": {"title": 0, "screen_id": 1, "작성일": 10,
                             "문서제목": 11},
            "shapes": {
                "title": "제목",
                "screen_id": "화면ID",
                "image": "화면이미지",
                "작성일": "작성일",
                "문서제목": "문서제목",
                "detail_tables": ["상세표1", "상세표2", "상세표3",
                                  "상세표4", "상세표5"],
            },
            "content_area": [-12319, 337940, 9957099, 6331421],
            "detail_tables": {"count": 5, "rows": 4},
            "table_columns": {"no": 0, "text": 1},
        },
        "options": {
            "detail_text_source": "desc",
            "overflow": "split",
            "clear_unused_slots": True,
        },
    }


def test_build_layout_mode_creates_slides(tmp_path: Path):
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    report = build(_screens(6), _layout_mapping(tpl), tmp_path, out, Warnings())

    assert report["slides"] == 1
    slide = Presentation(str(out)).slides[0]
    by_name = {s.name: s for s in slide.shapes}
    assert by_name["제목"].text_frame.text == "이용기관 목록"
    assert by_name["화면ID"].text_frame.text == "SCR001"


def test_build_layout_mode_fills_detail_tables(tmp_path: Path):
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    build(_screens(6), _layout_mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    tables = [s for s in slide.shapes if s.has_table]
    assert len(tables) == 5
    first = next(t for t in tables if t.name == "상세표1").table
    assert first.cell(0, 0).text == "1"
    assert first.cell(0, 1).text == "설명 1"


def test_build_layout_mode_splits_when_details_exceed_slots(tmp_path: Path):
    """슬롯은 5표 × 4행 = 20개다."""
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    warns = Warnings()
    report = build(_screens(25), _layout_mapping(tpl), tmp_path, out, warns)

    assert report["slides"] == 2
    assert report["split"] == ["SCR001"]
    assert any(w["code"] == "slide-split" for w in warns.to_list())


def test_build_layout_mode_places_image(tmp_path: Path):
    img = make_png(tmp_path / "images" / "SCR001.png")
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    build(_screens(3, image="images/SCR001.png"), _layout_mapping(tpl),
          tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    assert any("PICTURE" in str(s.shape_type) for s in slide.shapes)
    assert not any(s.name == "화면이미지" for s in slide.shapes)


def test_build_layout_mode_drops_empty_placeholders(tmp_path: Path):
    """값이 없는 문서제목 placeholder가 산출물에 남으면 안 된다."""
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    screens = _screens(3)
    screens["meta"] = {"source": "s.xlsx", "template": "t.pptx"}  # 문서제목 없음
    build(screens, _layout_mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    assert "문서제목" not in [s.name for s in slide.shapes]


def test_build_layout_mode_fills_doc_title_from_meta(tmp_path: Path):
    tpl = _layout_template(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    screens = _screens(3)
    screens["meta"]["문서제목"] = "발행기관관리"
    build(screens, _layout_mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    by_name = {s.name: s for s in slide.shapes}
    assert by_name["문서제목"].text_frame.text == "발행기관관리"


def test_build_layout_mode_raises_for_unknown_layout(tmp_path: Path):
    import pytest
    tpl = _layout_template(tmp_path / "t.pptx")
    mapping = _layout_mapping(tpl)
    mapping["template"]["layout"] = "없는레이아웃"
    with pytest.raises(ValueError) as exc:
        build(_screens(3), mapping, tmp_path, tmp_path / "out.pptx", Warnings())
    assert "없는레이아웃" in str(exc.value)


def test_build_clone_mode_still_works(tmp_path: Path):
    """layout 모드를 더해도 clone 경로는 그대로여야 한다."""
    tpl = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    report = build(_screens(6), _mapping(tpl), tmp_path, out, Warnings())
    assert report["slides"] == 1
```

- [ ] **Step 2: 테스트가 실패하는 것을 확인**

Run: `python -m pytest tests/test_build.py --override-ini="addopts=" -v -k layout`
Expected: FAIL — layout 모드가 `ValueError("template.source_slide가 없습니다...")`로 멈춘다

- [ ] **Step 3: build.py 수정**

import에 다음을 더한다.

```python
from slide_layout import (
    DEFAULT_CONTENT_AREA,
    add_detail_table,
    add_image_anchor,
    drop_empty_placeholders,
    find_layout,
    inherit_placeholders,
    name_placeholders,
    split_content_area,
)
```

새 함수를 `_fill_page` 앞에 넣는다.

```python
def _new_layout_slide(prs, layout, tpl, warns, screen_id):
    """레이아웃으로 슬라이드를 만들고, 레이아웃에 없는 자리만 채워 넣는다.

    placeholder에 mapping이 정한 이름을 붙이므로, 값을 채우는 _fill_page는
    clone 모드와 똑같은 코드를 쓴다.
    """
    shapes_cfg = tpl.get("shapes", {})
    tables_cfg = tpl.get("detail_tables", {}) or {}
    count = int(tables_cfg.get("count", 5))
    rows = int(tables_cfg.get("rows", 4))
    area = tuple(tpl.get("content_area") or DEFAULT_CONTENT_AREA)

    slide = prs.slides.add_slide(layout)
    inherit_placeholders(slide, layout)
    name_placeholders(slide, tpl.get("placeholders", {}), shapes_cfg,
                      warns, screen_id)

    image_box, table_boxes = split_content_area(area, count, rows)
    add_image_anchor(slide, image_box, shapes_cfg.get("image", "화면이미지"))

    names = shapes_cfg.get("detail_tables") or []
    for i, box in enumerate(table_boxes):
        name = names[i] if i < len(names) else "상세표%d" % (i + 1)
        add_detail_table(slide, box, rows, name)
    return slide
```

`build` 함수의 소스 슬라이드 결정 부분(현재 `source_slide is None`이면 `ValueError`를 던지는 블록)을 모드 분기로 바꾼다.

```python
def build(screens_data, mapping, work_dir, out_path, warns):
    tpl = mapping["template"]
    template_path = resolve_template_path(tpl, work_dir)

    prs = Presentation(str(template_path))
    mode = tpl.get("mode", "clone")
    originals = list(prs.slides)

    if mode == "layout":
        layout = find_layout(prs, tpl.get("layout", 0))
        tables_cfg = tpl.get("detail_tables", {}) or {}
        slot_count = int(tables_cfg.get("count", 5)) * int(tables_cfg.get("rows", 4))
        src = None
        # content_area가 슬라이드를 벗어나면 도형이 잘려 나간다. 화면마다
        # 같은 실패를 반복하기 전에 여기서 한 번 막는다.
        area = tuple(tpl.get("content_area") or DEFAULT_CONTENT_AREA)
        if area[1] < 0 or area[1] + area[3] > int(prs.slide_height):
            raise ValueError(
                "content_area가 슬라이드 높이를 벗어납니다: top=%d height=%d, "
                "슬라이드 높이=%d" % (area[1], area[3], int(prs.slide_height))
            )
    else:
        source_slide = tpl.get("source_slide", 0)
        if source_slide is None:
            raise ValueError(
                "template.source_slide가 없습니다. mode를 'layout'으로 두고 "
                "template.layout에 레이아웃 이름을 지정하거나, 예시 슬라이드가 "
                "있는 템플릿을 --template으로 지정하세요."
            )
        src = prs.slides[int(source_slide)]
        slot_count = count_slots(
            collect_tables(src, tpl.get("shapes", {}).get("detail_tables"), warns)
        )
    ...
```

화면 루프 안에서 슬라이드를 만드는 줄을 바꾼다.

```python
            for i, page in enumerate(pages):
                if mode == "layout":
                    slide = _new_layout_slide(prs, layout, tpl, warns, scr_id)
                else:
                    slide = clone_slide(prs, src)
                _fill_page(slide, scr, page,
                           page_title(scr["name"], i, len(pages)),
                           mapping, work_dir, warns, screens_data.get("meta"))
                if mode == "layout":
                    drop_empty_placeholders(slide)
                made += 1
```

**주의:** `drop_empty_placeholders`는 `_fill_page` **뒤에** 불러야 한다. 먼저 부르면 채우기도 전에 전부 지워진다.

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_build.py --override-ini="addopts=" -v`
Expected: 새 테스트 8개와 기존 clone 테스트 전부 PASS

- [ ] **Step 5: 전체 스위트**

Run: `python -m pytest --override-ini="addopts=" -q`

- [ ] **Step 6: 커밋**

```bash
git add skills/excel-wireframe/scripts/build.py tests/test_build.py
git commit -m "feat: build.py에 layout 모드 추가"
```

---

### Task 4: default_template.py를 레이아웃 기반으로 재작성

**Files:**
- Modify: `skills/excel-wireframe/scripts/default_template.py` (전면 재작성)
- Modify: `tests/test_default_template.py`
- Modify: `tests/fixtures.py:85-110` (`make_template_pptx`)

**Interfaces:**
- Consumes: `slide_layout.split_content_area`, `add_image_anchor`, `add_detail_table`, `find_layout`, `inherit_placeholders`, `name_placeholders`
- Produces:
  - `build_default_template(path, slide_width_emu=9906000, slide_height_emu=6858000) -> Path` — **`table_count`·`rows_per_table` 파라미터가 사라진다.** 표는 이제 템플릿이 아니라 빌드 시점에 만들어지므로 템플릿이 알 필요가 없다
  - `default_template_mapping(path, table_count=5, rows_per_table=4) -> dict` — `mode: "layout"`을 반환
  - `DEFAULT_LAYOUT_NAME = "화면"` — 기본 템플릿이 만드는 레이아웃 이름
  - `DEFAULT_SHAPE_NAMES` — 기존 5키에 `page_no` 추가
  - `PLACEHOLDER_IDX = {"title": 0, "screen_id": 1, "작성일": 10, "문서제목": 11, "쪽번호": 12}`

**핵심 설계 결정:** python-pptx 기본 템플릿의 `Title and Content` 레이아웃이 가진 placeholder 다섯 개(TITLE/OBJECT/DATE/FOOTER/SLIDE_NUMBER)를 실측 자리로 옮겨 쓴다. XML로 placeholder를 새로 만들 필요가 없다. 껍데기 도형은 임시 슬라이드에 그려 XML을 레이아웃으로 이식한 뒤 임시 슬라이드를 지운다 — `LayoutShapes`에는 `add_shape`가 없기 때문이다.

**남는 레이아웃에 대해:** python-pptx 기본 템플릿의 나머지 레이아웃 열 개는 그대로 남는다. `find_layout`이 이름으로 찾으므로 동작에는 영향이 없고, 지우려면 `sldLayoutIdLst`를 손대야 해서 얻는 것보다 위험이 크다. 사용자가 PowerPoint의 레이아웃 목록에서 `화면` 외의 것도 보게 된다는 점만 `mapping-schema.md`에 적어 둔다(Task 7).

- [ ] **Step 1: 기존 테스트 중 무엇이 깨지는지 목록을 만든다**

Run: `python -m pytest tests/test_default_template.py --override-ini="addopts=" -v 2>&1 | tail -40`

지금은 전부 통과한다. 아래 테스트들은 **삭제하거나 새 구조에 맞게 고쳐야 한다** — 표와 슬라이드가 템플릿에서 사라지기 때문이다.

| 테스트 | 조치 |
|---|---|
| `test_default_template_tables_have_20_slots` | **삭제** — 템플릿에 표가 없다 |
| `test_default_template_tables_do_not_overlap` | **삭제** |
| `test_default_template_table_height_scales_with_rows_per_table` | **삭제** — `split_content_area` 테스트가 대신한다(Task 2) |
| `test_default_template_raises_when_rows_per_table_too_large` | **삭제** — 같은 검증이 `split_content_area`로 옮겨 갔다 |
| `test_default_template_allows_max_workable_rows_per_table` | **삭제** |
| `test_default_template_image_slot_does_not_overlap_tables` | **삭제** |
| `test_default_template_tables_match_measurements` | **삭제** |
| `test_default_template_cell_formatting_matches` | **삭제** — `add_detail_table` 테스트가 대신한다 |
| `test_default_template_shell_shapes_are_on_the_slide` | **고침** — 껍데기가 레이아웃에 있는지 본다 |
| `test_default_template_content_shapes_match_measurements` | **고침** — placeholder 좌표를 본다 |
| `test_default_template_has_named_shapes` / `..._has_meta_shapes` / `..._meta_shapes_are_empty` / `..._meta_shapes_stay_inside_slide` / `..._shell_text_starts_empty` / `..._id_background_sits_behind_id_text` | **고침** — 검사 대상을 슬라이드에서 레이아웃으로 옮긴다 |
| `test_default_template_uses_measured_size` / `..._matches_measured_slide_size` / `..._size_is_configurable` | **유지** |
| `test_default_template_carries_no_third_party_copyright` | **고침** — 레이아웃까지 훑는다 |
| `test_default_template_mapping_matches_shapes` / `..._mapping_includes_meta_shapes` | **고침** — `mode: layout` 구조를 본다 |
| `test_default_template_falls_back_to_computed_layout` | **삭제** — 표가 없으니 폴백 계산도 없다 |

- [ ] **Step 2: 새 테스트를 작성한다**

`tests/test_default_template.py`를 아래 내용으로 만든다(위 표에서 삭제로 표시한 것은 넣지 않는다).

```python
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
```

- [ ] **Step 3: 테스트가 실패하는 것을 확인**

Run: `python -m pytest tests/test_default_template.py --override-ini="addopts=" -v`
Expected: FAIL — `ImportError: cannot import name 'DEFAULT_LAYOUT_NAME'`

- [ ] **Step 4: default_template.py 재작성**

```python
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
        geom = MEASURED_PLACEHOLDERS.get(ph.placeholder_format.idx)
        if geom is None:
            continue
        ph.left, ph.top, ph.width, ph.height = (
            Emu(v) for v in _scaled(geom, sx, sy)
        )
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
```

**주의:** `layout.shapes._spTree.append(ph._element)`는 이미 트리에 있는 요소를 옮기는 동작이다(lxml에서 append는 이동이다). 복사가 아니므로 placeholder가 중복되지 않는다.

- [ ] **Step 5: fixtures.py의 make_template_pptx 재배치**

기본 템플릿에 예시 슬라이드가 없어졌으므로, clone 모드 테스트용 픽스처는 layout 경로로 한 장을 만들어 쓴다.

```python
def make_template_pptx(path: Path, table_count: int = 5, rows_per_table: int = 4) -> Path:
    """실제 샘플과 같은 도형 이름·슬라이드 크기를 가진 *예시 슬라이드형* 템플릿.

    기본 템플릿은 슬라이드가 0장이므로, 여기서 layout 모드의 슬라이드 생성
    경로로 한 장을 만들어 예시가 채워진 실제 템플릿을 흉내 낸다 —
    clone 판정(suggest_mode)과 텍스트 교체 테스트들이 그런 파일을 가정한다.
    """
    from default_template import (
        DEFAULT_LAYOUT_NAME,
        DEFAULT_SHAPE_NAMES,
        build_default_template,
        default_template_mapping,
    )
    from slide_layout import (
        add_detail_table,
        add_image_anchor,
        find_layout,
        inherit_placeholders,
        name_placeholders,
        split_content_area,
        DEFAULT_CONTENT_AREA,
    )
    from common import Warnings

    build_default_template(path)
    prs = Presentation(str(path))
    layout = find_layout(prs, DEFAULT_LAYOUT_NAME)
    tpl = default_template_mapping(path, table_count, rows_per_table)

    slide = prs.slides.add_slide(layout)
    inherit_placeholders(slide, layout)
    name_placeholders(slide, tpl["placeholders"], tpl["shapes"],
                      Warnings(), "fixture")
    image_box, table_boxes = split_content_area(
        DEFAULT_CONTENT_AREA, table_count, rows_per_table)
    add_image_anchor(slide, image_box, DEFAULT_SHAPE_NAMES["image"])
    for i, box in enumerate(table_boxes):
        add_detail_table(slide, box, rows_per_table, "상세표%d" % (i + 1))

    # 실제 샘플의 도형 이름과 예시 텍스트를 흉내 낸다
    rename = {"제목": "제목 13", "화면ID": "텍스트 개체 틀 14", "화면이미지": "그림 18"}
    example_text = {"제목": "화면명", "화면ID": "SCR000", "작성일": "2024-01-01"}
    for shp in slide.shapes:
        if shp.name in example_text and shp.has_text_frame:
            shp.text_frame.text = example_text[shp.name]
        if shp.name in rename:
            shp.name = rename[shp.name]
        elif shp.name.startswith("상세표"):
            shp.name = "표 %d" % (6 + int(shp.name[len("상세표"):]))

    # 상세 표에 예시 문구를 넣는다 — clone 모드는 이것을 덮어써야 한다
    for shp in slide.shapes:
        if shp.has_table:
            for r in range(len(shp.table.rows)):
                shp.table.cell(r, 1).text_frame.text = "예시 설명"

    prs.save(str(path))
    return path
```

- [ ] **Step 6: 기본 템플릿 테스트 통과 확인**

Run: `python -m pytest tests/test_default_template.py --override-ini="addopts=" -v`
Expected: 전부 PASS

- [ ] **Step 7: 전체 스위트를 돌려 픽스처 여파를 잡는다**

Run: `python -m pytest --override-ini="addopts=" -q 2>&1 | tail -40`

`make_template_pptx`의 산출물이 바뀌므로 `test_pptx_scan.py`, `test_slide_fill_*.py`,
`test_build.py`, `test_slide_clone.py`, `test_verify.py`, `test_analyze.py`가 영향을
받을 수 있다. 특히 다음을 확인한다.

- 도형 개수를 세는 단언 — 껍데기가 슬라이드에서 레이아웃으로 갔으므로 줄어든다
- `suggest_mode`가 여전히 `clone`으로 판정하는지 — 텍스트 도형 3개 + 표/그림 1개 이상이 필요하다. 픽스처는 제목·화면ID·작성일 텍스트와 표 5개를 가지므로 조건을 만족한다
- 표 이름(`표 7`~`표 11`)이 기존 테스트의 기대와 맞는지

깨진 것은 새 구조 기준으로 고친다. **기대값을 옛 구조로 되돌려 통과시키지 않는다.**

- [ ] **Step 8: 커밋**

```bash
git add skills/excel-wireframe/scripts/default_template.py tests/test_default_template.py tests/fixtures.py tests/
git commit -m "feat: 기본 템플릿을 레이아웃 기반으로 재작성"
```

---

### Task 5: analyze.py의 content_area 추천

**Files:**
- Modify: `skills/excel-wireframe/scripts/pptx_scan.py`
- Modify: `skills/excel-wireframe/scripts/analyze.py`
- Modify: `tests/test_pptx_scan.py`
- Modify: `tests/test_analyze.py`

**Interfaces:**
- Consumes: `pptx_scan.scan_presentation`
- Produces:
  - `pptx_scan.scan_layouts(path) -> list[dict]` — 마스터의 모든 레이아웃과 각 placeholder idx·좌표, 일반 도형 좌표
  - `pptx_scan.suggest_content_area(layout_info, slide_width, slide_height) -> list[int]` — placeholder와 껍데기 도형이 비운 가장 큰 가로 띠를 `[left, top, width, height]`로 반환

**핵심 설계 결정:** 완벽한 빈 영역 탐색은 과하다. 껍데기는 대부분 상단과 하단에 가로로 깔리므로, **슬라이드를 가로 띠로 보고 위아래에서 잠식된 만큼을 깎는** 방식이면 충분하다. 추정이 빗나가면 사람이 mapping.json에서 고친다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_pptx_scan.py`에 덧붙인다.

```python
from pptx_scan import scan_layouts, suggest_content_area


def test_scan_layouts_lists_placeholders(tmp_path: Path):
    from default_template import DEFAULT_LAYOUT_NAME, build_default_template
    path = build_default_template(tmp_path / "d.pptx")
    layouts = scan_layouts(path)
    names = [lay["name"] for lay in layouts]
    assert DEFAULT_LAYOUT_NAME in names
    target = next(lay for lay in layouts if lay["name"] == DEFAULT_LAYOUT_NAME)
    idxs = [ph["idx"] for ph in target["placeholders"]]
    assert {0, 1, 10, 11, 12} <= set(idxs)
    assert any(s["name"] == "상단띠" for s in target["shapes"])


def test_suggest_content_area_avoids_header_and_footer(tmp_path: Path):
    from default_template import DEFAULT_LAYOUT_NAME, build_default_template
    path = build_default_template(tmp_path / "d.pptx")
    layout_info = next(lay for lay in scan_layouts(path)
                       if lay["name"] == DEFAULT_LAYOUT_NAME)

    area = suggest_content_area(layout_info, 9906000, 6858000)
    left, top, width, height = area

    # 구분선(top 404664 + height 216024 = 620688) 아래에서 시작한다
    assert top >= 620688
    # 하단바(top 6716266) 위에서 끝난다
    assert top + height <= 6716266
    assert width > 0 and height > 0
```

`tests/test_analyze.py`에 덧붙인다.

```python
def test_analyze_suggests_layout_mode_mapping(tmp_path: Path):
    from analyze import main
    from common import read_json
    from fixtures import make_sheet_per_screen_xlsx

    xlsx = make_sheet_per_screen_xlsx(
        tmp_path / "in.xlsx",
        [{"id": "SCR001", "name": "목록", "image": False,
          "details": [{"no": "1", "type": "버튼", "element": "[등록]",
                       "desc": "등록한다", "pos": "우상단"}]}],
    )
    out = tmp_path / "report.json"
    assert main(["--excel", str(xlsx), "--out", str(out)]) == 0

    report = read_json(out)
    assert report["template_generated"] is True
    tpl = report["suggested_template_mapping"]
    assert tpl["mode"] == "layout"
    assert "content_area" in tpl
    assert len(tpl["content_area"]) == 4
```

- [ ] **Step 2: 테스트가 실패하는 것을 확인**

Run: `python -m pytest tests/test_pptx_scan.py tests/test_analyze.py --override-ini="addopts=" -v`
Expected: FAIL — `ImportError: cannot import name 'scan_layouts'`

- [ ] **Step 3: pptx_scan.py에 구현 추가**

```python
def scan_layouts(path: Path) -> list[dict]:
    """모든 마스터의 레이아웃과 그 안의 자리·도형을 훑는다.

    어느 레이아웃을 쓸지, 본문 영역을 어디로 잡을지 사람이 판단할 재료다.
    """
    prs = Presentation(str(path))
    out = []
    for mi, master in enumerate(prs.slide_masters):
        for li, lay in enumerate(master.slide_layouts):
            out.append({
                "master": mi,
                "index": li,
                "name": lay.name,
                "placeholders": [
                    {
                        "idx": ph.placeholder_format.idx,
                        "type": str(ph.placeholder_format.type),
                        "left": int(ph.left or 0),
                        "top": int(ph.top or 0),
                        "width": int(ph.width or 0),
                        "height": int(ph.height or 0),
                    }
                    for ph in lay.placeholders
                ],
                "shapes": [
                    _scan_shape(s) for s in lay.shapes if not s.is_placeholder
                ],
            })
    return out


def suggest_content_area(layout_info: dict, slide_width: int,
                         slide_height: int) -> list[int]:
    """레이아웃에서 이미지와 상세표를 놓을 만한 가로 띠를 고른다.

    껍데기는 대개 상단과 하단에 가로로 깔린다. 그래서 완전한 빈 영역 탐색 대신
    '슬라이드 폭의 절반 이상을 덮는 도형'만 장애물로 보고, 위아래에서 잠식된
    만큼 깎는다. 추정이 빗나가도 mapping.json에서 고칠 수 있으므로 이 정도면
    충분하다.
    """
    half = slide_width // 2
    blockers = []
    for s in layout_info["shapes"] + layout_info["placeholders"]:
        w = s.get("width", 0)
        if w >= half:
            blockers.append((s.get("top", 0), s.get("top", 0) + s.get("height", 0)))

    top = 0
    bottom = slide_height
    for b_top, b_bottom in blockers:
        if b_bottom <= slide_height // 2:
            top = max(top, b_bottom)      # 상단 껍데기
        elif b_top >= slide_height // 2:
            bottom = min(bottom, b_top)   # 하단 껍데기

    return [0, int(top), int(slide_width), int(bottom - top)]
```

- [ ] **Step 4: analyze.py를 고쳐 추천값을 매핑에 담는다**

```python
from pptx_scan import scan_layouts, scan_presentation, suggest_content_area, suggest_mode


def build_report(excel_path: Path, template_path: Path) -> dict:
    excel = scan_workbook(excel_path)
    template = scan_presentation(template_path)
    return {
        "excel": excel,
        "template": template,
        "layouts": scan_layouts(template_path),
        "suggestion": suggest_mode(template),
    }
```

`main`의 제안 매핑 부분을 고친다.

```python
    if generated:
        from default_template import DEFAULT_LAYOUT_NAME
        tpl_mapping = default_template_mapping(template_path)
        layout_info = next(
            (lay for lay in report["layouts"] if lay["name"] == DEFAULT_LAYOUT_NAME),
            None,
        )
        if layout_info is not None:
            tpl_mapping["content_area"] = suggest_content_area(
                layout_info,
                report["template"]["slide_width"],
                report["template"]["slide_height"],
            )
        report["suggested_template_mapping"] = tpl_mapping
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `python -m pytest tests/test_pptx_scan.py tests/test_analyze.py --override-ini="addopts=" -v`
Expected: 전부 PASS

- [ ] **Step 6: 전체 스위트**

Run: `python -m pytest --override-ini="addopts=" -q`

- [ ] **Step 7: 커밋**

```bash
git add skills/excel-wireframe/scripts/pptx_scan.py skills/excel-wireframe/scripts/analyze.py tests/test_pptx_scan.py tests/test_analyze.py
git commit -m "feat: 레이아웃 스캔과 content_area 추천"
```

---

### Task 6: screens.json 강등

**Files:**
- Modify: `skills/excel-wireframe/scripts/extract.py:1-6, 20-56, 100-114`
- Modify: `tests/test_extract.py`

**Interfaces:**
- Produces: `extract.main(argv) -> int` — `screens.json`을 항상 덮어쓴다. `diff_screens`와 `_screen_key`는 제거된다

- [ ] **Step 1: 실패하는 테스트로 바꾼다**

`tests/test_extract.py`에서 다음 테스트를 **삭제**한다.

- `test_extract_does_not_overwrite_existing_screens_json` (또는 `screens.new.json`을 기대하는 이름의 테스트 — 62~63행 근처)
- `test_diff_screens_reports_changes`
- `test_diff_screens_empty_when_same`
- `test_diff_screens_reports_deletion`

import 줄에서 `_screen_key`, `diff_screens`를 뺀다.

```python
from extract import main
```

그 자리에 덮어쓰기 테스트를 넣는다.

```python
def test_extract_overwrites_existing_screens_json(tmp_path: Path):
    """screens.json은 중간 산출물이다. 재추출이 그냥 덮어쓴다."""
    xlsx = make_sheet_per_screen_xlsx(
        tmp_path / "in.xlsx",
        [{"id": "SCR001", "name": "이용기관 목록", "image": False,
          "details": [{"no": "1", "type": "버튼", "element": "[등록]",
                       "desc": "등록한다", "pos": "우상단"}]}],
    )
    work = tmp_path / "work"
    work.mkdir()
    mapping_path = work / "mapping.json"
    write_json(mapping_path, _mapping())

    write_json(work / "screens.json", {"meta": {}, "screens": [
        {"id": "OLD", "name": "옛 화면", "images": [], "fields": {}, "details": []}
    ]})

    assert main(["--excel", str(xlsx), "--mapping", str(mapping_path),
                 "--work", str(work)]) == 0

    data = read_json(work / "screens.json")
    assert [s["id"] for s in data["screens"]] == ["SCR001"]
    assert not (work / "screens.new.json").exists()
```

**주의:** `_mapping()`과 import(`read_json`, `write_json`, `make_sheet_per_screen_xlsx`)는 이 파일에 이미 있는 것을 쓴다. 없으면 파일 상단의 기존 헬퍼 이름에 맞춰 조정한다.

- [ ] **Step 2: 테스트가 실패하는 것을 확인**

Run: `python -m pytest tests/test_extract.py --override-ini="addopts=" -v`
Expected: FAIL — `screens.json`이 덮어써지지 않아 `OLD`가 그대로 남는다

- [ ] **Step 3: extract.py 수정**

모듈 docstring을 고친다.

```python
"""2단계: mapping.json에 따라 Excel에서 screens.json과 이미지를 뽑는다.

screens.json은 중간 산출물이다. 매핑이나 템플릿만 고쳐 생성 단계를 반복할 때
추출을 다시 하지 않아도 되게 하고, 결과가 이상할 때 추출과 배치 중 어느 쪽이
틀렸는지 갈라 보게 한다. 사람이 손으로 편집하는 파일이 아니므로 재추출은
그냥 덮어쓴다 — 추출 결과가 어긋나면 이 파일이 아니라 매핑을 고친다.
"""
```

`_screen_key`와 `diff_screens` 함수를 통째로 지운다(20~56행).

`main`의 저장 부분을 고친다.

```python
    target = work / "screens.json"
    write_json(target, payload)
    print("screens.json 저장: %s" % target)
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_extract.py --override-ini="addopts=" -v`
Expected: 전부 PASS

- [ ] **Step 5: 전체 스위트**

Run: `python -m pytest --override-ini="addopts=" -q`

- [ ] **Step 6: 커밋**

```bash
git add skills/excel-wireframe/scripts/extract.py tests/test_extract.py
git commit -m "refactor: screens.json을 중간 산출물로 강등"
```

---

### Task 7: 문서 갱신, 재설치, 실제 샘플 검증

**Files:**
- Modify: `skills/excel-wireframe/SKILL.md`
- Modify: `skills/excel-wireframe/references/mapping-schema.md`
- Modify: `README.md`
- Modify: `CLAUDE.md`
- Test: `tests/test_skill_docs.py`

**Interfaces:** 없음 (문서와 검증)

- [ ] **Step 1: 낡은 설명을 찾는다**

```bash
grep -rn "16:9\|13.33\|12192000\|SSOT\|screens.new\|source_slide\|clone" \
  skills/excel-wireframe/ README.md CLAUDE.md | grep -v "^docs/"
```

- [ ] **Step 2: SKILL.md 갱신**

세 곳을 고친다.

1. **원칙 절**(13~14행) — SSOT 문구를 바꾼다.

```markdown
`screens.json`은 Excel에서 뽑은 중간 산출물이다. 매핑이나 템플릿만 고쳐 생성 단계를
반복할 때 추출을 다시 하지 않아도 되고, 결과가 이상할 때 추출과 배치 중 어디가
틀렸는지 갈라 볼 수 있다. 재추출은 그냥 덮어쓴다 — 추출이 어긋나면 이 파일이 아니라
매핑을 고친다.
```

2. **1단계**(38~43행) — 기본 템플릿 설명을 바꾼다.

```markdown
**템플릿을 안 받았으면 `--template`을 생략한다.** 기본 템플릿이
`work/default-template.pptx`로 만들어진다 — 슬라이드 10.83 × 7.50in에 화면 페이지용
레이아웃(`화면`) 하나가 들어 있고, 상단 띠·구분선·하단 바 같은 껍데기는 그 레이아웃
위에, 글자가 들어가는 자리(제목·화면ID·작성일·문서제목·쪽번호)는 placeholder로 있다.
예시 슬라이드는 넣지 않는다 — 생성 단계가 레이아웃으로 슬라이드를 만들기 때문이다.
그에 맞는 매핑이 리포트의 `suggested_template_mapping`에 담긴다. 템플릿이 있는지
사용자에게 먼저 묻지 말고, 없으면 기본 템플릿으로 진행한 뒤 결과를 보여주며 "원하는
템플릿이 있으면 주시면 그대로 맞춰 드립니다"라고 알린다.
```

3. **3단계**(70~75행) — diff 확인 절차를 지우고 한 줄로 바꾼다.

```markdown
`work/screens.json`과 `work/images/`가 생긴다. 재추출은 둘 다 덮어쓴다.
```

4. **실패했을 때** 표의 마지막 줄(106행, "템플릿에 예시 슬라이드가 없다고 나옴")을 바꾼다.

```markdown
| 레이아웃을 찾지 못했다고 나옴 | `template.layout` 이름이 템플릿과 다르다. `structure-report.json`의 `layouts` 목록에서 실제 레이아웃 이름을 확인한다 |
```

5. **하지 않는 것** 절에 한 줄 더한다.

```markdown
- 기본 템플릿에 제3자 저작권 표기를 넣지 않는다
```

- [ ] **Step 3: references/mapping-schema.md 갱신**

`template` 절에 layout 모드를 더한다. 담아야 할 것:

- `mode`가 `clone`이면 `source_slide`, `layout`이면 `layout`·`placeholders`·`content_area`·`detail_tables`를 쓴다는 것
- `placeholders`는 **idx로** 지정한다는 것과 그 이유(이름은 슬라이드마다 달라진다)
- `shapes`는 채운 자리에 붙일 이름이며, `verify.py`가 그 이름으로 검증한다는 것
- `content_area`는 `[left, top, width, height]` EMU이고 `analyze.py`가 추천값을 계산한다는 것
- 표는 `content_area` 아래쪽에 폭을 균등 분할해 만들어지고, 행높이는 실측값을 쓴다는 것
- 기본 템플릿 절을 새 구조로 고친다 — 레이아웃 `화면` 하나, 슬라이드 0장, placeholder 다섯 개, 껍데기 도형 다섯 개

layout 모드 예시를 넣는다.

```json
"template": {
  "file": "work/default-template.pptx",
  "mode": "layout",
  "layout": "화면",
  "placeholders": { "title": 0, "screen_id": 1, "작성일": 10, "문서제목": 11, "쪽번호": 12 },
  "shapes": {
    "title": "제목",
    "screen_id": "화면ID",
    "image": "화면이미지",
    "작성일": "작성일",
    "문서제목": "문서제목",
    "쪽번호": "쪽번호",
    "detail_tables": ["상세표1", "상세표2", "상세표3", "상세표4", "상세표5"]
  },
  "content_area": [-12319, 337940, 9957099, 6331421],
  "detail_tables": { "count": 5, "rows": 4 },
  "table_columns": { "no": 0, "text": 1 }
}
```

`tests/test_skill_docs.py`가 `상세표1`과 `detail_tables`가 이 파일에 있는지 검사하므로, 위 예시를 그대로 두면 통과한다.

- [ ] **Step 4: README.md 갱신**

- 첫 줄의 오타 `ㄴ# excel-wireframe`을 `# excel-wireframe`으로 고친다
- "현재 한계"에서 기본 템플릿이 16:9라 실제와 다르다는 줄이 남아 있으면 지운다
- 사용자 템플릿의 레이아웃 요소는 여전히 건드리지 않는다는 줄은 유지하되, 기본 템플릿은 껍데기가 레이아웃에 있어 PowerPoint에서 레이아웃만 고치면 전 슬라이드에 반영된다는 것을 덧붙인다

- [ ] **Step 5: CLAUDE.md 갱신**

세 줄을 고친다.

1. SSOT 규칙:

```markdown
- **`screens.json`은 중간 산출물이다.** Excel은 임포트 소스, `screens.json`은 캐시다.
  `extract.py`는 재추출 때 그냥 덮어쓴다 — 추출이 어긋나면 매핑을 고친다.
  `build.py`는 openpyxl을 import하지 않는다 — 이 제약은 `tests/test_build.py`가
  소스를 검사해 강제한다.
```

2. `extract.py`가 덮어쓰지 않는다는 규칙을 지운다.

3. 픽스처 규칙:

```markdown
- 테스트 템플릿 픽스처는 `default_template.build_default_template`으로 템플릿을 만든 뒤
  `slide_layout`의 생성 경로로 예시 슬라이드를 한 장 얹는다. 픽스처용 pptx를 따로
  만들지 않는다.
```

4. 슬라이드 생성 규칙에 한 줄 더한다:

```markdown
- 레이아웃으로 슬라이드를 만들 때는 `slide_layout`의 함수만 쓴다. placeholder를
  직접 복제하면 date/footer/쪽번호가 빠지거나 빈 자리가 산출물에 남는다.
```

- [ ] **Step 6: 문서 테스트와 전체 스위트**

Run: `python -m pytest --override-ini="addopts=" -q`
Expected: 전부 PASS

- [ ] **Step 7: 스킬 재설치**

```bash
rm -rf /c/Users/user/.claude/skills/excel-wireframe
cp -r skills/excel-wireframe /c/Users/user/.claude/skills/
find /c/Users/user/.claude/skills/excel-wireframe -name __pycache__ -type d -exec rm -rf {} +
diff -r --strip-trailing-cr --exclude=__pycache__ skills/excel-wireframe /c/Users/user/.claude/skills/excel-wireframe && echo "설치 일치"
```

- [ ] **Step 8: 실제 샘플로 끝까지 확인**

```bash
export PYTHONIOENCODING=utf-8
S=~/.claude/skills/excel-wireframe/scripts
rm -rf work/layout && mkdir -p work/layout
python "$S/analyze.py" --excel "짧은 버전.xlsx" --out work/layout/structure-report.json
```

리포트에서 확인할 것:

- `template_generated`가 `true`
- `suggested_template_mapping.mode`가 `"layout"`
- `suggested_template_mapping.content_area`가 4개 값
- `template.slide_size_in`이 `10.83 x 7.50`

그다음 매핑을 조립해(리포트의 `suggested_template_mapping`을 `template`에 그대로 복사하고 `excel` 섹션만 판단) 추출과 생성을 돌린다.

```bash
python "$S/extract.py" --excel "짧은 버전.xlsx" --mapping work/layout/mapping.json --work work/layout
python "$S/build.py" --screens work/layout/screens.json --mapping work/layout/mapping.json --work work/layout --out work/layout/output/화면설계서.pptx
```

확인할 것:

- 검증 6항목이 전부 통과
- 슬라이드 크기가 9906000 × 6858000
- 슬라이드마다 제목·화면ID가 채워졌고 빈 placeholder가 남지 않았다
- 상세표 5개가 있고 번호가 Excel 값 그대로다

```bash
python - <<'PY'
import sys; sys.path.insert(0, "skills/excel-wireframe/scripts")
from pptx import Presentation
p = Presentation("work/layout/output/화면설계서.pptx")
print("슬라이드 %d장, 크기 %d x %d" % (len(p.slides._sldIdLst), p.slide_width, p.slide_height))
s = list(p.slides)[0]
print("layout:", s.slide_layout.name)
for shp in s.shapes:
    kind = "표" if shp.has_table else ("그림" if "PICTURE" in str(shp.shape_type) else "도형")
    text = shp.text_frame.text[:20] if shp.has_text_frame else ""
    print("   %-10s %-12s %s" % (kind, shp.name, text))
PY
```

- [ ] **Step 9: 커밋**

```bash
git add skills/excel-wireframe/ README.md CLAUDE.md tests/
git commit -m "docs: 레이아웃 기반 생성과 screens.json 강등을 문서에 반영"
```

---

## 검증 매트릭스

| 스펙 요구 | 커버 위치 |
|---|---|
| 레이아웃을 이름·인덱스로 찾는다 | Task 1 Step 1 |
| 마스터가 여러 개여도 훑는다 | Task 1 Step 3 |
| 없는 레이아웃은 `ValueError` | Task 1 Step 1, Task 3 Step 1 |
| date/footer/slidenumber 보충 상속 | Task 1 Step 1 |
| 채운 placeholder에 이름을 붙인다 | Task 1 Step 1 |
| 없는 idx는 `shape-not-found` 경고 | Task 1 Step 1 |
| 빈 placeholder 제거, 필드는 유지 | Task 1 Step 1 |
| `content_area`를 이미지·표로 분할 | Task 2 Step 1 |
| 표 폭 균등 분할, 행높이 실측 | Task 2 Step 1 |
| 이미지 자리 최소 높이 검증 | Task 2 Step 1 |
| 셀 서식 실측 일치 | Task 2 Step 1 |
| 표 셀이 비어 시작한다 | Task 2 Step 1 |
| layout 모드 빌드 | Task 3 Step 1 |
| 상세 분할·이미지 배치·meta 채우기 | Task 3 Step 1 |
| `content_area`가 슬라이드를 벗어나면 `ValueError` | Task 3 Step 3 |
| clone 모드 회귀 | Task 3 Step 1, Task 4 Step 7 |
| 기본 템플릿: 레이아웃 1개 + 슬라이드 0장 | Task 4 Step 2 |
| 기본 템플릿: placeholder 좌표 실측 일치 | Task 4 Step 2 |
| 기본 템플릿: 껍데기가 레이아웃에 | Task 4 Step 2 |
| 기본 템플릿: 자리표시 문구 없음 | Task 4 Step 2 |
| 제3자 저작권 부재 | Task 4 Step 2 |
| 제안 매핑이 그대로 빌드된다 | Task 4 Step 2 |
| 픽스처 재배치 | Task 4 Step 5 |
| 레이아웃 스캔과 `content_area` 추천 | Task 5 Step 1 |
| `screens.json` 덮어쓰기 | Task 6 Step 1 |
| 문서가 코드와 일치 | Task 7 |
| 실제 샘플 파이프라인 통과 | Task 7 Step 8 |

## 범위 밖

- 사용자가 준 템플릿의 레이아웃을 편집하지 않는다
- 표지·목차를 만들지 않는다
- 슬라이드 크기를 바꾸지 않는다
- 레이아웃에 표 placeholder를 심는 방식은 채택하지 않는다
- 경고 코드를 새로 만들지 않는다
