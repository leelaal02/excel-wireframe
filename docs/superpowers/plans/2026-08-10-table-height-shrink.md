# 표를 내용에 꼭 맞게 줄이기 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 표 높이를 장별로 잡고 줄간격을 좁혀, 표가 먹던 공간을 스크린샷에 돌려준다.

**Architecture:** 행 높이 계산을 화면 단위 최댓값에서 장별 최댓값으로 바꾼다(표 간 통일은 유지). 줄간격 계수와 셀 서식을 상수 하나로 묶어 계산과 산출물이 어긋나지 않게 한다. 2패스 구조는 그대로라 배분을 재계산하지 않으므로 순환은 생기지 않는다.

**Tech Stack:** Python 3.13, python-pptx, pytest

## Global Constraints

- 설계 문서: `docs/superpowers/specs/2026-08-10-table-height-shrink-design.md`
- 경고 코드는 아홉 개 그대로다. 늘리지도 줄이지도 않는다 — `tests/test_warning_codes.py`가 강제한다.
- `build.py`는 openpyxl을 import하지 않는다 — `tests/test_build.py`가 소스를 검사해 강제한다.
- 표 셀에 값을 쓸 때는 `slide_fill.set_cell_text`를 쓴다.
- 상세 표 번호는 Excel 값을 그대로 쓴다. 재부여·재정렬 금지.
- 슬라이드 크기와 `content_area`는 건드리지 않는다.
- 번호 칸 서식(6.5pt, 여백 18000)은 그대로 둔다.
- 폰트는 6pt 밑으로 내리지 않는다. `DETAIL_FONT_STEPS = (7.0, 6.5, 6.0)`.
- 전체 테스트: `python -m pytest -q` (현재 226개 통과)
- 스크립트는 `skills/excel-wireframe/scripts/`에 평면 배치. 모듈 간 import는 `from common import ...` 형식.

**변경 후 기대값** (실제 샘플로 프로토타입 검증 완료):

| | 현행 | 목표 |
|---|---|---|
| 긴 버전 표 | 3.31in 균일 | 2.35 / 2.21 / 1.62 / 2.35 / 2.00 / 2.79 |
| 긴 버전 1p 스크린샷 | 3.62in | 4.58in |
| 짧은 버전 표 | 3.19in | 3.01in |

---

## File Structure

| 파일 | 책임 | 변경 |
|---|---|---|
| `skills/excel-wireframe/scripts/text_metrics.py` | 줄 수·행 높이 계산. 줄간격·여백 상수의 단일 출처 | 상수 재편 |
| `skills/excel-wireframe/scripts/slide_layout.py` | 표를 그리고 서식을 입힌다 | 상수를 받아 셀 서식에 반영 |
| `skills/excel-wireframe/scripts/slide_fill.py` | 도형·셀에 값을 채운다 | 문단 속성 상속 |
| `skills/excel-wireframe/scripts/build.py` | 화면을 장으로 나누고 높이를 정한다 | 장별 계산 |

`text_metrics`가 상수의 출처이고 `slide_layout`이 그것을 가져다 쓴다. 반대 방향 의존은 없다.

---

### Task 1: 줄간격과 여백을 상수 하나로 모으고 셀 서식에 반영한다

계산 계수만 낮추면 렌더링 줄높이는 그대로라 행이 스스로 자란다 — 직전 작업이 고친 넘침이 되돌아온다. 그래서 상수 재편과 셀 서식 반영을 한 태스크로 묶는다.

**Files:**
- Modify: `skills/excel-wireframe/scripts/text_metrics.py:18-20`
- Modify: `skills/excel-wireframe/scripts/slide_layout.py` (import부, `TEXT_CELL_MARGIN`, `_format_cell`, `add_detail_table`)
- Test: `tests/test_text_metrics.py`, `tests/test_slide_layout.py`

**Interfaces:**
- Produces: `text_metrics.BASE_LINE_SPACING = 1.2`, `text_metrics.LINE_SPACING_RATIO = 0.95`, `text_metrics.LINE_SPACING = BASE_LINE_SPACING * LINE_SPACING_RATIO`, `text_metrics.CELL_MARGIN = 4762`, `text_metrics.DEFAULT_MARGIN_TOP = CELL_MARGIN`
- Produces: `_format_cell(cell, size_pt, bold, align, anchor, margin, margin_bottom, font=None, line_spacing=None)` — 인자 하나 추가, 기본값 `None`이라 기존 호출은 그대로 동작

- [ ] **Step 1: 실패하는 테스트를 쓴다**

`tests/test_text_metrics.py` 끝에 추가:

```python
def test_line_spacing_derives_from_ratio():
    """계수는 유도값이다. 셀 서식에 들어가는 배수와 따로 놀면 안 된다."""
    from text_metrics import BASE_LINE_SPACING, LINE_SPACING, LINE_SPACING_RATIO
    assert LINE_SPACING == BASE_LINE_SPACING * LINE_SPACING_RATIO
    assert LINE_SPACING < BASE_LINE_SPACING


def test_default_margin_top_is_the_shared_cell_margin():
    """행 높이 계산과 셀 서식이 같은 여백 값을 봐야 표가 자리에 맞는다."""
    from text_metrics import CELL_MARGIN, DEFAULT_MARGIN_TOP
    assert DEFAULT_MARGIN_TOP == CELL_MARGIN
```

`tests/test_slide_layout.py` 끝에 추가:

```python
def test_add_detail_table_sets_line_spacing_on_text_cells(tmp_path: Path):
    """계산에 쓰는 줄간격이 셀 서식에도 들어가야 렌더링과 계산이 맞는다."""
    from text_metrics import LINE_SPACING_RATIO
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    slide = prs.slides[0]
    table = add_detail_table(slide, (0, 0, 1971135, 1416117), 4, "상세표1").table
    para = table.cell(0, 1).text_frame.paragraphs[0]
    assert para.line_spacing == LINE_SPACING_RATIO


def test_add_detail_table_text_margin_comes_from_the_shared_constant(tmp_path: Path):
    """여백이 세 곳에 흩어져 있으면 계산과 산출물이 어긋난다."""
    from text_metrics import CELL_MARGIN
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    slide = prs.slides[0]
    table = add_detail_table(slide, (0, 0, 1971135, 1416117), 4, "상세표1").table
    txt = table.cell(0, 1)
    assert txt.margin_left == CELL_MARGIN
    assert txt.margin_top == CELL_MARGIN
```

- [ ] **Step 2: 테스트가 실패하는 것을 확인한다**

Run: `python -m pytest tests/test_text_metrics.py::test_line_spacing_derives_from_ratio tests/test_slide_layout.py::test_add_detail_table_sets_line_spacing_on_text_cells -v`

Expected: FAIL — `ImportError: cannot import name 'BASE_LINE_SPACING'`, `LINE_SPACING_RATIO`

- [ ] **Step 3: text_metrics의 상수를 재편한다**

`skills/excel-wireframe/scripts/text_metrics.py`의 18~20행을 바꾼다:

```python
EMU_PER_PT = 12700
BASE_LINE_SPACING = 1.2      # 맑은 고딕이 기본으로 잡는 줄높이 (폰트 크기 대비)
LINE_SPACING_RATIO = 0.95    # 셀 문단에 실제로 넣는 줄간격 배수
# 계산에 쓰는 줄높이. slide_layout이 LINE_SPACING_RATIO를 셀 서식에 그대로
# 넣으므로 이 값과 산출물이 일치한다 — 한쪽만 바꾸면 표가 자리에 안 맞는다.
LINE_SPACING = BASE_LINE_SPACING * LINE_SPACING_RATIO
CELL_MARGIN = 4762           # 설명 칸 여백. slide_layout이 셀 서식에 그대로 쓴다
DEFAULT_MARGIN_TOP = CELL_MARGIN
```

- [ ] **Step 4: slide_layout이 그 상수를 쓰게 한다**

import부(13행 `from common import EMU_PER_INCH` 다음 줄)에 추가:

```python
from text_metrics import CELL_MARGIN, LINE_SPACING_RATIO
```

`TEXT_CELL_MARGIN` 정의(144행)를 바꾼다:

```python
TEXT_CELL_MARGIN = CELL_MARGIN       # 설명 칸 여백. 계산과 서식이 같은 값을 쓴다
```

`_format_cell`(306행)에 인자를 더하고 줄간격을 넣는다:

```python
def _format_cell(cell, size_pt, bold, align, anchor, margin, margin_bottom,
                 font=None, line_spacing=None):
    cell.margin_left = Emu(margin)
    cell.margin_right = Emu(margin)
    cell.margin_top = Emu(margin)
    cell.margin_bottom = Emu(margin_bottom)
    if anchor is not None:
        cell.vertical_anchor = anchor
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    # 줄간격을 명시하면 text_metrics의 LINE_SPACING과 렌더링이 맞는다.
    # 안 넣으면 폰트 기본 줄높이가 쓰여 계산보다 행이 커진다.
    if line_spacing is not None:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = ""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if font:
        run.font.name = font
```

`add_detail_table`의 설명 칸 호출(352~353행)에서 하드코딩된 `9525`를 없앤다:

```python
        _format_cell(table.cell(r, 1), size_pt, False, PP_ALIGN.LEFT, None,
                     TEXT_CELL_MARGIN, 0, font="맑은 고딕",
                     line_spacing=LINE_SPACING_RATIO)
```

- [ ] **Step 5: 여백을 하드코딩한 기존 테스트를 상수로 바꾼다**

`tests/test_slide_layout.py:221`의 `assert txt.margin_left == 9525`를 찾아 바꾼다:

```python
    assert txt.margin_left == CELL_MARGIN
```

그 파일 상단 import에 `from text_metrics import CELL_MARGIN`을 더한다. 같은 파일 289행 근처에 `margin_bottom == 0` 검사가 있는데 그건 그대로 둔다.

`tests/test_text_metrics.py:50-58`의 두 테스트에서 하드코딩된 `1.2`를 상수로 바꾼다:

```python
def test_row_height_grows_with_lines():
    from text_metrics import LINE_SPACING
    one = row_height(1, 7.0)
    two = row_height(2, 7.0)
    assert two - one == int(7.0 * LINE_SPACING * EMU_PER_PT)


def test_row_height_includes_margins():
    from text_metrics import LINE_SPACING
    assert row_height(1, 7.0, margin_top=9525, margin_bottom=0) == \
        int(7.0 * LINE_SPACING * EMU_PER_PT) + 9525
```

- [ ] **Step 6: 테스트를 돌려 통과를 확인한다**

Run: `python -m pytest tests/test_text_metrics.py tests/test_slide_layout.py -q`
Expected: 전부 PASS

- [ ] **Step 7: 전체 스위트를 돌린다**

Run: `python -m pytest -q`

Expected: 전부 PASS. `test_fits_lines_counts_what_a_row_holds`가 기대하는 `[3, 2, 4, 2]`는 새 계수에서도 그대로 나온다 — 검산: 줄높이 `int(7 × 1.14 × 12700) = 101346`, `(382457 − 4762) // 101346 = 3`, `(268746 − 4762) // 101346 = 2`, `(496168 − 4762) // 101346 = 4`.

깨지는 테스트가 있으면 멈추고 원인을 본다. 이 태스크는 계산 계수를 바꾸므로 표 높이를 단언하는 테스트가 걸릴 수 있다.

- [ ] **Step 8: 커밋**

```bash
git add skills/excel-wireframe/scripts/text_metrics.py skills/excel-wireframe/scripts/slide_layout.py tests/test_text_metrics.py tests/test_slide_layout.py
git commit -m "refactor: 줄간격과 여백을 상수 하나로 모으고 셀 서식에 넣는다"
```

---

### Task 2: 이어붙인 문단에도 문단 속성을 물려준다

`_fill_text_frame`은 런 속성만 복사하고 문단 속성은 복사하지 않는다. Task 1이 넣은 줄간격이 줄바꿈 뒤 문단에서 사라져 계산과 어긋난다.

**Files:**
- Modify: `skills/excel-wireframe/scripts/slide_fill.py:41-66`
- Test: `tests/test_slide_fill_text.py`

**Interfaces:**
- Consumes: Task 1의 `line_spacing` 셀 서식
- Produces: 동작 변경만. 시그니처는 그대로다

- [ ] **Step 1: 실패하는 테스트를 쓴다**

`tests/test_slide_fill_text.py` 끝에 추가:

```python
def test_set_cell_text_keeps_paragraph_spacing_on_extra_lines(tmp_path: Path):
    """줄바꿈으로 늘어난 문단도 첫 문단의 줄간격을 물려받아야 한다.

    안 물려받으면 둘째 줄부터 폰트 기본 줄높이가 쓰여, 계산한 행 높이보다
    실제 렌더링이 커진다 — 표가 슬라이드를 넘는 원인이다.
    """
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    table = next(s for s in prs.slides[0].shapes if s.has_table).table
    cell = table.cell(0, 1)
    cell.text_frame.paragraphs[0].line_spacing = 0.95

    set_cell_text(cell, "첫 줄\n둘째 줄\n셋째 줄")

    paras = cell.text_frame.paragraphs
    assert len(paras) == 3
    assert all(p.line_spacing == 0.95 for p in paras)
```

- [ ] **Step 2: 테스트가 실패하는 것을 확인한다**

Run: `python -m pytest tests/test_slide_fill_text.py::test_set_cell_text_keeps_paragraph_spacing_on_extra_lines -v`

Expected: FAIL — 둘째·셋째 문단의 `line_spacing`이 `None`이다

- [ ] **Step 3: 문단 속성을 복사한다**

`skills/excel-wireframe/scripts/slide_fill.py`의 `_fill_text_frame`에서, `base_rPr`을 만든 다음 줄에 문단 속성 복사를 더하고, 새 문단을 만들 때 그것을 넣는다:

```python
    # 문단 속성(줄간격·정렬)도 물려줘야 한다. 런 속성만 옮기면 둘째 줄부터
    # 줄간격이 폰트 기본값으로 돌아가 행이 계산보다 커진다.
    pPr = p0._p.find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}pPr"
    )
    base_pPr = copy.deepcopy(pPr) if pPr is not None else None

    base_run.text = lines[0]
    for extra in p0.runs[1:]:
        extra._r.getparent().remove(extra._r)
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)

    for line in lines[1:]:
        p = tf.add_paragraph()
        if base_pPr is not None:
            p._p.insert(0, copy.deepcopy(base_pPr))
        run = p.add_run()
        run.text = line
        if base_rPr is not None:
            run._r.insert(0, copy.deepcopy(base_rPr))
```

`pPr`은 `a:p`의 첫 자식이어야 하므로 `insert(0, ...)`으로 넣는다.

- [ ] **Step 4: 테스트 통과를 확인한다**

Run: `python -m pytest tests/test_slide_fill_text.py -q`
Expected: 전부 PASS. 기존 `test_set_cell_text_splits_newlines`도 그대로 통과해야 한다

- [ ] **Step 5: 전체 스위트**

Run: `python -m pytest -q`
Expected: 전부 PASS

- [ ] **Step 6: 커밋**

```bash
git add skills/excel-wireframe/scripts/slide_fill.py tests/test_slide_fill_text.py
git commit -m "fix: 줄바꿈으로 늘어난 문단에도 줄간격을 물려준다"
```

---

### Task 3: 표 높이를 장별로 잡는다

**Files:**
- Modify: `skills/excel-wireframe/scripts/build.py` (`_fit_tables`, 그 호출부)
- Test: `tests/test_build.py`

**Interfaces:**
- Consumes: `text_metrics.plan_row_heights(pages, rows_per_table, width_emu, size_pt, floors)` — 시그니처 변경 없음. 페이지를 하나만 담은 목록을 넘겨 그 장만의 최댓값을 얻는다
- Produces: `_fit_tables(pages, area, count, rows, text_key) -> (list[list[int]], float, bool)` — 첫 값이 **장별** 행 높이 목록으로 바뀐다. 기존에는 행 높이 하나였다

- [ ] **Step 1: 실패하는 테스트를 쓴다**

`tests/test_build.py`의 `test_build_layout_mode_uses_one_table_height_across_pages`(676행)를 **통째로 지우고** 그 자리에 아래 셋을 넣는다:

```python
def test_build_layout_mode_gives_each_page_its_own_height(tmp_path: Path):
    """장마다 자기 내용에 맞는 높이를 쓴다.

    전 장을 최악값으로 통일하면 긴 상세 하나가 모든 장의 표를 밀어 올려
    스크린샷을 눌러 버린다.
    """
    descs = ["짧게"] * 40
    descs[25] = LONG_DESC          # 2장에만 긴 항목이 있다
    prs = _build_layout(tmp_path, _screens_with_desc(descs))

    assert len(prs.slides) == 2
    first = _tables_of(prs.slides[0])[0]
    second = _tables_of(prs.slides[1])[0]
    # 1장은 짧은 상세뿐이라 실측 하한 그대로다
    assert [r.height for r in first.table.rows] == [382457, 268746, 496168, 268746]
    # 2장만 긴 항목 때문에 자란다
    assert second.height > first.height
    assert second.top < first.top


def test_build_layout_mode_bottom_aligns_every_page(tmp_path: Path):
    """장별 높이가 달라도 아래 끝은 모든 장에서 본문 영역 하단이다."""
    area_bottom = 337940 + 6331421
    descs = ["짧게"] * 40
    descs[25] = LONG_DESC
    prs = _build_layout(tmp_path, _screens_with_desc(descs))

    for slide in prs.slides:
        for tbl in _tables_of(slide):
            assert tbl.top + tbl.height == area_bottom


def test_build_layout_mode_uses_one_font_size_across_pages(tmp_path: Path):
    """높이는 장별이지만 글자 크기는 화면 안에서 통일한다.

    한 장만 작으면 장을 넘길 때 글자가 커졌다 작아졌다 한다.
    """
    prs = _build_layout(tmp_path, _screens_with_desc([LONG_DESC] * 40))

    assert len(prs.slides) == 2
    sizes = set()
    for slide in prs.slides:
        cell = _tables_of(slide)[0].table.cell(0, 1)
        sizes.add(cell.text_frame.paragraphs[0].runs[0].font.size)
    assert len(sizes) == 1
```

- [ ] **Step 2: 테스트가 실패하는 것을 확인한다**

Run: `python -m pytest tests/test_build.py::test_build_layout_mode_gives_each_page_its_own_height -v`

Expected: FAIL — 두 장의 행 높이가 같아서 `second.height > first.height`가 깨진다

- [ ] **Step 3: `_fit_tables`를 장별 계산으로 바꾼다**

`skills/excel-wireframe/scripts/build.py`의 `_fit_tables`를 통째로 바꾼다:

```python
def _fit_tables(pages, area, count: int, rows: int, text_key: str):
    """화면의 장별 행 높이와, 화면 전체가 공유할 설명 칸 글자 크기를 정한다.

    2패스의 두 번째다. 배분(pages)은 이미 확정됐고, 여기서는 그 배분에 맞는
    표 높이만 구한다 — 배분을 다시 돌리면 이미지 자리가 바뀌어 조각 수가 바뀌고,
    조각 수가 바뀌면 배분이 또 바뀌어 순환에 빠진다.

    높이는 장마다 자기 내용에 맞춰 잡는다. 전 장을 최악값으로 통일하면 긴 상세
    하나가 모든 장의 표를 밀어 올려 이미지 자리를 통째로 잡아먹는다. 배분을
    재계산하지 않으므로 장마다 높이가 달라도 순환은 생기지 않고, 이미지는
    place_image가 비율을 지켜 축소 배치한다.

    글자 크기만은 화면 단위로 통일한다 — 한 장만 작으면 장을 넘길 때 글자가
    커졌다 작아졌다 한다.

    (장별 행 높이 목록, 글자 크기, 낮췄는지)를 돌려준다.
    """
    floors = measured_row_heights(rows)
    width = detail_text_width(area[2], count)
    limit = area[3] - MIN_IMAGE_HEIGHT_EMU
    texts = [[str(d.get(text_key, "") or "") for d in page] for page, _ in pages]

    for size_pt in DETAIL_FONT_STEPS:
        per_page = [plan_row_heights([t], rows, width, size_pt, floors)
                    for t in texts]
        if all(sum(h) <= limit for h in per_page):
            return per_page, size_pt, size_pt != DETAIL_FONT_STEPS[0]
    return ([_cap_heights(h, floors, limit) for h in per_page],
            DETAIL_FONT_STEPS[-1], True)
```

- [ ] **Step 4: 호출부가 장별 높이를 넘기게 한다**

같은 파일에서 `row_heights = size_pt = None`으로 시작하는 블록(377~391행 근처)을 바꾼다:

```python
            # 2패스: 배분이 확정됐으니 그 상세에 맞는 표 높이를 구한다.
            # 높이는 장별로, 글자 크기는 화면 단위로 정해진다.
            page_heights = size_pt = None
            if mode == "layout":
                page_heights, size_pt, shrunk = _fit_tables(
                    pages, area, count, rows, text_key)
                if shrunk:
                    print("  [%s] 상세가 길어 설명 글자를 %.1fpt로 낮췄습니다"
                          % (scr_id, size_pt))

            for i, (page, page_image) in enumerate(pages):
                if mode == "layout":
                    slide = _new_layout_slide(prs, layout, tpl, warns, scr_id,
                                              page_heights[i], size_pt)
```

`clone` 모드는 `page_heights`를 쓰지 않으므로 그대로 둔다.

- [ ] **Step 5: 테스트 통과를 확인한다**

Run: `python -m pytest tests/test_build.py -q`
Expected: 전부 PASS

- [ ] **Step 6: 전체 스위트**

Run: `python -m pytest -q`
Expected: 전부 PASS

- [ ] **Step 7: 커밋**

```bash
git add skills/excel-wireframe/scripts/build.py tests/test_build.py
git commit -m "feat: 표 높이를 장별로 잡아 스크린샷 자리를 돌려준다"
```

---

### Task 4: 실제 샘플로 확인하고 설치본을 갱신한다

**Files:**
- Test: 실제 샘플 파일 (`C:\Users\user\Desktop\화면설계서\긴 버전.xlsx`, `짧은 버전.xlsx`)
- Modify: `~/.claude/skills/excel-wireframe/scripts/` (설치본)

**Interfaces:**
- Consumes: Task 1~3의 모든 변경

- [ ] **Step 1: 두 샘플을 저장소 코드로 빌드한다**

작업 폴더에 `screens.json`이 없으면 먼저 추출한다:

```bash
cd "/c/Users/user/Desktop/화면설계서"
PYTHONIOENCODING=utf-8 python "/c/Users/user/Desktop/wireframe/skills/excel-wireframe/scripts/extract.py" \
  --excel "긴 버전.xlsx" --mapping work-long/mapping.json --work work-long
PYTHONIOENCODING=utf-8 python "/c/Users/user/Desktop/wireframe/skills/excel-wireframe/scripts/extract.py" \
  --excel "짧은 버전.xlsx" --mapping work-short/mapping.json --work work-short
```

빌드한다:

```bash
cd "/c/Users/user/Desktop/화면설계서"
PYTHONIOENCODING=utf-8 python "/c/Users/user/Desktop/wireframe/skills/excel-wireframe/scripts/build.py" \
  --screens work-long/screens.json --mapping work-long/mapping.json \
  --work work-long --out "work-long/output/화면설계서_긴버전.pptx"
PYTHONIOENCODING=utf-8 python "/c/Users/user/Desktop/wireframe/skills/excel-wireframe/scripts/build.py" \
  --screens work-short/screens.json --mapping work-short/mapping.json \
  --work work-short --out "work-short/output/화면설계서_짧은버전.pptx"
```

Expected: 검증 6항목 통과. 경고는 긴 버전 `slide-split` 1건뿐, `text-overflow` 0건

- [ ] **Step 2: 표 높이를 잰다**

Bash에서 한글 경로를 인수로 넘기면 깨지므로 스크립트 파일로 만들어 실행한다. 스크래치패드에 `check.py`를 쓴다:

```python
# -*- coding: utf-8 -*-
import glob
import os
import sys
from pathlib import Path

sys.path.insert(0, r"C:\Users\user\Desktop\wireframe\skills\excel-wireframe\scripts")
sys.stdout.reconfigure(encoding="utf-8")
from pptx import Presentation

CA_BOT = 337940 + 6331421
desk = Path(os.path.expanduser("~")) / "Desktop" / "화면설계서"
for path in sorted(glob.glob(str(desk / "work-*" / "output" / "*.pptx"))):
    prs = Presentation(path)
    print(os.path.basename(path))
    for si, s in enumerate(prs.slides):
        tabs = [sh for sh in s.shapes if sh.has_table]
        imgs = [sh for sh in s.shapes if sh.shape_type == 13]
        bot = max(t.top + t.height for t in tabs)
        print("  s%d 표 %.2fin | 하단차 %+d | 이미지 %.2fin"
              % (si, tabs[0].height / 914400, bot - CA_BOT,
                 imgs[0].height / 914400 if imgs else 0))
```

Run: `python <스크래치패드>/check.py`

Expected: 긴 버전 표가 `2.35 / 2.21 / 1.62 / 2.35 / 2.00 / 2.79`in, 짧은 버전 `3.01`in. 하단차는 전부 `+0`. 1페이지 이미지 `4.58`in

값이 다르면 멈추고 원인을 본다 — 설계 문서의 "검증된 결과"와 어긋난다는 뜻이다.

- [ ] **Step 3: 설치본을 갱신한다**

```bash
cd "/c/Users/user/Desktop/wireframe"
cp -r skills/excel-wireframe "/c/Users/user/.claude/skills/"
diff -rq "/c/Users/user/.claude/skills/excel-wireframe/scripts/" skills/excel-wireframe/scripts/ | grep -v __pycache__
```

Expected: diff 출력 없음. `user-default.json`과 `assets/`는 저장소에 없으므로 지워지지 않는다 — `ls ~/.claude/skills/excel-wireframe/`로 둘 다 남아 있는지 확인한다

- [ ] **Step 4: 커밋**

이 태스크는 저장소 파일을 바꾸지 않으므로 커밋할 것이 없다. `git status`로 확인만 한다.

```bash
cd "/c/Users/user/Desktop/wireframe" && git status --short
```

Expected: `skills/excel-wireframe/scripts/image_split.py`의 미커밋 주석 변경만 남아 있다 (의도된 상태 — 사용자가 남겨 두기로 했다)

---

## 검증 매트릭스

| 설계 요구 | 태스크 | 확인 방법 |
|---|---|---|
| 표 높이를 장별로 | Task 3 | `test_build_layout_mode_gives_each_page_its_own_height` |
| 표 간(가로) 통일 유지 | Task 3 | 기존 `test_build_layout_mode_aligns_row_heights_across_tables` |
| 줄간격 0.95배 | Task 1 | `test_add_detail_table_sets_line_spacing_on_text_cells` |
| 계산과 서식이 같은 상수 | Task 1 | `test_line_spacing_derives_from_ratio`, `test_add_detail_table_text_margin_comes_from_the_shared_constant` |
| 여백 4762 | Task 1 | `test_default_margin_top_is_the_shared_cell_margin` |
| 문단 속성 상속 | Task 2 | `test_set_cell_text_keeps_paragraph_spacing_on_extra_lines` |
| 폰트는 화면 단위 통일 | Task 3 | `test_build_layout_mode_uses_one_font_size_across_pages` |
| 하단 고정 | Task 3 | `test_build_layout_mode_bottom_aligns_every_page` |
| 하한 유지(회귀 방지) | Task 1, 3 | 기존 `test_build_layout_mode_short_details_keep_measured_heights` |
| 경고 코드 아홉 개 | 전체 | 기존 `test_warning_codes.py` |
| 실제 샘플 수치 | Task 4 | 표 2.35/2.21/1.62/2.35/2.00/2.79in, 넘침 0 |

## 범위 밖

- 표 간(가로) 행 높이 통일 해제
- 상세 번호 재부여·재정렬
- `content_area`·슬라이드 크기 변경
- 번호 칸 서식(6.5pt, 여백 18000)
- 폰트 6pt 미만
- `image_split.py`의 미커밋 주석 — 손대지 않는다
