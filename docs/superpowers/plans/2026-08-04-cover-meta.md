# Excel 표지 메타 반영 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Excel 표지 시트의 프로젝트명·작성일 같은 문서 단위 정보를 읽어 `screens.json`의 `meta`에 담고, PPT의 동명 도형에 채운다.

**Architecture:** 새 모듈 `xlsx_meta.py`가 표지에서 라벨-값 쌍을 추출한다. `extract.py`가 이를 `meta`에 담고(현재 하드코딩된 `"title": "화면설계서"` 제거), `build.py`가 `template.shapes`의 키와 같은 이름의 `meta` 값을 도형에 채운다. 화면별 `fields`가 문서 `meta`를 이긴다. `build.py`는 여전히 `screens.json`만 읽으므로 openpyxl 비의존 제약이 유지된다.

**Tech Stack:** Python 3.13, openpyxl 3.1.5, python-pptx 1.0.2, pytest

## 범위

설계 스펙 `docs/superpowers/specs/2026-08-04-default-template-sample-parity-design.md`는 두 가지를 담고 있다. **이 계획은 그중 (A) 표지 메타 반영만 구현한다.**

- **구현함** — 5절(표지 메타 추출), 6절(`build.py`의 meta 채우기), 7절 중 기본 템플릿에 meta 도형 2개 추가
- **구현하지 않음** — 4절의 기본 템플릿 실측 기하 재작성(슬라이드 크기 9906000, 표 좌표·열폭·행높이, 껍데기 요소). 기본 템플릿은 현재의 16:9 비율 계산 방식을 유지한다.

(B)를 빼면 채울 자리가 없어지므로, **기본 템플릿에 `문서제목`·`작성일` 도형 두 개만 추가**한다. 현재 기하 안에서 배치하고, 좌표는 실측값이 아니라 현재 레이아웃에 맞춰 계산한다.

## Global Constraints

- 프로젝트 루트: `C:\Users\user\Desktop\wireframe`. 스킬 본체는 `skills/excel-to-wireframe-ppt/`.
- 스크립트는 `skills/excel-to-wireframe-ppt/scripts/`에 평면 배치. 패키지를 만들지 않는다. import는 `from common import ...` 형태.
- **`build.py`는 openpyxl을 import하지 않는다.** `tests/test_build.py`가 소스를 검사해 강제한다.
- **경고 코드는 아홉 개뿐이다:** `no-image`, `no-detail`, `text-overflow`, `shape-not-found`, `slide-split`, `slot-shortage`, `screen-failed`, `image-convert-failed`, `orphan-row`. 표지를 못 찾은 상황에 맞는 코드가 없으므로 **새 코드를 만들지 않고** `extract.py`의 출력 한 줄로 알린다. `tests/test_warning_codes.py`가 이를 강제한다.
- 모든 CLI 진입점은 `setup_stdio()`를 먼저 호출한다.
- 화면 단위로 예외를 격리한다. 표지 파싱이 실패해도 화면 페이지 생성은 정상 진행된다.
- `meta`의 `source`·`template`은 예약 키다. 표지에 같은 이름의 라벨이 있어도 덮어쓰지 않는다.
- **우선순위: 화면별 `fields` > 문서 `meta`.**
- `tests/fixtures.py`의 기존 생성기는 여러 테스트가 의존한다. 확장은 하되 형태를 바꾸지 않는다.
- 테스트는 `pytest -q` 한 줄로 전부 돌아야 한다. `pytest.ini`가 `addopts = -q`를 설정하므로 명령줄 `-v`가 무시된다 — 개별 PASSED 줄이 필요하면 `--override-ini="addopts="`.
- 스킬 파일을 수정하면 `C:\Users\user\.claude\skills\excel-to-wireframe-ppt\`에 재설치하고 `__pycache__`를 제외한다.

## 실측 데이터

두 샘플 Excel의 표지 시트 구조는 동일하다.

| 셀 | 내용 |
|---|---|
| `B3` | `화면설계서 & 단위테스트 케이스 리스트` (단독 셀) |
| `B4` | `멀티분석 기반 (Xplatform + Web HTML + 스크린샷)` (단독 셀) |
| `C7` / `E7` | `프로젝트명` / `누리미디어 통합관리시스템(BIS)` |
| `C8` / `E8` | `전환 내용` / `Xplatform → JSP 기반 Web 전환` |
| `C9` / `E9` | `분석 모드` / `멀티모달 분석 (코드 + 스크린샷)` |
| `C10` / `E10` | `분석 화면 수` / `1개` |
| `C11` / `E11` | `화면설계서 요소` / `16건` (긴 버전은 `93건`) |
| `C12` / `E12` | `단위테스트 항목` / `17건` (긴 버전은 `56건`) |
| `C13` / `E13` | `작성일` / `2026-06-11` (긴 버전은 `2026-06-25`) |
| `C14` / `E14` | `작성 도구` / `LLM 기반 자동 생성 (Claude API 멀티모달)` |
| `C15` / `E15` | `상태` / `1차 생성 - 검토 필요` |
| `B18` | `* 본 문서는 LLM(Claude API) 멀티모달 분석으로 자동 생성되…` (단독 셀, 무시 대상) |

라벨은 `C:D` 병합, 값은 `E:G` 병합이다. 라벨 없이 단독으로 선 셀은 `B3`, `B4`, `B18` 세 개다.

## 파일 구조

| 파일 | 책임 |
|---|---|
| `scripts/xlsx_meta.py` | 신규. 표지 시트에서 라벨-값 쌍과 단독 셀을 추출 |
| `scripts/extract.py` | `meta` 하드코딩 제거, `xlsx_meta` 호출, 표지 미발견 시 안내 출력 |
| `scripts/build.py` | `template.shapes` 키에 대해 `fields` → `meta` 순으로 도형 채우기 |
| `scripts/default_template.py` | `문서제목`·`작성일` 도형 추가 |
| `scripts/analyze.py` | `suggested_template_mapping.shapes`에 두 키 포함 |
| `references/mapping-schema.md` | `cover`, `meta_overrides`, `meta` 문서화 |
| `SKILL.md` | 표지 값이 반영된다는 사실 추가 |

---

### Task 1: 표지 메타 추출

**Files:**
- Create: `skills/excel-to-wireframe-ppt/scripts/xlsx_meta.py`
- Test: `tests/test_xlsx_meta.py`

**Interfaces:**
- Consumes: `Warnings` (사용하지 않지만 시그니처 일관성을 위해 받는다)
- Produces:
  - `find_cover_sheet(wb, mapping) -> str | None` — 표지 시트 이름
  - `read_cover_meta(wb, mapping, warns) -> dict[str, str]`

`find_cover_sheet` 규칙: `mapping["excel"].get("cover", {}).get("sheet")`가 있으면 그것. 없으면 `sheet_include` 정규식에 **맞지 않는** 첫 시트. `sheet_include`가 없으면 첫 시트.

`read_cover_meta` 규칙:
1. 텍스트가 있는 셀에 대해 같은 행 오른쪽을 훑어 첫 비어 있지 않은 값을 찾는다 → `meta[라벨] = 값`
2. 오른쪽에 값이 없는 단독 셀 중 최상단을 `문서제목`, 그다음을 `부제`로 둔다. 세 번째부터는 무시한다
3. `mapping["excel"].get("meta_overrides")`가 있으면 그 값이 이긴다
4. 값은 문자열로 정규화한다. `datetime`/`date`는 `YYYY-MM-DD`

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_xlsx_meta.py`:

```python
import datetime
from pathlib import Path

from common import Warnings
from openpyxl import Workbook, load_workbook
from xlsx_meta import find_cover_sheet, read_cover_meta

MAPPING = {"excel": {"layout": "sheet-per-screen", "sheet_include": "^설계_"}}


def _cover_xlsx(path: Path, rows=None, title_cells=None) -> Path:
    """실제 샘플과 같은 배치의 표지: 라벨 C열, 값 E열, 단독 셀 B열."""
    wb = Workbook()
    ws = wb.active
    ws.title = "표지"
    for cell, text in (title_cells or {"B3": "화면설계서", "B4": "부제목입니다"}).items():
        ws[cell] = text
    r = 7
    for label, value in (rows or [("프로젝트명", "통합관리시스템"), ("작성일", "2026-06-11")]):
        ws.cell(row=r, column=3, value=label)
        ws.cell(row=r, column=5, value=value)
        r += 1
    wb.create_sheet("설계_SCR001")["A1"] = "화면설계서 - SCR001 (목록)"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def test_find_cover_sheet_skips_screen_sheets(tmp_path: Path):
    wb = load_workbook(_cover_xlsx(tmp_path / "s.xlsx"))
    assert find_cover_sheet(wb, MAPPING) == "표지"


def test_find_cover_sheet_honors_explicit_mapping(tmp_path: Path):
    wb = load_workbook(_cover_xlsx(tmp_path / "s.xlsx"))
    mapping = {"excel": {"sheet_include": "^설계_", "cover": {"sheet": "설계_SCR001"}}}
    assert find_cover_sheet(wb, mapping) == "설계_SCR001"


def test_read_cover_meta_pairs_label_and_value(tmp_path: Path):
    wb = load_workbook(_cover_xlsx(tmp_path / "s.xlsx"))
    meta = read_cover_meta(wb, MAPPING, Warnings())
    assert meta["프로젝트명"] == "통합관리시스템"
    assert meta["작성일"] == "2026-06-11"


def test_read_cover_meta_takes_standalone_cells_as_title_and_subtitle(tmp_path: Path):
    wb = load_workbook(_cover_xlsx(tmp_path / "s.xlsx"))
    meta = read_cover_meta(wb, MAPPING, Warnings())
    assert meta["문서제목"] == "화면설계서"
    assert meta["부제"] == "부제목입니다"


def test_read_cover_meta_ignores_third_standalone_cell(tmp_path: Path):
    xlsx = _cover_xlsx(
        tmp_path / "s.xlsx",
        title_cells={"B3": "제목", "B4": "부제", "B18": "* 주석 문구입니다"},
    )
    meta = read_cover_meta(load_workbook(xlsx), MAPPING, Warnings())
    assert meta["문서제목"] == "제목"
    assert meta["부제"] == "부제"
    assert "* 주석 문구입니다" not in meta.values()


def test_read_cover_meta_normalizes_dates(tmp_path: Path):
    xlsx = _cover_xlsx(tmp_path / "s.xlsx", rows=[("작성일", datetime.date(2026, 6, 11))])
    meta = read_cover_meta(load_workbook(xlsx), MAPPING, Warnings())
    assert meta["작성일"] == "2026-06-11"


def test_meta_overrides_win(tmp_path: Path):
    wb = load_workbook(_cover_xlsx(tmp_path / "s.xlsx"))
    mapping = {"excel": {"sheet_include": "^설계_",
                         "meta_overrides": {"프로젝트명": "손으로 지정한 값"}}}
    meta = read_cover_meta(wb, mapping, Warnings())
    assert meta["프로젝트명"] == "손으로 지정한 값"
    assert meta["작성일"] == "2026-06-11"


def test_read_cover_meta_returns_empty_when_no_cover(tmp_path: Path):
    wb = Workbook()
    wb.active.title = "설계_SCR001"
    wb.active["A1"] = "화면설계서 - SCR001 (목록)"
    p = tmp_path / "only-screens.xlsx"
    wb.save(p)
    loaded = load_workbook(p)
    assert find_cover_sheet(loaded, MAPPING) is None
    assert read_cover_meta(loaded, MAPPING, Warnings()) == {}
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_xlsx_meta.py --override-ini="addopts=" -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xlsx_meta'`

- [ ] **Step 3: xlsx_meta.py 구현**

```python
# -*- coding: utf-8 -*-
"""Excel 표지 시트에서 문서 단위 정보를 읽는다.

화면 페이지는 화면마다 값이 다르지만, 프로젝트명·작성일 같은 값은 문서 전체에
하나뿐이다. 표지는 그 값들이 모여 있는 유일한 자리다.
"""
from __future__ import annotations

import datetime
import re

RESERVED = {"source", "template"}
MAX_SCAN_ROW = 40
MAX_SCAN_COL = 12
MAX_STANDALONE = 2


def _text(ws, row: int, col: int) -> str:
    v = ws.cell(row=row, column=col).value
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def find_cover_sheet(wb, mapping: dict) -> str | None:
    """표지 시트를 고른다.

    명시 지정이 있으면 그것을 쓴다. 없으면 화면 시트로 쓰이지 않은 첫 시트다 —
    표지는 보통 맨 앞에 있고, 화면 시트는 sheet_include로 이미 식별된다.
    """
    cfg = mapping.get("excel", {})
    named = (cfg.get("cover") or {}).get("sheet")
    if named:
        return named if named in wb.sheetnames else None

    include = cfg.get("sheet_include")
    if not include:
        return wb.sheetnames[0] if wb.sheetnames else None
    rx = re.compile(include)
    for name in wb.sheetnames:
        if not rx.search(name):
            return name
    return None


def read_cover_meta(wb, mapping: dict, warns) -> dict[str, str]:
    """표지에서 라벨-값 쌍과 단독 셀을 뽑아 문서 정보 사전을 만든다."""
    sheet = find_cover_sheet(wb, mapping)
    if sheet is None:
        return {}
    ws = wb[sheet]

    meta: dict[str, str] = {}
    standalone: list[str] = []
    last_row = min(ws.max_row or 0, MAX_SCAN_ROW)
    last_col = min(ws.max_column or 0, MAX_SCAN_COL)

    for r in range(1, last_row + 1):
        c = 1
        while c <= last_col:
            label = _text(ws, r, c)
            if not label:
                c += 1
                continue
            # 같은 행 오른쪽에서 첫 값을 찾는다. 병합 셀은 두 번째 칸부터 비어
            # 있으므로 자연히 건너뛰어진다.
            value = ""
            vc = c + 1
            while vc <= last_col:
                value = _text(ws, r, vc)
                if value:
                    break
                vc += 1
            if value:
                if label not in RESERVED:
                    meta[label] = value
                c = vc + 1
            else:
                if len(standalone) < MAX_STANDALONE:
                    standalone.append(label)
                break
    if standalone:
        meta.setdefault("문서제목", standalone[0])
    if len(standalone) > 1:
        meta.setdefault("부제", standalone[1])

    overrides = mapping.get("excel", {}).get("meta_overrides") or {}
    for k, v in overrides.items():
        if k not in RESERVED:
            meta[k] = "" if v is None else str(v)
    return meta
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_xlsx_meta.py --override-ini="addopts=" -v`
Expected: PASS (8 passed)

- [ ] **Step 5: 실제 두 샘플로 손수 확인**

```bash
export PYTHONIOENCODING=utf-8
python - <<'PY'
import sys; sys.path.insert(0, "skills/excel-to-wireframe-ppt/scripts")
from openpyxl import load_workbook
from common import Warnings
from xlsx_meta import read_cover_meta
MAPPING = {"excel": {"sheet_include": "^설계_"}}
for f in ["짧은 버전.xlsx", "긴 버전.xlsx"]:
    wb = load_workbook(f, data_only=True)
    for k, v in read_cover_meta(wb, MAPPING, Warnings()).items():
        print(f, "|", k, "=", v)
    wb.close()
PY
```

Expected: 두 파일 모두 `문서제목`, `부제`, `프로젝트명`, `전환 내용`, `분석 모드`, `분석 화면 수`, `화면설계서 요소`, `단위테스트 항목`, `작성일`, `작성 도구`, `상태`가 나온다. `* 본 문서는 …` 주석은 나오지 않는다.

- [ ] **Step 6: 커밋**

```bash
git add tests/test_xlsx_meta.py skills/excel-to-wireframe-ppt/scripts/xlsx_meta.py
git commit -m "feat: Excel 표지에서 문서 단위 정보 추출"
```

---

### Task 2: extract.py 연결

**Files:**
- Modify: `skills/excel-to-wireframe-ppt/scripts/extract.py`
- Test: `tests/test_extract.py` (기존 파일에 추가)

**Interfaces:**
- Consumes: `read_cover_meta(wb, mapping, warns)`
- Produces: `screens.json`의 `meta`가 `{**cover_meta, "source": ..., "template": ...}` 형태가 된다

`"title": "화면설계서"` 하드코딩을 제거한다. 표지에 `문서제목`이 있으면 그것이 문서 제목이고, 없으면 `title` 키 자체가 없다. `source`·`template`은 예약 키로 언제나 마지막에 덮어쓴다.

표지를 못 찾으면 `meta`에 `source`·`template`만 남고, 출력에 안내 한 줄이 나온다. **새 경고 코드를 만들지 않는다.**

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_extract.py` 끝에 추가:

```python
def test_extract_puts_cover_meta_into_screens_json(tmp_path: Path):
    xlsx, work, mp = _setup(tmp_path)
    # _setup의 픽스처에는 표지 시트가 '표지' 이름으로 들어 있다
    from openpyxl import load_workbook
    wb = load_workbook(xlsx)
    ws = wb["표지"]
    ws["C7"] = "프로젝트명"
    ws["E7"] = "통합관리시스템"
    ws["C8"] = "작성일"
    ws["E8"] = "2026-06-11"
    wb.save(xlsx)

    assert main(["--excel", str(xlsx), "--mapping", str(mp), "--work", str(work)]) == 0
    meta = read_json(work / "screens.json")["meta"]
    assert meta["프로젝트명"] == "통합관리시스템"
    assert meta["작성일"] == "2026-06-11"
    assert meta["source"].endswith("s.xlsx")


def test_extract_no_longer_hardcodes_title(tmp_path: Path):
    xlsx, work, mp = _setup(tmp_path)
    main(["--excel", str(xlsx), "--mapping", str(mp), "--work", str(work)])
    meta = read_json(work / "screens.json")["meta"]
    assert meta.get("title") != "화면설계서"


def test_extract_reports_when_cover_missing(tmp_path: Path, capsys):
    xlsx, work, mp = _setup(tmp_path)
    from openpyxl import load_workbook
    wb = load_workbook(xlsx)
    del wb["표지"]
    wb.save(xlsx)

    assert main(["--excel", str(xlsx), "--mapping", str(mp), "--work", str(work)]) == 0
    out = capsys.readouterr().out
    assert "표지" in out
    meta = read_json(work / "screens.json")["meta"]
    assert set(meta.keys()) == {"source", "template"}


def test_extract_reserved_keys_survive_cover_labels(tmp_path: Path):
    xlsx, work, mp = _setup(tmp_path)
    from openpyxl import load_workbook
    wb = load_workbook(xlsx)
    ws = wb["표지"]
    ws["C7"] = "source"
    ws["E7"] = "표지가 지정한 엉뚱한 값"
    wb.save(xlsx)

    main(["--excel", str(xlsx), "--mapping", str(mp), "--work", str(work)])
    meta = read_json(work / "screens.json")["meta"]
    assert meta["source"].endswith("s.xlsx")
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_extract.py --override-ini="addopts=" -v`
Expected: FAIL — `KeyError: '프로젝트명'` (그리고 `title` 테스트는 하드코딩 때문에 실패)

- [ ] **Step 3: extract.py 수정**

import에 추가:

```python
from xlsx_meta import read_cover_meta
```

`main`에서 `read_screens` 호출부 근처를 다음으로 바꾼다:

```python
    wb = load_workbook(excel_path, data_only=True)
    try:
        screens = read_screens(wb, mapping, warns)
        cover_meta = read_cover_meta(wb, mapping, warns)
        extract_images(excel_path, wb, mapping, screens, work, warns)
    finally:
        wb.close()

    if not cover_meta:
        print("표지 시트를 찾지 못해 문서 정보를 비웠습니다 "
              "(mapping.excel.cover.sheet로 지정할 수 있습니다)")

    payload = {
        "meta": {
            **cover_meta,
            "source": str(excel_path),
            "template": mapping.get("template", {}).get("file", ""),
        },
        "screens": screens,
    }
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_extract.py --override-ini="addopts=" -v`
Expected: PASS

- [ ] **Step 5: 전체 스위트 확인**

Run: `python -m pytest --override-ini="addopts=" -q`
Expected: PASS. 기존 `test_extract_creates_screens_json`이 `meta["title"]`을 검사하고 있으면 그 단언을 새 동작에 맞게 고친다.

- [ ] **Step 6: 커밋**

```bash
git add tests/test_extract.py skills/excel-to-wireframe-ppt/scripts/extract.py
git commit -m "feat: extract.py가 표지 정보를 meta에 담는다"
```

---

### Task 3: build.py의 meta 채우기

**Files:**
- Modify: `skills/excel-to-wireframe-ppt/scripts/build.py`
- Test: `tests/test_build.py` (기존 파일에 추가)

**Interfaces:**
- Consumes: `screens_data["meta"]`
- Produces: `_fill_page`가 `template.shapes`의 각 키에 대해 `screen["fields"]` → `meta` 순으로 값을 찾아 동명 도형에 쓴다

`title`·`screen_id`·`image`·`detail_tables`는 이미 전용 처리가 있으므로 이 경로에서 제외한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_build.py` 끝에 추가:

```python
def _mapping_with_meta_shape(template: Path) -> dict:
    m = _mapping(template)
    m["template"]["shapes"]["프로젝트명"] = "텍스트 개체 틀 14"
    return m


def test_build_fills_meta_into_named_shape(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["meta"]["프로젝트명"] = "통합관리시스템"
    out = tmp_path / "out.pptx"
    build(data, _mapping_with_meta_shape(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    shp = next(s for s in slide.shapes if s.name == "텍스트 개체 틀 14")
    assert shp.text_frame.text == "통합관리시스템"


def test_build_screen_fields_beat_document_meta(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["meta"]["프로젝트명"] = "문서 전체 값"
    data["screens"][0]["fields"]["프로젝트명"] = "화면별 값"
    out = tmp_path / "out.pptx"
    build(data, _mapping_with_meta_shape(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    shp = next(s for s in slide.shapes if s.name == "텍스트 개체 틀 14")
    assert shp.text_frame.text == "화면별 값"


def test_build_ignores_meta_without_matching_shape(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["meta"]["아무도_안_쓰는_키"] = "값"
    out = tmp_path / "out.pptx"
    warns = Warnings()
    build(data, _mapping(tpl), tmp_path, out, warns)
    assert "shape-not-found" not in [w["code"] for w in warns.to_list()]


def test_build_meta_does_not_override_title_shape(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["meta"]["title"] = "문서 제목이 화면명을 덮으면 안 된다"
    out = tmp_path / "out.pptx"
    build(data, _mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    shp = next(s for s in slide.shapes if s.name == "제목 13")
    assert shp.text_frame.text == "이용기관 목록"
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_build.py --override-ini="addopts=" -v`
Expected: FAIL — `프로젝트명` 도형이 채워지지 않아 텍스트가 예시값 그대로다

- [ ] **Step 3: build.py 수정**

`build`가 `_fill_page`에 `meta`를 넘기도록 시그니처를 넓히고(`meta: dict | None = None`), `_fill_page`의 `fields` 처리 블록을 다음으로 교체한다:

```python
    # 화면별 fields가 문서 meta를 이긴다. 화면마다 다른 값이 있으면 그게 더 구체적이다.
    reserved = {"title", "screen_id", "image", "detail_tables"}
    doc_meta = meta or {}
    for key, name in shapes_cfg.items():
        if key in reserved:
            continue
        if key in (scr.get("fields") or {}):
            value = (scr.get("fields") or {})[key]
        elif key in doc_meta:
            value = doc_meta[key]
        else:
            continue
        shp = find_shape(slide, name)
        if shp is not None:
            set_text(shp, value)
```

`build` 안의 `_fill_page` 호출부에 `screens_data.get("meta")`를 넘긴다.

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_build.py --override-ini="addopts=" -v`
Expected: PASS

- [ ] **Step 5: openpyxl 비의존 재확인**

Run: `python -m pytest tests/test_build.py::test_build_does_not_import_openpyxl --override-ini="addopts=" -v`
Expected: PASS

- [ ] **Step 6: 커밋**

```bash
git add tests/test_build.py skills/excel-to-wireframe-ppt/scripts/build.py
git commit -m "feat: build.py가 문서 meta를 동명 도형에 채운다"
```

---

### Task 4: 기본 템플릿 도형 추가, analyze 연결, 문서 갱신

**Files:**
- Modify: `skills/excel-to-wireframe-ppt/scripts/default_template.py`
- Modify: `skills/excel-to-wireframe-ppt/scripts/analyze.py`
- Modify: `skills/excel-to-wireframe-ppt/references/mapping-schema.md`
- Modify: `skills/excel-to-wireframe-ppt/SKILL.md`
- Test: `tests/test_default_template.py`, `tests/test_analyze.py` (기존 파일에 추가)

**Interfaces:**
- `DEFAULT_SHAPE_NAMES`에 `"doc_title": "문서제목"`, `"date": "작성일"` 추가
- `default_template_mapping(path, table_count=5)`의 `shapes`에 `"문서제목": "문서제목"`, `"작성일": "작성일"` 추가

기본 템플릿의 현재 기하를 유지한 채 도형 둘을 얹는다. `문서제목`은 슬라이드 하단 좌측, `작성일`은 상단 바 우측 끝이다. 둘 다 실제 값이 없으면 빈 문자열로 남는다 — 자리표시 문구를 넣으면 표지가 없는 Excel에서 그 문구가 그대로 산출물에 남는다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_default_template.py` 끝에 추가:

```python
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
```

`tests/test_analyze.py` 끝에 추가:

```python
def test_suggested_mapping_carries_meta_shapes(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    out = tmp_path / "work" / "structure-report.json"
    main(["--excel", str(xlsx), "--out", str(out)])
    shapes = read_json(out)["suggested_template_mapping"]["shapes"]
    assert shapes["문서제목"] == "문서제목"
    assert shapes["작성일"] == "작성일"
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_default_template.py tests/test_analyze.py --override-ini="addopts=" -v`
Expected: FAIL — `StopIteration` (도형 없음), `KeyError: '문서제목'`

- [ ] **Step 3: default_template.py 수정**

`DEFAULT_SHAPE_NAMES`를 확장한다:

```python
DEFAULT_SHAPE_NAMES = {
    "title": "제목",
    "screen_id": "화면ID",
    "image": "화면이미지",
    "doc_title": "문서제목",
    "date": "작성일",
}
```

`build_default_template`에서 상단 바를 만든 뒤에 `작성일`을, 표를 만든 뒤에 `문서제목`을 추가한다. 두 도형 모두 빈 텍스트로 둔다:

```python
    _textbox(slide, DEFAULT_SHAPE_NAMES["date"],
             margin + int(inner_w * 0.82), int(0.10 * EMU_PER_INCH),
             int(inner_w * 0.18), int(0.30 * EMU_PER_INCH),
             "", 9, TEXT_ON_BAR)
```

```python
    foot_top = tables_top + table_h + int(0.08 * EMU_PER_INCH)
    foot_h = min(int(0.22 * EMU_PER_INCH), max(slide_height_emu - foot_top, 0))
    _textbox(slide, DEFAULT_SHAPE_NAMES["doc_title"],
             margin, foot_top, int(inner_w * 0.5), foot_h,
             "", 8, SLOT_BORDER)
```

`default_template_mapping`의 `shapes`에 두 항목을 추가한다:

```python
            "문서제목": DEFAULT_SHAPE_NAMES["doc_title"],
            "작성일": DEFAULT_SHAPE_NAMES["date"],
```

`작성일`이 화면ID 도형과 겹치지 않도록 화면ID의 폭을 `inner_w * 0.38`에서 `inner_w * 0.20`으로 줄인다.

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_default_template.py tests/test_analyze.py --override-ini="addopts=" -v`
Expected: PASS

`test_default_template_meta_shapes_stay_inside_slide`가 실패하면 `foot_top` 계산이 슬라이드 밖으로 나간 것이다. 표 하단(`tables_top + table_h`)과 슬라이드 높이를 출력해 확인한다.

- [ ] **Step 5: 문서 갱신**

`references/mapping-schema.md`의 `excel` 표에 추가:

| 필드 | 값 | 설명 |
|---|---|---|
| `cover.sheet` | 시트명 | 표지 시트를 명시 지정. 생략하면 화면 시트가 아닌 첫 시트 |
| `meta_overrides` | `{"프로젝트명": "값"}` | 표지 인식 결과를 사람이 덮어쓴다 |

`screens.json` 절에 `meta` 설명을 추가한다 — 표지에서 읽은 라벨-값이 그대로 담기고, `source`·`template`은 예약 키이며, `template.shapes`에 같은 이름이 있으면 그 도형에 채워진다는 것. 화면별 `fields`가 우선한다는 것도 명시한다.

기본 템플릿 절에 `문서제목`·`작성일` 도형이 추가됐음을 적는다.

`SKILL.md`의 3단계(추출) 뒤에 한 줄 추가: 표지 시트의 프로젝트명·작성일 등이 `meta`로 들어가며, 템플릿에 같은 이름의 도형이 있으면 자동으로 채워진다.

- [ ] **Step 6: 전체 스위트와 실제 샘플 확인**

```bash
python -m pytest --override-ini="addopts=" -q
export PYTHONIOENCODING=utf-8
rm -rf work/meta-check && mkdir -p work/meta-check
python skills/excel-to-wireframe-ppt/scripts/analyze.py --excel "짧은 버전.xlsx" --out work/meta-check/structure-report.json
```

Expected: 전체 통과. 리포트의 `suggested_template_mapping.shapes`에 `문서제목`·`작성일`이 있다.

- [ ] **Step 7: 재설치와 커밋**

```bash
rm -rf /c/Users/user/.claude/skills/excel-to-wireframe-ppt
cp -r skills/excel-to-wireframe-ppt /c/Users/user/.claude/skills/
find /c/Users/user/.claude/skills/excel-to-wireframe-ppt -name __pycache__ -type d -exec rm -rf {} +
diff -r skills/excel-to-wireframe-ppt /c/Users/user/.claude/skills/excel-to-wireframe-ppt && echo "설치 일치"

git add skills/ tests/ docs/
git commit -m "feat: 기본 템플릿 meta 도형과 문서 갱신"
```

---

## 검증 매트릭스

| 스펙 항목 | 커버 위치 |
|---|---|
| 표지 시트 선택 (명시 / 자동) | Task 1 |
| 라벨-값 쌍 인식 | Task 1 |
| 단독 셀 → 문서제목·부제, 세 번째 무시 | Task 1 |
| 날짜 정규화 | Task 1 |
| `meta_overrides` 우선 | Task 1 |
| 표지 없음 → 빈 meta + 안내 출력, 새 경고 코드 없음 | Task 1, Task 2 |
| `"title": "화면설계서"` 하드코딩 제거 | Task 2 |
| `source`·`template` 예약 키 보호 | Task 1, Task 2 |
| `meta` → 동명 도형 | Task 3 |
| `fields` > `meta` 우선순위 | Task 3 |
| openpyxl 비의존 유지 | Task 3 |
| 기본 템플릿 meta 도형 | Task 4 |
| `suggested_template_mapping` 전달 | Task 4 |

스펙 4절(기본 템플릿 실측 기하 재작성)은 이 계획의 범위 밖이다. 슬라이드 크기·표 좌표·껍데기 요소는 현재 구현을 그대로 둔다.
