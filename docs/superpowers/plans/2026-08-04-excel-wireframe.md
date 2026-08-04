# Excel 기반 화면설계서 PPT 생성 Skill 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 임의 양식의 화면설계서 Excel과 사용자 제공 PowerPoint 템플릿을 입력받아, 템플릿 디자인을 유지한 화면설계서 PPT를 자동 생성하는 Claude Skill을 만든다.

**Architecture:** `analyze.py`(구조 스캔) → Claude가 `mapping.json` 작성 → `extract.py`(Excel → `screens.json` + 이미지) → `build.py`(`screens.json` + 템플릿 → pptx)의 3단 파이프라인. `screens.json`이 SSOT이며 `build.py`는 Excel을 전혀 모른다. 템플릿 디자인 보존은 python-pptx에 없는 슬라이드 복제를 XML deepcopy + rId 재매핑으로 직접 구현해 달성한다.

**Tech Stack:** Python 3.13, python-pptx 1.0.2, openpyxl 3.1.5, Pillow, pytest

## Global Constraints

- 프로젝트 루트: `C:\Users\user\Desktop\wireframe`. 스킬 본체는 `skills/excel-wireframe/`에서 개발하고, 최종 태스크에서 `C:\Users\user\.claude\skills\`로 복사한다.
- 스크립트는 `skills/excel-wireframe/scripts/`에 평면 배치한다. 패키지(`__init__.py`)를 만들지 않는다. 모듈 간 import는 `from common import ...` 형태의 절대 import를 쓰고, 테스트는 `conftest.py`에서 `sys.path`에 scripts 디렉토리를 넣어 해결한다. CLI로도 모듈로도 같은 코드가 동작해야 하기 때문이다.
- **모든 CLI 진입점은 첫 줄에서 `setup_stdio()`를 호출한다.** Windows 기본 코드페이지(cp949)에서 한글 출력이 깨지는 것을 샘플 분석 중 실제로 확인했다.
- **`build.py`는 openpyxl을 import하지 않는다.** `mapping.json`의 `template`·`options` 섹션만 읽고 `excel` 섹션은 보지 않는다. 이 제약은 Task 12에서 테스트로 강제한다.
- **`extract.py`는 기존 `screens.json`을 덮어쓰지 않는다.** 존재하면 `screens.new.json`에 쓰고 diff를 출력한다.
- 화면 단위로 예외를 격리한다. 한 화면이 실패해도 나머지는 완성한다.
- 경고는 `Warnings` 수집기에 모아 최종 리포트에 화면ID와 함께 출력한다. 코드 값은 이 아홉 개뿐이다: `no-image`, `no-detail`, `text-overflow`, `shape-not-found`, `slide-split`, `slot-shortage`, `screen-failed`, `image-convert-failed`, `orphan-row`. 새 코드를 만들지 말고 `references/mapping-schema.md`의 표와 일치시킨다.
- 슬라이드 크기 변환을 하지 않는다. 결과물은 `template.file`의 크기를 그대로 승계한다.
- **템플릿은 선택 입력이다.** 사용자가 PPT 템플릿을 주지 않으면 `default_template.py`가 표준 화면 페이지 구조를 생성해 그것을 쓴다. 기본 템플릿은 16:9(12192000 × 6858000 EMU), 상세 표 5개 × 4행 = 20슬롯이다. 사용자가 준 템플릿이 있으면 그것이 언제나 우선한다.
- 표지·목차 슬라이드를 만들지 않는다. 화면 페이지만 생성한다.
- 상세 표의 번호는 Excel의 `No.` 값을 그대로 쓴다. 스크린샷의 SoM 뱃지와 대응하므로 재부여하면 안 된다.
- 테스트는 `pytest -q` 한 줄로 전부 돌아야 한다. 실제 샘플 파일을 쓰는 회귀 테스트는 파일이 없으면 `skip`한다.

---

## 파일 구조

| 파일 | 책임 |
|---|---|
| `skills/excel-wireframe/SKILL.md` | Claude가 읽는 절차서. 3단 파이프라인과 판단 지점 3곳 |
| `skills/excel-wireframe/references/mapping-schema.md` | `mapping.json` / `screens.json` 필드 레퍼런스 |
| `scripts/common.py` | UTF-8 stdio, JSON 입출력, `Warnings` 수집기 |
| `scripts/xlsx_scan.py` | Excel 구조 스캔(시트·셀·병합·이미지 앵커) → dict |
| `scripts/pptx_scan.py` | PPTX 구조 스캔(슬라이드 크기·도형·표) → dict |
| `scripts/default_template.py` | 템플릿 미제공 시 쓸 기본 템플릿 생성 + 대응 매핑 |
| `scripts/analyze.py` | CLI. 위 둘을 합쳐 `structure-report.json` 출력 |
| `scripts/xlsx_read.py` | mapping 기반 화면·상세 읽기. 두 레이아웃 모두 담당 |
| `scripts/xlsx_images.py` | 삽입 이미지 추출. openpyxl 1차 + zip 폴백 |
| `scripts/extract.py` | CLI. `screens.json` + `images/` 생성, 덮어쓰기 금지 + diff |
| `scripts/slide_clone.py` | 슬라이드 복제 + rId 재매핑 (스파이크 검증 완료) |
| `scripts/slide_fill.py` | 도형 찾기, 텍스트 주입(서식 보존), 이미지 배치, 표 슬롯 채우기 |
| `scripts/verify.py` | 생성된 pptx 재파싱 검증 |
| `scripts/build.py` | CLI. 넘침 분할, 화면별 예외 격리, 최종 리포트 |
| `tests/fixtures.py` | 테스트용 xlsx·pptx·png 생성 함수 |
| `tests/conftest.py` | `sys.path` 설정, 공용 fixture |
| `tests/test_*.py` | 모듈별 테스트 |

---

### Task 1: 프로젝트 초기화와 공용 유틸

**Files:**
- Create: `requirements.txt`, `pytest.ini`, `.gitignore`, `CLAUDE.md`
- Create: `skills/excel-wireframe/scripts/common.py`
- Create: `tests/conftest.py`
- Test: `tests/test_common.py`

CLAUDE.md를 여기서 만드는 이유는, 이후 모든 태스크의 작업자가 그 규칙을 읽고 따라야 하기 때문이다. 마지막에 쓰면 아무 역할을 하지 못한다.

**Interfaces:**
- Consumes: 없음
- Produces:
  - `setup_stdio() -> None`
  - `read_json(path: Path) -> dict`
  - `write_json(path: Path, data: dict) -> None`
  - `class Warnings` — `add(screen_id: str | None, code: str, message: str) -> None`, `to_list() -> list[dict]`, `format() -> str`, `__len__() -> int`

- [ ] **Step 1: git 저장소와 의존성 준비**

```bash
cd /c/Users/user/Desktop/wireframe
git init
python -m pip install python-pptx openpyxl Pillow pytest
```

`.gitignore`:

```
__pycache__/
*.pyc
.pytest_cache/
work/
output/
.superpowers/
```

`requirements.txt`:

```
python-pptx>=1.0.2
openpyxl>=3.1.5
Pillow>=10.0.0
pytest>=8.0.0
```

`pytest.ini`:

```ini
[pytest]
testpaths = tests
addopts = -q
```

`CLAUDE.md`:

```markdown
# 프로젝트 규칙

Excel 화면설계서 → PPT 생성 Claude Skill을 만드는 저장소다.

## 설계 원칙

- **`screens.json`이 SSOT다.** Excel은 임포트 소스일 뿐이다. `build.py`는 openpyxl을
  import하지 않는다 — 이 제약은 `tests/test_build.py`가 소스를 검사해 강제한다.
- **`extract.py`는 `screens.json`을 덮어쓰지 않는다.** 존재하면 `screens.new.json` + diff.
  사람이 손본 내용이 재추출로 소실되면 안 된다.
- **화면 단위로 예외를 격리한다.** 한 화면이 실패해도 나머지는 완성하고, 실패는 경고로 남긴다.
- **상세 표 번호는 Excel 값을 그대로 쓴다.** 스크린샷의 SoM 뱃지와 대응하므로 재부여 금지.
- **슬라이드 크기를 바꾸지 않는다.** 템플릿 크기를 승계한다.
- **템플릿은 선택 입력이다.** 없으면 `default_template.py`가 만든다. 사용자에게 템플릿을
  먼저 요구하지 말고, 기본 템플릿으로 결과를 만들어 보여준 뒤 물어본다.

## 코드 규칙

- 스크립트는 `skills/excel-wireframe/scripts/`에 평면 배치. 패키지로 만들지 않는다.
  모듈 간 import는 `from common import ...`.
- 모든 CLI 진입점은 `setup_stdio()`를 먼저 호출한다. Windows cp949에서 한글이 깨진다.
- 표 셀에 값을 쓸 때는 `slide_fill.set_cell_text`를 쓴다. 런을 새로 만들면 서식이 초기화된다.
- 슬라이드 복제는 `slide_clone.clone_slide`만 쓴다. rId 재매핑을 빼면 그림이 사라진다.
- 테스트 템플릿 픽스처는 `default_template.build_default_template`을 재사용한다.
  픽스처용 pptx를 따로 만들지 않는다.
- 경고 코드는 아홉 개뿐이다: `no-image`, `no-detail`, `text-overflow`, `shape-not-found`,
  `slide-split`, `slot-shortage`, `screen-failed`, `image-convert-failed`, `orphan-row`.

## 테스트

```bash
python -m pytest -q                      # 전체
python -m pytest tests/test_build.py -v  # 개별
```

실제 샘플(`짧은 버전.xlsx`, `화면설계서_저작권_...pptx`)을 쓰는 회귀 테스트는
파일이 없으면 자동으로 skip된다. 픽스처는 `tests/fixtures.py`가 생성한다.

## 문서

- 설계: `docs/superpowers/specs/2026-08-03-excel-wireframe-design.md`
- 계획: `docs/superpowers/plans/2026-08-04-excel-wireframe.md`
```

- [ ] **Step 2: 실패하는 테스트 작성**

`tests/conftest.py`:

```python
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "skills" / "excel-wireframe" / "scripts"
sys.path.insert(0, str(SCRIPTS))
```

`tests/test_common.py`:

```python
from pathlib import Path

from common import Warnings, read_json, write_json


def test_json_roundtrip_keeps_korean(tmp_path: Path):
    p = tmp_path / "a.json"
    write_json(p, {"name": "이용기관 목록"})
    assert read_json(p) == {"name": "이용기관 목록"}
    assert "이용기관" in p.read_text(encoding="utf-8")


def test_warnings_collects_and_formats():
    w = Warnings()
    w.add("B2BISMT1001", "no-image", "이미지를 찾지 못했습니다")
    w.add(None, "shape-not-found", "제목 도형 없음")
    assert len(w) == 2
    assert w.to_list()[0] == {
        "screen_id": "B2BISMT1001",
        "code": "no-image",
        "message": "이미지를 찾지 못했습니다",
    }
    text = w.format()
    assert "B2BISMT1001" in text
    assert "no-image" in text


def test_warnings_format_is_empty_when_none():
    assert Warnings().format() == ""
```

- [ ] **Step 3: 테스트 실패 확인**

Run: `python -m pytest tests/test_common.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'common'`

- [ ] **Step 4: common.py 구현**

```python
# -*- coding: utf-8 -*-
"""공용 유틸: UTF-8 stdio, JSON 입출력, 경고 수집."""
from __future__ import annotations

import json
import sys
from pathlib import Path


def setup_stdio() -> None:
    """Windows cp949 콘솔에서 한글이 깨지지 않도록 UTF-8을 강제한다."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


def read_json(path: Path) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path: Path, data: dict) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


class Warnings:
    """화면 단위 경고 수집기. 최종 리포트에서 한 번에 출력한다."""

    def __init__(self) -> None:
        self._items: list[dict] = []

    def add(self, screen_id: str | None, code: str, message: str) -> None:
        self._items.append(
            {"screen_id": screen_id, "code": code, "message": message}
        )

    def to_list(self) -> list[dict]:
        return list(self._items)

    def format(self) -> str:
        if not self._items:
            return ""
        lines = ["경고 %d건:" % len(self._items)]
        for it in self._items:
            sid = it["screen_id"] or "-"
            lines.append("  [%s] %s: %s" % (sid, it["code"], it["message"]))
        return "\n".join(lines)

    def __len__(self) -> int:
        return len(self._items)
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `python -m pytest tests/test_common.py -v`
Expected: PASS (3 passed)

- [ ] **Step 6: 커밋**

```bash
git add .gitignore requirements.txt pytest.ini CLAUDE.md tests/conftest.py tests/test_common.py skills/excel-wireframe/scripts/common.py docs/
git commit -m "feat: 프로젝트 초기화와 공용 유틸(UTF-8 stdio, JSON, 경고 수집)"
```

---

### Task 2: Excel 구조 스캔

**Files:**
- Create: `skills/excel-wireframe/scripts/xlsx_scan.py`
- Create: `tests/fixtures.py`
- Test: `tests/test_xlsx_scan.py`

**Interfaces:**
- Consumes: 없음
- Produces:
  - `scan_workbook(path: Path, max_rows: int = 60, max_cols: int = 10) -> dict`
    반환 구조: `{"file": str, "sheets": [{"name": str, "max_row": int, "max_column": int, "merged": [str], "image_count": int, "cells": [{"ref": str, "value": str}]}]}`
  - `tests/fixtures.py`: `make_png(path: Path, size=(400, 300)) -> Path`, `make_sheet_per_screen_xlsx(path: Path, screens: list[dict]) -> Path`, `make_table_xlsx(path: Path) -> Path`

`make_sheet_per_screen_xlsx`의 `screens` 항목 형식: `{"id": "SCR001", "name": "화면명", "details": [{"no": "1", "type": "버튼", "element": "[등록]", "desc": "설명", "pos": "상단"}], "image": True}`

- [ ] **Step 1: 픽스처 생성기 작성**

`tests/fixtures.py`:

```python
# -*- coding: utf-8 -*-
"""테스트용 xlsx / pptx / png 생성기."""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

DETAIL_HEADER = ["No.", "요소타입", "요소명", "상세 설명", "위치"]


def make_png(path: Path, size=(400, 300), color=(200, 220, 255)) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    PILImage.new("RGB", size, color).save(path)
    return path


def make_sheet_per_screen_xlsx(path: Path, screens: list[dict]) -> Path:
    """1시트 = 1화면 양식. 화면이 아닌 시트도 섞어 시트 필터를 검증할 수 있게 한다."""
    wb = Workbook()
    cover = wb.active
    cover.title = "표지"
    cover["A1"] = "화면설계서"

    for scr in screens:
        ws = wb.create_sheet("설계_%s" % scr["id"])
        ws["A1"] = "화면설계서 - %s (%s)" % (scr["id"], scr["name"])
        ws["A3"] = "[웹 스크린샷 (SoM 뱃지)]"
        if scr.get("image"):
            img_path = path.parent / ("_fx_%s.png" % scr["id"])
            # 화면마다 색을 달리한다. 내용이 같으면 저장 과정에서 하나로 합쳐져
            # 이미지 귀속 테스트가 무의미해질 수 있다.
            tint = 40 * (ord(scr["id"][-1]) % 5)
            make_png(img_path, color=(200, 220 - tint, 255 - tint))
            ws.add_image(XLImage(str(img_path)), "A4")
        header_row = 28
        for i, name in enumerate(DETAIL_HEADER):
            ws.cell(row=header_row, column=i + 1, value=name)
        for j, d in enumerate(scr["details"]):
            r = header_row + 1 + j
            ws.cell(row=r, column=1, value=d["no"])
            ws.cell(row=r, column=2, value=d["type"])
            ws.cell(row=r, column=3, value=d["element"])
            ws.cell(row=r, column=4, value=d["desc"])
            ws.cell(row=r, column=5, value=d["pos"])

    test_ws = wb.create_sheet("테스트_무시대상")
    test_ws["A1"] = "이 시트는 무시되어야 한다"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def make_table_xlsx(path: Path) -> Path:
    """1행 = 1화면 양식. 화면ID가 빈 후속 행은 직전 화면의 상세 행이다."""
    wb = Workbook()
    ws = wb.active
    ws.title = "화면목록"
    ws["A1"] = "화면 정의서"
    headers = ["화면ID", "화면명", "화면설명", "No.", "요소명", "상세 설명"]
    for i, h in enumerate(headers):
        ws.cell(row=3, column=i + 1, value=h)

    rows = [
        ["SCR001", "이용기관 목록", "기관을 조회한다", "1", "[등록] 버튼", "등록 팝업을 연다"],
        ["", "", "", "2", "[삭제] 버튼", "선택 항목을 삭제한다"],
        ["SCR002", "이용기관 상세", "상세를 본다", "1", "[저장] 버튼", "변경 내용을 저장한다"],
    ]
    for j, row in enumerate(rows):
        for i, v in enumerate(row):
            ws.cell(row=4 + j, column=i + 1, value=v)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
```

- [ ] **Step 2: 실패하는 테스트 작성**

`tests/test_xlsx_scan.py`:

```python
from pathlib import Path

from fixtures import make_sheet_per_screen_xlsx
from xlsx_scan import scan_workbook

SCREENS = [
    {
        "id": "SCR001",
        "name": "이용기관 목록",
        "image": True,
        "details": [
            {"no": "1", "type": "버튼", "element": "[등록]", "desc": "등록한다", "pos": "상단"},
            {"no": "2", "type": "버튼", "element": "[삭제]", "desc": "삭제한다", "pos": "상단"},
        ],
    }
]


def test_scan_lists_all_sheets(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    report = scan_workbook(xlsx)
    names = [s["name"] for s in report["sheets"]]
    assert names == ["표지", "설계_SCR001", "테스트_무시대상"]


def test_scan_reports_cells_and_images(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    report = scan_workbook(xlsx)
    sheet = next(s for s in report["sheets"] if s["name"] == "설계_SCR001")
    cells = {c["ref"]: c["value"] for c in sheet["cells"]}
    assert cells["A1"] == "화면설계서 - SCR001 (이용기관 목록)"
    assert cells["A28"] == "No."
    assert cells["D29"] == "등록한다"
    assert sheet["image_count"] == 1


def test_scan_skips_empty_cells(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    report = scan_workbook(xlsx)
    sheet = next(s for s in report["sheets"] if s["name"] == "설계_SCR001")
    refs = [c["ref"] for c in sheet["cells"]]
    assert "B1" not in refs
```

- [ ] **Step 3: 테스트 실패 확인**

Run: `python -m pytest tests/test_xlsx_scan.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xlsx_scan'`

- [ ] **Step 4: xlsx_scan.py 구현**

```python
# -*- coding: utf-8 -*-
"""Excel 구조 스캔. Claude가 양식을 판단할 재료를 만든다."""
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def scan_workbook(path: Path, max_rows: int = 60, max_cols: int = 10) -> dict:
    """시트별 셀 값·병합 범위·이미지 수를 훑는다.

    전체를 읽지 않고 앞쪽 일부만 보는 것은, 양식 판단에 필요한 단서(제목 셀,
    헤더 행)가 상단에 몰려 있고 대형 파일에서 리포트가 폭발하는 것을 막기 위해서다.
    """
    path = Path(path)
    wb = load_workbook(path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        cells = []
        row_limit = min(ws.max_row or 0, max_rows)
        col_limit = min(ws.max_column or 0, max_cols)
        for r in range(1, row_limit + 1):
            for c in range(1, col_limit + 1):
                v = ws.cell(row=r, column=c).value
                if v is None:
                    continue
                text = str(v).strip()
                if not text:
                    continue
                cells.append({"ref": "%s%d" % (get_column_letter(c), r), "value": text})
        sheets.append(
            {
                "name": ws.title,
                "max_row": ws.max_row or 0,
                "max_column": ws.max_column or 0,
                "merged": [str(rng) for rng in ws.merged_cells.ranges],
                "image_count": len(getattr(ws, "_images", [])),
                "cells": cells,
            }
        )
    wb.close()
    return {"file": str(path), "sheets": sheets}
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `python -m pytest tests/test_xlsx_scan.py -v`
Expected: PASS (3 passed)

- [ ] **Step 6: 커밋**

```bash
git add tests/fixtures.py tests/test_xlsx_scan.py skills/excel-wireframe/scripts/xlsx_scan.py
git commit -m "feat: Excel 구조 스캔과 테스트 픽스처 생성기"
```

---

### Task 3: 기본 템플릿 생성과 PPTX 구조 스캔

**Files:**
- Create: `skills/excel-wireframe/scripts/default_template.py`
- Create: `skills/excel-wireframe/scripts/pptx_scan.py`
- Modify: `tests/fixtures.py` (템플릿 픽스처 추가)
- Test: `tests/test_default_template.py`, `tests/test_pptx_scan.py`

**Interfaces:**
- Consumes: 없음
- Produces:
  - `build_default_template(path: Path, slide_width_emu: int = 12192000, slide_height_emu: int = 6858000, table_count: int = 5, rows_per_table: int = 4) -> Path`
  - `default_template_mapping(path: Path, table_count: int = 5) -> dict` — `mapping.json`의 `template` 섹션에 그대로 넣을 수 있는 dict
  - `DEFAULT_SHAPE_NAMES: dict` — `{"title": "제목", "screen_id": "화면ID", "image": "화면이미지"}`
  - `scan_presentation(path: Path) -> dict`
    반환: `{"file": str, "slide_width": int, "slide_height": int, "slide_size_in": [float, float], "slides": [{"index": int, "layout": str, "shape_count": int, "text_shape_count": int, "shapes": [...]}]}`
    도형 항목: `{"name": str, "shape_type": str, "is_placeholder": bool, "left": int, "top": int, "width": int, "height": int, "text": str, "table": {"rows": int, "cols": int} | None}`
  - `suggest_mode(report: dict) -> dict` — `{"mode": "clone"|"layout", "source_slide": int | None, "reason": str}`
  - `tests/fixtures.py`: `make_template_pptx(path, table_count=5, rows_per_table=4) -> Path`, `make_empty_layout_pptx(path) -> Path`

`suggest_mode` 판정 규칙: 텍스트가 채워진 도형이 3개 이상이고 표 또는 그림을 하나 이상 가진 슬라이드 중 **가장 앞선 것**을 `source_slide`로 하는 `clone`. 그런 슬라이드가 없으면 `layout`.

**왜 이 모듈이 필요한가** — 사용자가 Excel만 주고 PPT 템플릿은 안 줄 수 있다. 그때 저작권 있는 실제 템플릿을 번들해 둘 수는 없으므로, 표준 화면 페이지 구조(제목 바 + 화면ID + 이미지 자리 + 하단 상세 표 N개)를 코드로 생성한다. 테스트 픽스처도 같은 생성기를 재사용해 중복을 없앤다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_default_template.py`:

```python
from pathlib import Path

from default_template import (
    DEFAULT_SHAPE_NAMES,
    build_default_template,
    default_template_mapping,
)
from pptx import Presentation


def test_default_template_is_16_9(tmp_path: Path):
    p = build_default_template(tmp_path / "d.pptx")
    prs = Presentation(str(p))
    assert prs.slide_width == 12192000
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
    p = build_default_template(
        tmp_path / "d.pptx", slide_width_emu=9906000, table_count=3
    )
    prs = Presentation(str(p))
    assert prs.slide_width == 9906000
    assert sum(1 for s in prs.slides[0].shapes if s.has_table) == 3


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
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_default_template.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'default_template'`

- [ ] **Step 3: default_template.py 구현**

```python
# -*- coding: utf-8 -*-
"""PPT 템플릿을 사용자가 주지 않았을 때 쓸 기본 템플릿을 만든다.

실제 화면설계서의 표준 페이지 구조를 재현한다 — 상단 제목 바, 우측 화면ID,
가운데 큰 이미지 자리, 하단에 가로로 나란한 상세 표. 도형 이름을 의미 있게 붙여
mapping.json을 자동으로 채울 수 있게 한다.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

DEFAULT_SHAPE_NAMES = {
    "title": "제목",
    "screen_id": "화면ID",
    "image": "화면이미지",
}

BAR_COLOR = RGBColor(0x1F, 0x3B, 0x63)
TEXT_ON_BAR = RGBColor(0xFF, 0xFF, 0xFF)
SLOT_BORDER = RGBColor(0xBB, 0xBB, 0xBB)
SLOT_FILL = RGBColor(0xF5, 0xF6, 0xF8)
EMU_PER_INCH = 914400


def _textbox(slide, name, left, top, width, height, text, size_pt, color=None,
             bold=False):
    box = slide.shapes.add_textbox(Emu(int(left)), Emu(int(top)),
                                   Emu(int(width)), Emu(int(height)))
    box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    return box


def build_default_template(
    path: Path,
    slide_width_emu: int = 12192000,
    slide_height_emu: int = 6858000,
    table_count: int = 5,
    rows_per_table: int = 4,
) -> Path:
    prs = Presentation()
    prs.slide_width = Emu(slide_width_emu)
    prs.slide_height = Emu(slide_height_emu)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    margin = int(0.25 * EMU_PER_INCH)
    inner_w = slide_width_emu - margin * 2
    bar_h = int(0.45 * EMU_PER_INCH)

    bar = slide.shapes.add_shape(1, Emu(0), Emu(0), Emu(slide_width_emu), Emu(bar_h))
    bar.name = "제목바"
    bar.fill.solid()
    bar.fill.fore_color.rgb = BAR_COLOR
    bar.line.fill.background()
    bar.text_frame.text = ""

    _textbox(slide, DEFAULT_SHAPE_NAMES["title"], margin, int(0.06 * EMU_PER_INCH),
             int(inner_w * 0.6), int(0.35 * EMU_PER_INCH),
             "화면명", 18, TEXT_ON_BAR, bold=True)
    _textbox(slide, DEFAULT_SHAPE_NAMES["screen_id"],
             margin + int(inner_w * 0.62), int(0.10 * EMU_PER_INCH),
             int(inner_w * 0.38), int(0.30 * EMU_PER_INCH),
             "화면ID", 11, TEXT_ON_BAR)

    tables_top = int(slide_height_emu - 1.85 * EMU_PER_INCH)
    img_top = bar_h + int(0.15 * EMU_PER_INCH)
    img_h = tables_top - img_top - int(0.15 * EMU_PER_INCH)

    img_slot = slide.shapes.add_shape(1, Emu(margin), Emu(img_top),
                                      Emu(inner_w), Emu(img_h))
    img_slot.name = DEFAULT_SHAPE_NAMES["image"]
    img_slot.fill.solid()
    img_slot.fill.fore_color.rgb = SLOT_FILL
    img_slot.line.color.rgb = SLOT_BORDER
    tf = img_slot.text_frame
    run = tf.paragraphs[0].add_run()
    run.text = "[화면 이미지]"
    run.font.size = Pt(14)
    run.font.color.rgb = SLOT_BORDER

    gap = int(0.06 * EMU_PER_INCH)
    table_w = (inner_w - gap * (table_count - 1)) // table_count
    table_h = int(1.55 * EMU_PER_INCH)
    for t in range(table_count):
        left = margin + (table_w + gap) * t
        shp = slide.shapes.add_table(
            rows_per_table, 2, Emu(left), Emu(tables_top), Emu(table_w), Emu(table_h)
        )
        shp.name = "상세표%d" % (t + 1)
        table = shp.table
        table.columns[0].width = Emu(int(table_w * 0.18))
        table.columns[1].width = Emu(table_w - int(table_w * 0.18))
        for r in range(rows_per_table):
            n = t * rows_per_table + r + 1
            table.cell(r, 0).text = str(n)
            table.cell(r, 1).text = "예시 설명 %d" % n
            for c in range(2):
                for p in table.cell(r, c).text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(9)

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
            "detail_tables": ["상세표%d" % (i + 1) for i in range(table_count)],
        },
        "table_columns": {"no": 0, "text": 1},
    }
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_default_template.py -v`
Expected: PASS (6 passed)

- [ ] **Step 5: 템플릿 픽스처 추가**

`tests/fixtures.py` 끝에 추가. 생성기를 재사용하되 도형 이름만 실제 샘플의 이름(`제목 13` 등)으로 바꾼다 — 이름 기반 매핑이 임의의 도형 이름에서도 동작하는지 검증하기 위해서다.

```python
from default_template import build_default_template
from pptx import Presentation
from pptx.util import Emu


def make_template_pptx(path: Path, table_count: int = 5, rows_per_table: int = 4) -> Path:
    """실제 샘플과 같은 도형 이름·슬라이드 크기를 가진 예시 슬라이드형 템플릿."""
    build_default_template(
        path,
        slide_width_emu=9906000,   # 10.83in — 실제 샘플과 동일
        slide_height_emu=6858000,  # 7.50in
        table_count=table_count,
        rows_per_table=rows_per_table,
    )
    prs = Presentation(str(path))
    rename = {"제목": "제목 13", "화면ID": "텍스트 개체 틀 14", "화면이미지": "그림 18"}
    for shp in prs.slides[0].shapes:
        if shp.name in rename:
            shp.name = rename[shp.name]
        elif shp.name.startswith("상세표"):
            shp.name = "표 %d" % (6 + int(shp.name[len("상세표"):]))
    prs.save(str(path))
    return path


def make_empty_layout_pptx(path: Path) -> Path:
    """빈 레이아웃형 템플릿 — clone 모드로 오판되면 안 된다."""
    prs = Presentation()
    prs.slide_width = Emu(9906000)
    prs.slide_height = Emu(6858000)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path
```

- [ ] **Step 6: 실패하는 스캔 테스트 작성**

`tests/test_pptx_scan.py`:

```python
from pathlib import Path

from fixtures import make_empty_layout_pptx, make_template_pptx
from pptx_scan import scan_presentation, suggest_mode


def test_scan_reports_slide_size_and_shapes(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    report = scan_presentation(pptx)
    assert report["slide_width"] == 9906000
    assert round(report["slide_size_in"][0], 2) == 10.83
    slide = report["slides"][0]
    names = [s["name"] for s in slide["shapes"]]
    assert "제목 13" in names
    assert "그림 18" in names
    assert names.count("표 7") == 1


def test_scan_reports_table_dimensions(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    report = scan_presentation(pptx)
    tables = [s for s in report["slides"][0]["shapes"] if s["table"]]
    assert len(tables) == 5
    assert tables[0]["table"] == {"rows": 4, "cols": 2}


def test_suggest_mode_picks_clone_for_example_slide(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    got = suggest_mode(scan_presentation(pptx))
    assert got["mode"] == "clone"
    assert got["source_slide"] == 0


def test_suggest_mode_falls_back_to_layout(tmp_path: Path):
    pptx = make_empty_layout_pptx(tmp_path / "e.pptx")
    got = suggest_mode(scan_presentation(pptx))
    assert got["mode"] == "layout"
    assert got["source_slide"] is None
```

- [ ] **Step 7: 테스트 실패 확인**

Run: `python -m pytest tests/test_pptx_scan.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'pptx_scan'`

- [ ] **Step 8: pptx_scan.py 구현**

```python
# -*- coding: utf-8 -*-
"""PPTX 구조 스캔. 어느 슬라이드를 복제 소스로 쓸지 판단할 재료를 만든다."""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation

EMU_PER_INCH = 914400


def _shape_text(shape) -> str:
    if not shape.has_text_frame:
        return ""
    return shape.text_frame.text.strip()


def _scan_shape(shape) -> dict:
    table = None
    if shape.has_table:
        table = {"rows": len(shape.table.rows), "cols": len(shape.table.columns)}
    return {
        "name": shape.name,
        "shape_type": str(shape.shape_type),
        "is_placeholder": bool(shape.is_placeholder),
        "left": int(shape.left) if shape.left is not None else 0,
        "top": int(shape.top) if shape.top is not None else 0,
        "width": int(shape.width) if shape.width is not None else 0,
        "height": int(shape.height) if shape.height is not None else 0,
        "text": _shape_text(shape)[:200],
        "table": table,
    }


def scan_presentation(path: Path) -> dict:
    path = Path(path)
    prs = Presentation(str(path))
    slides = []
    for i, slide in enumerate(prs.slides):
        shapes = [_scan_shape(s) for s in slide.shapes]
        slides.append(
            {
                "index": i,
                "layout": slide.slide_layout.name,
                "shape_count": len(shapes),
                "text_shape_count": sum(1 for s in shapes if s["text"]),
                "shapes": shapes,
            }
        )
    return {
        "file": str(path),
        "slide_width": int(prs.slide_width),
        "slide_height": int(prs.slide_height),
        "slide_size_in": [
            prs.slide_width / EMU_PER_INCH,
            prs.slide_height / EMU_PER_INCH,
        ],
        "slides": slides,
    }


def suggest_mode(report: dict) -> dict:
    """예시 슬라이드가 있으면 clone, 없으면 layout.

    '텍스트가 채워진 도형 3개 이상 + 표나 그림 1개 이상'을 예시 슬라이드의 신호로 본다.
    빈 레이아웃만 있는 템플릿은 이 조건을 통과하지 못한다.
    """
    for s in report["slides"]:
        has_visual = any(
            sh["table"] is not None or "PICTURE" in sh["shape_type"]
            for sh in s["shapes"]
        )
        if s["text_shape_count"] >= 3 and has_visual:
            return {
                "mode": "clone",
                "source_slide": s["index"],
                "reason": "슬라이드 %d에 채워진 텍스트 %d개와 표/그림이 있어 예시 슬라이드로 판단"
                % (s["index"], s["text_shape_count"]),
            }
    return {
        "mode": "layout",
        "source_slide": None,
        "reason": "예시 슬라이드를 찾지 못해 빈 레이아웃 모드로 판단",
    }
```

- [ ] **Step 9: 테스트 통과 확인**

Run: `python -m pytest tests/test_pptx_scan.py tests/test_default_template.py -v`
Expected: PASS (10 passed)

- [ ] **Step 10: 기본 템플릿을 눈으로 확인**

```bash
python -c "import sys; sys.path.insert(0,'skills/excel-wireframe/scripts'); from default_template import build_default_template; build_default_template('work/default-template.pptx')"
```

`work/default-template.pptx`를 PowerPoint로 열어 표가 슬라이드 밖으로 나가지 않고 이미지 자리와 겹치지 않는지 확인한다. 좌표 계산이 틀리면 모든 생성물이 같은 방식으로 어긋난다.

- [ ] **Step 11: 커밋**

```bash
git add tests/fixtures.py tests/test_pptx_scan.py tests/test_default_template.py skills/excel-wireframe/scripts/pptx_scan.py skills/excel-wireframe/scripts/default_template.py
git commit -m "feat: 기본 템플릿 생성기와 PPTX 구조 스캔"
```

---

### Task 4: analyze.py CLI

**Files:**
- Create: `skills/excel-wireframe/scripts/analyze.py`
- Test: `tests/test_analyze.py`

**Interfaces:**
- Consumes: `scan_workbook`, `scan_presentation`, `suggest_mode`, `build_default_template`, `default_template_mapping`, `setup_stdio`, `write_json`
- Produces:
  - CLI: `python analyze.py --excel <xlsx> [--template <pptx>] --out <structure-report.json>`
  - `build_report(excel_path: Path, template_path: Path) -> dict` — `{"excel": {...}, "template": {...}, "suggestion": {...}}`
  - `resolve_template(template_arg: str | None, out_path: Path) -> tuple[Path, bool]` — 템플릿 경로와 "생성했는가" 플래그
  - `main(argv: list[str] | None = None) -> int`

**`--template` 생략 시**: `<out의 부모 디렉토리>/default-template.pptx`를 생성해 그것을 쓴다. 리포트에 `template_generated: true`와 `suggested_template_mapping`(그대로 `mapping.json`의 `template` 섹션에 넣을 수 있는 dict)을 담아, Claude가 도형 이름을 추측하지 않아도 되게 한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_analyze.py`:

```python
from pathlib import Path

from analyze import build_report, main, resolve_template
from common import read_json
from fixtures import make_sheet_per_screen_xlsx, make_template_pptx

SCREENS = [
    {
        "id": "SCR001",
        "name": "이용기관 목록",
        "image": True,
        "details": [{"no": "1", "type": "버튼", "element": "[등록]", "desc": "등록한다", "pos": "상단"}],
    }
]


def test_build_report_joins_both_sides(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    pptx = make_template_pptx(tmp_path / "t.pptx")
    report = build_report(xlsx, pptx)
    assert [s["name"] for s in report["excel"]["sheets"]][1] == "설계_SCR001"
    assert report["template"]["slide_width"] == 9906000
    assert report["suggestion"]["mode"] == "clone"


def test_main_writes_report_file(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    pptx = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "work" / "structure-report.json"
    code = main(["--excel", str(xlsx), "--template", str(pptx), "--out", str(out)])
    assert code == 0
    data = read_json(out)
    assert data["suggestion"]["source_slide"] == 0
    assert data["template_generated"] is False


def test_resolve_template_uses_given_path(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    path, generated = resolve_template(str(pptx), tmp_path / "work" / "r.json")
    assert path == pptx
    assert generated is False


def test_resolve_template_generates_when_missing(tmp_path: Path):
    out = tmp_path / "work" / "r.json"
    path, generated = resolve_template(None, out)
    assert generated is True
    assert path == tmp_path / "work" / "default-template.pptx"
    assert path.exists()


def test_main_without_template_generates_one(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    out = tmp_path / "work" / "structure-report.json"
    code = main(["--excel", str(xlsx), "--out", str(out)])
    assert code == 0

    data = read_json(out)
    assert data["template_generated"] is True
    assert (tmp_path / "work" / "default-template.pptx").exists()
    assert data["suggestion"]["mode"] == "clone"

    suggested = data["suggested_template_mapping"]
    assert suggested["shapes"]["title"] == "제목"
    assert suggested["shapes"]["detail_tables"][0] == "상세표1"
    assert data["template"]["slide_width"] == 12192000


def test_main_with_given_template_has_no_suggestion(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    pptx = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "work" / "structure-report.json"
    main(["--excel", str(xlsx), "--template", str(pptx), "--out", str(out)])
    assert "suggested_template_mapping" not in read_json(out)
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_analyze.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'analyze'`

- [ ] **Step 3: analyze.py 구현**

```python
# -*- coding: utf-8 -*-
"""1단계: Excel과 템플릿의 구조를 스캔해 리포트로 남긴다.

이 리포트를 Claude가 읽고 mapping.json을 작성한다. 스크립트는 판단하지 않고
관찰만 한다 — suggest_mode의 제안조차 사용자 확인을 거친다.

템플릿을 주지 않으면 기본 템플릿을 만들어 쓴다. 그 경우 도형 이름을 우리가 정했으므로
추측할 필요가 없고, 바로 쓸 수 있는 매핑을 리포트에 함께 담는다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from common import setup_stdio, write_json
from default_template import build_default_template, default_template_mapping
from pptx_scan import scan_presentation, suggest_mode
from xlsx_scan import scan_workbook


def build_report(excel_path: Path, template_path: Path) -> dict:
    excel = scan_workbook(excel_path)
    template = scan_presentation(template_path)
    return {
        "excel": excel,
        "template": template,
        "suggestion": suggest_mode(template),
    }


def resolve_template(template_arg: str | None, out_path: Path) -> tuple[Path, bool]:
    if template_arg:
        return Path(template_arg), False
    generated = Path(out_path).parent / "default-template.pptx"
    build_default_template(generated)
    return generated, True


def main(argv: list[str] | None = None) -> int:
    setup_stdio()
    ap = argparse.ArgumentParser(description="Excel과 PPT 템플릿 구조를 스캔한다")
    ap.add_argument("--excel", required=True)
    ap.add_argument("--template", default=None,
                    help="생략하면 기본 템플릿을 만들어 사용한다")
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)

    out_path = Path(args.out)
    template_path, generated = resolve_template(args.template, out_path)

    report = build_report(Path(args.excel), template_path)
    report["template_generated"] = generated
    if generated:
        report["suggested_template_mapping"] = default_template_mapping(template_path)
    write_json(out_path, report)

    sheets = report["excel"]["sheets"]
    print("Excel 시트 %d개: %s" % (len(sheets), ", ".join(s["name"] for s in sheets)))
    if generated:
        print("템플릿 미제공 → 기본 템플릿 생성: %s" % template_path)
        print("  제안 매핑이 리포트의 suggested_template_mapping에 있습니다")
    w, h = report["template"]["slide_size_in"]
    print("템플릿 슬라이드 %d장, 크기 %.2f x %.2f in"
          % (len(report["template"]["slides"]), w, h))
    print("제안 모드: %s (%s)"
          % (report["suggestion"]["mode"], report["suggestion"]["reason"]))
    print("리포트 저장: %s" % args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_analyze.py -v`
Expected: PASS (6 passed)

- [ ] **Step 5: 실제 샘플로 손수 확인**

```bash
cd /c/Users/user/Desktop/wireframe
python skills/excel-wireframe/scripts/analyze.py \
  --excel "짧은 버전.xlsx" \
  --template "화면설계서_저작권_발행기관.정산처_발행기관관리_v1.0_20260427.pptx" \
  --out work/structure-report.json
```

Expected: 시트 4개(`표지`, `설계_B2BISMT1001`, `테스트_B2BISMT1001`, `비교결과요약`), 크기 `10.83 x 7.50 in`, 모드 `clone` 출력. 한글이 깨지지 않아야 한다.

템플릿 없는 경로도 확인한다:

```bash
python skills/excel-wireframe/scripts/analyze.py \
  --excel "짧은 버전.xlsx" \
  --out work/no-template/structure-report.json
```

Expected: `템플릿 미제공 → 기본 템플릿 생성` 줄이 나오고 `13.33 x 7.50 in`으로 보고된다.

- [ ] **Step 6: 커밋**

```bash
git add tests/test_analyze.py skills/excel-wireframe/scripts/analyze.py
git commit -m "feat: analyze.py — 구조 리포트 생성 CLI"
```

---

### Task 5: sheet-per-screen 레이아웃 읽기

**Files:**
- Create: `skills/excel-wireframe/scripts/xlsx_read.py`
- Test: `tests/test_xlsx_read_sheet.py`

**Interfaces:**
- Consumes: `Warnings`
- Produces:
  - `find_header_row(ws, column: str, marker: str, limit: int = 200) -> int | None`
  - `parse_screen_meta(text: str, pattern: str | None, fallback_id: str) -> tuple[str, str]`
  - `read_screens(wb, mapping: dict, warns: Warnings) -> list[dict]`
    화면 dict: `{"id": str, "name": str, "sheet": str, "images": [], "fields": {}, "details": [dict]}`

`read_screens`는 `mapping["excel"]["layout"]`으로 분기한다. Task 5는 `sheet-per-screen`만 구현하고 `table`은 Task 6에서 채운다. `images`는 빈 리스트로 두고 Task 7의 `xlsx_images`가 채운다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_xlsx_read_sheet.py`:

```python
from pathlib import Path

from common import Warnings
from fixtures import make_sheet_per_screen_xlsx
from openpyxl import load_workbook
from xlsx_read import find_header_row, parse_screen_meta, read_screens

MAPPING = {
    "excel": {
        "layout": "sheet-per-screen",
        "sheet_include": "^설계_",
        "screen_meta": {
            "cell": "A1",
            "pattern": r"화면설계서\s*-\s*(?P<id>\S+)\s*\((?P<name>.+)\)",
        },
        "detail": {
            "header_scan_column": "A",
            "header_marker": "No.",
            "columns": {"no": "A", "type": "B", "element": "C", "desc": "D", "pos": "E"},
        },
    }
}

SCREENS = [
    {
        "id": "SCR001",
        "name": "이용기관 목록",
        "image": False,
        "details": [
            {"no": "1", "type": "버튼", "element": "[등록]", "desc": "등록한다", "pos": "상단"},
            {"no": "2", "type": "버튼", "element": "[삭제]", "desc": "삭제한다", "pos": "상단"},
        ],
    },
    {
        "id": "SCR002",
        "name": "이용기관 상세",
        "image": False,
        "details": [
            {"no": "1", "type": "버튼", "element": "[저장]", "desc": "저장한다", "pos": "하단"},
        ],
    },
]


def test_parse_screen_meta_splits_id_and_name():
    pat = MAPPING["excel"]["screen_meta"]["pattern"]
    assert parse_screen_meta("화면설계서 - SCR001 (이용기관 목록)", pat, "설계_SCR001") == (
        "SCR001",
        "이용기관 목록",
    )


def test_parse_screen_meta_falls_back_when_pattern_misses():
    pat = MAPPING["excel"]["screen_meta"]["pattern"]
    assert parse_screen_meta("그냥 제목", pat, "설계_SCR009") == ("설계_SCR009", "그냥 제목")


def test_find_header_row(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    wb = load_workbook(xlsx, data_only=True)
    ws = wb["설계_SCR001"]
    assert find_header_row(ws, "A", "No.") == 28
    assert find_header_row(ws, "A", "존재하지않음") is None


def test_read_screens_filters_sheets_and_reads_details(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    wb = load_workbook(xlsx, data_only=True)
    warns = Warnings()
    screens = read_screens(wb, MAPPING, warns)

    assert [s["id"] for s in screens] == ["SCR001", "SCR002"]
    assert screens[0]["name"] == "이용기관 목록"
    assert len(screens[0]["details"]) == 2
    assert screens[0]["details"][1] == {
        "no": "2", "type": "버튼", "element": "[삭제]", "desc": "삭제한다", "pos": "상단",
    }
    assert screens[0]["images"] == []
    assert screens[0]["fields"] == {}


def test_read_screens_warns_when_header_missing(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    wb = load_workbook(xlsx, data_only=True)
    mapping = {"excel": dict(MAPPING["excel"])}
    mapping["excel"]["detail"] = dict(MAPPING["excel"]["detail"])
    mapping["excel"]["detail"]["header_marker"] = "없는마커"
    warns = Warnings()
    screens = read_screens(wb, mapping, warns)
    assert screens[0]["details"] == []
    codes = [w["code"] for w in warns.to_list()]
    assert "shape-not-found" not in codes
    assert any(w["message"].startswith("상세 표 헤더") for w in warns.to_list())
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_xlsx_read_sheet.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xlsx_read'`

- [ ] **Step 3: xlsx_read.py 구현 (sheet-per-screen)**

```python
# -*- coding: utf-8 -*-
"""mapping.json에 따라 Excel에서 화면과 상세를 읽는다."""
from __future__ import annotations

import re

from common import Warnings
from openpyxl.utils import column_index_from_string


def _cell_text(ws, row: int, col: int) -> str:
    v = ws.cell(row=row, column=col).value
    return "" if v is None else str(v).strip()


def find_header_row(ws, column: str, marker: str, limit: int = 200) -> int | None:
    """상세 표 헤더 행을 위에서부터 훑어 찾는다.

    헤더 위치는 상단 이미지 크기에 따라 밀리므로 고정할 수 없다.
    """
    col = column_index_from_string(column)
    last = min(ws.max_row or 0, limit)
    for r in range(1, last + 1):
        if _cell_text(ws, r, col) == marker:
            return r
    return None


def parse_screen_meta(text: str, pattern: str | None, fallback_id: str) -> tuple[str, str]:
    """제목 셀에서 화면ID와 화면명을 분리한다.

    패턴이 안 맞으면 셀 전체를 이름으로 두고 ID는 시트명에서 딴다. 양식이 조금
    달라도 화면 자체가 통째로 사라지지 않게 하기 위한 방어다.
    """
    if pattern:
        m = re.search(pattern, text or "")
        if m:
            gd = m.groupdict()
            sid = (gd.get("id") or "").strip()
            name = (gd.get("name") or "").strip()
            if sid or name:
                return sid or fallback_id, name or fallback_id
    return fallback_id, (text or "").strip() or fallback_id


def _read_details(ws, detail_cfg: dict, warns: Warnings, screen_id: str) -> list[dict]:
    marker = detail_cfg.get("header_marker", "No.")
    scan_col = detail_cfg.get("header_scan_column", "A")
    header_row = find_header_row(ws, scan_col, marker)
    if header_row is None:
        warns.add(screen_id, "no-detail",
                  "상세 표 헤더('%s')를 찾지 못했습니다" % marker)
        return []

    cols = {k: column_index_from_string(v) for k, v in detail_cfg["columns"].items()}
    key_col = cols.get("no") or column_index_from_string(scan_col)
    details = []
    r = header_row + 1
    last = ws.max_row or 0
    while r <= last:
        row_vals = {k: _cell_text(ws, r, c) for k, c in cols.items()}
        if not any(row_vals.values()):
            break
        if not row_vals.get("no") and not _cell_text(ws, r, key_col):
            break
        details.append(row_vals)
        r += 1
    return details


def _read_sheet_per_screen(wb, cfg: dict, warns: Warnings) -> list[dict]:
    include = cfg.get("sheet_include")
    rx = re.compile(include) if include else None
    meta_cfg = cfg.get("screen_meta", {})
    meta_cell = meta_cfg.get("cell", "A1")
    pattern = meta_cfg.get("pattern")

    screens = []
    for ws in wb.worksheets:
        if rx and not rx.search(ws.title):
            continue
        raw = ws[meta_cell].value
        text = "" if raw is None else str(raw).strip()
        sid, name = parse_screen_meta(text, pattern, ws.title)
        details = _read_details(ws, cfg["detail"], warns, sid)
        screens.append(
            {
                "id": sid,
                "name": name,
                "sheet": ws.title,
                "images": [],
                "fields": {},
                "details": details,
            }
        )
    return screens


def read_screens(wb, mapping: dict, warns: Warnings) -> list[dict]:
    cfg = mapping["excel"]
    layout = cfg.get("layout", "sheet-per-screen")
    if layout == "sheet-per-screen":
        return _read_sheet_per_screen(wb, cfg, warns)
    raise ValueError("지원하지 않는 excel.layout: %s" % layout)
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_xlsx_read_sheet.py -v`
Expected: PASS (5 passed)

- [ ] **Step 5: 커밋**

```bash
git add tests/test_xlsx_read_sheet.py skills/excel-wireframe/scripts/xlsx_read.py
git commit -m "feat: sheet-per-screen 레이아웃 읽기"
```

---

### Task 6: table 레이아웃 읽기

**Files:**
- Modify: `skills/excel-wireframe/scripts/xlsx_read.py`
- Test: `tests/test_xlsx_read_table.py`

**Interfaces:**
- Consumes: Task 5의 `read_screens`, `_cell_text`
- Produces: `read_screens`가 `layout == "table"`을 처리. `_read_table(wb, cfg, warns) -> list[dict]` 추가

`table` mapping 형식:

```json
{
  "excel": {
    "layout": "table",
    "sheet": "화면목록",
    "header_row": 3,
    "columns": { "id": "A", "name": "B" },
    "fields": { "설명": "C" },
    "detail": {
      "mode": "grouped-rows",
      "columns": { "no": "D", "element": "E", "desc": "F" }
    }
  }
}
```

`detail.mode`: `grouped-rows`(화면ID가 빈 후속 행은 직전 화면에 귀속) | `none`(상세 없음).

**스펙의 `merged-cells` 모드는 구현하지 않는다.** 병합 셀로 화면 경계를 나누는 양식은
openpyxl로 읽으면 병합 범위의 첫 셀에만 값이 있고 나머지는 `None`이 되므로, `grouped-rows`가
그대로 같은 결과를 낸다. 모드를 하나 더 두면 매핑 작성자가 어느 쪽을 고를지 고민만 늘어난다.
`merged-cells`로 적힌 매핑이 들어와도 `grouped-rows`처럼 동작하게 받아준다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_xlsx_read_table.py`:

```python
from pathlib import Path

from common import Warnings
from fixtures import make_table_xlsx
from openpyxl import load_workbook
from xlsx_read import read_screens

MAPPING = {
    "excel": {
        "layout": "table",
        "sheet": "화면목록",
        "header_row": 3,
        "columns": {"id": "A", "name": "B"},
        "fields": {"설명": "C"},
        "detail": {
            "mode": "grouped-rows",
            "columns": {"no": "D", "element": "E", "desc": "F"},
        },
    }
}


def test_table_layout_groups_continuation_rows(tmp_path: Path):
    xlsx = make_table_xlsx(tmp_path / "t.xlsx")
    wb = load_workbook(xlsx, data_only=True)
    screens = read_screens(wb, MAPPING, Warnings())

    assert [s["id"] for s in screens] == ["SCR001", "SCR002"]
    assert screens[0]["name"] == "이용기관 목록"
    assert len(screens[0]["details"]) == 2
    assert screens[0]["details"][1]["element"] == "[삭제] 버튼"
    assert len(screens[1]["details"]) == 1


def test_table_layout_reads_screen_fields(tmp_path: Path):
    xlsx = make_table_xlsx(tmp_path / "t.xlsx")
    wb = load_workbook(xlsx, data_only=True)
    screens = read_screens(wb, MAPPING, Warnings())
    assert screens[0]["fields"] == {"설명": "기관을 조회한다"}


def test_table_layout_without_details(tmp_path: Path):
    xlsx = make_table_xlsx(tmp_path / "t.xlsx")
    wb = load_workbook(xlsx, data_only=True)
    mapping = {"excel": dict(MAPPING["excel"])}
    mapping["excel"]["detail"] = {"mode": "none"}
    screens = read_screens(wb, mapping, Warnings())
    assert screens[0]["details"] == []
    assert len(screens) == 2
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_xlsx_read_table.py -v`
Expected: FAIL — `ValueError: 지원하지 않는 excel.layout: table`

- [ ] **Step 3: `_read_table` 구현**

`xlsx_read.py`의 `read_screens` 바로 앞에 추가:

```python
def _read_table(wb, cfg: dict, warns: Warnings) -> list[dict]:
    """1행 = 1화면 양식. 화면ID가 빈 행은 직전 화면의 상세 행으로 귀속시킨다."""
    sheet_name = cfg.get("sheet")
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]

    id_col = column_index_from_string(cfg["columns"]["id"])
    name_col = column_index_from_string(cfg["columns"]["name"])
    field_cols = {
        k: column_index_from_string(v) for k, v in (cfg.get("fields") or {}).items()
    }
    detail_cfg = cfg.get("detail") or {"mode": "none"}
    detail_mode = detail_cfg.get("mode", "none")
    detail_cols = {
        k: column_index_from_string(v)
        for k, v in (detail_cfg.get("columns") or {}).items()
    }

    screens: list[dict] = []
    start = int(cfg["header_row"]) + 1
    last = ws.max_row or 0
    blank_streak = 0
    for r in range(start, last + 1):
        sid = _cell_text(ws, r, id_col)
        name = _cell_text(ws, r, name_col)
        detail_vals = {k: _cell_text(ws, r, c) for k, c in detail_cols.items()}
        has_detail = any(detail_vals.values())

        if not sid and not name and not has_detail:
            blank_streak += 1
            if blank_streak >= 3:
                break
            continue
        blank_streak = 0

        if sid:
            screens.append(
                {
                    "id": sid,
                    "name": name or sid,
                    "sheet": ws.title,
                    "images": [],
                    "fields": {
                        k: _cell_text(ws, r, c) for k, c in field_cols.items()
                    },
                    "details": [],
                }
            )
        elif not screens:
            warns.add(None, "orphan-row", "%d행에 화면ID가 없어 건너뜁니다" % r)
            continue

        # merged-cells 양식도 openpyxl로 읽으면 병합 범위의 둘째 행부터 빈 셀이므로
        # grouped-rows와 결과가 같다. 같은 분기로 받는다.
        if detail_mode in ("grouped-rows", "merged-cells") and has_detail:
            screens[-1]["details"].append(detail_vals)

    return screens
```

`read_screens`를 다음으로 교체:

```python
def read_screens(wb, mapping: dict, warns: Warnings) -> list[dict]:
    cfg = mapping["excel"]
    layout = cfg.get("layout", "sheet-per-screen")
    if layout == "sheet-per-screen":
        return _read_sheet_per_screen(wb, cfg, warns)
    if layout == "table":
        return _read_table(wb, cfg, warns)
    raise ValueError("지원하지 않는 excel.layout: %s" % layout)
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_xlsx_read_table.py tests/test_xlsx_read_sheet.py -v`
Expected: PASS (8 passed)

- [ ] **Step 5: 커밋**

```bash
git add tests/test_xlsx_read_table.py skills/excel-wireframe/scripts/xlsx_read.py
git commit -m "feat: table 레이아웃 읽기(grouped-rows 상세 귀속)"
```

---

### Task 7: 삽입 이미지 추출

**Files:**
- Create: `skills/excel-wireframe/scripts/xlsx_images.py`
- Test: `tests/test_xlsx_images.py`

**Interfaces:**
- Consumes: `Warnings`
- Produces:
  - `collect_openpyxl_images(wb) -> list[dict]` — `[{"sheet": str, "row": int, "col": int, "data": bytes, "ext": str}]` (row/col은 0-based 앵커)
  - `collect_zip_images(xlsx_path: Path) -> list[dict]` — 같은 형식. 앵커를 못 구하면 `row`/`col`은 `-1`
  - `extract_images(xlsx_path: Path, wb, mapping: dict, screens: list[dict], out_dir: Path, warns: Warnings) -> None`
    각 화면의 `images` 리스트를 `images/<화면ID>.png` 같은 상대 경로 문자열로 채운다.

귀속 규칙: `sheet-per-screen`이면 시트명으로, `table`이면 앵커 행이 그 화면 행 이상이고 다음 화면 행 미만인 이미지를 귀속시킨다(`table` 레이아웃 화면 dict에 `_row` 키를 `_read_table`이 넣어두지 않으므로, 이미지가 하나뿐이면 첫 화면에 붙이고 여러 개면 순서대로 화면에 1:1 배분한다).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_xlsx_images.py`:

```python
import zipfile
from pathlib import Path

from common import Warnings
from fixtures import make_sheet_per_screen_xlsx
from openpyxl import load_workbook
from xlsx_images import collect_openpyxl_images, collect_zip_images, extract_images

MAPPING = {"excel": {"layout": "sheet-per-screen"}}

SCREENS_SPEC = [
    {"id": "SCR001", "name": "목록", "image": True,
     "details": [{"no": "1", "type": "버튼", "element": "[등록]", "desc": "등록", "pos": "상단"}]},
    {"id": "SCR002", "name": "상세", "image": True,
     "details": [{"no": "1", "type": "버튼", "element": "[저장]", "desc": "저장", "pos": "하단"}]},
]


def _screens():
    return [
        {"id": "SCR001", "name": "목록", "sheet": "설계_SCR001", "images": [], "fields": {}, "details": []},
        {"id": "SCR002", "name": "상세", "sheet": "설계_SCR002", "images": [], "fields": {}, "details": []},
    ]


def test_collect_openpyxl_images_reads_anchors(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS_SPEC)
    wb = load_workbook(xlsx)
    found = collect_openpyxl_images(wb)
    assert len(found) == 2
    assert {f["sheet"] for f in found} == {"설계_SCR001", "설계_SCR002"}
    assert found[0]["row"] == 3  # A4 앵커 = 0-based row 3
    assert found[0]["data"][:4] == b"\x89PNG"


def test_collect_zip_images_finds_media(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS_SPEC)
    found = collect_zip_images(xlsx)
    assert len(found) == 2
    assert all(f["ext"] == "png" for f in found)
    with zipfile.ZipFile(xlsx) as z:
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
    assert len(media) == 2


def test_extract_images_writes_files_and_fills_screens(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS_SPEC)
    wb = load_workbook(xlsx)
    screens = _screens()
    out = tmp_path / "work"
    extract_images(xlsx, wb, MAPPING, screens, out, Warnings())

    assert screens[0]["images"] == ["images/SCR001.png"]
    assert screens[1]["images"] == ["images/SCR002.png"]
    assert (out / "images" / "SCR001.png").exists()
    assert (out / "images" / "SCR001.png").stat().st_size > 0


def test_extract_images_warns_when_screen_has_none(tmp_path: Path):
    spec = [dict(SCREENS_SPEC[0], image=False)]
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", spec)
    wb = load_workbook(xlsx)
    screens = [_screens()[0]]
    warns = Warnings()
    extract_images(xlsx, wb, MAPPING, screens, tmp_path / "work", warns)
    assert screens[0]["images"] == []
    assert [w["code"] for w in warns.to_list()] == ["no-image"]
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_xlsx_images.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xlsx_images'`

- [ ] **Step 3: xlsx_images.py 구현**

```python
# -*- coding: utf-8 -*-
"""Excel에 삽입된 와이어프레임 이미지를 추출한다.

가장 깨지기 쉬운 부분이라 2단 방어로 간다. openpyxl이 놓치는 이미지가
xl/media에 남아 있으면 zip 폴백이 주워 담는다.
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from common import Warnings

RASTER = {"png", "jpg", "jpeg", "gif", "bmp"}
XDR_NS = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"


def _image_bytes(img) -> bytes:
    ref = img.ref
    if isinstance(ref, (str, Path)):
        return Path(ref).read_bytes()
    if isinstance(ref, (bytes, bytearray)):
        return bytes(ref)
    if hasattr(ref, "read"):
        pos = ref.tell() if hasattr(ref, "tell") else None
        try:
            if hasattr(ref, "seek"):
                ref.seek(0)
            return ref.read()
        finally:
            if pos is not None and hasattr(ref, "seek"):
                ref.seek(pos)
    buf = io.BytesIO()
    img.image.save(buf, format=(img.format or "PNG").upper())
    return buf.getvalue()


def collect_openpyxl_images(wb) -> list[dict]:
    out = []
    for ws in wb.worksheets:
        for img in getattr(ws, "_images", []):
            anchor = getattr(img, "anchor", None)
            frm = getattr(anchor, "_from", None)
            row = int(getattr(frm, "row", -1)) if frm is not None else -1
            col = int(getattr(frm, "col", -1)) if frm is not None else -1
            try:
                data = _image_bytes(img)
            except Exception:
                continue
            ext = (getattr(img, "format", None) or "png").lower()
            out.append(
                {"sheet": ws.title, "row": row, "col": col, "data": data, "ext": ext}
            )
    return out


def _drawing_anchor_map(z: zipfile.ZipFile) -> dict[str, tuple[str, int, int]]:
    """xl/drawings/*.xml에서 (미디어 파일명 → 시트 미상, row, col)을 만든다.

    시트 귀속은 drawing → sheet 관계를 거꾸로 타야 해서 비용이 크다. 폴백 경로에서는
    순서 기반 배분으로 충분하므로 앵커만 뽑는다.
    """
    result: dict[str, tuple[str, int, int]] = {}
    for name in z.namelist():
        if not re.match(r"xl/drawings/drawing\d+\.xml$", name):
            continue
        rels_name = "xl/drawings/_rels/%s.rels" % Path(name).name
        rid_to_media: dict[str, str] = {}
        if rels_name in z.namelist():
            rels = ET.fromstring(z.read(rels_name))
            for rel in rels:
                target = rel.get("Target", "")
                rid_to_media[rel.get("Id", "")] = Path(target).name
        root = ET.fromstring(z.read(name))
        for anchor in root:
            frm = anchor.find("%sfrom" % XDR_NS)
            row = col = -1
            if frm is not None:
                row_el = frm.find("%srow" % XDR_NS)
                col_el = frm.find("%scol" % XDR_NS)
                row = int(row_el.text) if row_el is not None else -1
                col = int(col_el.text) if col_el is not None else -1
            for blip in anchor.iter():
                embed = blip.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                )
                if embed and embed in rid_to_media:
                    result[rid_to_media[embed]] = ("", row, col)
    return result


def collect_zip_images(xlsx_path: Path) -> list[dict]:
    out = []
    with zipfile.ZipFile(xlsx_path) as z:
        anchors = _drawing_anchor_map(z)
        media = sorted(n for n in z.namelist() if n.startswith("xl/media/"))
        for name in media:
            base = Path(name).name
            sheet, row, col = anchors.get(base, ("", -1, -1))
            out.append(
                {
                    "sheet": sheet,
                    "row": row,
                    "col": col,
                    "data": z.read(name),
                    "ext": Path(name).suffix.lstrip(".").lower(),
                }
            )
    return out


def _to_png(data: bytes, ext: str, warns: Warnings, screen_id: str) -> tuple[bytes, str]:
    """EMF/WMF는 PPT에서 안 보이는 환경이 있어 PNG 변환을 시도한다."""
    if ext in RASTER:
        return data, "png" if ext == "png" else ext
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as im:
            buf = io.BytesIO()
            im.convert("RGB").save(buf, format="PNG")
            return buf.getvalue(), "png"
    except Exception as exc:
        warns.add(screen_id, "image-convert-failed",
                  "%s 이미지를 PNG로 변환하지 못해 원본을 사용합니다 (%s)" % (ext, exc))
        return data, ext


def _assign(mapping: dict, screens: list[dict], found: list[dict]) -> dict[str, list[dict]]:
    layout = mapping["excel"].get("layout", "sheet-per-screen")
    by_screen: dict[str, list[dict]] = {s["id"]: [] for s in screens}
    if layout == "sheet-per-screen":
        sheet_to_id = {s.get("sheet"): s["id"] for s in screens}
        leftovers = []
        for f in found:
            sid = sheet_to_id.get(f["sheet"])
            if sid:
                by_screen[sid].append(f)
            else:
                leftovers.append(f)
        # 시트명을 못 구한 폴백 결과는 순서대로 이미지 없는 화면에 배분한다
        empty = [s["id"] for s in screens if not by_screen[s["id"]]]
        for sid, f in zip(empty, leftovers):
            by_screen[sid].append(f)
    else:
        ordered = sorted(found, key=lambda f: (f["row"], f["col"]))
        for i, f in enumerate(ordered):
            if i < len(screens):
                by_screen[screens[i]["id"]].append(f)
    return by_screen


def extract_images(
    xlsx_path: Path,
    wb,
    mapping: dict,
    screens: list[dict],
    out_dir: Path,
    warns: Warnings,
) -> None:
    found = collect_openpyxl_images(wb)
    with zipfile.ZipFile(xlsx_path) as z:
        media_count = sum(1 for n in z.namelist() if n.startswith("xl/media/"))
    if media_count > len(found):
        found = collect_zip_images(Path(xlsx_path))

    by_screen = _assign(mapping, screens, found)
    img_dir = Path(out_dir) / "images"

    for scr in screens:
        items = by_screen.get(scr["id"], [])
        if not items:
            warns.add(scr["id"], "no-image", "이 화면에 연결된 이미지를 찾지 못했습니다")
            continue
        for i, f in enumerate(items):
            data, ext = _to_png(f["data"], f["ext"], warns, scr["id"])
            suffix = "" if len(items) == 1 else "-%d" % (i + 1)
            fname = "%s%s.%s" % (scr["id"], suffix, ext)
            img_dir.mkdir(parents=True, exist_ok=True)
            (img_dir / fname).write_bytes(data)
            scr["images"].append("images/%s" % fname)
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_xlsx_images.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: 커밋**

```bash
git add tests/test_xlsx_images.py skills/excel-wireframe/scripts/xlsx_images.py
git commit -m "feat: Excel 삽입 이미지 추출(openpyxl + zip 폴백)"
```

---

### Task 8: extract.py CLI — SSOT 생성과 덮어쓰기 금지

**Files:**
- Create: `skills/excel-wireframe/scripts/extract.py`
- Test: `tests/test_extract.py`

**Interfaces:**
- Consumes: `read_screens`, `extract_images`, `read_json`, `write_json`, `Warnings`, `setup_stdio`
- Produces:
  - `diff_screens(old: dict, new: dict) -> list[str]` — 사람이 읽는 diff 줄 목록
  - `main(argv: list[str] | None = None) -> int`
  - CLI: `python extract.py --excel <xlsx> --mapping <mapping.json> --work <작업폴더>`

`screens.json`이 이미 있으면 `screens.new.json`을 쓰고 diff를 출력한 뒤 종료 코드 `0`을 반환한다(실패가 아니라 사용자 판단 대기 상태이므로).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_extract.py`:

```python
from pathlib import Path

from common import read_json, write_json
from extract import diff_screens, main
from fixtures import make_sheet_per_screen_xlsx

MAPPING = {
    "version": 1,
    "excel": {
        "layout": "sheet-per-screen",
        "sheet_include": "^설계_",
        "screen_meta": {
            "cell": "A1",
            "pattern": r"화면설계서\s*-\s*(?P<id>\S+)\s*\((?P<name>.+)\)",
        },
        "detail": {
            "header_scan_column": "A",
            "header_marker": "No.",
            "columns": {"no": "A", "type": "B", "element": "C", "desc": "D", "pos": "E"},
        },
    },
    "template": {"file": "t.pptx", "mode": "clone", "source_slide": 0},
    "options": {"detail_text_source": "desc"},
}

SPEC = [
    {"id": "SCR001", "name": "이용기관 목록", "image": True,
     "details": [{"no": "1", "type": "버튼", "element": "[등록]", "desc": "등록한다", "pos": "상단"}]},
]


def _setup(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SPEC)
    work = tmp_path / "work"
    mp = work / "mapping.json"
    write_json(mp, MAPPING)
    return xlsx, work, mp


def test_extract_creates_screens_json(tmp_path: Path):
    xlsx, work, mp = _setup(tmp_path)
    code = main(["--excel", str(xlsx), "--mapping", str(mp), "--work", str(work)])
    assert code == 0
    data = read_json(work / "screens.json")
    assert data["screens"][0]["id"] == "SCR001"
    assert data["screens"][0]["details"][0]["desc"] == "등록한다"
    assert data["screens"][0]["images"] == ["images/SCR001.png"]
    assert data["meta"]["source"].endswith("s.xlsx")


def test_extract_does_not_overwrite_existing(tmp_path: Path):
    xlsx, work, mp = _setup(tmp_path)
    main(["--excel", str(xlsx), "--mapping", str(mp), "--work", str(work)])

    edited = read_json(work / "screens.json")
    edited["screens"][0]["name"] = "사람이 고친 이름"
    write_json(work / "screens.json", edited)

    code = main(["--excel", str(xlsx), "--mapping", str(mp), "--work", str(work)])
    assert code == 0
    assert read_json(work / "screens.json")["screens"][0]["name"] == "사람이 고친 이름"
    assert (work / "screens.new.json").exists()
    assert read_json(work / "screens.new.json")["screens"][0]["name"] == "이용기관 목록"


def test_diff_screens_reports_changes():
    old = {"screens": [{"id": "A", "name": "옛 이름", "details": [{"desc": "x"}]}]}
    new = {"screens": [
        {"id": "A", "name": "새 이름", "details": [{"desc": "x"}, {"desc": "y"}]},
        {"id": "B", "name": "추가된 화면", "details": []},
    ]}
    lines = "\n".join(diff_screens(old, new))
    assert "A" in lines and "새 이름" in lines
    assert "B" in lines
    assert "1 -> 2" in lines


def test_diff_screens_empty_when_same():
    same = {"screens": [{"id": "A", "name": "n", "details": []}]}
    assert diff_screens(same, same) == []
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_extract.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'extract'`

- [ ] **Step 3: extract.py 구현**

```python
# -*- coding: utf-8 -*-
"""2단계: mapping.json에 따라 Excel에서 screens.json과 이미지를 뽑는다.

screens.json은 SSOT다. 이미 있으면 절대 덮어쓰지 않는다 — 사람이 손본 내용이
Excel 재추출로 소실되면 안 되기 때문이다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from common import Warnings, read_json, setup_stdio, write_json
from openpyxl import load_workbook
from xlsx_images import extract_images
from xlsx_read import read_screens


def diff_screens(old: dict, new: dict) -> list[str]:
    old_map = {s["id"]: s for s in old.get("screens", [])}
    new_map = {s["id"]: s for s in new.get("screens", [])}
    lines: list[str] = []

    for sid in new_map:
        if sid not in old_map:
            lines.append("+ 화면 추가: %s (%s)" % (sid, new_map[sid].get("name", "")))
    for sid in old_map:
        if sid not in new_map:
            lines.append("- 화면 삭제: %s (%s)" % (sid, old_map[sid].get("name", "")))
    for sid, new_scr in new_map.items():
        old_scr = old_map.get(sid)
        if old_scr is None:
            continue
        if old_scr.get("name") != new_scr.get("name"):
            lines.append(
                "~ %s 화면명: %s -> %s"
                % (sid, old_scr.get("name", ""), new_scr.get("name", ""))
            )
        n_old, n_new = len(old_scr.get("details", [])), len(new_scr.get("details", []))
        if n_old != n_new:
            lines.append("~ %s 상세 건수: %d -> %d" % (sid, n_old, n_new))
        elif old_scr.get("details") != new_scr.get("details"):
            lines.append("~ %s 상세 내용 변경" % sid)
    return lines


def main(argv: list[str] | None = None) -> int:
    setup_stdio()
    ap = argparse.ArgumentParser(description="Excel에서 screens.json과 이미지를 추출한다")
    ap.add_argument("--excel", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--work", required=True)
    args = ap.parse_args(argv)

    excel_path = Path(args.excel)
    work = Path(args.work)
    mapping = read_json(Path(args.mapping))
    warns = Warnings()

    wb = load_workbook(excel_path, data_only=True)
    try:
        screens = read_screens(wb, mapping, warns)
        extract_images(excel_path, wb, mapping, screens, work, warns)
    finally:
        wb.close()

    payload = {
        "meta": {
            "title": "화면설계서",
            "source": str(excel_path),
            "template": mapping.get("template", {}).get("file", ""),
        },
        "screens": screens,
    }

    target = work / "screens.json"
    if target.exists():
        new_path = work / "screens.new.json"
        write_json(new_path, payload)
        lines = diff_screens(read_json(target), payload)
        print("screens.json이 이미 있어 덮어쓰지 않았습니다. -> %s" % new_path)
        if lines:
            print("변경 사항 %d건:" % len(lines))
            for line in lines:
                print("  " + line)
        else:
            print("변경 사항 없음")
    else:
        write_json(target, payload)
        print("screens.json 생성: %s" % target)

    total_details = sum(len(s["details"]) for s in screens)
    with_image = sum(1 for s in screens if s["images"])
    print("화면 %d개, 상세 %d건, 이미지 있는 화면 %d개"
          % (len(screens), total_details, with_image))
    if len(warns):
        print(warns.format())
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_extract.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: 커밋**

```bash
git add tests/test_extract.py skills/excel-wireframe/scripts/extract.py
git commit -m "feat: extract.py — screens.json 생성과 덮어쓰기 금지 + diff"
```

---

### Task 9: 슬라이드 복제

**Files:**
- Create: `skills/excel-wireframe/scripts/slide_clone.py`
- Test: `tests/test_slide_clone.py`

**Interfaces:**
- Consumes: 없음
- Produces: `clone_slide(prs: Presentation, src) -> slide`

이 코드는 스파이크로 이미 검증했다(도형 8개, 그림 183KB, 표 5개 라운드트립 생존). python-pptx의 `_add_relationship`은 rId를 지정할 수 없고 `_next_rId`로 자동 할당하므로, 복제한 XML의 r-네임스페이스 속성을 새 rId로 치환하는 단계가 필수다. 이 단계를 빼면 그림이 사라지거나 파일이 손상된다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_slide_clone.py`:

```python
from pathlib import Path

from fixtures import make_template_pptx
from pptx import Presentation
from slide_clone import clone_slide


def test_clone_copies_all_shapes(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    src = prs.slides[0]
    new = clone_slide(prs, src)
    assert len(new.shapes) == len(src.shapes)
    assert [s.name for s in new.shapes] == [s.name for s in src.shapes]


def test_clone_survives_save_and_reload(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    clone_slide(prs, prs.slides[0])
    out = tmp_path / "out.pptx"
    prs.save(out)

    chk = Presentation(str(out))
    assert len(chk.slides) == 2
    last = chk.slides[-1]
    tables = [s for s in last.shapes if s.has_table]
    assert len(tables) == 5
    assert tables[0].table.cell(0, 0).text == "1"


def test_clone_twice_produces_three_slides(tmp_path: Path):
    pptx = make_template_pptx(tmp_path / "t.pptx")
    prs = Presentation(str(pptx))
    src = prs.slides[0]
    clone_slide(prs, src)
    clone_slide(prs, src)
    out = tmp_path / "out.pptx"
    prs.save(out)
    assert len(Presentation(str(out)).slides) == 3
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_slide_clone.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'slide_clone'`

- [ ] **Step 3: slide_clone.py 구현**

```python
# -*- coding: utf-8 -*-
"""슬라이드 복제. python-pptx에 복제 API가 없어 직접 구현한다.

관계(rels)를 새 슬라이드에 재등록할 때 rId를 지정할 수 없고 자동 할당되므로,
복제한 XML 안의 r:embed / r:id 등을 새 rId로 치환해야 한다. 이 재매핑을 빼면
그림이 사라지거나 파일이 손상된다.
"""
from __future__ import annotations

import copy

from pptx.opc.constants import RELATIONSHIP_TYPE as RT

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_ATTRS = [
    "{%s}%s" % (R_NS, n)
    for n in ("id", "embed", "link", "pict", "dm", "lo", "qs", "cs")
]
SKIP = {RT.SLIDE_LAYOUT, RT.SLIDE_MASTER}


def clone_slide(prs, src):
    """src 슬라이드를 프레젠테이션 끝에 복제하고 새 슬라이드를 반환한다."""
    new = prs.slides.add_slide(src.slide_layout)

    # 레이아웃이 자동 삽입한 플레이스홀더를 걷어낸다. 원본 도형만 남겨야
    # 템플릿 모양이 정확히 재현된다.
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)

    rid_map: dict[str, str] = {}
    for old_rid, rel in src.part.rels.items():
        if rel.reltype in SKIP:
            continue
        if rel.is_external:
            new_rid = new.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_rid = new.part.rels.get_or_add(rel.reltype, rel.target_part)
        rid_map[old_rid] = new_rid

    spTree = new.shapes._spTree
    for shp in src.shapes:
        el = copy.deepcopy(shp._element)
        for node in el.iter():
            for attr in R_ATTRS:
                v = node.get(attr)
                if v is not None and v in rid_map:
                    node.set(attr, rid_map[v])
        spTree.append(el)

    return new
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_slide_clone.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: 실제 템플릿으로 확인**

```bash
python - <<'PY'
import sys
sys.path.insert(0, "skills/excel-wireframe/scripts")
from pptx import Presentation
from slide_clone import clone_slide

prs = Presentation("화면설계서_저작권_발행기관.정산처_발행기관관리_v1.0_20260427.pptx")
new = clone_slide(prs, prs.slides[1])
prs.save("work/clone-check.pptx")
chk = Presentation("work/clone-check.pptx")
last = chk.slides[-1]
pics = [s for s in last.shapes if s.shape_type == 13]
print("shapes=%d pics=%d tables=%d" % (
    len(last.shapes), len(pics), sum(1 for s in last.shapes if s.has_table)))
print("image bytes =", len(pics[0].image.blob))
PY
```

Expected: `shapes=8 pics=1 tables=5`, `image bytes = 183426`

- [ ] **Step 6: 커밋**

```bash
git add tests/test_slide_clone.py skills/excel-wireframe/scripts/slide_clone.py
git commit -m "feat: 슬라이드 복제(XML deepcopy + rId 재매핑)"
```

---

### Task 10: 도형 텍스트 주입과 서식 보존

**Files:**
- Create: `skills/excel-wireframe/scripts/slide_fill.py`
- Test: `tests/test_slide_fill_text.py`

**Interfaces:**
- Consumes: 없음
- Produces:
  - `find_shape(slide, name: str)` → 도형 또는 `None`
  - `set_text(shape, text: str) -> None` — 첫 런의 서식을 유지한 채 텍스트 교체
  - `set_cell_text(cell, text: str) -> None` — 표 셀용. `\n`은 문단으로 분리
  - `estimate_overflow(text: str, cell_width_emu: int, limit_chars: int = 60) -> bool`

서식 보존 방식: 첫 문단 첫 런의 텍스트만 교체하고 나머지 런을 제거한다. 런을 전부 지우고 새로 만들면 폰트·크기·색이 초기화되어 템플릿 디자인이 무너진다. 줄이 여러 개면 첫 런의 `rPr`을 복사해 새 문단에 붙인다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_slide_fill_text.py`:

```python
from pathlib import Path

from fixtures import make_template_pptx
from pptx import Presentation
from pptx.util import Pt
from slide_fill import estimate_overflow, find_shape, set_cell_text, set_text


def test_find_shape_by_name(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    slide = prs.slides[0]
    assert find_shape(slide, "제목 13") is not None
    assert find_shape(slide, "없는도형") is None


def test_set_text_keeps_font_size(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    slide = prs.slides[0]
    shape = find_shape(slide, "제목 13")
    shape.text_frame.paragraphs[0].runs[0].font.size = Pt(18)

    set_text(shape, "이용기관 목록")

    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.text == "이용기관 목록"
    assert run.font.size == Pt(18)
    assert len(shape.text_frame.paragraphs) == 1


def test_set_cell_text_splits_newlines(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    table = next(s for s in prs.slides[0].shapes if s.has_table).table
    cell = table.cell(0, 1)
    cell.text_frame.paragraphs[0].runs[0].font.size = Pt(9)

    set_cell_text(cell, "- 구분값 : 전체, Y, N\n- 디폴트 : 전체")

    paras = cell.text_frame.paragraphs
    assert len(paras) == 2
    assert paras[0].runs[0].text == "- 구분값 : 전체, Y, N"
    assert paras[1].runs[0].text == "- 디폴트 : 전체"
    assert paras[0].runs[0].font.size == Pt(9)
    assert paras[1].runs[0].font.size == Pt(9)


def test_set_cell_text_clears_previous_content(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    table = next(s for s in prs.slides[0].shapes if s.has_table).table
    cell = table.cell(0, 1)
    set_cell_text(cell, "첫 줄\n둘째 줄")
    set_cell_text(cell, "짧게")
    assert cell.text == "짧게"


def test_set_cell_text_accepts_empty(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    table = next(s for s in prs.slides[0].shapes if s.has_table).table
    cell = table.cell(0, 1)
    set_cell_text(cell, "")
    assert cell.text == ""


def test_estimate_overflow():
    assert estimate_overflow("짧은 글", 1974850) is False
    assert estimate_overflow("가" * 400, 1974850) is True
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_slide_fill_text.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'slide_fill'`

- [ ] **Step 3: slide_fill.py 구현 (텍스트 부분)**

```python
# -*- coding: utf-8 -*-
"""슬라이드 도형에 값을 채운다.

핵심은 서식 보존이다. 런을 전부 지우고 새로 만들면 폰트·크기·색이 초기화되어
템플릿 디자인이 무너지므로, 첫 런의 텍스트만 갈아끼우는 방식을 쓴다.
"""
from __future__ import annotations

import copy

EMU_PER_INCH = 914400


def find_shape(slide, name: str):
    for shp in slide.shapes:
        if shp.name == name:
            return shp
    return None


def _fill_text_frame(tf, text: str) -> None:
    lines = str(text).split("\n") if text else [""]
    p0 = tf.paragraphs[0]

    if p0.runs:
        base_run = p0.runs[0]
    else:
        base_run = p0.add_run()
    rPr = base_run._r.find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"
    )
    base_rPr = copy.deepcopy(rPr) if rPr is not None else None

    base_run.text = lines[0]
    for extra in p0.runs[1:]:
        extra._r.getparent().remove(extra._r)
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)

    for line in lines[1:]:
        p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        if base_rPr is not None:
            run._r.insert(0, copy.deepcopy(base_rPr))


def set_text(shape, text: str) -> None:
    if not shape.has_text_frame:
        return
    _fill_text_frame(shape.text_frame, text)


def set_cell_text(cell, text: str) -> None:
    _fill_text_frame(cell.text_frame, text)


def estimate_overflow(text: str, cell_width_emu: int, limit_chars: int = 60) -> bool:
    """셀 폭 대비 글자 수로 잘림 가능성을 추정한다.

    정확한 텍스트 측정은 폰트 메트릭이 필요해 과하다. 자동 축소로 서식을 무너뜨리는
    것보다 사람이 확인하도록 경고만 올리는 편이 낫다.
    """
    if not text:
        return False
    inches = max(cell_width_emu / EMU_PER_INCH, 0.1)
    return len(str(text)) > limit_chars * inches
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_slide_fill_text.py -v`
Expected: PASS (6 passed)

- [ ] **Step 5: 커밋**

```bash
git add tests/test_slide_fill_text.py skills/excel-wireframe/scripts/slide_fill.py
git commit -m "feat: 도형 텍스트 주입(서식 보존, 줄바꿈 문단 분리)"
```

---

### Task 11: 이미지 배치와 표 슬롯 채우기

**Files:**
- Modify: `skills/excel-wireframe/scripts/slide_fill.py`
- Test: `tests/test_slide_fill_slots.py`

**Interfaces:**
- Consumes: Task 10의 `find_shape`, `set_cell_text`, `estimate_overflow`
- Produces:
  - `collect_tables(slide, names: list[str] | None) -> list` — 표 도형 목록. `names`가 있으면 그 순서, 없으면 좌→우(`left`, 같으면 `top`)
  - `count_slots(tables) -> int`
  - `place_image(slide, anchor_shape, image_path: Path)` → 삽입된 그림 도형. 앵커 도형은 제거
  - `fill_slots(tables, details, cols: dict, text_key: str, clear_unused: bool, warns, screen_id: str) -> int` — 채운 슬롯 수 반환

`cols`는 `{"no": 0, "text": 1}` 형태의 열 인덱스다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_slide_fill_slots.py`:

```python
from pathlib import Path

from common import Warnings
from fixtures import make_png, make_template_pptx
from pptx import Presentation
from slide_fill import collect_tables, count_slots, fill_slots, find_shape, place_image

COLS = {"no": 0, "text": 1}


def _details(n: int) -> list[dict]:
    return [{"no": str(i + 1), "desc": "설명 %d" % (i + 1)} for i in range(n)]


def test_collect_tables_by_name(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    tables = collect_tables(prs.slides[0], ["표 8", "표 7"])
    assert [t.name for t in tables] == ["표 8", "표 7"]


def test_collect_tables_left_to_right_when_no_names(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    tables = collect_tables(prs.slides[0], None)
    assert [t.name for t in tables] == ["표 7", "표 8", "표 9", "표 10", "표 11"]
    assert count_slots(tables) == 20


def test_fill_slots_fills_in_order(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    tables = collect_tables(prs.slides[0], None)
    filled = fill_slots(tables, _details(6), COLS, "desc", True, Warnings(), "S1")

    assert filled == 6
    assert tables[0].table.cell(0, 0).text == "1"
    assert tables[0].table.cell(0, 1).text == "설명 1"
    assert tables[1].table.cell(1, 1).text == "설명 6"


def test_fill_slots_clears_unused(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    tables = collect_tables(prs.slides[0], None)
    fill_slots(tables, _details(6), COLS, "desc", True, Warnings(), "S1")
    assert tables[4].table.cell(3, 0).text == ""
    assert tables[4].table.cell(3, 1).text == ""


def test_fill_slots_keeps_unused_when_flag_off(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    tables = collect_tables(prs.slides[0], None)
    fill_slots(tables, _details(2), COLS, "desc", False, Warnings(), "S1")
    assert tables[4].table.cell(3, 1).text == "예시 설명 20"


def test_fill_slots_warns_on_shortage(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    tables = collect_tables(prs.slides[0], None)
    warns = Warnings()
    filled = fill_slots(tables, _details(25), COLS, "desc", True, warns, "S1")
    assert filled == 20
    assert [w["code"] for w in warns.to_list()] == ["slot-shortage"]


def test_fill_slots_uses_original_no(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    tables = collect_tables(prs.slides[0], None)
    details = [{"no": "17", "desc": "열일곱"}, {"no": "18", "desc": "열여덟"}]
    fill_slots(tables, details, COLS, "desc", True, Warnings(), "S1")
    assert tables[0].table.cell(0, 0).text == "17"
    assert tables[0].table.cell(1, 0).text == "18"


def test_place_image_fits_and_centers(tmp_path: Path):
    prs = Presentation(str(make_template_pptx(tmp_path / "t.pptx")))
    slide = prs.slides[0]
    anchor = find_shape(slide, "그림 18")
    left, top, w, h = anchor.left, anchor.top, anchor.width, anchor.height

    png = make_png(tmp_path / "shot.png", size=(1000, 200))
    pic = place_image(slide, anchor, png)

    assert find_shape(slide, "그림 18") is None
    assert pic.width <= w and pic.height <= h
    assert abs((pic.width / pic.height) - 5.0) < 0.05
    assert pic.left >= left and pic.top >= top
    assert abs((pic.left - left) - (left + w - (pic.left + pic.width))) <= 2
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_slide_fill_slots.py -v`
Expected: FAIL — `ImportError: cannot import name 'collect_tables' from 'slide_fill'`

- [ ] **Step 3: slide_fill.py에 추가**

파일 상단 import에 추가:

```python
from pathlib import Path

from PIL import Image
```

파일 끝에 추가:

```python
def collect_tables(slide, names: list[str] | None):
    """상세 표를 슬롯 순서대로 모은다.

    이름이 지정되면 그 순서를 그대로 따른다. 없으면 좌→우로 정렬한다 —
    화면상 왼쪽 표가 앞 번호를 담는 것이 사람의 읽기 순서와 맞기 때문이다.
    """
    tables = [s for s in slide.shapes if s.has_table]
    if names:
        by_name = {t.name: t for t in tables}
        return [by_name[n] for n in names if n in by_name]
    return sorted(tables, key=lambda t: (t.left or 0, t.top or 0))


def count_slots(tables) -> int:
    return sum(len(t.table.rows) for t in tables)


def place_image(slide, anchor_shape, image_path: Path):
    """앵커 도형 자리에 종횡비를 유지해 이미지를 넣고 앵커는 지운다."""
    left = anchor_shape.left
    top = anchor_shape.top
    box_w = anchor_shape.width
    box_h = anchor_shape.height

    with Image.open(image_path) as im:
        iw, ih = im.size

    scale = min(box_w / iw, box_h / ih)
    new_w = int(iw * scale)
    new_h = int(ih * scale)
    new_left = left + (box_w - new_w) // 2
    new_top = top + (box_h - new_h) // 2

    pic = slide.shapes.add_picture(str(image_path), new_left, new_top, new_w, new_h)
    anchor_shape._element.getparent().remove(anchor_shape._element)
    return pic


def fill_slots(
    tables,
    details: list[dict],
    cols: dict,
    text_key: str,
    clear_unused: bool,
    warns,
    screen_id: str,
) -> int:
    """상세 항목을 표1 r0…rN → 표2 r0…rN 순으로 흘려 넣는다."""
    no_col = int(cols.get("no", 0))
    text_col = int(cols.get("text", 1))

    slots = []
    for t in tables:
        table = t.table
        width = 0
        if text_col < len(table.columns):
            width = int(table.columns[text_col].width or 0)
        for r in range(len(table.rows)):
            slots.append((table, r, width))

    filled = 0
    for i, (table, row, width) in enumerate(slots):
        if i < len(details):
            d = details[i]
            text = str(d.get(text_key, "") or "")
            if no_col < len(table.columns):
                set_cell_text(table.cell(row, no_col), str(d.get("no", "") or ""))
            if text_col < len(table.columns):
                set_cell_text(table.cell(row, text_col), text)
            if estimate_overflow(text, width):
                warns.add(screen_id, "text-overflow",
                          "%s번 항목의 설명이 셀 폭을 넘길 수 있습니다" % d.get("no", "?"))
            filled += 1
        elif clear_unused:
            if no_col < len(table.columns):
                set_cell_text(table.cell(row, no_col), "")
            if text_col < len(table.columns):
                set_cell_text(table.cell(row, text_col), "")

    if len(details) > len(slots):
        warns.add(screen_id, "slot-shortage",
                  "상세 %d건 중 %d건만 이 슬라이드에 들어갔습니다"
                  % (len(details), len(slots)))
    return filled
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_slide_fill_slots.py -v`
Expected: PASS (8 passed)

- [ ] **Step 5: 커밋**

```bash
git add tests/test_slide_fill_slots.py skills/excel-wireframe/scripts/slide_fill.py
git commit -m "feat: 이미지 배치와 다중 표 슬롯 채우기"
```

---

### Task 12: build.py — 넘침 분할과 PPT 생성

**Files:**
- Create: `skills/excel-wireframe/scripts/build.py`
- Test: `tests/test_build.py`

**Interfaces:**
- Consumes: `clone_slide`, `collect_tables`, `count_slots`, `fill_slots`, `find_shape`, `place_image`, `set_text`, `Warnings`, `read_json`, `setup_stdio`
- Produces:
  - `chunk_details(details: list[dict], slot_count: int) -> list[list[dict]]`
  - `page_title(name: str, index: int, total: int) -> str`
  - `build(screens_data: dict, mapping: dict, work_dir: Path, out_path: Path, warns: Warnings) -> dict`
    반환: `{"slides": int, "screens": int, "split": [str], "failed": [str]}`
  - `main(argv: list[str] | None = None) -> int`
  - CLI: `python build.py --screens <screens.json> --mapping <mapping.json> --work <작업폴더> --out <출력.pptx>`

생성 절차: 소스 슬라이드를 화면 수(+분할분)만큼 복제 → 각 복제본을 채움 → 마지막에 원본 슬라이드들을 전부 제거. 원본을 먼저 지우면 복제 소스가 사라지므로 순서가 중요하다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_build.py`:

```python
from pathlib import Path

from build import build, chunk_details, page_title
from common import Warnings, write_json
from fixtures import make_png, make_template_pptx
from pptx import Presentation


def _mapping(template: Path) -> dict:
    return {
        "version": 1,
        "excel": {"layout": "sheet-per-screen"},
        "template": {
            "file": str(template),
            "mode": "clone",
            "source_slide": 0,
            "shapes": {
                "title": "제목 13",
                "screen_id": "텍스트 개체 틀 14",
                "image": "그림 18",
                "detail_tables": ["표 7", "표 8", "표 9", "표 10", "표 11"],
            },
            "table_columns": {"no": 0, "text": 1},
        },
        "options": {
            "detail_text_source": "desc",
            "overflow": "split",
            "clear_unused_slots": True,
        },
    }


def _screens(n_details: int, image: str | None = None) -> dict:
    return {
        "meta": {"title": "화면설계서", "source": "s.xlsx", "template": "t.pptx"},
        "screens": [
            {
                "id": "SCR001",
                "name": "이용기관 목록",
                "images": [image] if image else [],
                "fields": {},
                "details": [
                    {"no": str(i + 1), "desc": "설명 %d" % (i + 1)}
                    for i in range(n_details)
                ],
            }
        ],
    }


def test_chunk_details_splits_by_slot_count():
    d = [{"no": str(i)} for i in range(25)]
    chunks = chunk_details(d, 20)
    assert [len(c) for c in chunks] == [20, 5]
    assert chunk_details([], 20) == [[]]
    assert len(chunk_details(d[:20], 20)) == 1


def test_page_title_marks_split():
    assert page_title("목록", 0, 1) == "목록"
    assert page_title("목록", 0, 2) == "목록 (1/2)"
    assert page_title("목록", 1, 2) == "목록 (2/2)"


def test_build_creates_one_slide_per_screen(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    report = build(_screens(6), _mapping(tpl), tmp_path, out, Warnings())

    assert report["slides"] == 1
    prs = Presentation(str(out))
    assert len(prs.slides) == 1
    slide = prs.slides[0]
    title = next(s for s in slide.shapes if s.name == "제목 13")
    assert title.text_frame.text == "이용기관 목록"
    sid = next(s for s in slide.shapes if s.name == "텍스트 개체 틀 14")
    assert sid.text_frame.text == "SCR001"


def test_build_keeps_slide_size(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    build(_screens(3), _mapping(tpl), tmp_path, out, Warnings())
    assert Presentation(str(out)).slide_width == 9906000


def test_build_splits_on_overflow(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "out.pptx"
    warns = Warnings()
    report = build(_screens(25), _mapping(tpl), tmp_path, out, warns)

    assert report["slides"] == 2
    assert report["split"] == ["SCR001"]
    prs = Presentation(str(out))
    titles = [
        next(s for s in sl.shapes if s.name == "제목 13").text_frame.text
        for sl in prs.slides
    ]
    assert titles == ["이용기관 목록 (1/2)", "이용기관 목록 (2/2)"]
    assert "slide-split" in [w["code"] for w in warns.to_list()]


def test_build_places_image(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    make_png(tmp_path / "images" / "SCR001.png", size=(800, 400))
    out = tmp_path / "out.pptx"
    build(_screens(4, "images/SCR001.png"), _mapping(tpl), tmp_path, out, Warnings())

    slide = Presentation(str(out)).slides[0]
    pics = [s for s in slide.shapes if s.shape_type == 13]
    assert len(pics) == 1
    assert not any(s.name == "그림 18" for s in slide.shapes)


def test_build_isolates_screen_failure(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _screens(2)
    data["screens"].append(
        {"id": "BROKEN", "name": "깨진 화면", "images": ["images/없는파일.png"],
         "fields": {}, "details": [{"no": "1", "desc": "x"}]}
    )
    out = tmp_path / "out.pptx"
    warns = Warnings()
    report = build(data, _mapping(tpl), tmp_path, out, warns)

    assert report["screens"] == 2
    assert report["failed"] == ["BROKEN"]
    assert out.exists()
    assert "screen-failed" in [w["code"] for w in warns.to_list()]

    prs = Presentation(str(out))
    assert len(prs.slides) == report["slides"]
    titles = [
        next(s for s in sl.shapes if s.name == "제목 13").text_frame.text
        for sl in prs.slides
    ]
    assert titles == ["이용기관 목록", "[생성 실패] 깨진 화면"]


def test_build_does_not_import_openpyxl():
    import build as build_mod
    import inspect

    src = inspect.getsource(build_mod)
    assert "openpyxl" not in src
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_build.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'build'`

- [ ] **Step 3: build.py 구현**

```python
# -*- coding: utf-8 -*-
"""3단계: screens.json과 템플릿으로 PPT를 만든다.

screens.json이 SSOT이므로 이 모듈은 Excel을 전혀 모른다. mapping.json에서도
template / options 섹션만 읽는다. 덕분에 Excel 픽스처 없이 빌드 로직을 검증할 수 있다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from common import Warnings, read_json, setup_stdio
from pptx import Presentation
from slide_clone import clone_slide
from slide_fill import (
    collect_tables,
    count_slots,
    fill_slots,
    find_shape,
    place_image,
    set_text,
)


def chunk_details(details: list[dict], slot_count: int) -> list[list[dict]]:
    """슬롯 총량을 넘는 상세를 다음 슬라이드 분량으로 쪼갠다."""
    if not details:
        return [[]]
    if slot_count <= 0:
        return [details]
    return [
        details[i : i + slot_count] for i in range(0, len(details), slot_count)
    ]


def page_title(name: str, index: int, total: int) -> str:
    if total <= 1:
        return name
    return "%s (%d/%d)" % (name, index + 1, total)


def _drop_slide(prs, slide) -> None:
    """프레젠테이션에서 슬라이드를 제거한다.

    python-pptx에 삭제 API가 없어 sldIdLst 항목과 관계를 직접 지운다.
    복제가 모두 끝난 뒤에 원본을 지워야 한다 — 먼저 지우면 복제 소스가 사라진다.
    """
    xml_slides = prs.slides._sldIdLst
    for sld_id in list(xml_slides):
        if prs.part.rels[sld_id.rId].target_part is slide.part:
            prs.part.drop_rel(sld_id.rId)
            xml_slides.remove(sld_id)
            return


def _fill_page(slide, scr: dict, page: list[dict], title: str, mapping: dict,
               work_dir: Path, warns: Warnings) -> None:
    tpl = mapping["template"]
    shapes_cfg = tpl.get("shapes", {})
    opts = mapping.get("options", {})
    cols = tpl.get("table_columns", {"no": 0, "text": 1})
    text_key = opts.get("detail_text_source", "desc")
    clear_unused = bool(opts.get("clear_unused_slots", True))

    title_name = shapes_cfg.get("title")
    if title_name:
        shp = find_shape(slide, title_name)
        if shp is None:
            warns.add(scr["id"], "shape-not-found", "제목 도형 '%s' 없음" % title_name)
        else:
            set_text(shp, title)

    sid_name = shapes_cfg.get("screen_id")
    if sid_name:
        shp = find_shape(slide, sid_name)
        if shp is None:
            warns.add(scr["id"], "shape-not-found", "화면ID 도형 '%s' 없음" % sid_name)
        else:
            set_text(shp, scr["id"])

    for key, value in (scr.get("fields") or {}).items():
        name = shapes_cfg.get(key)
        if not name:
            continue
        shp = find_shape(slide, name)
        if shp is not None:
            set_text(shp, value)

    img_name = shapes_cfg.get("image")
    if img_name:
        anchor = find_shape(slide, img_name)
        images = scr.get("images") or []
        if anchor is None:
            warns.add(scr["id"], "shape-not-found", "이미지 자리 '%s' 없음" % img_name)
        elif not images:
            warns.add(scr["id"], "no-image", "배치할 이미지가 없습니다")
        else:
            place_image(slide, anchor, work_dir / images[0])

    tables = collect_tables(slide, shapes_cfg.get("detail_tables"))
    if tables:
        fill_slots(tables, page, cols, text_key, clear_unused, warns, scr["id"])


def build(screens_data: dict, mapping: dict, work_dir: Path, out_path: Path,
          warns: Warnings) -> dict:
    tpl = mapping["template"]
    template_path = Path(tpl["file"])
    if not template_path.is_absolute() and not template_path.exists():
        template_path = Path(work_dir) / tpl["file"]

    prs = Presentation(str(template_path))
    source_index = int(tpl.get("source_slide", 0))
    src = prs.slides[source_index]
    originals = list(prs.slides)

    slot_count = count_slots(collect_tables(src, tpl.get("shapes", {}).get("detail_tables")))

    made = 0
    split_ids: list[str] = []
    failed_ids: list[str] = []
    ok_screens = 0

    for scr in screens_data.get("screens", []):
        try:
            pages = chunk_details(scr.get("details", []), slot_count)
            if len(pages) > 1:
                split_ids.append(scr["id"])
                warns.add(scr["id"], "slide-split",
                          "상세 %d건이 슬롯 %d개를 넘어 %d장으로 나눴습니다"
                          % (len(scr["details"]), slot_count, len(pages)))
            for i, page in enumerate(pages):
                slide = clone_slide(prs, src)
                _fill_page(slide, scr, page,
                           page_title(scr["name"], i, len(pages)),
                           mapping, work_dir, warns)
                made += 1
            ok_screens += 1
        except Exception as exc:
            failed_ids.append(scr["id"])
            warns.add(scr["id"], "screen-failed", "슬라이드 생성 실패: %s" % exc)
            if len(prs.slides) > len(originals) + made:
                # 예외 직전에 만들어진 슬라이드가 남아 있다. 어디가 실패했는지
                # 결과물에서 바로 보이도록 제목만 실패 표시로 바꾼다.
                title_name = tpl.get("shapes", {}).get("title")
                shp = find_shape(prs.slides[-1], title_name) if title_name else None
                if shp is not None:
                    set_text(shp, "[생성 실패] %s" % scr.get("name", scr["id"]))
                made += 1

    for slide in originals:
        _drop_slide(prs, slide)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))

    return {
        "slides": made,
        "screens": ok_screens,
        "split": split_ids,
        "failed": failed_ids,
    }


def main(argv: list[str] | None = None) -> int:
    setup_stdio()
    ap = argparse.ArgumentParser(description="screens.json으로 화면설계서 PPT를 만든다")
    ap.add_argument("--screens", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--work", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)

    screens_data = read_json(Path(args.screens))
    mapping = read_json(Path(args.mapping))
    warns = Warnings()

    report = build(screens_data, mapping, Path(args.work), Path(args.out), warns)

    print("슬라이드 %d장 생성 (화면 %d개)" % (report["slides"], report["screens"]))
    if report["split"]:
        print("분할된 화면: %s" % ", ".join(report["split"]))
    if report["failed"]:
        print("실패한 화면: %s" % ", ".join(report["failed"]))
    if len(warns):
        print(warns.format())
    print("저장: %s" % args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_build.py -v`
Expected: PASS (8 passed)

- [ ] **Step 5: 전체 테스트 확인**

Run: `python -m pytest -q`
Expected: PASS (전체 통과)

- [ ] **Step 6: 커밋**

```bash
git add tests/test_build.py skills/excel-wireframe/scripts/build.py
git commit -m "feat: build.py — PPT 생성, 넘침 분할, 화면별 예외 격리"
```

---

### Task 13: 생성물 검증과 실제 샘플 E2E

**Files:**
- Create: `skills/excel-wireframe/scripts/verify.py`
- Modify: `skills/excel-wireframe/scripts/build.py` (검증 호출 추가)
- Test: `tests/test_verify.py`, `tests/test_sample_e2e.py`

**Interfaces:**
- Consumes: `scan_presentation`
- Produces:
  - `verify_output(out_path: Path, screens_data: dict, mapping: dict, expected_slides: int) -> dict`
    반환: `{"ok": bool, "checks": [{"name": str, "ok": bool, "detail": str}]}`

검사 항목: 슬라이드 수 일치, 각 화면명이 어느 슬라이드 제목에든 들어감, 이미지가 있는 화면의 슬라이드에 그림 도형 존재, 슬라이드 크기가 템플릿과 동일.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_verify.py`:

```python
from pathlib import Path

from build import build
from common import Warnings
from fixtures import make_png, make_template_pptx
from verify import verify_output


def _mapping(template: Path) -> dict:
    return {
        "template": {
            "file": str(template),
            "mode": "clone",
            "source_slide": 0,
            "shapes": {
                "title": "제목 13",
                "screen_id": "텍스트 개체 틀 14",
                "image": "그림 18",
                "detail_tables": ["표 7", "표 8", "표 9", "표 10", "표 11"],
            },
            "table_columns": {"no": 0, "text": 1},
        },
        "options": {"detail_text_source": "desc", "clear_unused_slots": True},
    }


def _data(image: str | None) -> dict:
    return {
        "meta": {"title": "화면설계서"},
        "screens": [
            {"id": "SCR001", "name": "이용기관 목록",
             "images": [image] if image else [], "fields": {},
             "details": [{"no": "1", "desc": "등록한다"}]}
        ],
    }


def test_verify_passes_for_good_output(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    make_png(tmp_path / "images" / "SCR001.png")
    data = _data("images/SCR001.png")
    out = tmp_path / "out.pptx"
    report = build(data, _mapping(tpl), tmp_path, out, Warnings())

    result = verify_output(out, data, _mapping(tpl), report["slides"])
    assert result["ok"] is True
    assert all(c["ok"] for c in result["checks"])


def test_verify_detects_slide_count_mismatch(tmp_path: Path):
    tpl = make_template_pptx(tmp_path / "t.pptx")
    data = _data(None)
    out = tmp_path / "out.pptx"
    build(data, _mapping(tpl), tmp_path, out, Warnings())

    result = verify_output(out, data, _mapping(tpl), 5)
    assert result["ok"] is False
    failed = [c["name"] for c in result["checks"] if not c["ok"]]
    assert "슬라이드 수" in failed
```

`tests/test_sample_e2e.py`:

```python
import subprocess
import sys
from pathlib import Path

import pytest
from common import read_json, write_json
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "skills" / "excel-wireframe" / "scripts"
SAMPLE_XLSX = ROOT / "짧은 버전.xlsx"
SAMPLE_PPTX = ROOT / "화면설계서_저작권_발행기관.정산처_발행기관관리_v1.0_20260427.pptx"

pytestmark = pytest.mark.skipif(
    not (SAMPLE_XLSX.exists() and SAMPLE_PPTX.exists()),
    reason="실제 샘플 파일이 없습니다",
)

MAPPING = {
    "version": 1,
    "excel": {
        "layout": "sheet-per-screen",
        "sheet_include": "^설계_",
        "screen_meta": {
            "cell": "A1",
            "pattern": r"화면설계서\s*-\s*(?P<id>\S+)\s*\((?P<name>.+)\)",
        },
        "detail": {
            "header_scan_column": "A",
            "header_marker": "No.",
            "columns": {"no": "A", "type": "B", "element": "C", "desc": "D", "pos": "E"},
        },
    },
    "template": {
        "file": str(SAMPLE_PPTX),
        "mode": "clone",
        "source_slide": 1,
        "shapes": {
            "title": "제목 13",
            "screen_id": "텍스트 개체 틀 14",
            "image": "그림 18",
            "detail_tables": ["표 7", "표 10", "표 11", "표 12", "표 15"],
        },
        "table_columns": {"no": 0, "text": 1},
    },
    "options": {
        "detail_text_source": "desc",
        "overflow": "split",
        "clear_unused_slots": True,
    },
}


def _run(script: str, *args: str):
    cmd = [sys.executable, str(SCRIPTS / script), *args]
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", cwd=ROOT)


def test_sample_end_to_end(tmp_path: Path):
    work = tmp_path / "work"
    mapping_path = work / "mapping.json"
    write_json(mapping_path, MAPPING)

    r = _run("extract.py", "--excel", str(SAMPLE_XLSX),
             "--mapping", str(mapping_path), "--work", str(work))
    assert r.returncode == 0, r.stderr

    screens = read_json(work / "screens.json")
    assert len(screens["screens"]) == 1
    scr = screens["screens"][0]
    assert scr["id"] == "B2BISMT1001"
    assert scr["name"] == "이용기관 목록"
    assert len(scr["details"]) == 16
    assert scr["images"] == ["images/B2BISMT1001.png"]

    out = work / "output" / "화면설계서.pptx"
    r = _run("build.py", "--screens", str(work / "screens.json"),
             "--mapping", str(mapping_path), "--work", str(work), "--out", str(out))
    assert r.returncode == 0, r.stderr
    assert out.exists()

    prs = Presentation(str(out))
    assert len(prs.slides) == 1
    assert prs.slide_width == 9906000

    slide = prs.slides[0]
    title = next(s for s in slide.shapes if s.name == "제목 13")
    assert title.text_frame.text == "이용기관 목록"

    tables = sorted((s for s in slide.shapes if s.has_table), key=lambda s: s.left)
    assert len(tables) == 5
    assert tables[0].table.cell(0, 0).text == "1"
    # 상세 16건 → 마지막 표 마지막 행(20번 슬롯)은 비어 있어야 한다
    assert tables[4].table.cell(3, 1).text == ""

    pics = [s for s in slide.shapes if s.shape_type == 13]
    assert len(pics) == 1


def test_analyze_on_sample(tmp_path: Path):
    out = tmp_path / "structure-report.json"
    r = _run("analyze.py", "--excel", str(SAMPLE_XLSX),
             "--template", str(SAMPLE_PPTX), "--out", str(out))
    assert r.returncode == 0, r.stderr
    report = read_json(out)
    assert [s["name"] for s in report["excel"]["sheets"]] == [
        "표지", "설계_B2BISMT1001", "테스트_B2BISMT1001", "비교결과요약",
    ]
    assert report["suggestion"]["mode"] == "clone"
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_verify.py tests/test_sample_e2e.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'verify'`

- [ ] **Step 3: verify.py 구현**

```python
# -*- coding: utf-8 -*-
"""생성된 pptx를 다시 파싱해 기대와 대조한다.

만들었다는 사실만으로는 부족하다. 도형 이름이 안 맞거나 rId가 깨지면
파일은 생기지만 내용이 비어 있을 수 있다.
"""
from __future__ import annotations

from pathlib import Path

from pptx_scan import scan_presentation


def verify_output(out_path: Path, screens_data: dict, mapping: dict,
                  expected_slides: int) -> dict:
    report = scan_presentation(Path(out_path))
    slides = report["slides"]
    checks: list[dict] = []

    checks.append(
        {
            "name": "슬라이드 수",
            "ok": len(slides) == expected_slides,
            "detail": "기대 %d장, 실제 %d장" % (expected_slides, len(slides)),
        }
    )

    all_text = "\n".join(
        sh["text"] for s in slides for sh in s["shapes"] if sh["text"]
    )
    missing = [
        scr["name"] for scr in screens_data.get("screens", [])
        if scr["name"] not in all_text
    ]
    checks.append(
        {
            "name": "화면명 반영",
            "ok": not missing,
            "detail": "누락 없음" if not missing else "누락: %s" % ", ".join(missing),
        }
    )

    want_pics = sum(1 for scr in screens_data.get("screens", []) if scr.get("images"))
    got_pics = sum(
        1 for s in slides for sh in s["shapes"] if "PICTURE" in sh["shape_type"]
    )
    checks.append(
        {
            "name": "이미지 배치",
            "ok": got_pics >= want_pics,
            "detail": "이미지 있는 화면 %d개, 슬라이드의 그림 도형 %d개"
            % (want_pics, got_pics),
        }
    )

    template = Path(mapping["template"]["file"])
    if template.exists():
        tpl_report = scan_presentation(template)
        same = (
            tpl_report["slide_width"] == report["slide_width"]
            and tpl_report["slide_height"] == report["slide_height"]
        )
        checks.append(
            {
                "name": "슬라이드 크기",
                "ok": same,
                "detail": "%.2f x %.2f in" % tuple(report["slide_size_in"]),
            }
        )

    return {"ok": all(c["ok"] for c in checks), "checks": checks}
```

- [ ] **Step 4: build.py의 main에 검증 호출 추가**

`build.py`의 import에 추가:

```python
from verify import verify_output
```

`main`에서 `print("저장: %s" % args.out)` 바로 앞에 삽입:

```python
    result = verify_output(Path(args.out), screens_data, mapping, report["slides"])
    print("검증: %s" % ("통과" if result["ok"] else "실패"))
    for c in result["checks"]:
        print("  [%s] %s — %s" % ("O" if c["ok"] else "X", c["name"], c["detail"]))
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `python -m pytest tests/test_verify.py tests/test_sample_e2e.py -v`
Expected: PASS (4 passed)

만약 `test_sample_e2e`에서 상세 건수가 16이 아니거나 이미지가 안 잡히면, 실패한 지점을 `python -m pytest tests/test_sample_e2e.py -v -s`로 출력을 보며 좁힌다. 실제 파일의 구조는 `work/structure-report.json`에 이미 있으니 그것과 대조한다.

- [ ] **Step 6: 전체 테스트와 커밋**

```bash
python -m pytest -q
git add tests/test_verify.py tests/test_sample_e2e.py skills/excel-wireframe/scripts/verify.py skills/excel-wireframe/scripts/build.py
git commit -m "feat: 생성물 검증과 실제 샘플 E2E 회귀 테스트"
```

---

### Task 14: SKILL.md, 레퍼런스, 설치

**Files:**
- Create: `skills/excel-wireframe/SKILL.md`
- Create: `skills/excel-wireframe/references/mapping-schema.md`
- Test: `tests/test_skill_docs.py`

**Interfaces:**
- Consumes: 전 태스크의 CLI 인터페이스
- Produces: 배포 가능한 스킬 디렉토리

- [ ] **Step 1: 문서 존재를 확인하는 테스트 작성**

`tests/test_skill_docs.py`:

```python
from pathlib import Path

import re

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "skills" / "excel-wireframe"


def test_skill_md_has_frontmatter():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    assert text.startswith("---\n")
    fm = text.split("---", 2)[1]
    assert re.search(r"^name:\s*excel-wireframe\s*$", fm, re.M)
    assert re.search(r"^description:\s*\S", fm, re.M)


def test_skill_md_documents_three_scripts():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    for script in ("analyze.py", "extract.py", "build.py"):
        assert script in text


def test_skill_md_covers_missing_template():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    assert "기본 템플릿" in text
    assert "--template" in text


def test_reference_exists():
    ref = SKILL / "references" / "mapping-schema.md"
    text = ref.read_text(encoding="utf-8")
    assert "mapping.json" in text
    assert "screens.json" in text
    assert "detail_tables" in text
    assert "상세표1" in text
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_skill_docs.py -v`
Expected: FAIL — `FileNotFoundError: ...SKILL.md`

- [ ] **Step 3: SKILL.md 작성**

```markdown
---
name: excel-wireframe
description: Excel로 작성된 화면설계서를 PowerPoint 화면설계서로 자동 생성한다. 사용자가 화면설계서, 화면정의서, 스크린 설계, 와이어프레임 문서를 Excel에서 PPT로 만들어 달라고 하거나, xlsx와 pptx 템플릿을 함께 건네며 "이 양식대로 만들어줘"라고 할 때 반드시 사용한다. 처음 보는 Excel 양식이어도 구조를 분석해 매핑을 도출하고, PPT 템플릿을 안 줘도 기본 템플릿을 만들어 진행하므로, 양식이 낯설거나 템플릿이 없다는 이유로 건너뛰지 말 것.
---

# Excel 기반 화면설계서 PPT 생성

Excel 화면설계서를 받아 화면 페이지 PPT를 생성한다. PPT 템플릿을 함께 받으면
그 디자인을 그대로 유지하고, 템플릿이 없으면 표준 구조의 기본 템플릿을 만들어 쓴다.

## 원칙

`screens.json`이 SSOT다. Excel은 최초 임포트 소스일 뿐이고, PPT의 내용은 `screens.json`에서 온다.
사람이 `screens.json`을 손봤을 수 있으므로 재추출이 그것을 덮어쓰지 않는다.

판단은 세 곳에서만 한다. 나머지는 스크립트가 결정론적으로 처리한다.

## 절차

### 0. 준비

작업 디렉토리를 정한다(기본: 입력 Excel 옆의 `work/`). 의존성을 확인한다.

```bash
python -c "import pptx, openpyxl, PIL"
```

없으면 `python -m pip install python-pptx openpyxl Pillow`.

`work/mapping.json`이 이미 있으면 1~2단계를 건너뛰고 3단계로 간다.

### 1. 구조 분석

```bash
python <스킬>/scripts/analyze.py --excel <입력.xlsx> --template <템플릿.pptx> --out work/structure-report.json
```

**템플릿을 안 받았으면 `--template`을 생략한다.** 기본 템플릿(16:9, 제목 바 + 이미지 자리 +
상세 표 5개 × 4행)이 `work/default-template.pptx`로 만들어지고, 그에 맞는 매핑이
리포트의 `suggested_template_mapping`에 담긴다. 템플릿이 있는지 사용자에게 먼저 묻지 말고,
없으면 기본 템플릿으로 진행한 뒤 결과를 보여주며 "원하는 템플릿이 있으면 주시면 그대로
맞춰 드립니다"라고 알린다. 빈손으로 되묻는 것보다 결과물을 보고 판단하는 편이 빠르다.

### 2. [판단①] 매핑 작성과 확인

`work/structure-report.json`을 읽고 `work/mapping.json`을 작성한다.
필드 의미는 `references/mapping-schema.md`를 참고한다.

`template_generated`가 `true`면 `template` 섹션은 `suggested_template_mapping`을
그대로 복사한다. 도형 이름을 우리가 정했으므로 추측할 여지가 없다. Excel 쪽만 판단하면 된다.

판단할 것:

- **Excel 레이아웃** — 시트마다 화면 하나면 `sheet-per-screen`, 한 시트에 화면이 행으로 나열되면 `table`
- **화면이 아닌 시트** — `표지`, `테스트_*`, `요약` 같은 시트를 `sheet_include` 정규식으로 걸러낸다
- **상세 표 헤더 마커** — 헤더 행의 첫 칸 값(보통 `No.`). 헤더 위치는 이미지 크기에 따라 밀리므로 행 번호로 고정하지 말고 마커로 찾게 한다
- **템플릿 도형 이름** — 제목·화면ID·이미지 자리·상세 표. 상세 표가 여러 개면 슬롯 순서대로 나열한다
- **상세 표에 넣을 컬럼** — `detail_text_source`. 보통 상세 설명 컬럼 하나다

작성한 매핑을 **표로 사용자에게 보여주고 확인을 받는다.** 도형 이름을 위치·크기 휴리스틱으로
추측했다면 그 사실을 명시한다. 확인 없이 3단계로 가지 않는다.

### 3. 추출

```bash
python <스킬>/scripts/extract.py --excel <입력.xlsx> --mapping work/mapping.json --work work
```

`work/screens.json`과 `work/images/`가 생긴다. 이미 `screens.json`이 있으면
`screens.new.json`에 쓰고 diff를 출력한다 — 그 경우 diff를 사용자에게 보여주고
반영 여부를 물은 뒤, 반영하기로 하면 파일을 교체한다.

### 4. [판단②] 추출 결과 보고

화면 수, 상세 건수, 이미지가 붙은 화면 수, 경고를 사용자에게 요약한다.
`no-image` 경고가 많으면 이미지 귀속이 잘못됐을 수 있으니 매핑을 재검토한다.

### 5. 생성

```bash
python <스킬>/scripts/build.py --screens work/screens.json --mapping work/mapping.json --work work --out work/output/화면설계서.pptx
```

### 6. [판단③] 검증 결과 보고

`build.py`가 생성물을 재파싱해 검증 결과를 출력한다. 실패 항목이 있으면 원인을 짚어
매핑을 고치고 5단계를 다시 돌린다. 분할된 화면과 잘림 경고는 사용자에게 그대로 전달한다.

## 실패했을 때

| 증상 | 원인과 조치 |
|---|---|
| 상세가 0건 | `header_marker`가 실제 헤더 값과 다르다. `structure-report.json`에서 헤더 셀 값을 확인한다 |
| 이미지가 안 붙음 | `sheet-per-screen`인데 시트명 매칭이 안 되거나, Excel이 이미지를 도형으로 넣었다. `structure-report.json`의 `image_count`를 먼저 본다 |
| 제목이 안 채워짐 | `shapes.title`의 도형 이름이 템플릿과 다르다. 리포트의 도형 이름 목록과 대조한다 |
| 표가 예시 텍스트 그대로 | `detail_tables` 이름이 틀렸다. 이름을 비우면 좌→우 자동 정렬로 대체된다 |
| 한글이 깨져 출력됨 | 콘솔 인코딩 문제다. `PYTHONIOENCODING=utf-8`을 설정하고 재실행한다 |
| 템플릿에 예시 슬라이드가 없다고 나옴 | `analyze.py`가 `layout` 모드로 판정한 경우다. 빈 레이아웃에는 상세를 넣을 표가 없으므로, 사용자에게 예시 페이지가 있는 템플릿을 요청하거나 `--template` 없이 기본 템플릿으로 진행한다 |

## 하지 않는 것

- 와이어프레임을 새로 그리지 않는다. Excel에 삽입된 이미지를 쓴다
- 슬라이드 크기를 바꾸지 않는다. 템플릿 크기를 그대로 승계한다
- 표지·목차를 만들지 않는다. 화면 페이지만 만든다
- 상세 표 번호를 재부여하지 않는다. Excel의 번호가 스크린샷 뱃지와 대응한다
- 템플릿을 달라고 먼저 요구하지 않는다. 없으면 기본 템플릿으로 만들어 보여준 뒤 물어본다
```

- [ ] **Step 4: references/mapping-schema.md 작성**

```markdown
# mapping.json / screens.json 레퍼런스

## mapping.json

양식 해석 결과. 같은 양식을 다시 처리할 때 분석 단계를 건너뛰게 해준다.

### excel

| 필드 | 값 | 설명 |
|---|---|---|
| `layout` | `sheet-per-screen` \| `table` | 1시트=1화면인지, 1행=1화면인지 |
| `sheet_include` | 정규식 | `sheet-per-screen`에서 화면 시트만 고르는 필터. 예: `^설계_` |
| `screen_meta.cell` | 셀 주소 | 화면ID·화면명이 든 셀. 예: `A1` |
| `screen_meta.pattern` | 정규식 | 명명 그룹 `id`, `name`으로 분리. 안 맞으면 시트명이 ID가 된다 |
| `sheet` | 시트명 | `table` 레이아웃에서 화면 목록이 있는 시트 |
| `header_row` | 정수 | `table` 레이아웃의 헤더 행 번호 |
| `columns` | `{"id": "A", "name": "B"}` | `table` 레이아웃의 화면ID·화면명 열 |
| `fields` | `{"설명": "C"}` | `table` 레이아웃의 화면 단위 부가 값. 키는 자유 |
| `detail.header_scan_column` | 열 문자 | 상세 헤더를 찾을 열. 예: `A` |
| `detail.header_marker` | 문자열 | 헤더 행을 식별하는 값. 예: `No.` |
| `detail.mode` | `grouped-rows` \| `merged-cells` \| `none` | `table` 레이아웃 전용. 화면ID가 빈 행을 직전 화면 상세로 귀속. 앞의 두 값은 동작이 같다 |
| `detail.columns` | `{"no": "A", "desc": "D"}` | 상세 열 매핑. 키는 자유이나 `no`는 표 번호 칸에 쓰인다 |

### template

| 필드 | 값 | 설명 |
|---|---|---|
| `file` | 경로 | 복제 소스 pptx. 결과물의 슬라이드 크기는 이 파일을 따른다 |
| `mode` | `clone` \| `layout` | 예시 슬라이드 복제인지, 빈 레이아웃인지 |
| `source_slide` | 정수 | 복제할 슬라이드 인덱스(0-based) |
| `shapes.title` | 도형 이름 | 화면명이 들어갈 자리 |
| `shapes.screen_id` | 도형 이름 | 화면ID가 들어갈 자리 |
| `shapes.image` | 도형 이름 | 스크린샷이 들어갈 자리. 이 도형은 제거되고 그 위치·크기에 이미지가 들어간다 |
| `shapes.detail_tables` | 도형 이름 배열 | **슬롯 순서대로** 나열. 생략하면 좌→우 자동 정렬 |
| `table_columns` | `{"no": 0, "text": 1}` | 표 안의 번호 열·설명 열 인덱스 |

`shapes`에 `fields`의 키와 같은 이름을 넣으면 그 값이 해당 도형에 기록된다.

### 기본 템플릿

사용자가 템플릿을 주지 않으면 `analyze.py`가 `default-template.pptx`를 만든다.
구조는 16:9 슬라이드에 상단 제목 바(`제목`, `화면ID`), 가운데 이미지 자리(`화면이미지`),
하단 상세 표 5개(`상세표1`~`상세표5`, 각 4행 2열 = 20슬롯)다.

이 경우 `structure-report.json`의 `suggested_template_mapping`을 `template` 섹션에
그대로 복사하면 된다. 도형 이름이 고정이라 추측이 필요 없다.

```json
{
  "file": "work/default-template.pptx",
  "mode": "clone",
  "source_slide": 0,
  "shapes": {
    "title": "제목",
    "screen_id": "화면ID",
    "image": "화면이미지",
    "detail_tables": ["상세표1", "상세표2", "상세표3", "상세표4", "상세표5"]
  },
  "table_columns": { "no": 0, "text": 1 }
}
```

`build_default_template(path, slide_width_emu=..., table_count=..., rows_per_table=...)`로
크기와 슬롯 수를 바꿀 수 있다. 상세가 항상 20건을 넘는 양식이라면 `rows_per_table`을
키워 분할을 줄인다.

### options

| 필드 | 기본 | 설명 |
|---|---|---|
| `detail_text_source` | `desc` | 상세 표에 넣을 `details` 키 |
| `overflow` | `split` | 슬롯을 넘치면 다음 슬라이드로 분할하고 제목에 `(1/2)` 표시 |
| `clear_unused_slots` | `true` | 남는 슬롯의 예시 텍스트를 비운다 |

## screens.json (SSOT)

```json
{
  "meta": { "title": "화면설계서", "source": "입력.xlsx", "template": "템플릿.pptx" },
  "screens": [
    {
      "id": "B2BISMT1001",
      "name": "이용기관 목록",
      "sheet": "설계_B2BISMT1001",
      "images": ["images/B2BISMT1001.png"],
      "fields": {},
      "details": [
        { "no": "1", "type": "버튼", "element": "[등록] 버튼",
          "desc": "이용기관을 신규로 등록한다", "pos": "상단 우측" }
      ]
    }
  ]
}
```

`details`의 키는 양식마다 다르다. `build.py`는 `options.detail_text_source`가 가리키는
키와 `no`만 사용하므로, 나머지 키는 사람이 참고하도록 그대로 보존된다.

`images`는 작업 디렉토리 기준 상대 경로다.

## 경고 코드

| 코드 | 의미 |
|---|---|
| `no-image` | 화면에 연결된 이미지를 찾지 못함 |
| `no-detail` | 상세 표 헤더를 찾지 못함 |
| `text-overflow` | 셀 폭 대비 글자 수가 많아 잘릴 수 있음 |
| `shape-not-found` | 템플릿에서 지정한 도형 이름을 찾지 못함 |
| `slide-split` | 상세가 슬롯을 넘쳐 슬라이드를 나눔 |
| `slot-shortage` | 이 슬라이드에 다 못 담음 |
| `screen-failed` | 해당 화면 생성 실패. 나머지 화면은 계속 생성됨 |
| `image-convert-failed` | EMF/WMF를 PNG로 변환하지 못해 원본을 사용 |
| `orphan-row` | `table` 레이아웃에서 화면ID 없이 시작하는 행 |
```

- [ ] **Step 5: 테스트 통과 확인과 전역 설치**

Run: `python -m pytest -q`
Expected: PASS (전체 통과)

```bash
mkdir -p /c/Users/user/.claude/skills
cp -r skills/excel-wireframe /c/Users/user/.claude/skills/
ls /c/Users/user/.claude/skills/excel-wireframe
```

Expected: `SKILL.md`, `references`, `scripts`가 보인다.

- [ ] **Step 6: 커밋**

```bash
git add tests/test_skill_docs.py skills/excel-wireframe/SKILL.md skills/excel-wireframe/references/mapping-schema.md
git commit -m "docs: SKILL.md와 매핑 레퍼런스 작성, 스킬 전역 설치"
```

---

## 검증 매트릭스

스펙 11절의 테스트 조합이 어느 태스크에서 커버되는지:

| 조합 | 커버 위치 |
|---|---|
| `sheet-per-screen` × 예시 슬라이드형 | Task 5, 12, 13 (`test_sample_e2e`) |
| `table` × 예시 슬라이드형 | Task 6 + Task 12의 `build` (screens.json 형식이 동일하므로 빌드 경로는 공유) |
| 이미지 포함 추출 | Task 7 |
| 빈 레이아웃형 템플릿 판정 | Task 3 (`test_suggest_mode_falls_back_to_layout`) |
| 템플릿 미제공 → 기본 템플릿 생성 | Task 3 (`test_default_template.py`), Task 4 (`test_main_without_template_generates_one`) |
| 넘침 분할 | Task 12 (`test_build_splits_on_overflow`) |
| 슬라이드 크기 승계 | Task 12, 13 |
| 화면별 실패 격리 | Task 12 (`test_build_isolates_screen_failure`) |

`layout` 모드(빈 레이아웃에 플레이스홀더 채우기)는 판정까지만 구현하고 생성 경로는 만들지 않는다.
스펙 8절이 `layout` 모드를 언급하지만, 실제 샘플이 전부 `clone`이고 빈 레이아웃 템플릿에서는
표 슬롯 자체가 없어 상세를 넣을 자리가 없다. `analyze.py`가 `layout`으로 판정하면 SKILL.md의
절차상 사용자에게 "이 템플릿에는 예시 슬라이드가 없어 자동 생성이 어렵다"고 알리고,
예시 슬라이드가 있는 템플릿을 요청한다. 이 한계를 SKILL.md 실패 표에 명시할 필요는 없다 —
`analyze.py` 출력의 `reason`이 이미 설명한다.
