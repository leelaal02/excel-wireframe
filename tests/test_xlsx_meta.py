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
