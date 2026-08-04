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
