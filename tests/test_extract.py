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
