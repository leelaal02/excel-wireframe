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


def test_diff_screens_reports_deletion():
    """추가/변경 분기는 이미 돈다. 삭제 분기(옛 화면이 새 추출 결과에서
    사라진 경우)는 지금까지 어떤 테스트도 실행하지 않았다."""
    old = {"screens": [
        {"id": "A", "name": "남는 화면", "details": []},
        {"id": "B", "name": "삭제될 화면", "details": []},
    ]}
    new = {"screens": [{"id": "A", "name": "남는 화면", "details": []}]}

    lines = diff_screens(old, new)
    assert "- 화면 삭제: B (삭제될 화면)" in lines
    assert not any(line.startswith("+") for line in lines)
    assert not any(line.startswith("~") for line in lines)


def test_extract_warns_and_prints_on_zero_screens(tmp_path: Path):
    """sheet_include가 아무 시트에도 안 맞아 화면이 0개 추출되면, 다음 단계인
    build.py의 검증이 잡아주는 것과는 별개로 이 단계에서 바로 원인을 알려줘야
    한다 — 여기가 사람이 sheet_include를 실제로 고칠 수 있는 지점이다."""
    xlsx, work, mp = _setup(tmp_path)
    mapping = read_json(mp)
    mapping["excel"]["sheet_include"] = "^존재하지않는패턴_"
    write_json(mp, mapping)

    import io
    import contextlib

    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        code = main(["--excel", str(xlsx), "--mapping", str(mp), "--work", str(work)])
    assert code == 0

    out = buf.getvalue()
    assert "화면 0개" in out
    assert "sheet_include" in out
    data = read_json(work / "screens.json")
    assert data["screens"] == []
