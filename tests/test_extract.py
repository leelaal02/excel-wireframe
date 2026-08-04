from pathlib import Path

from common import read_json, write_json
from extract import _screen_key, diff_screens, main
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


def test_screen_key_fallback_label_matches_build_py_numbering():
    """회귀 5: build.py는 screen_count를 화면마다 증가시킨 뒤(1부터 시작) id
    없는 화면의 폴백 라벨을 만든다. extract.py의 _screen_key는 enumerate()의
    0-based 인덱스를 그대로 썼어서 같은 첫 화면이 build.py에서는
    '(id 없음 #1)', extract.py에서는 '(id 없음 #0)'으로 서로 다르게
    불렸다 — 두 단계 출력을 사람이 대조할 수 없었다."""
    scr = {"name": "id 없는 화면"}

    # build.py 쪽 공식 재현: for 루프에서 screen_count를 증가시킨 뒤(1-based)
    # scr.get("id") or "(id 없음 #%d)" % screen_count 로 계산한다.
    screen_count_for_first_screen = 1
    build_style_label = scr.get("id") or "(id 없음 #%d)" % screen_count_for_first_screen

    assert _screen_key(scr, 0) == build_style_label == "(id 없음 #1)"


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

    # 픽스처에는 화면 시트가 아닌 "테스트_무시대상" 시트도 섞여 있어, 자동
    # 탐지(첫 비-화면 시트)에 맡기면 표지 대신 그 시트를 표지로 오인해버려
    # "표지 없음" 경로를 실제로 타지 못한다. find_cover_sheet는 명시 지정된
    # 시트가 없을 때만 None을 반환하므로(tests/test_xlsx_meta.py의
    # test_find_cover_sheet_returns_none_when_named_sheet_missing 참고),
    # 표지를 명시 지정해 방금 지운 시트를 가리키게 한다.
    mapping = read_json(mp)
    mapping["excel"]["cover"] = {"sheet": "표지"}
    write_json(mp, mapping)

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
