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
