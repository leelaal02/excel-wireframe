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
    """generated=True일 때 실제로 쓰이는 제안은 `suggested_template_mapping`
    이다(build_default_template의 도형 이름을 이미 알고 있으므로 항상
    clone/source_slide=0로 채워진다) — `suggestion`(범용 휴리스틱)이 아니다.

    새 기본 템플릿은 채워질 도형의 글자를 일부러 비워 둔다(빈 Excel에서
    자리표시 문구가 산출물에 그대로 찍히는 것을 막기 위해). 그래서
    `suggest_mode`의 "글자가 채워진 도형 3개 이상" 휴리스틱은 더는 우리
    기본 템플릿을 example 슬라이드로 보지 않고 layout으로 본다 — 실제
    파이프라인 동작에는 영향이 없다(`suggested_template_mapping`이 이미
    clone으로 확정돼 있으므로), 판정 결과만 정직하게 반영한다."""
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    out = tmp_path / "work" / "structure-report.json"
    code = main(["--excel", str(xlsx), "--out", str(out)])
    assert code == 0

    data = read_json(out)
    assert data["template_generated"] is True
    assert (tmp_path / "work" / "default-template.pptx").exists()
    assert data["suggestion"]["mode"] == "layout"

    suggested = data["suggested_template_mapping"]
    assert suggested["shapes"]["title"] == "제목"
    assert suggested["shapes"]["detail_tables"][0] == "상세표1"
    assert data["template"]["slide_width"] == 9906000


def test_main_with_given_template_has_no_suggestion(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    pptx = make_template_pptx(tmp_path / "t.pptx")
    out = tmp_path / "work" / "structure-report.json"
    main(["--excel", str(xlsx), "--template", str(pptx), "--out", str(out)])
    assert "suggested_template_mapping" not in read_json(out)


def test_suggested_mapping_carries_meta_shapes(tmp_path: Path):
    xlsx = make_sheet_per_screen_xlsx(tmp_path / "s.xlsx", SCREENS)
    out = tmp_path / "work" / "structure-report.json"
    main(["--excel", str(xlsx), "--out", str(out)])
    shapes = read_json(out)["suggested_template_mapping"]["shapes"]
    assert shapes["문서제목"] == "문서제목"
    assert shapes["작성일"] == "작성일"
