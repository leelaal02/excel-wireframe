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
