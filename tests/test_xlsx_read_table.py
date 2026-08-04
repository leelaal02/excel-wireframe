# -*- coding: utf-8 -*-
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
