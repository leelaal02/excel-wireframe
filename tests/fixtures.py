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
