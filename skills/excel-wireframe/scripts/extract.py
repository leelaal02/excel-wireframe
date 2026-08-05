# -*- coding: utf-8 -*-
"""2단계: mapping.json에 따라 Excel에서 screens.json과 이미지를 뽑는다.

screens.json은 중간 산출물이다. 매핑이나 템플릿만 고쳐 생성 단계를 반복할 때
추출을 다시 하지 않아도 되게 하고, 결과가 이상할 때 추출과 배치 중 어느 쪽이
틀렸는지 갈라 보게 한다. 사람이 손으로 편집하는 파일이 아니므로 재추출은
그냥 덮어쓴다 — 추출 결과가 어긋나면 이 파일이 아니라 매핑을 고친다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from common import Warnings, read_json, setup_stdio, write_json
from openpyxl import load_workbook
from xlsx_images import extract_images
from xlsx_meta import read_cover_meta
from xlsx_read import read_screens


def main(argv: list[str] | None = None) -> int:
    setup_stdio()
    ap = argparse.ArgumentParser(description="Excel에서 screens.json과 이미지를 추출한다")
    ap.add_argument("--excel", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--work", required=True)
    args = ap.parse_args(argv)

    excel_path = Path(args.excel)
    work = Path(args.work)
    mapping = read_json(Path(args.mapping))
    warns = Warnings()

    wb = load_workbook(excel_path, data_only=True)
    try:
        screens = read_screens(wb, mapping, warns)
        try:
            cover_meta = read_cover_meta(wb, mapping, warns)
        except Exception:
            # mapping.json의 cover/meta_overrides 모양이 잘못돼도(예: cover가
            # 문자열, meta_overrides가 리스트) 표지 파싱 실패가 화면 페이지
            # 생성 전체를 무너뜨리면 안 된다 — 표지 없음과 같은 취급으로
            # 내려가 아래 안내 문구가 그대로 출력되게 한다.
            cover_meta = {}
        extract_images(excel_path, wb, mapping, screens, work, warns)
    finally:
        wb.close()

    if not cover_meta:
        print("표지 시트를 찾지 못해 문서 정보를 비웠습니다 "
              "(mapping.excel.cover.sheet로 지정할 수 있습니다)")

    payload = {
        "meta": {
            **cover_meta,
            "source": str(excel_path),
            "template": mapping.get("template", {}).get("file", ""),
        },
        "screens": screens,
    }

    target = work / "screens.json"
    write_json(target, payload)
    print("screens.json 저장: %s" % target)

    total_details = sum(len(s["details"]) for s in screens)
    with_image = sum(1 for s in screens if s["images"])
    print("화면 %d개, 상세 %d건, 이미지 있는 화면 %d개"
          % (len(screens), total_details, with_image))
    if not screens:
        # 화면이 통째로 0개인 가장 흔한 원인은 sheet_include 정규식이 실제 시트명과
        # 안 맞는 매핑 오타다. build.py는 이 상태에서도 0개짜리 결과물을 "성공"으로
        # 보고할 수 있으므로(검증 단계에서 별도로 잡는다), 이 단계에서 바로 원인을
        # 짚어 사람이 다음 실행 전에 알아챌 수 있게 한다.
        print("경고: 화면을 0개 추출했습니다 — mapping.json의 excel.sheet_include"
              "(sheet-per-screen) 또는 excel 레이아웃 설정을 확인하세요")
    if len(warns):
        print(warns.format())
    return 0


if __name__ == "__main__":
    sys.exit(main())
