# -*- coding: utf-8 -*-
"""2단계: mapping.json에 따라 Excel에서 screens.json과 이미지를 뽑는다.

screens.json은 SSOT다. 이미 있으면 절대 덮어쓰지 않는다 — 사람이 손본 내용이
Excel 재추출로 소실되면 안 되기 때문이다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from common import Warnings, read_json, setup_stdio, write_json
from openpyxl import load_workbook
from xlsx_images import extract_images
from xlsx_read import read_screens


def _screen_key(scr: dict, index: int) -> str:
    """screens.json은 사람이 손으로 편집한다 — id가 빠진 화면 dict가 섞여도
    KeyError로 diff 전체를 무너뜨리지 않도록 인덱스 기반 대체 키를 쓴다.

    index는 enumerate()의 0-based 값이므로 1을 더해 build.py의 표기(스크린
    카운터를 증가시킨 뒤 쓰므로 1-based)와 맞춘다 — 같은 화면이 두 단계
    출력에서 다른 번호로 불리면 사람이 대조하기 어렵다.
    """
    return scr.get("id") or "(id 없음 #%d)" % (index + 1)


def diff_screens(old: dict, new: dict) -> list[str]:
    old_map = {_screen_key(s, i): s for i, s in enumerate(old.get("screens", []))}
    new_map = {_screen_key(s, i): s for i, s in enumerate(new.get("screens", []))}
    lines: list[str] = []

    for sid in new_map:
        if sid not in old_map:
            lines.append("+ 화면 추가: %s (%s)" % (sid, new_map[sid].get("name", "")))
    for sid in old_map:
        if sid not in new_map:
            lines.append("- 화면 삭제: %s (%s)" % (sid, old_map[sid].get("name", "")))
    for sid, new_scr in new_map.items():
        old_scr = old_map.get(sid)
        if old_scr is None:
            continue
        if old_scr.get("name") != new_scr.get("name"):
            lines.append(
                "~ %s 화면명: %s -> %s"
                % (sid, old_scr.get("name", ""), new_scr.get("name", ""))
            )
        n_old, n_new = len(old_scr.get("details", [])), len(new_scr.get("details", []))
        if n_old != n_new:
            lines.append("~ %s 상세 건수: %d -> %d" % (sid, n_old, n_new))
        elif old_scr.get("details") != new_scr.get("details"):
            lines.append("~ %s 상세 내용 변경" % sid)
    return lines


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
        extract_images(excel_path, wb, mapping, screens, work, warns)
    finally:
        wb.close()

    payload = {
        "meta": {
            "title": "화면설계서",
            "source": str(excel_path),
            "template": mapping.get("template", {}).get("file", ""),
        },
        "screens": screens,
    }

    target = work / "screens.json"
    if target.exists():
        new_path = work / "screens.new.json"
        write_json(new_path, payload)
        lines = diff_screens(read_json(target), payload)
        print("screens.json이 이미 있어 덮어쓰지 않았습니다. -> %s" % new_path)
        if lines:
            print("변경 사항 %d건:" % len(lines))
            for line in lines:
                print("  " + line)
        else:
            print("변경 사항 없음")
    else:
        write_json(target, payload)
        print("screens.json 생성: %s" % target)

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
