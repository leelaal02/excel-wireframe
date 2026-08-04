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


def diff_screens(old: dict, new: dict) -> list[str]:
    old_map = {s["id"]: s for s in old.get("screens", [])}
    new_map = {s["id"]: s for s in new.get("screens", [])}
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
    if len(warns):
        print(warns.format())
    return 0


if __name__ == "__main__":
    sys.exit(main())
