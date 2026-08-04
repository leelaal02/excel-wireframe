# -*- coding: utf-8 -*-
"""Excel 표지 시트에서 문서 단위 정보를 읽는다.

화면 페이지는 화면마다 값이 다르지만, 프로젝트명·작성일 같은 값은 문서 전체에
하나뿐이다. 표지는 그 값들이 모여 있는 유일한 자리다.
"""
from __future__ import annotations

import datetime
import re

RESERVED = {"source", "template"}
MAX_SCAN_ROW = 40
MAX_SCAN_COL = 12
MAX_STANDALONE = 2


def _text(ws, row: int, col: int) -> str:
    v = ws.cell(row=row, column=col).value
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def find_cover_sheet(wb, mapping: dict) -> str | None:
    """표지 시트를 고른다.

    명시 지정이 있으면 그것을 쓴다. 없으면 화면 시트로 쓰이지 않은 첫 시트다 —
    표지는 보통 맨 앞에 있고, 화면 시트는 sheet_include로 이미 식별된다.
    """
    cfg = mapping.get("excel", {})
    named = (cfg.get("cover") or {}).get("sheet")
    if named:
        return named if named in wb.sheetnames else None

    include = cfg.get("sheet_include")
    if not include:
        return wb.sheetnames[0] if wb.sheetnames else None
    rx = re.compile(include)
    for name in wb.sheetnames:
        if not rx.search(name):
            return name
    return None


def read_cover_meta(wb, mapping: dict, warns) -> dict[str, str]:
    """표지에서 라벨-값 쌍과 단독 셀을 뽑아 문서 정보 사전을 만든다."""
    sheet = find_cover_sheet(wb, mapping)
    if sheet is None:
        return {}
    ws = wb[sheet]

    meta: dict[str, str] = {}
    standalone: list[str] = []
    last_row = min(ws.max_row or 0, MAX_SCAN_ROW)
    last_col = min(ws.max_column or 0, MAX_SCAN_COL)

    for r in range(1, last_row + 1):
        c = 1
        while c <= last_col:
            label = _text(ws, r, c)
            if not label:
                c += 1
                continue
            # 같은 행 오른쪽에서 첫 값을 찾는다. 병합 셀은 두 번째 칸부터 비어
            # 있으므로 자연히 건너뛰어진다.
            value = ""
            vc = c + 1
            while vc <= last_col:
                value = _text(ws, r, vc)
                if value:
                    break
                vc += 1
            if value:
                if label not in RESERVED:
                    meta[label] = value
                c = vc + 1
            else:
                if len(standalone) < MAX_STANDALONE:
                    standalone.append(label)
                break
    if standalone:
        meta.setdefault("문서제목", standalone[0])
    if len(standalone) > 1:
        meta.setdefault("부제", standalone[1])

    overrides = mapping.get("excel", {}).get("meta_overrides") or {}
    for k, v in overrides.items():
        if k not in RESERVED:
            meta[k] = "" if v is None else str(v)
    return meta
