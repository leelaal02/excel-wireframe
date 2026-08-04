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
MIN_COVER_PAIRS = 2
COVER_NAME_HINTS = ("표지", "표제", "개요", "cover", "front")


def _text(ws, row: int, col: int) -> str:
    v = ws.cell(row=row, column=col).value
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def _scan_cover_cells(ws) -> tuple[list[tuple[str, str]], list[str]]:
    """시트를 훑어 라벨-값 쌍과 단독 셀 라벨을 뽑는다.

    find_cover_sheet의 후보 검증(쌍이 몇 개인지)과 read_cover_meta의 실제
    파싱이 같은 스캔 규칙을 쓰도록 여기 하나로 모은다 — 규칙이 갈라지면
    "표지로 채택됐는데 정작 못 읽는다" 류의 불일치가 생긴다.
    """
    pairs: list[tuple[str, str]] = []
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
                pairs.append((label, value))
                c = vc + 1
            else:
                if len(standalone) < MAX_STANDALONE:
                    standalone.append(label)
                break
    return pairs, standalone


def _has_cover_name_hint(name: str) -> bool:
    lower = name.lower()
    return any(hint in lower for hint in COVER_NAME_HINTS)


def find_cover_sheet(wb, mapping: dict) -> str | None:
    """표지 시트를 고른다.

    1. 명시 지정(``excel.cover.sheet``)이 있으면 그것을 쓴다(없는 이름이면 None).
    2. 시트 이름에 표지를 가리키는 힌트(표지/표제/개요/cover/front)가 있으면
       그 첫 시트를 쓴다 — 실제 샘플의 '표지' 시트가 여기서 바로 잡힌다.
    3. 힌트가 없으면 화면 시트로 쓰이지 않은 첫 시트(``sheet_include`` 미설정
       이면 첫 시트)를 후보로 삼되, 그 시트에서 라벨-값 쌍이 2개 이상 나올
       때만 표지로 채택한다. 표지가 아예 없는 워크북은 화면이 아닌 시트도
       한둘 섞여 있을 수 있는데(부록, 비교표 등), 쌍이 1개 이하면 그런
       시트를 표지로 오인한 것으로 보고 None을 반환한다.
    """
    cfg = mapping.get("excel", {})
    named = (cfg.get("cover") or {}).get("sheet")
    if named:
        return named if named in wb.sheetnames else None

    for name in wb.sheetnames:
        if _has_cover_name_hint(name):
            return name

    include = cfg.get("sheet_include")
    candidate: str | None = None
    if not include:
        candidate = wb.sheetnames[0] if wb.sheetnames else None
    else:
        rx = re.compile(include)
        for name in wb.sheetnames:
            if not rx.search(name):
                candidate = name
                break

    if candidate is None:
        return None
    pairs, _ = _scan_cover_cells(wb[candidate])
    return candidate if len(pairs) >= MIN_COVER_PAIRS else None


def read_cover_meta(wb, mapping: dict, warns) -> dict[str, str]:
    """표지에서 라벨-값 쌍과 단독 셀을 뽑아 문서 정보 사전을 만든다."""
    sheet = find_cover_sheet(wb, mapping)
    if sheet is None:
        return {}
    ws = wb[sheet]

    meta: dict[str, str] = {}
    pairs, standalone = _scan_cover_cells(ws)
    for label, value in pairs:
        if label not in RESERVED:
            meta[label] = value
    if standalone:
        meta.setdefault("문서제목", standalone[0])
    if len(standalone) > 1:
        meta.setdefault("부제", standalone[1])

    overrides = mapping.get("excel", {}).get("meta_overrides") or {}
    for k, v in overrides.items():
        if k not in RESERVED:
            meta[k] = "" if v is None else str(v)
    return meta
