# -*- coding: utf-8 -*-
"""공용 유틸: UTF-8 stdio, JSON 입출력, 경고 수집."""
from __future__ import annotations

import json
import sys
from pathlib import Path

EMU_PER_INCH = 914400


def resolve_template_path(tpl: dict, work_dir: Path | None = None) -> Path:
    """template.file 경로를 해석한다.

    절대경로거나 cwd 기준으로 이미 존재하면 그대로 쓴다. 아니면 작업 디렉토리
    기준으로 다시 찾는다 — mapping.json에는 work 디렉토리 상대 경로를 쓰는 것이
    자연스럽고, mapping-schema.md의 예시도 그렇게 쓰여 있다. build.py와
    verify.py가 이 함수를 함께 써야 같은 파일을 같은 파일로 판단한다.
    """
    path = Path(tpl["file"])
    if not path.is_absolute() and not path.exists() and work_dir is not None:
        path = Path(work_dir) / tpl["file"]
    return path


def setup_stdio() -> None:
    """Windows cp949 콘솔에서 한글이 깨지지 않도록 UTF-8을 강제한다."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


def read_json(path: Path) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path: Path, data: dict) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


class Warnings:
    """화면 단위 경고 수집기. 최종 리포트에서 한 번에 출력한다."""

    def __init__(self) -> None:
        self._items: list[dict] = []

    def add(self, screen_id: str | None, code: str, message: str) -> None:
        self._items.append(
            {"screen_id": screen_id, "code": code, "message": message}
        )

    def to_list(self) -> list[dict]:
        return list(self._items)

    def format(self) -> str:
        if not self._items:
            return ""
        lines = ["경고 %d건:" % len(self._items)]
        for it in self._items:
            sid = it["screen_id"] or "-"
            lines.append("  [%s] %s: %s" % (sid, it["code"], it["message"]))
        return "\n".join(lines)

    def __len__(self) -> int:
        return len(self._items)
