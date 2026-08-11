# -*- coding: utf-8 -*-
"""공용 유틸: UTF-8 stdio, JSON 입출력, 경고 수집."""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

EMU_PER_INCH = 914400

DEFAULT_OUTPUT_STEM = "화면설계서"

WORK_DIR_NAME = ".work"

# `.work/`로 옮길 구버전 중간 산출물. 이름을 못 박는다 — "pptx가 아닌 것"
# 같은 규칙으로는 default-template.pptx를 걸러낼 수 없고, 사용자가 결과물
# 이름을 직접 지어 놨을 수도 있어 결과물이 딸려 들어간다.
LEGACY_WORK_NAMES = (
    "structure-report.json",
    "mapping.json",
    "screens.json",
    "default-template.pptx",
    "images",
)


def resolve_template_path(tpl: dict, work_dir: Path | None = None) -> Path:
    """template.file 경로를 해석한다.

    절대경로거나 cwd 기준으로 이미 존재하면 그대로 쓴다. 아니면 작업 폴더
    기준으로 다시 찾는다 — 만들어 쓰는 기본 템플릿이 거기 있고,
    mapping-schema.md의 예시도 그 상대 경로로 쓰여 있다. build.py와
    verify.py가 이 함수를 함께 써야 같은 파일을 같은 파일로 판단한다.
    """
    path = Path(tpl["file"])
    if not path.is_absolute() and not path.exists() and work_dir is not None:
        path = Path(work_dir) / tpl["file"]
    return path


def output_stem(source_excel: str) -> str:
    """결과물 파일명의 기본 stem. 원본 Excel 파일명을 그대로 쓴다.

    `meta.source`가 비었으면 기본 이름으로 떨어진다 — 손으로 만든 screens.json이나
    meta.source가 없던 시절의 파일에도 이름을 내줘야 한다. 여기서 죽으면 생성
    단계 전체가 멈춘다.
    """
    stem = Path(source_excel).stem if source_excel else ""
    return stem or DEFAULT_OUTPUT_STEM


def resolve_output_path(output_dir: Path, source_excel: str) -> Path:
    """결과물 pptx 경로를 정한다.

    이름은 원본 Excel 파일명이다 — 어느 입력에서 나온 결과물인지 파일명만 보고
    알 수 있어야 한다. 같은 이름이 이미 있으면 뒤에 2, 3…을 붙여 비어 있는 첫
    번호를 쓴다. 재실행이 직전 산출물을 덮어쓰면 결과를 비교할 수단이 없어진다.

    디렉토리는 만들지 않는다 — 경로를 정하는 일과 만드는 일을 섞지 않는다.
    저장 직전 mkdir은 build()가 한다.
    """
    output_dir = Path(output_dir)
    stem = output_stem(source_excel)
    candidate = output_dir / (stem + ".pptx")
    n = 2
    while candidate.exists():
        candidate = output_dir / ("%s%d.pptx" % (stem, n))
        n += 1
    return candidate


def work_dir(output_dir: Path) -> Path:
    """중간 산출물을 두는 작업 폴더. `output/.work`.

    결과물 폴더에는 pptx만 보이게 하되, 작업 폴더를 그 **안**에 둬서 "이 Excel의
    산출물은 여기 전부"가 그대로 성립하게 한다 — 폴더째 지우면 전부 사라진다.

    디렉토리는 만들지 않는다. 경로를 정하는 일과 만드는 일을 섞지 않는다.
    """
    return Path(output_dir) / WORK_DIR_NAME


def migrate_legacy_work(output_dir: Path,
                        keep: list[Path] | None = None) -> list[str]:
    """`output/` 루트에 남은 구버전 중간 산출물을 `.work/`로 옮긴다.

    옮긴 이름 목록을 돌려준다. 옮길 게 없으면 빈 리스트이고, 이때 `.work/`를
    만들지도 않는다 — 아무 일도 안 했으면 흔적도 남기지 않는다.

    `.work/`에 같은 이름이 이미 있으면 옮기지 않는다. 그쪽이 새 경로에서 만들어진
    최신 파일이고, 구버전이 덮어쓰면 방금 한 작업이 사라진다.

    `keep`은 CLI가 `--mapping` 같은 인자로 직접 가리킨 경로다. 그 파일을 옮기면
    바로 그 실행이 자기가 읽을 파일을 잃는다. 사용자가 짚은 경로는 건드리지 않는다.
    """
    output_dir = Path(output_dir)
    target_dir = work_dir(output_dir)
    kept = {Path(p).resolve() for p in keep or []}
    moved: list[str] = []
    for name in LEGACY_WORK_NAMES:
        src = output_dir / name
        if not src.exists():
            continue
        if src.resolve() in kept:
            continue
        dst = target_dir / name
        if dst.exists():
            continue
        target_dir.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(dst))
        moved.append(name)
    return moved


def migration_notice(moved: list[str]) -> str:
    """마이그레이션 결과를 사람에게 알릴 한 줄. 옮긴 게 없으면 빈 문자열이다.

    조용히 파일을 움직이면 사용자가 사라졌다고 오해한다. 세 CLI가 같은 문구를
    쓰도록 여기 둔다 — `Warnings.format`과 같은 규약이다.
    """
    if not moved:
        return ""
    return ("기존 작업 파일 %d개를 %s/로 옮겼습니다: %s"
            % (len(moved), WORK_DIR_NAME, ", ".join(moved)))


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
