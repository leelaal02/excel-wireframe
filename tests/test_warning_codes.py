# -*- coding: utf-8 -*-
"""경고 코드는 아홉 개로 닫힌 집합이다(CLAUDE.md). 새 코드를 추가하고 싶은
유혹이 스크립트 곳곳에 생길 수 있으므로, warns.add(...)가 실제로 넘기는 코드
리터럴을 소스에서 뽑아 그 집합 밖으로 나가지 않는지를 기계적으로 검사한다.

Warnings.add 안에서 검증하지 않는다 — 거기서 raise하면 부분 성공(경고만 남기고
나머지 화면은 계속 만든다)이 크래시로 뒤집힌다. 이 프로젝트의 핵심 원칙이다.
"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "skills" / "excel-to-wireframe-ppt" / "scripts"

ALLOWED_CODES = {
    "no-image",
    "no-detail",
    "text-overflow",
    "shape-not-found",
    "slide-split",
    "slot-shortage",
    "screen-failed",
    "image-convert-failed",
    "orphan-row",
}

# warns.add(screen_id, "code", message) — 코드는 항상 두 번째 위치 인자다.
# 호출이 여러 줄에 걸쳐 있어도(예: xlsx_images.py) 잡아내야 하므로 첫 인자와
# 코드 문자열 사이의 공백/개행은 자유롭게 허용한다.
CALL_PATTERN = re.compile(r'warns\.add\(\s*[^,]*,\s*"([^"]+)"')


def _script_files():
    return sorted(p for p in SCRIPTS.glob("*.py") if p.name != "__init__.py")


def test_every_warns_add_code_is_in_the_closed_set():
    offenders = []
    seen_any = False
    for path in _script_files():
        text = path.read_text(encoding="utf-8")
        for m in CALL_PATTERN.finditer(text):
            seen_any = True
            code = m.group(1)
            if code not in ALLOWED_CODES:
                offenders.append("%s: %r" % (path.name, code))

    assert seen_any, "warns.add(...) 호출을 하나도 못 찾았습니다 — 정규식이 깨졌을 수 있습니다"
    assert offenders == [], "닫힌 집합 밖의 경고 코드: %s" % "; ".join(offenders)


def test_closed_set_matches_claude_md():
    """CLAUDE.md에 적힌 아홉 개와 이 테스트의 허용 목록이 어긋나면 둘 중 하나가
    낡은 것이니 바로 드러나야 한다."""
    text = (ROOT / "CLAUDE.md").read_text(encoding="utf-8")
    for code in ALLOWED_CODES:
        assert code in text, "CLAUDE.md에 %s가 없습니다" % code
    assert len(ALLOWED_CODES) == 9
