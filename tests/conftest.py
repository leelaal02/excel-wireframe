import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "skills" / "excel-wireframe" / "scripts"
sys.path.insert(0, str(SCRIPTS))


@pytest.fixture(autouse=True)
def isolate_user_default(monkeypatch):
    """설치본의 user-default.json이 테스트 결과를 바꾸지 못하게 막는다.

    사용자가 자기 조직 템플릿을 기본으로 등록해 두면 analyze가 그것을 집는다.
    테스트가 그 설정을 읽으면 개발자 환경마다 결과가 달라진다 — 없는 디렉토리를
    보게 해서 항상 '설정 없음' 경로를 타게 한다. 설정 자체를 검증하는
    test_user_default.py는 base를 직접 넘기므로 이 픽스처의 영향을 받지 않는다.
    """
    import user_default

    monkeypatch.setattr(
        user_default, "skill_dir", lambda: ROOT / "__no_user_default__"
    )
