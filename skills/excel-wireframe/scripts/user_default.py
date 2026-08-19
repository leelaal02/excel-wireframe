# -*- coding: utf-8 -*-
"""사용자가 자기 조직 템플릿을 기본으로 두는 설정.

이 스킬은 코드째 배포되므로 남의 저작권 표기와 로고가 박힌 실물 템플릿을
저장소에 담을 수 없다. 대신 설치본에 `user-default.json`을 두면 `--template`
없이 실행할 때 그 템플릿과 레이아웃을 쓴다. 설정이 없으면 지금까지처럼
`default_template.py`가 기본 템플릿을 만든다.

설정 파일은 스킬 디렉토리(`scripts/`의 부모)에 둔다. `template`은 그 디렉토리
기준 상대경로이거나 절대경로다 — 파일을 스킬 안에 복사해 두든, 다른 곳의 원본을
가리키든 둘 다 된다.
"""
from __future__ import annotations

from pathlib import Path

from common import read_json

CONFIG_NAME = "user-default.json"

# mapping의 template 섹션으로 그대로 옮길 키. 나머지는 내부용이다.
_MAPPING_KEYS = (
    "mode",
    "layout",
    "source_slide",
    "placeholders",
    "shapes",
    "meta_table",
    "detail_tables",
    "table_columns",
    "content_area",
)


def skill_dir() -> Path:
    """스킬 디렉토리. scripts/user_default.py의 부모의 부모다."""
    return Path(__file__).resolve().parent.parent


def load_user_default(base: Path | None = None) -> dict | None:
    """설치본의 user-default.json을 읽는다. 없으면 None.

    파일이 있는데 내용이 잘못됐으면 ValueError로 알린다. 조용히 기본 템플릿으로
    떨어지면 왜 내 템플릿이 안 쓰이는지 알 길이 없다.
    """
    base = Path(base) if base is not None else skill_dir()
    path = base / CONFIG_NAME
    if not path.exists():
        return None

    cfg = dict(read_json(path))

    raw = cfg.get("template")
    if not raw:
        raise ValueError("%s에 template 경로가 없습니다" % path)
    tpl = Path(raw)
    if not tpl.is_absolute():
        tpl = base / tpl
    if not tpl.exists():
        raise ValueError(
            "%s가 가리키는 템플릿을 찾지 못했습니다: %s" % (path, tpl)
        )

    mode = cfg.get("mode", "layout")
    if mode == "layout" and not cfg.get("layout"):
        raise ValueError(
            "%s에 layout이 없습니다. 레이아웃 이름을 지정하지 않으면 "
            "빌드가 첫 번째 레이아웃을 집어 엉뚱한 디자인이 나옵니다." % path
        )
    if mode == "clone" and cfg.get("source_slide") is None:
        raise ValueError("%s에 source_slide가 없습니다 (mode: clone)" % path)

    cfg["mode"] = mode
    cfg["template_path"] = tpl
    return cfg


def user_default_mapping(cfg: dict) -> dict:
    """설정을 mapping.json의 template 섹션으로 바꾼다."""
    out = {"file": str(cfg["template_path"])}
    for key in _MAPPING_KEYS:
        if key in cfg and cfg[key] is not None:
            out[key] = cfg[key]
    return out
