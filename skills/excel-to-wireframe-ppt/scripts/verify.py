# -*- coding: utf-8 -*-
"""생성된 pptx를 다시 파싱해 기대와 대조한다.

만들었다는 사실만으로는 부족하다. 도형 이름이 안 맞거나 rId가 깨지면
파일은 생기지만 내용이 비어 있을 수 있다.
"""
from __future__ import annotations

from pathlib import Path

from pptx_scan import scan_presentation


def _title_texts(slide: dict, title_name: str | None) -> list[str]:
    """슬라이드에서 화면명 판단에 쓸 텍스트를 뽑는다.

    제목 도형 이름이 지정되면 그 도형만 본다. 없으면(매핑이 제목을 지정하지
    않으면) 어느 도형이든 검사하던 예전 방식으로 물러선다 — 크래시를 피하기
    위함이지, 정확도를 위한 것은 아니다.
    """
    if title_name:
        return [sh["text"] for sh in slide["shapes"] if sh["name"] == title_name]
    return [sh["text"] for sh in slide["shapes"] if sh["text"]]


def _slides_for_screen(slides: list[dict], title_name: str | None,
                       screen_name: str) -> list[dict]:
    """화면명이 제목(또는 대체 텍스트)에 들어간 슬라이드를 모두 찾는다.

    분할된 화면은 '이름 (1/2)', '이름 (2/2)'처럼 슬라이드가 여러 장이므로
    첫 번째 매치만 보면 안 된다.
    """
    return [
        s for s in slides
        if any(screen_name in t for t in _title_texts(s, title_name))
    ]


def verify_output(out_path: Path, screens_data: dict, mapping: dict,
                  expected_slides: int) -> dict:
    report = scan_presentation(Path(out_path))
    slides = report["slides"]
    checks: list[dict] = []

    checks.append(
        {
            "name": "슬라이드 수",
            "ok": len(slides) == expected_slides,
            "detail": "기대 %d장, 실제 %d장" % (expected_slides, len(slides)),
        }
    )

    title_name = mapping.get("template", {}).get("shapes", {}).get("title")
    screens = screens_data.get("screens", [])

    missing_titles = [
        scr["name"] for scr in screens
        if not _slides_for_screen(slides, title_name, scr["name"])
    ]
    checks.append(
        {
            "name": "화면명 반영",
            "ok": not missing_titles,
            "detail": "누락 없음" if not missing_titles
            else "누락: %s" % ", ".join(missing_titles),
        }
    )

    image_issues: list[str] = []
    for scr in screens:
        matches = _slides_for_screen(slides, title_name, scr["name"])
        if not matches:
            image_issues.append("%s: 화면에 해당하는 슬라이드를 찾지 못함" % scr["id"])
            continue
        if scr.get("images"):
            has_pic = any(
                any("PICTURE" in sh["shape_type"] for sh in s["shapes"])
                for s in matches
            )
            if not has_pic:
                image_issues.append("%s: 이미지가 배치되지 않음" % scr["id"])
    checks.append(
        {
            "name": "이미지 배치",
            "ok": not image_issues,
            "detail": "문제 없음" if not image_issues else "; ".join(image_issues),
        }
    )

    template = Path(mapping["template"]["file"])
    if template.exists():
        tpl_report = scan_presentation(template)
        same = (
            tpl_report["slide_width"] == report["slide_width"]
            and tpl_report["slide_height"] == report["slide_height"]
        )
        checks.append(
            {
                "name": "슬라이드 크기",
                "ok": same,
                "detail": "%.2f x %.2f in" % tuple(report["slide_size_in"]),
            }
        )
    else:
        checks.append(
            {
                "name": "슬라이드 크기",
                "ok": False,
                "detail": "템플릿 파일을 찾을 수 없어 비교 불가: %s" % template,
            }
        )

    return {"ok": all(c["ok"] for c in checks), "checks": checks}
