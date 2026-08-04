# -*- coding: utf-8 -*-
"""생성된 pptx를 다시 파싱해 기대와 대조한다.

만들었다는 사실만으로는 부족하다. 도형 이름이 안 맞거나 rId가 깨지면
파일은 생기지만 내용이 비어 있을 수 있다.
"""
from __future__ import annotations

from pathlib import Path

from pptx_scan import scan_presentation


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

    all_text = "\n".join(
        sh["text"] for s in slides for sh in s["shapes"] if sh["text"]
    )
    missing = [
        scr["name"] for scr in screens_data.get("screens", [])
        if scr["name"] not in all_text
    ]
    checks.append(
        {
            "name": "화면명 반영",
            "ok": not missing,
            "detail": "누락 없음" if not missing else "누락: %s" % ", ".join(missing),
        }
    )

    want_pics = sum(1 for scr in screens_data.get("screens", []) if scr.get("images"))
    got_pics = sum(
        1 for s in slides for sh in s["shapes"] if "PICTURE" in sh["shape_type"]
    )
    checks.append(
        {
            "name": "이미지 배치",
            "ok": got_pics >= want_pics,
            "detail": "이미지 있는 화면 %d개, 슬라이드의 그림 도형 %d개"
            % (want_pics, got_pics),
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

    return {"ok": all(c["ok"] for c in checks), "checks": checks}
