# -*- coding: utf-8 -*-
"""1단계: Excel과 템플릿의 구조를 스캔해 리포트로 남긴다.

이 리포트를 Claude가 읽고 mapping.json을 작성한다. 스크립트는 판단하지 않고
관찰만 한다 — suggest_mode의 제안조차 사용자 확인을 거친다.

템플릿을 주지 않으면 기본 템플릿을 만들어 쓴다. 그 경우 도형 이름을 우리가 정했으므로
추측할 필요가 없고, 바로 쓸 수 있는 매핑을 리포트에 함께 담는다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from common import setup_stdio, write_json
from default_template import build_default_template, default_template_mapping
from pptx_scan import scan_presentation, suggest_mode
from xlsx_scan import scan_workbook


def build_report(excel_path: Path, template_path: Path) -> dict:
    excel = scan_workbook(excel_path)
    template = scan_presentation(template_path)
    return {
        "excel": excel,
        "template": template,
        "suggestion": suggest_mode(template),
    }


def resolve_template(template_arg: str | None, out_path: Path) -> tuple[Path, bool]:
    if template_arg:
        return Path(template_arg), False
    generated = Path(out_path).parent / "default-template.pptx"
    build_default_template(generated)
    return generated, True


def main(argv: list[str] | None = None) -> int:
    setup_stdio()
    ap = argparse.ArgumentParser(description="Excel과 PPT 템플릿 구조를 스캔한다")
    ap.add_argument("--excel", required=True)
    ap.add_argument("--template", default=None,
                    help="생략하면 기본 템플릿을 만들어 사용한다")
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)

    out_path = Path(args.out)
    template_path, generated = resolve_template(args.template, out_path)

    report = build_report(Path(args.excel), template_path)
    report["template_generated"] = generated
    if generated:
        report["suggested_template_mapping"] = default_template_mapping(template_path)
    write_json(out_path, report)

    sheets = report["excel"]["sheets"]
    print("Excel 시트 %d개: %s" % (len(sheets), ", ".join(s["name"] for s in sheets)))
    if generated:
        print("템플릿 미제공 → 기본 템플릿 생성: %s" % template_path)
        print("  제안 매핑이 리포트의 suggested_template_mapping에 있습니다")
    w, h = report["template"]["slide_size_in"]
    print("템플릿 슬라이드 %d장, 크기 %.2f x %.2f in"
          % (len(report["template"]["slides"]), w, h))
    print("제안 모드: %s (%s)"
          % (report["suggestion"]["mode"], report["suggestion"]["reason"]))
    print("리포트 저장: %s" % args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
