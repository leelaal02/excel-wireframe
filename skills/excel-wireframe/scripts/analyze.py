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

from common import (
    migrate_legacy_work,
    migration_notice,
    setup_stdio,
    work_dir,
    write_json,
)
from default_template import build_default_template, default_template_mapping
from pptx_scan import scan_layouts, scan_presentation, suggest_content_area, suggest_mode
from user_default import load_user_default, user_default_mapping
from xlsx_scan import scan_workbook


def build_report(excel_path: Path, template_path: Path) -> dict:
    excel = scan_workbook(excel_path)
    template = scan_presentation(template_path)
    return {
        "excel": excel,
        "template": template,
        "layouts": scan_layouts(template_path),
        "suggestion": suggest_mode(template),
    }


def resolve_template(template_arg: str | None, out_path: Path,
                     skill_base: Path | None = None):
    """쓸 템플릿을 정한다. (경로, 생성했는가, 사용자 기본 설정) 셋을 돌려준다.

    우선순위는 명시 지정 > 사용자 기본 설정 > 생성이다. 사용자가 자기 조직
    템플릿을 설치본에 등록해 두었으면 --template 없이도 그 디자인이 나온다.
    """
    if template_arg:
        return Path(template_arg), False, None

    cfg = load_user_default(skill_base)
    if cfg is not None:
        return cfg["template_path"], False, cfg

    generated = Path(out_path).parent / "default-template.pptx"
    build_default_template(generated)
    return generated, True, None


def main(argv: list[str] | None = None) -> int:
    setup_stdio()
    ap = argparse.ArgumentParser(description="Excel과 PPT 템플릿 구조를 스캔한다")
    ap.add_argument("--excel", required=True)
    ap.add_argument("--template", default=None,
                    help="생략하면 기본 템플릿을 만들어 사용한다")
    ap.add_argument("--output", default=None,
                    help="결과물 폴더. 리포트는 그 안 .work/에 쓴다")
    ap.add_argument("--out", default=None,
                    help="리포트 경로를 직접 지정한다. --output보다 우선한다")
    args = ap.parse_args(argv)

    if not args.out and not args.output:
        # 어디에 쓸지 모르는 채로 스캔을 시작하지 않는다. Excel과 템플릿을 다 읽고
        # 나서 저장할 곳이 없다고 죽으면 시간만 버린다.
        ap.error("--output 또는 --out 중 하나가 필요합니다")

    if args.out:
        out_path = Path(args.out)
    else:
        output_dir = Path(args.output)
        notice = migration_notice(migrate_legacy_work(output_dir))
        if notice:
            print(notice)
        out_path = work_dir(output_dir) / "structure-report.json"
    template_path, generated, user_cfg = resolve_template(args.template, out_path)

    report = build_report(Path(args.excel), template_path)
    report["template_generated"] = generated
    report["user_default"] = user_cfg is not None
    if user_cfg is not None:
        report["suggested_template_mapping"] = user_default_mapping(user_cfg)
    if generated:
        from default_template import DEFAULT_LAYOUT_NAME
        tpl_mapping = default_template_mapping(template_path)
        layout_info = next(
            (lay for lay in report["layouts"] if lay["name"] == DEFAULT_LAYOUT_NAME),
            None,
        )
        if layout_info is not None:
            tpl_mapping["content_area"] = suggest_content_area(
                layout_info,
                report["template"]["slide_width"],
                report["template"]["slide_height"],
            )
        report["suggested_template_mapping"] = tpl_mapping
    write_json(out_path, report)

    sheets = report["excel"]["sheets"]
    print("Excel 시트 %d개: %s" % (len(sheets), ", ".join(s["name"] for s in sheets)))
    if generated:
        print("템플릿 미제공 → 기본 템플릿 생성: %s" % template_path)
        print("  제안 매핑이 리포트의 suggested_template_mapping에 있습니다")
    elif user_cfg is not None:
        print("템플릿 미제공 → 사용자 기본 템플릿 사용: %s" % template_path)
        print("  레이아웃 '%s', 제안 매핑이 리포트의 "
              "suggested_template_mapping에 있습니다" % user_cfg.get("layout"))
    w, h = report["template"]["slide_size_in"]
    print("템플릿 슬라이드 %d장, 크기 %.2f x %.2f in"
          % (len(report["template"]["slides"]), w, h))
    print("제안 모드: %s (%s)"
          % (report["suggestion"]["mode"], report["suggestion"]["reason"]))
    print("리포트 저장: %s" % out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
