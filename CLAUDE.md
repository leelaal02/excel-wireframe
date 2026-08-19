# 프로젝트 규칙

Excel 화면설계서 → PPT 생성 Claude Skill을 만드는 저장소다.

## 설계 원칙

- **`screens.json`은 중간 산출물이다.** Excel은 임포트 소스, `screens.json`은 캐시다.
  `extract.py`는 재추출 때 그냥 덮어쓴다 — 추출이 어긋나면 매핑을 고친다.
  `build.py`는 openpyxl을 import하지 않는다 — 이 제약은 `tests/test_build.py`가
  소스를 검사해 강제한다.
- **화면 단위로 예외를 격리한다.** 한 화면이 실패해도 나머지는 완성하고, 실패는 경고로 남긴다.
- **사진은 자리 위쪽에 붙인다(가로는 가운데).** 스크린샷은 위에서부터 읽으므로
  장을 넘길 때 상단 기준선이 흔들리면 안 된다.
- **사진과 표 자리는 모든 화면에서 같다.** 본문 영역 높이에 대한 고정 비율이다.
  내용이 길다고 표를 늘리지 않는다 — 늘리면 화면마다 사진 크기가 달라진다.
  넘치면 글자를 줄이고(7 → 6.5 → 6pt), 그래도 넘치면 `text-overflow`로 알린다.
- **긴 상세는 `summaries.json`으로 줄여 넣는다.** Excel 원문은 화면설계서에 그대로
  싣기엔 길다. 요약은 Claude가 추출과 생성 사이에서 만들고, 파일이 없으면 원문을 쓴다.
- **상세 표 번호는 Excel 값을 그대로 쓴다.** 스크린샷의 SoM 뱃지와 대응하므로 재부여 금지.
- **슬라이드 크기를 바꾸지 않는다.** 템플릿 크기를 승계한다.
- **템플릿은 선택 입력이다.** 없으면 `default_template.py`가 만든다. 사용자에게 템플릿을
  먼저 요구하지 말고, 기본 템플릿으로 결과를 만들어 보여준 뒤 물어본다.

## 코드 규칙

- 스크립트는 `skills/excel-wireframe/scripts/`에 평면 배치. 패키지로 만들지 않는다.
  모듈 간 import는 `from common import ...`.
- 모든 CLI 진입점은 `setup_stdio()`를 먼저 호출한다. Windows cp949에서 한글이 깨진다.
- 표 셀에 값을 쓸 때는 `slide_fill.set_cell_text`를 쓴다. 런을 새로 만들면 서식이 초기화된다.
- 값이 없는 자리는 비우지 않는다. `slide_fill.set_text_or_required`가 빨강·굵게
  `입력필요`를 남긴다 — 자동 필드(쪽번호)만 예외다.
- 레이아웃 위의 표에 값을 넣을 때는 표를 복제해 덮지 않는다. 원본은 그대로 두고
  `slide_layout.add_meta_text_slots`로 칸 자리에 글자만 올린다 — 복제하면 표가
  두 겹이 되어 원본을 클릭할 수 없고 테두리가 두 번 그려진다. (PowerPoint는
  레이아웃 도형 중 placeholder만 상속하므로 표에 직접 쓸 수는 없다.)
- 빈 자리에 값을 쓸 때는 그 자리가 `endParaRPr`에 품고 있는 서식을 물려받는다.
  그냥 런을 만들면 크기가 상속(기본 18pt)으로 떨어져 자리가 부푼다.
- 슬라이드 복제는 `slide_clone.clone_slide`만 쓴다. rId 재매핑을 빼면 그림이 사라진다.
- 레이아웃으로 슬라이드를 만들 때는 `slide_layout`의 함수만 쓴다. placeholder를
  직접 복제하면 date/footer/쪽번호가 빠지거나 빈 자리가 산출물에 남는다.
- 테스트 템플릿 픽스처는 `default_template.build_default_template`으로 템플릿을 만든 뒤
  `slide_layout`의 생성 경로로 예시 슬라이드를 한 장 얹는다. 픽스처용 pptx를 따로
  만들지 않는다.
- 글자가 몇 줄이 되는지는 `text_metrics`로만 센다. 폭·여백 기준은
  `slide_layout.detail_text_width`, 행 높이 기준은 `slide_layout.ratio_row_heights`가
  단일 출처다 — 계산과 산출물이 다른 식을 쓰면 표가 자리에 안 맞는다.
- 결과물 pptx 이름은 `common.resolve_output_path`가 정한다 — 원본 Excel 파일명을 쓰고
  같은 이름이 있으면 뒤에 `2`, `3`을 붙인다. 스킬 지시문이나 호출부에 파일명을
  하드코딩하지 않는다. 재실행이 직전 산출물을 덮어쓰면 결과를 비교할 수 없다.
  `--out-file`은 사용자가 이름을 직접 지정했을 때만 쓴다.
- 결과물 폴더에는 pptx만 둔다. 중간 산출물은 그 안의 `.work/`에 모인다 —
  경로는 `common.work_dir`이 유도하므로 호출부가 `.work`를 타이핑하지 않는다.
  CLI 인자는 `--output`(결과물 폴더) 하나이고, `--mapping`·`--screens`는 다른
  위치의 파일을 쓸 때만 붙이는 선택 인자다. 코드 안의 이름은 결과물 기준이면
  `output_dir`, 중간 산출물 기준이면 `work_dir`로 가른다.
- 경고 코드는 아홉 개뿐이다: `no-image`, `no-detail`, `text-overflow`, `shape-not-found`,
  `slide-split`, `slot-shortage`, `screen-failed`, `image-convert-failed`, `orphan-row`.

## 테스트

```bash
python -m pytest -q                      # 전체
python -m pytest tests/test_build.py -v  # 개별
```

실제 샘플(`짧은 버전.xlsx`, `화면설계서_저작권_...pptx`)을 쓰는 회귀 테스트는
파일이 없으면 자동으로 skip된다. 픽스처는 `tests/fixtures.py`가 생성한다.

## 문서

- 설계: `docs/superpowers/specs/2026-08-03-excel-wireframe-design.md`
- 계획: `docs/superpowers/plans/2026-08-04-excel-wireframe.md`
