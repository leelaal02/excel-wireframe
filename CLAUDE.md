# 프로젝트 규칙

Excel 화면설계서 → PPT 생성 Claude Skill을 만드는 저장소다.

## 설계 원칙

- **`screens.json`은 중간 산출물이다.** Excel은 임포트 소스, `screens.json`은 캐시다.
  `extract.py`는 재추출 때 그냥 덮어쓴다 — 추출이 어긋나면 매핑을 고친다.
  `build.py`는 openpyxl을 import하지 않는다 — 이 제약은 `tests/test_build.py`가
  소스를 검사해 강제한다.
- **화면 단위로 예외를 격리한다.** 한 화면이 실패해도 나머지는 완성하고, 실패는 경고로 남긴다.
- **상세 표 번호는 Excel 값을 그대로 쓴다.** 스크린샷의 SoM 뱃지와 대응하므로 재부여 금지.
- **슬라이드 크기를 바꾸지 않는다.** 템플릿 크기를 승계한다.
- **템플릿은 선택 입력이다.** 없으면 `default_template.py`가 만든다. 사용자에게 템플릿을
  먼저 요구하지 말고, 기본 템플릿으로 결과를 만들어 보여준 뒤 물어본다.

## 코드 규칙

- 스크립트는 `skills/excel-wireframe/scripts/`에 평면 배치. 패키지로 만들지 않는다.
  모듈 간 import는 `from common import ...`.
- 모든 CLI 진입점은 `setup_stdio()`를 먼저 호출한다. Windows cp949에서 한글이 깨진다.
- 표 셀에 값을 쓸 때는 `slide_fill.set_cell_text`를 쓴다. 런을 새로 만들면 서식이 초기화된다.
- 슬라이드 복제는 `slide_clone.clone_slide`만 쓴다. rId 재매핑을 빼면 그림이 사라진다.
- 레이아웃으로 슬라이드를 만들 때는 `slide_layout`의 함수만 쓴다. placeholder를
  직접 복제하면 date/footer/쪽번호가 빠지거나 빈 자리가 산출물에 남는다.
- 테스트 템플릿 픽스처는 `default_template.build_default_template`으로 템플릿을 만든 뒤
  `slide_layout`의 생성 경로로 예시 슬라이드를 한 장 얹는다. 픽스처용 pptx를 따로
  만들지 않는다.
- 글자가 몇 줄이 되고 행이 얼마나 높아야 하는지는 `text_metrics`로만 센다.
  폭·여백 기준은 `slide_layout.detail_text_width`가 단일 출처다 — 계산과 산출물이
  다른 식을 쓰면 표가 자리에 안 맞는다.
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
