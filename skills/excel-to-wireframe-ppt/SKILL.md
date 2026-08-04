---
name: excel-to-wireframe-ppt
description: Excel로 작성된 화면설계서를 PowerPoint 화면설계서로 자동 생성한다. 사용자가 화면설계서, 화면정의서, 스크린 설계, 와이어프레임 문서를 Excel에서 PPT로 만들어 달라고 하거나, xlsx와 pptx 템플릿을 함께 건네며 "이 양식대로 만들어줘"라고 할 때 반드시 사용한다. 처음 보는 Excel 양식이어도 구조를 분석해 매핑을 도출하고, PPT 템플릿을 안 줘도 기본 템플릿을 만들어 진행하므로, 양식이 낯설거나 템플릿이 없다는 이유로 건너뛰지 말 것.
---

# Excel 기반 화면설계서 PPT 생성

Excel 화면설계서를 받아 화면 페이지 PPT를 생성한다. PPT 템플릿을 함께 받으면
그 디자인을 그대로 유지하고, 템플릿이 없으면 표준 구조의 기본 템플릿을 만들어 쓴다.

## 원칙

`screens.json`이 SSOT다. Excel은 최초 임포트 소스일 뿐이고, PPT의 내용은 `screens.json`에서 온다.
사람이 `screens.json`을 손봤을 수 있으므로 재추출이 그것을 덮어쓰지 않는다.

판단은 세 곳에서만 한다. 나머지는 스크립트가 결정론적으로 처리한다.

## 절차

### 0. 준비

작업 디렉토리를 정한다(기본: 입력 Excel 옆의 `work/`). 의존성을 확인한다.

```bash
python -c "import pptx, openpyxl, PIL"
```

없으면 `python -m pip install python-pptx openpyxl Pillow`.

`work/mapping.json`이 이미 있으면 1~2단계를 건너뛰고 3단계로 간다.

### 1. 구조 분석

```bash
python <스킬>/scripts/analyze.py --excel <입력.xlsx> --template <템플릿.pptx> --out work/structure-report.json
```

**템플릿을 안 받았으면 `--template`을 생략한다.** 기본 템플릿(16:9, 제목 바 + 이미지 자리 +
상세 표 5개 × 4행)이 `work/default-template.pptx`로 만들어지고, 그에 맞는 매핑이
리포트의 `suggested_template_mapping`에 담긴다. 템플릿이 있는지 사용자에게 먼저 묻지 말고,
없으면 기본 템플릿으로 진행한 뒤 결과를 보여주며 "원하는 템플릿이 있으면 주시면 그대로
맞춰 드립니다"라고 알린다. 빈손으로 되묻는 것보다 결과물을 보고 판단하는 편이 빠르다.

### 2. [판단①] 매핑 작성과 확인

`work/structure-report.json`을 읽고 `work/mapping.json`을 작성한다.
필드 의미는 `references/mapping-schema.md`를 참고한다.

`template_generated`가 `true`면 `template` 섹션은 `suggested_template_mapping`을
그대로 복사한다. 도형 이름을 우리가 정했으므로 추측할 여지가 없다. Excel 쪽만 판단하면 된다.

판단할 것:

- **Excel 레이아웃** — 시트마다 화면 하나면 `sheet-per-screen`, 한 시트에 화면이 행으로 나열되면 `table`
- **화면이 아닌 시트** — `표지`, `테스트_*`, `요약` 같은 시트를 `sheet_include` 정규식으로 걸러낸다
- **상세 표 헤더 마커** — 헤더 행의 첫 칸 값(보통 `No.`). 헤더 위치는 이미지 크기에 따라 밀리므로 행 번호로 고정하지 말고 마커로 찾게 한다
- **템플릿 도형 이름** — 제목·화면ID·이미지 자리·상세 표. 상세 표가 여러 개면 슬롯 순서대로 나열한다
- **상세 표에 넣을 컬럼** — `detail_text_source`. 보통 상세 설명 컬럼 하나다

작성한 매핑을 **표로 사용자에게 보여주고 확인을 받는다.** 도형 이름을 위치·크기 휴리스틱으로
추측했다면 그 사실을 명시한다. 확인 없이 3단계로 가지 않는다.

### 3. 추출

```bash
python <스킬>/scripts/extract.py --excel <입력.xlsx> --mapping work/mapping.json --work work
```

`work/screens.json`과 `work/images/`가 생긴다. 이미 `screens.json`이 있으면
`screens.new.json`에 쓰고 diff를 출력한다 — 그 경우 diff를 사용자에게 보여주고
반영 여부를 물은 뒤, 반영하기로 하면 파일을 교체한다.

재추출은 `screens.json`은 보호하지만 `work/images/` 아래 파일은 그대로 덮어쓴다.
SSOT는 `screens.json`뿐이다 — 스크린샷을 손으로 교체했다면 재추출 시 원래 이미지로 되돌아간다는 것을 사용자에게 알린다.

### 4. [판단②] 추출 결과 보고

화면 수, 상세 건수, 이미지가 붙은 화면 수, 경고를 사용자에게 요약한다.
`no-image` 경고가 많으면 이미지 귀속이 잘못됐을 수 있으니 매핑을 재검토한다.

### 5. 생성

```bash
python <스킬>/scripts/build.py --screens work/screens.json --mapping work/mapping.json --work work --out work/output/화면설계서.pptx
```

### 6. [판단③] 검증 결과 보고

`build.py`가 생성물을 재파싱해 검증 결과를 출력한다. 실패 항목이 있으면 원인을 짚어
매핑을 고치고 5단계를 다시 돌린다. 분할된 화면과 잘림 경고는 사용자에게 그대로 전달한다.

## 실패했을 때

| 증상 | 원인과 조치 |
|---|---|
| 상세가 0건 | `header_marker`가 실제 헤더 값과 다르다. `structure-report.json`에서 헤더 셀 값을 확인한다 |
| 이미지가 안 붙음 | `sheet-per-screen`인데 시트명 매칭이 안 되거나, Excel이 이미지를 도형으로 넣었다. `structure-report.json`의 `image_count`를 먼저 본다 |
| 제목이 안 채워짐 | `shapes.title`의 도형 이름이 템플릿과 다르다. 리포트의 도형 이름 목록과 대조한다 |
| 표가 예시 텍스트 그대로 | `detail_tables` 이름이 틀렸다. 이름을 비우면 좌→우 자동 정렬로 대체된다. 이 경우 `shape-not-found` 경고가 함께 떠서 어떤 표 이름을 못 찾았는지 알려주므로, 눈에 보이는 증상만 보지 말고 경고 목록을 먼저 확인한다 |
| 한글이 깨져 출력됨 | 콘솔 인코딩 문제다. `PYTHONIOENCODING=utf-8`을 설정하고 재실행한다 |
| 템플릿에 예시 슬라이드가 없다고 나옴 | `analyze.py`가 `layout` 모드로 판정한 경우다. 빈 레이아웃에는 상세를 넣을 표가 없으므로, 사용자에게 예시 페이지가 있는 템플릿을 요청하거나 `--template` 없이 기본 템플릿으로 진행한다 |

## 하지 않는 것

- 와이어프레임을 새로 그리지 않는다. Excel에 삽입된 이미지를 쓴다
- 슬라이드 크기를 바꾸지 않는다. 템플릿 크기를 그대로 승계한다
- 표지·목차를 만들지 않는다. 화면 페이지만 만든다
- 상세 표 번호를 재부여하지 않는다. Excel의 번호가 스크린샷 뱃지와 대응한다
- 템플릿을 달라고 먼저 요구하지 않는다. 없으면 기본 템플릿으로 만들어 보여준 뒤 물어본다
