# mapping.json / screens.json 레퍼런스

## mapping.json

양식 해석 결과. 같은 양식을 다시 처리할 때 분석 단계를 건너뛰게 해준다.

### excel

| 필드 | 값 | 설명 |
|---|---|---|
| `layout` | `sheet-per-screen` \| `table` | 1시트=1화면인지, 1행=1화면인지 |
| `sheet_include` | 정규식 | `sheet-per-screen`에서 화면 시트만 고르는 필터. 예: `^설계_` |
| `screen_meta.cell` | 셀 주소 | 화면ID·화면명이 든 셀. 예: `A1` |
| `screen_meta.pattern` | 정규식 | 명명 그룹 `id`, `name`으로 분리. 안 맞으면 시트명이 ID가 된다 |
| `sheet` | 시트명 | `table` 레이아웃에서 화면 목록이 있는 시트 |
| `header_row` | 정수 | `table` 레이아웃의 헤더 행 번호 |
| `columns` | `{"id": "A", "name": "B"}` | `table` 레이아웃의 화면ID·화면명 열 |
| `fields` | `{"설명": "C"}` | `table` 레이아웃의 화면 단위 부가 값. 키는 자유 |
| `detail.header_scan_column` | 열 문자 | 상세 헤더를 찾을 열. 예: `A` |
| `detail.header_marker` | 문자열 | 헤더 행을 식별하는 값. 예: `No.` |
| `detail.mode` | `grouped-rows` \| `merged-cells` \| `none` | `table` 레이아웃 전용. 화면ID가 빈 행을 직전 화면 상세로 귀속. 앞의 두 값은 동작이 같다 |
| `detail.columns` | `{"no": "A", "desc": "D"}` | 상세 열 매핑. 키는 자유이나 `no`는 표 번호 칸에 쓰인다 |

### template

| 필드 | 값 | 설명 |
|---|---|---|
| `file` | 경로 | 복제 소스 pptx. 결과물의 슬라이드 크기는 이 파일을 따른다 |
| `mode` | `clone` \| `layout` | 예시 슬라이드 복제인지, 빈 레이아웃인지 |
| `source_slide` | 정수 | 복제할 슬라이드 인덱스(0-based) |
| `shapes.title` | 도형 이름 | 화면명이 들어갈 자리 |
| `shapes.screen_id` | 도형 이름 | 화면ID가 들어갈 자리 |
| `shapes.image` | 도형 이름 | 스크린샷이 들어갈 자리. 이 도형은 제거되고 그 위치·크기에 이미지가 들어간다 |
| `shapes.detail_tables` | 도형 이름 배열 | **슬롯 순서대로** 나열. 생략하면 좌→우 자동 정렬 |
| `table_columns` | `{"no": 0, "text": 1}` | 표 안의 번호 열·설명 열 인덱스 |

`shapes`에 `fields`의 키와 같은 이름을 넣으면 그 값이 해당 도형에 기록된다.

### 기본 템플릿

사용자가 템플릿을 주지 않으면 `analyze.py`가 `default-template.pptx`를 만든다.
구조는 16:9 슬라이드에 상단 제목 바(`제목`, `화면ID`), 가운데 이미지 자리(`화면이미지`),
하단 상세 표 5개(`상세표1`~`상세표5`, 각 4행 2열 = 20슬롯)다.

이 경우 `structure-report.json`의 `suggested_template_mapping`을 `template` 섹션에
그대로 복사하면 된다. 도형 이름이 고정이라 추측이 필요 없다.

```json
{
  "file": "work/default-template.pptx",
  "mode": "clone",
  "source_slide": 0,
  "shapes": {
    "title": "제목",
    "screen_id": "화면ID",
    "image": "화면이미지",
    "detail_tables": ["상세표1", "상세표2", "상세표3", "상세표4", "상세표5"]
  },
  "table_columns": { "no": 0, "text": 1 }
}
```

`build_default_template(path, slide_width_emu=..., slide_height_emu=..., table_count=..., rows_per_table=...)`로
크기와 슬롯 수를 바꿀 수 있다. 상세가 항상 20건을 넘는 양식이라면 `rows_per_table`을
키워 분할을 줄인다.

### options

| 필드 | 기본 | 설명 |
|---|---|---|
| `detail_text_source` | `desc` | 상세 표에 넣을 `details` 키 |
| `overflow` | `split` | 슬롯을 넘치면 다음 슬라이드로 분할하고 제목에 `(1/2)` 표시. `build.py`는 이 값을 읽지 않는다 — 분할은 항상 일어난다. 다른 값을 넣어도(예: `none`) 동작은 바뀌지 않는다. 이 필드는 향후 확장을 위해 자리만 있다 |
| `clear_unused_slots` | `true` | 남는 슬롯의 예시 텍스트를 비운다 |

## screens.json (SSOT)

```json
{
  "meta": { "title": "화면설계서", "source": "입력.xlsx", "template": "템플릿.pptx" },
  "screens": [
    {
      "id": "B2BISMT1001",
      "name": "이용기관 목록",
      "sheet": "설계_B2BISMT1001",
      "images": ["images/B2BISMT1001.png"],
      "fields": {},
      "details": [
        { "no": "1", "type": "버튼", "element": "[등록] 버튼",
          "desc": "이용기관을 신규로 등록한다", "pos": "상단 우측" }
      ]
    }
  ]
}
```

`details`의 키는 양식마다 다르다. `build.py`는 `options.detail_text_source`가 가리키는
키와 `no`만 사용하므로, 나머지 키는 사람이 참고하도록 그대로 보존된다.

`images`는 작업 디렉토리 기준 상대 경로다.

## 경고 코드

| 코드 | 의미 |
|---|---|
| `no-image` | 화면에 연결된 이미지를 찾지 못함 |
| `no-detail` | 상세 표 헤더를 찾지 못함 |
| `text-overflow` | 셀 폭 대비 글자 수가 많아 잘릴 수 있음 |
| `shape-not-found` | 템플릿에서 지정한 도형 이름을 찾지 못함 |
| `slide-split` | 상세가 슬롯을 넘쳐 슬라이드를 나눔 |
| `slot-shortage` | 이 슬라이드에 다 못 담음 |
| `screen-failed` | 해당 화면 생성 실패. 나머지 화면은 계속 생성됨 |
| `image-convert-failed` | EMF/WMF를 PNG로 변환하지 못해 원본을 사용 |
| `orphan-row` | `table` 레이아웃에서 화면ID 없이 시작하는 행 |
