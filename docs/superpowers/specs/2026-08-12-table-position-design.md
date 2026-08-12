# 상세표 위치를 레이아웃에 맞춰 옮긴다

작성 2026-08-12. **상태: 설계 승인 대기** — 마지막 확인 질문에 답을 못 받고 중단했다
(맨 아래 "미결" 참고). 구현은 시작하지 않았다.

## 문제

`layout` 모드는 `slide_layout.split_content_area`가 본문 영역을 항상 같은 방식으로
가른다 — 위는 이미지, 아래는 표 5개를 가로로 균등 분할. 스크린샷이 가로로 넓으면
잘 맞지만, 세로로 긴 화면은 `place_image`가 비율을 지켜 축소하면서 좌우가 통째로
빈다. 그 빈 폭이 바로 표가 들어갈 자리다.

기본 템플릿의 레이아웃 이름이 `_1366_768`(가로)과 `_freeHeight_noScroll`(세로로 긴
스크롤 페이지)로 갈리는 것도 같은 사정이다. 레이아웃마다 스크린샷 모양이 다르므로
표 자리도 달라야 한다.

### 실측 근거 (`output/긴 버전7.pptx`, `output/.work/images`)

슬라이드 10.83 × 7.5in, 본문 영역 10.89 × 6.92in.

| 화면 | 원본 비율 | 아래 배치에서 실제 이미지 |
|---|---|---|
| B2BISMT1001 | 2.38 | 9.13 × 4.58 — 거의 꽉 참 |
| CPTIPMT1002 조각 | 2.00 | 9.41 × 4.72 — 잘 맞음 |
| (6번째 슬라이드) | 1.51 | 6.13 × 4.05 — 좌우에 2.37in씩 버려짐 |

원본 `CPTIPMT1002`는 1904 × 3164(비율 0.60)라 `image_split`이 4조각으로 잘라
슬라이드 4장에 싣는다. 이미지 자리의 비율이 바뀌면 이 조각 수가 바뀐다.

| 배치 | 이미지 자리 | 자리 비율 | CPTIPMT1002 조각 수 |
|---|---|---|---|
| 아래(현재) | 10.89 × 4.5in | 2.42 | 4장 |
| 옆 4.0in | 6.89 × 6.92in | 1.00 | 2장 |

반대로 가로로 넓은 B2BISMT1001은 옆 배치에서 6.89 × 2.89로 작아진다. 어느 한쪽이
항상 옳지 않으므로 **설정으로 고른다**.

## 설정

`mapping.json`과 `user-default.json`의 `template.detail_tables`에 두 키를 더한다.

```json
"detail_tables": { "count": 4, "rows": 4, "position": "right", "width_inch": 4.0 }
```

- `position`: `bottom`(기본) | `right` | `left`. 없으면 지금과 동일하게 동작한다.
- `width_inch`: 표 영역 폭. `right`/`left`에서만 의미가 있고 기본 4.0in.
- `top`은 넣지 않는다. 쓸 자리가 없다.
- 문서 전체에 하나의 배치다. 화면마다 다르게 하지 않는다 — 지금도 문서당 레이아웃이
  하나이므로 레이아웃을 바꿔 실행하면 배치도 같이 바뀐다.
- `user_default._MAPPING_KEYS`는 `detail_tables`를 통째로 넘기므로 **`user_default.py`는
  고치지 않는다.** 사용자 기본 템플릿에서도 그대로 먹는다.

## 배치 규칙

```
bottom (현재)                      right
┌──────────────────────┐          ┌──────────────┬────────┐
│      이미지            │          │              │  표1   │
│  (표 위 나머지 전부)     │          │              ├────────┤
├────┬────┬────┬────┬──┤          │   이미지       │  표2   │
│표1 │표2 │표3 │표4 │표5│          │ (남는 폭 전부)  ├────────┤
└────┴────┴────┴────┴──┘          │              │  표3   │
표 높이 = 행높이 합, 폭 균등분할        └──────────────┴────────┘
아래 끝에 붙고 위로 자란다              표 폭 = width_inch 고정, 위부터 세로로 쌓임
```

`bottom`이 "표 높이 먼저, 이미지는 나머지"인 것과 대칭으로 `right`는 "표 폭 먼저,
이미지는 나머지"다.

- 표는 세로로 쌓는다. 옆에서 다시 가로로 나누면 표 하나가 너무 좁아진다.
- 남는 세로는 마지막 표에 몰아주지 않고 비워 둔다. 아래 배치에서 마지막 표가 폭
  나머지를 흡수하는 것은 열 폭이라 행 높이 계산과 무관했지만, 세로는 행 높이의
  결과값이라 여기서 늘리면 계산과 산출물이 어긋난다.
- `left`는 `right`의 좌우 대칭이다.

기본값 기준 수치:

| | bottom | right (4.0in) |
|---|---|---|
| 이미지 자리 | 10.89 × 4.5in | 6.89 × 6.92in |
| 표 하나 폭 | 2.18in | 4.00in |
| 설명 칸 글자 폭 | 1.99in | 3.66in (1.84배) |

## 한계 검사

표 하나(4행 실측 `ROW_HEIGHTS` 합)가 1.55in이므로 `right`에서 5개면 7.74in로 본문
6.92in를 넘는다. 표 폭이 넓어져 줄바꿈이 줄어도 실측 행 높이가 하한으로 남기 때문에
글자 크기를 낮춰도 해결되지 않는다.

- **초과하면 ValueError로 멈춘다.** "이 content_area 높이(6.92in)에서는 count를 4
  이하로 쓰세요"처럼 대안을 적는다. `rows_per_table` 과다와 같은 자리 — `build()`의
  화면 루프 **밖** 사전 호출에서 터뜨려야 화면 단위 예외 격리에 걸려 같은 오류가
  화면 수만큼 `screen-failed`로 흩어지지 않는다.
- 이미지 폭이 1in 밑으로 내려가면(`width_inch` 과다) 같은 방식으로 막는다.
  `MIN_IMAGE_HEIGHT_EMU`의 가로 짝으로 `MIN_IMAGE_WIDTH_EMU`를 둔다.

## 코드 변경 지점

| 파일 | 변경 |
|---|---|
| `slide_layout.py` | `split_content_area(area, table_count, rows_per_table, row_heights=None, position="bottom", table_width=None)` 축 분기. `detail_text_width`에 `position`/`table_width` 추가. `MIN_IMAGE_WIDTH_EMU` 신설 |
| `build.py` | `_new_layout_slide`와 `build`가 설정을 읽어 전달. `_fit_tables`의 한계가 갈림 |
| `user_default.py` | 변경 없음 |
| `SKILL.md` | 설정 설명 추가 (`tests/test_skill_docs.py`가 `detail_tables` 언급을 검사한다) |

`_fit_tables`의 한계:

- `bottom`: `limit = area[3] - MIN_IMAGE_HEIGHT_EMU` (표 전체 높이의 한계, 현행)
- `right`/`left`: `limit = area[3] // count` (표 **하나**의 높이 한계)

`detail_text_width`는 폭 계산의 단일 출처다(CLAUDE.md). `right`에서는 표 폭이
`area_width // count`가 아니라 `table_width` 전체이므로 이 함수가 `position`을 알아야
계산과 산출물이 어긋나지 않는다.

글자 크기 축소(`DETAIL_FONT_STEPS` 7 → 6.5 → 6pt)와 `_cap_heights` 폴백은 축만 바뀌고
로직은 그대로 재사용한다.

## 테스트

- `tests/test_slide_layout.py`
  - `right`에서 이미지 자리가 본문 높이를 전부 쓰고 폭만 줄어드는지
  - 표 박스가 세로로 쌓이고 서로 겹치지 않는지, 폭이 `width_inch`인지
  - `left`가 `right`의 좌우 대칭인지
  - `position`을 안 주면 기존 결과와 완전히 같은지 (회귀 방지)
  - count 초과 / `width_inch` 과다에서 ValueError와 그 메시지에 대안이 담기는지
- `tests/test_build.py`
  - `position: right` 매핑으로 빌드해 표가 오른쪽에 서로 겹치지 않게 놓이는지
  - count 초과 매핑이 화면별 `screen-failed`가 아니라 빌드 전체 ValueError가 되는지
- `tests/test_user_default.py`
  - `user-default.json`의 `detail_tables`에 넣은 `position`/`width_inch`가 mapping으로
    그대로 넘어오는지

## 미결 — 다음 세션에서 먼저 답을 받을 것

**`right`에서 표가 5개 → 4개로 줄어드는 것을 받아들일 수 있는가?** 슬롯이 20 → 16으로
줄어 상세가 많은 화면은 슬라이드가 한 장 더 늘어난다. 받아들이기 어렵다면 대안은
`ROW_HEIGHTS` 실측 하한을 옆 배치에서만 낮추는 것인데, 그러면 원본 화면설계서와
행 높이가 달라진다.
