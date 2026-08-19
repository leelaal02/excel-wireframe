# excel-wireframe

Excel로 작성된 화면설계서를 PowerPoint 화면설계서로 자동 생성하는 Claude Skill.

특정 Excel 양식에 하드코딩하지 않는다. 처음 보는 양식이 들어와도 구조를 분석해 매핑을 도출하고, 사용자 확인을 거쳐 PPT를 만든다. PPT 템플릿을 주면 그 디자인을 그대로 유지하고, 주지 않으면 표준 구조의 기본 템플릿을 생성해 쓴다.

## 설치

이 저장소는 스킬을 **개발하는** 곳이라 소스를 `skills/excel-wireframe/`에 둔다.
Claude Code가 스킬을 **읽는** 곳은 `.claude/skills/`다 — 저장소 경로는 자동으로
인식되지 않으므로 복사해야 걸린다.

```bash
# 전역 설치 — 모든 프로젝트에서 쓴다
cp -r skills/excel-wireframe ~/.claude/skills/

# 특정 프로젝트에서만 쓰려면
cp -r skills/excel-wireframe <프로젝트>/.claude/skills/
```

복사본은 저장소와 따로 논다. 스크립트를 고쳤으면 다시 복사해야 설치본에 반영된다.

의존성:

```bash
python -m pip install python-pptx openpyxl Pillow
```

Python 3.13 기준. Windows에서는 콘솔 인코딩이 cp949라 한글이 깨질 수 있다 — 스크립트가 `setup_stdio()`로 UTF-8을 강제하지만, 직접 파이프를 쓸 때는 `PYTHONIOENCODING=utf-8`을 붙인다.

## 쓰는 법

Claude에게 Excel 파일을 주면서 "화면설계서 PPT 만들어줘"라고 하면 스킬이 자동으로 걸린다. CLI를 직접 쓸 수도 있다.

```bash
S=~/.claude/skills/excel-wireframe/scripts

# 1. 구조 분석 (--template 생략 시 기본 템플릿 생성)
python $S/analyze.py --excel 입력.xlsx --template 템플릿.pptx --output output

# 2. 리포트를 읽고 output/.work/mapping.json 작성 — 사람이 판단하는 단계

# 3. 추출
python $S/extract.py --excel 입력.xlsx --output output

# 4. 상세를 output/.work/summaries.json으로 줄이기 — Claude가 판단하는 단계 (생략 가능)

# 5. 생성 (이름은 원본 Excel 파일명, 같은 이름이 있으면 뒤에 2, 3이 붙는다)
python $S/build.py --output output
#   → output/입력.pptx

# 이름을 직접 정하려면 (이때는 번호를 붙이지 않고 덮어쓴다)
python $S/build.py --output output --out-file 납품용.pptx
```

`--output`은 결과물 폴더다. 여기엔 pptx만 놓인다. 중간 산출물
(`structure-report.json`, `mapping.json`, `screens.json`, `summaries.json`,
`default-template.pptx`, `images/`)은 그 안의 `.work/`에 모인다 — 폴더째 지우면
이 Excel의 산출물이 전부 사라진다.

```
output/
├─ 입력.pptx          ← 결과물
└─ .work/             ← 중간 산출물
```

경로는 스크립트가 정한다. 다른 위치의 매핑이나 screens를 쓰려면 `--mapping`,
`--screens`로 짚어 줄 수 있고, 그때는 그 경로를 그대로 쓴다. 예전 구조(`output/`
루트에 중간 산출물이 있던 폴더)로 돌리면 첫 실행 때 `.work/`로 옮기고 이어서
진행한다.

## 설계

**`screens.json`은 Excel에서 뽑은 중간 산출물이다.** 사람이 손으로 고치는 파일이 아니라서 `extract.py`는 재추출 때 그냥 덮어쓴다 — 추출이 어긋나면 이 파일이 아니라 매핑을 고친다. 그래도 파일로 남기는 이유는 셋이다:

- `build.py`는 Excel을 모른다. openpyxl 의존은 추출 계열 모듈에만 갇힌다 (테스트가 소스를 검사해 강제). 덕분에 빌드 로직을 Excel 픽스처 없이 단위 테스트할 수 있다
- 이미지 추출이 느리다. 매핑이나 템플릿만 고쳤으면 생성 단계만 다시 돌린다
- 결과가 이상할 때 추출이 틀렸는지 배치가 틀렸는지 갈라 볼 수 있다

**슬라이드를 만드는 방법이 두 가지다.** `clone` 모드는 템플릿의 예시 슬라이드를 복제하고, `layout` 모드는 레이아웃으로 빈 슬라이드를 만든 뒤 이미지 자리와 상세표를 그 자리에서 그린다. 값을 채우는 코드는 둘이 공유한다 — `layout` 모드가 새로 만든 자리에 매핑이 정한 이름을 붙이기 때문이다. 기본 템플릿은 `layout` 모드를 쓴다.

**사진과 표 자리는 모든 화면에서 같다.** 본문 영역 높이에 대한 고정 비율이다(`slide_layout.ratio_row_heights`). 내용이 길다고 표를 늘리지 않는다 — 늘리면 화면마다 사진 크기가 달라져 장을 넘길 때 상단 기준선이 흔들린다. 넘치면 글자를 줄이고(7 → 6.5 → 6pt), 그래도 넘치면 `text-overflow` 경고로 알린다.

**긴 상세는 `summaries.json`으로 줄여 넣는다.** Excel 원문은 화면설계서에 그대로 싣기엔 길다(실제 샘플 평균 94자). 추출과 생성 사이에서 Claude가 요소명 + 개조식 한 줄로 줄여 `.work/summaries.json`에 남기면 `build.py`가 반영한다. `{"화면ID": {"상세번호": "요약문"}}` 꼴이고, 파일이 없으면 원문을 그대로 쓴다. 상세번호는 Excel 값이라 재추출해도 키가 유지된다.

**값이 없는 자리는 비우지 않는다.** 빨강·굵게 `입력필요`를 남긴다 — 빈 칸은 채우는 걸 잊은 건지 원래 값이 없는 건지 구별되지 않는다. 쪽번호 같은 자동 필드만 예외다.

**화면 단위로 예외를 격리한다.** 화면 40개 중 3개가 실패해도 나머지는 완성되고, 실패한 화면은 `[생성 실패]` 표시로 남는다.

**생성물을 재파싱해 검증한다.** 파일이 만들어졌다는 사실만으로는 부족하다 — 도형 이름이 안 맞거나 관계가 깨지면 파일은 생기지만 내용이 비어 있을 수 있다.

## 구조

```
skills/excel-wireframe/
├── SKILL.md                      # Claude가 읽는 절차서
├── references/mapping-schema.md  # mapping.json / screens.json 레퍼런스
├── user-default.example.json     # 조직 기본 템플릿 경로 설정의 예시
├── user-default.json             # 실제 설정 (gitignore — 남의 저작권 pptx를 가리킨다)
├── assets/                       # 설정이 가리키는 템플릿 사본 (gitignore)
└── scripts/
    ├── common.py            # UTF-8 stdio, JSON, 경고 수집, EMU 상수
    ├── xlsx_scan.py         # Excel 구조 스캔
    ├── pptx_scan.py         # PPTX 구조 스캔, clone/layout 모드 판정
    ├── default_template.py  # 템플릿 미제공 시 쓸 기본 템플릿 생성
    ├── analyze.py           # [CLI] 구조 리포트
    ├── xlsx_read.py         # 화면·상세 읽기 (sheet-per-screen / table)
    ├── xlsx_meta.py         # 표지 시트에서 문서 단위 정보 추출
    ├── xlsx_images.py       # 삽입 이미지 추출 (openpyxl + zip 폴백)
    ├── extract.py           # [CLI] screens.json + images/
    ├── slide_clone.py       # 슬라이드 복제·삭제 (XML deepcopy + rId 재매핑)
    ├── slide_layout.py      # 레이아웃 탐색, placeholder 상속·명명·정리, 메타 표 자리,
    │                        #   본문 영역 분할 (행 높이·글자 폭의 단일 출처)
    ├── slide_fill.py        # 도형 텍스트 주입, 이미지 배치, 표 슬롯 채우기
    ├── text_metrics.py      # 줄 수·넘침 계산 (순수 함수)
    ├── image_split.py       # 긴 스크린샷 자동 분할
    ├── user_default.py      # user-default.json 기본 템플릿 설정
    ├── verify.py            # 생성물 재파싱 검증
    └── build.py             # [CLI] PPT 생성, 넘침 분할
```

## 주요 동작

**Excel 양식 두 가지를 지원한다.** `sheet-per-screen`(1시트 = 1화면)과 `table`(1행 = 1화면). `mapping.excel.layout` 한 필드로 구분한다.

**상세 표는 슬롯 모델이다.** 실무 템플릿은 상세 표가 하나가 아니라 4행짜리 표 5개가 가로로 놓여 20슬롯을 이룬다. 상세 항목은 표1 r0…r3 → 표2 r0…r3 순으로 흘러 들어간다. 슬롯을 넘으면 다음 슬라이드로 분할되고 제목에 `(1/2)`, `(2/2)`가 붙는다.

**상세 번호는 Excel 값을 그대로 쓴다.** 스크린샷의 SoM 뱃지와 대응하므로 재부여하지 않는다.

**서식을 보존한다.** 셀에 값을 쓸 때 런을 새로 만들지 않고 첫 런의 텍스트만 교체한다. 런을 새로 만들면 폰트·크기·색이 초기화되어 템플릿 디자인이 무너진다.

**표지 시트의 문서 정보를 반영한다.** 프로젝트명·작성일 같은 값을 읽어 `meta`에 담고, `template.shapes`에 같은 이름의 도형이 있으면 채운다. 화면별 `fields`가 문서 `meta`보다 우선한다.

**기본 템플릿의 상단은 메타 정보 표다.** 단색 띠가 아니라 프로젝트명·산출물명·화면명·ID·버전·작성자·검토자·작성일·수정일·네비게이션·화면유형·알림여부를 담는 두 개의 표(`메타표1`, `메타표2`)다. 레이아웃 위의 표는 슬라이드로 상속되지 않으므로(PowerPoint는 placeholder만 물려준다) 표를 복제해 덮지 않고 `slide_layout.add_meta_text_slots`로 칸 자리에 글자만 올린다 — 복제하면 표가 두 겹이 되어 원본을 클릭할 수 없고 테두리가 두 번 그려진다.

**조직 템플릿을 기본으로 둘 수 있다.** 스킬 디렉토리에 `user-default.json`을 두면 `--template` 없이 실행할 때 거기 적힌 템플릿과 레이아웃을 쓴다. 설정이 없을 때만 `default_template.py`가 기본 템플릿을 만든다. 남의 저작권 표기가 박힌 실물 템플릿을 저장소에 담을 수 없어서 생긴 우회로다 — 형식은 `user-default.example.json` 참고.

## 경고 코드

한 화면이 실패해도 나머지는 완성하고, 실패는 경고로 남긴다. 코드는 아홉 개뿐이다.

| 코드 | 뜻 |
|---|---|
| `no-image` | 화면에 연결된 이미지를 찾지 못했다 |
| `no-detail` | 상세 표 헤더를 찾지 못했다 |
| `text-overflow` | 글자를 줄여도 설명이 셀을 넘길 수 있다 |
| `shape-not-found` | 매핑이 가리킨 도형·표·placeholder가 템플릿에 없다 |
| `slide-split` | 상세가 슬롯을 넘어 여러 장으로 나눴다 |
| `slot-shortage` | 상세 중 일부만 이 슬라이드에 들어갔다 |
| `screen-failed` | 이 화면의 슬라이드 생성이 실패했다 |
| `image-convert-failed` | 이미지 변환에 실패했다 |
| `orphan-row` | 화면ID가 없는 행이라 건너뛰었다 |

## 테스트

```bash
python -m pytest -q                              # 전체
python -m pytest tests/test_build.py -v          # 개별
python -m pytest --override-ini="addopts=" -v    # 개별 PASSED 줄 보기
```

`pytest.ini`가 `addopts = -q`를 설정하므로 명령줄 `-v`가 무시된다.

실제 문서를 쓰는 회귀 테스트(`tests/test_sample_e2e.py`)는 파일이 없으면 자동으로 skip한다. 저장소에는 저작권 문제로 포함하지 않는다.

## 현재 한계

- **사용자 템플릿의 레이아웃에 박힌 요소는 건드리지 못한다.** 상단 띠·하단 바 같은 것이 슬라이드가 아니라 레이아웃에 있으면 PowerPoint에서 직접 고쳐야 한다. 다만 기본 템플릿은 껍데기가 애초에 레이아웃에 있으므로, PowerPoint에서 레이아웃 `내용설명연결` 하나만 고치면 전 슬라이드에 반영된다.
- **메타 표의 칸은 열두 개로 고정이다.** 표지에서 그보다 많은 값을 읽어와도 꽂을 자리가 없다. 템플릿의 표에 칸을 늘리고 `meta_table.labels`에 라벨을 추가하면 코드 수정 없이 채워진다.
- **기본 템플릿에는 python-pptx 기본 레이아웃 열 개가 함께 남는다.** 레이아웃을 이름으로 찾으므로 동작에는 영향이 없지만, PowerPoint에서 템플릿을 열면 쓰지 않는 레이아웃이 보인다.
- **상세표는 항상 사진 아래에 놓인다.** 스크린샷이 세로로 길면 좌우가 통째로 비고 `image_split`이 불필요하게 여러 조각으로 자른다. 표를 옆으로 옮기는 설계는 `docs/superpowers/specs/2026-08-12-table-position-design.md`에 적혀 있으나 구현하지 않았다 — 표가 5개에서 4개로 줄어드는 것을 받아들일지가 미결이다.
- `options.overflow`는 매핑에 자리만 있고 코드가 읽지 않는다. 분할은 항상 일어난다.

## 문서

- 설계: `docs/superpowers/specs/` — 최근 것은 기본 템플릿 재구조화(`2026-08-18-out-template-restructure-design.md`)와 아직 구현하지 않은 상세표 위치(`2026-08-12-table-position-design.md`)다
- 구현 계획: `docs/superpowers/plans/`
- 프로젝트 규칙: `CLAUDE.md`
- Claude가 읽는 절차서: `skills/excel-wireframe/SKILL.md`
