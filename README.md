# excel-wireframe

Excel로 작성된 화면설계서를 PowerPoint 화면설계서로 자동 생성하는 Claude Skill.

특정 Excel 양식에 하드코딩하지 않는다. 처음 보는 양식이 들어와도 구조를 분석해 매핑을 도출하고, 사용자 확인을 거쳐 PPT를 만든다. PPT 템플릿을 주면 그 디자인을 그대로 유지하고, 주지 않으면 표준 구조의 기본 템플릿을 생성해 쓴다.

## 설치

```bash
cp -r skills/excel-wireframe ~/.claude/skills/
```

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
python $S/analyze.py --excel 입력.xlsx --template 템플릿.pptx --out work/structure-report.json

# 2. 리포트를 읽고 work/mapping.json 작성 — 사람이 판단하는 단계

# 3. 추출
python $S/extract.py --excel 입력.xlsx --mapping work/mapping.json --work work

# 4. 생성
python $S/build.py --screens work/screens.json --mapping work/mapping.json --work work --out 결과.pptx
```

## 설계

**`screens.json`이 SSOT(Single Source of Truth)다.** Excel은 최초 임포트 소스일 뿐이고, PPT의 내용은 `screens.json`에서 온다. 따라서:

- `build.py`는 Excel을 모른다. openpyxl 의존은 추출 계열 모듈에만 갇힌다 (테스트가 소스를 검사해 강제)
- Excel 없이 `screens.json`만 직접 써도 PPT가 나온다
- 빌드 로직을 Excel 픽스처 없이 단위 테스트할 수 있다

**`extract.py`는 기존 `screens.json`을 덮어쓰지 않는다.** 이미 있으면 `screens.new.json`에 쓰고 차이를 보고한다. 사람이 손본 내용이 재추출로 소실되지 않는다.

**화면 단위로 예외를 격리한다.** 화면 40개 중 3개가 실패해도 나머지는 완성되고, 실패한 화면은 `[생성 실패]` 표시로 남는다.

**생성물을 재파싱해 검증한다.** 파일이 만들어졌다는 사실만으로는 부족하다 — 도형 이름이 안 맞거나 관계가 깨지면 파일은 생기지만 내용이 비어 있을 수 있다.

## 구조

```
skills/excel-wireframe/
├── SKILL.md                      # Claude가 읽는 절차서
├── references/mapping-schema.md  # mapping.json / screens.json 레퍼런스
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
    ├── slide_clone.py       # 슬라이드 복제 (XML deepcopy + rId 재매핑)
    ├── slide_fill.py        # 도형 텍스트 주입, 이미지 배치, 표 슬롯 채우기
    ├── verify.py            # 생성물 재파싱 검증
    └── build.py             # [CLI] PPT 생성, 넘침 분할
```

## 주요 동작

**Excel 양식 두 가지를 지원한다.** `sheet-per-screen`(1시트 = 1화면)과 `table`(1행 = 1화면). `mapping.excel.layout` 한 필드로 구분한다.

**상세 표는 슬롯 모델이다.** 실무 템플릿은 상세 표가 하나가 아니라 4행짜리 표 5개가 가로로 놓여 20슬롯을 이룬다. 상세 항목은 표1 r0…r3 → 표2 r0…r3 순으로 흘러 들어간다. 슬롯을 넘으면 다음 슬라이드로 분할되고 제목에 `(1/2)`, `(2/2)`가 붙는다.

**상세 번호는 Excel 값을 그대로 쓴다.** 스크린샷의 SoM 뱃지와 대응하므로 재부여하지 않는다.

**서식을 보존한다.** 셀에 값을 쓸 때 런을 새로 만들지 않고 첫 런의 텍스트만 교체한다. 런을 새로 만들면 폰트·크기·색이 초기화되어 템플릿 디자인이 무너진다.

**표지 시트의 문서 정보를 반영한다.** 프로젝트명·작성일 같은 값을 읽어 `meta`에 담고, `template.shapes`에 같은 이름의 도형이 있으면 채운다. 화면별 `fields`가 문서 `meta`보다 우선한다.

## 테스트

```bash
python -m pytest -q                              # 전체
python -m pytest tests/test_build.py -v          # 개별
python -m pytest --override-ini="addopts=" -v    # 개별 PASSED 줄 보기
```

`pytest.ini`가 `addopts = -q`를 설정하므로 명령줄 `-v`가 무시된다.

실제 문서를 쓰는 회귀 테스트(`tests/test_sample_e2e.py`)는 파일이 없으면 자동으로 skip한다. 저장소에는 저작권 문제로 포함하지 않는다.

## 현재 한계

- **`layout` 모드는 생성을 지원하지 않는다.** 빈 레이아웃만 있는 템플릿에는 상세를 넣을 표가 없다. `analyze.py`가 그렇게 판정하면 예시 페이지가 있는 템플릿을 쓰거나 `--template`을 생략한다.
- **사용자 템플릿의 레이아웃에 박힌 요소는 건드리지 못한다.** 상단 띠·하단 저작권 바 같은 것이 슬라이드가 아니라 레이아웃에 있으면 PowerPoint에서 직접 고쳐야 한다.
- **기본 템플릿의 meta 도형은 두 개뿐이다** (`문서제목`, `작성일`). 표지에서 더 많은 값을 읽어와도 꽂을 자리가 없다. 템플릿에 같은 이름의 도형을 추가하면 코드 수정 없이 채워진다.
- `options.overflow`는 매핑에 자리만 있고 코드가 읽지 않는다. 분할은 항상 일어난다.

## 문서

- 설계: `docs/superpowers/specs/`
- 구현 계획: `docs/superpowers/plans/`
- 프로젝트 규칙: `CLAUDE.md`
