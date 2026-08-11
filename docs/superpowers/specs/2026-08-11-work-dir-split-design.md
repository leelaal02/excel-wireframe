# 결과물 폴더에는 pptx만 두고 나머지는 `.work/`로 모은다

## 문제

`output/` 한 폴더에 여섯 가지가 섞여 있다.

```
output/
├─ structure-report.json
├─ mapping.json
├─ screens.json
├─ default-template.pptx
├─ images/
└─ 짧은 버전.pptx      ← 사용자가 찾는 건 이거 하나
```

사용자가 열어 볼 파일은 pptx 하나인데 폴더를 열면 여섯 개가 보인다. 특히
`default-template.pptx`는 결과물과 확장자가 같아서 어느 쪽이 산출물인지 파일
목록만 보고 알 수 없다. 결과물을 두 번 만들어 `짧은 버전2.pptx`가 생기면 pptx가
세 개가 되고, 그 중 하나만 열어 볼 대상이다.

바로 앞 스펙(`2026-08-11-output-path-naming-design.md`)은 "중간 산출물과 결과물을
별도 폴더로 가르지 않는다"고 정했다. **이 결정을 뒤집는다.** 그때의 근거는 "한
폴더에 두면 이 Excel의 산출물은 여기 전부가 성립해서 지우기 쉽다"였는데, 작업
폴더를 `output/` **안에** 두면 그 성질은 그대로 유지된다. 폴더째 지우면 전부
사라진다. 가르는 비용 없이 결과물만 보이게 할 수 있다.

## 설계

### 1. `output/.work/`에 중간 산출물을 모은다

```
입력폴더/
├─ 짧은 버전.xlsx
└─ output/                 ← 결과물 폴더 (--output)
   ├─ 짧은 버전.pptx        ← 여기엔 pptx만
   └─ .work/               ← 작업 폴더 (코드가 유도)
      ├─ structure-report.json
      ├─ mapping.json
      ├─ screens.json
      ├─ default-template.pptx
      └─ images/
```

`.work`는 점으로 시작해 탐색기와 `ls`에서 기본으로 숨는다. 의도한 것이다 —
사람이 평소에 열어 볼 폴더가 아니다.

### 2. 기준 경로가 둘로 갈린다

앞 스펙에서 `work_dir`를 `output_dir`로 통일했는데, 이제 두 개념이 다시 갈린다.

- **`output_dir`** — 결과물 pptx가 놓이는 곳. `--output`으로 받는다.
- **`work_dir`** — 이미지·템플릿·중간 JSON을 푸는 기준. `output_dir / ".work"`.

지금 `output_dir`를 받는 함수들은 **전부 후자**다. 하는 일이 이미지 상대경로와
템플릿 경로를 푸는 것이기 때문이다. 그래서 아래를 `work_dir`로 되돌린다.

| 함수 | 지금 | 이후 |
|---|---|---|
| `build()` | `output_dir` | `work_dir` |
| `plan_pages()` | `output_dir` | `work_dir` |
| `fill_slide()` | `output_dir` | `work_dir` |
| `verify_output()` | `output_dir` | `work_dir` |
| `resolve_template_path()` | `output_dir` | `work_dir` |

`build()`는 결과물 경로를 이미 `out_path`로 따로 받고 있다. 개명만으로 두 개념이
시그니처에서 갈라지고, "이 함수가 푸는 건 작업 파일이지 결과물이 아니다"가
이름에 드러난다. `resolve_output_path()`만 `output_dir`를 받는다.

`screens.json`에 적히는 이미지 경로 표기(`images/xxx.png`)는 그대로 둔다. 상대
경로의 기준 폴더만 바뀌므로 파일 내용은 손대지 않는다.

### 3. `common.py`에 두 함수를 추가한다

```python
def work_dir(output_dir: Path) -> Path:
    """작업 폴더 경로. output/.work.

    디렉토리를 만들지 않는다 — 경로를 정하는 일과 만드는 일을 섞지 않는다.
    """


def migrate_legacy_work(output_dir: Path) -> list[str]:
    """output/ 루트에 남은 구버전 중간 산출물을 .work/로 옮긴다.

    옮긴 이름 목록을 돌려준다. 옮길 게 없으면 빈 리스트.
    """
```

`work_dir`는 `resolve_output_path` 옆에 둔다. 둘 다 순수 함수라 파일시스템만
놓고 단위 테스트할 수 있다.

`migrate_legacy_work`가 옮기는 대상은 **고정된 화이트리스트뿐이다.**

```
structure-report.json, mapping.json, screens.json, default-template.pptx, images/
```

목록에 없는 파일은 건드리지 않는다. 결과물 pptx가 딸려 들어갈 여지를 없애려면
"pptx가 아닌 것을 옮긴다" 같은 규칙으로는 안 된다 — `default-template.pptx`가
pptx이고, 사용자가 결과물 이름을 직접 지어 놨을 수도 있다. 옮길 것을 이름으로
못 박는 편이 안전하다.

같은 이름이 `.work/`에 이미 있으면 옮기지 않고 루트 쪽을 그대로 둔다. `.work/`
쪽이 새 경로에서 만들어진 최신 파일이므로 구버전이 덮어써서는 안 된다.

세 CLI 진입점이 작업 폴더를 얻을 때 한 번씩 호출한다. 옮긴 게 있으면 한 줄
출력한다.

```
기존 작업 파일 3개를 .work/로 옮겼습니다: mapping.json, screens.json, images/
```

조용히 파일을 움직이면 사용자가 사라졌다고 오해한다. 경고 코드는 늘리지
않는다 — 마이그레이션은 정상 동작이다. **경고 코드는 아홉 개를 유지한다.**

### 4. CLI 인자를 선택으로 낮춘다

| 스크립트 | 이전 | 이후 |
|---|---|---|
| `analyze.py` | `--out <파일>` (필수) | `--output <디렉토리>`와 `--out <파일>` 중 하나 |
| `extract.py` | `--mapping <파일>` (필수) | 선택 — 생략하면 `.work/mapping.json` |
| `extract.py` | `--output <디렉토리>` (필수) | 변경 없음 |
| `build.py` | `--screens <파일>` (필수) | 선택 — 생략하면 `.work/screens.json` |
| `build.py` | `--mapping <파일>` (필수) | 선택 — 생략하면 `.work/mapping.json` |
| `build.py` | `--output <디렉토리>` (필수) | 변경 없음 |
| `build.py` | `--out-file <파일>` (선택) | 변경 없음 |

경로를 지시문에 적어 두면 Claude가 매번 타이핑한다. CLAUDE.md는 "판단은 세
곳에서만 하고 나머지는 스크립트가 결정론적으로 처리한다"고 정하고 있다. 중간
산출물이 어디 있는지는 판단할 거리가 아니라 규칙이므로 코드가 유도한다.

인자를 없애지 않고 선택으로 남기는 이유는 다른 위치의 매핑을 써 보는 일이
실제로 있기 때문이다. 주면 그 경로를 그대로 쓰고, 생략하면 `.work/`에서 찾는다.

`analyze.py`는 `--output`과 `--out` 중 하나를 받는다. 둘 다 없으면 인자 오류로
죽인다 — 리포트를 어디에 쓸지 모르는 채로 스캔을 시작하지 않는다.

- `--output`을 주면 리포트는 `.work/structure-report.json`, 기본 템플릿은
  `.work/default-template.pptx`로 간다. 스킬이 쓰는 경로다.
- `--out`을 주면 지금처럼 그 파일 경로에 리포트를 쓰고, 기본 템플릿은 그
  파일과 같은 폴더에 만든다. 리포트 경로를 직접 받아 쓰는 기존 호출부와
  테스트가 있어서 남긴다.
- 둘 다 주면 `--out`이 이긴다. 명시 지정이 유도된 경로를 이긴다는 규칙은
  `build.py`의 `--out-file`과 같다.

## 테스트

`tests/test_common.py`

- `work_dir(Path("output"))`이 `output/.work`
- `migrate_legacy_work`가 화이트리스트 다섯 항목을 옮긴다
- 결과물 pptx(`짧은 버전.pptx`)는 루트에 남는다
- `.work/`에 같은 이름이 있으면 덮어쓰지 않고 루트 쪽을 남긴다
- 옮길 게 없으면 빈 리스트를 돌려주고 `.work/`를 만들지 않는다

`tests/test_build.py`

- `--screens`/`--mapping` 없이 `--output`만 주면 `.work/`에서 읽는다
- 결과물 pptx가 `output/` 루트에 생기고 `.work/`에는 안 생긴다
- `--screens`/`--mapping`을 주면 그 경로를 쓴다

`tests/test_extract.py`

- `--mapping` 없이 돌리면 `.work/mapping.json`을 읽고 `.work/screens.json`과
  `.work/images/`를 쓴다

기존 테스트의 경로 기대값을 `.work/` 기준으로 갱신한다: `test_extract.py`,
`test_extract_build_meta.py`, `test_sample_e2e.py`, `test_analyze.py`,
`test_verify.py`, `test_xlsx_images.py`.

## 문서

`SKILL.md`의 세 명령이 짧아진다.

```bash
python <스킬>/scripts/analyze.py --excel <입력.xlsx> --output output
python <스킬>/scripts/extract.py --excel <입력.xlsx> --output output
python <스킬>/scripts/build.py --output output
```

본문의 경로 언급(`output/mapping.json`, `output/screens.json`,
`output/structure-report.json`, `output/default-template.pptx`, `output/images/`)을
`output/.work/...`로 고친다. 결과물 경로 `output/<Excel 이름>.pptx`는 그대로다.

`CLAUDE.md`의 "작업 디렉토리는 `output/` 하나다" 규칙을 "결과물은 `output/`,
중간 산출물은 `output/.work/`"로 갱신한다. `references/mapping-schema.md`의
`template.file` 예시도 `.work/` 기준으로 맞춘다.

`docs/superpowers/plans/*.md`와 앞 스펙 문서는 과거 기록이므로 손대지 않는다.

## 하지 않는 것

- `.work` 이름을 설정 가능하게 만들지 않는다
- 중간 산출물에 번호를 붙이거나 이력을 남기지 않는다 — 캐시라 덮어쓰는 게 맞다
- `output/`을 자동으로 청소하지 않는다
- 마이그레이션을 위해 새 경고 코드를 만들지 않는다
