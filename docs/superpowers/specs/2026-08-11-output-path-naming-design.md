# 결과물 저장 위치와 파일 이름을 코드가 정한다

## 문제

결과물 경로가 SKILL.md의 지시문 한 줄로만 정해져 있다.

```bash
python <스킬>/scripts/build.py --screens work/screens.json --mapping work/mapping.json \
  --work work --out work/output/화면설계서.pptx
```

`화면설계서.pptx`가 고정 문자열이다. 세 가지가 문제다.

1. **재실행이 이전 산출물을 조용히 덮어쓴다.** 매핑을 고쳐 5단계를 다시 돌리면 직전
   결과물이 사라진다. 비교할 수단이 없다.
2. **어느 Excel에서 나온 결과물인지 이름에 없다.** 입력이 여러 개면 파일명만 보고
   구별할 수 없다.
3. **이름을 정하는 일이 LLM 판단으로 들어가 있다.** 스킬 지시문에 적힌 경로를
   Claude가 매번 타이핑한다. CLAUDE.md는 "판단은 세 곳에서만 한다, 나머지는
   스크립트가 결정론적으로 처리한다"고 정하고 있는데 이건 네 번째 판단이다.

작업 디렉토리 이름도 어긋나 있다. `work/` 안에 중간 산출물과 `output/` 하위
결과물이 섞여 있어 결과물이 한 단계 더 깊이 묻힌다.

## 설계

### 1. 작업 디렉토리를 `output/`으로 통일한다

`work/`를 `output/`으로 바꾸고, 결과물 pptx를 그 안에 **직접** 놓는다. `output/output/`
같은 중첩을 만들지 않는다.

```
입력폴더/
├─ 짧은 버전.xlsx
└─ output/
   ├─ structure-report.json
   ├─ mapping.json
   ├─ screens.json
   ├─ images/
   ├─ default-template.pptx
   └─ 짧은 버전.pptx      ← 결과물
```

중간 산출물과 결과물을 굳이 가르지 않는다. 중간 산출물은 재추출 때 그냥 덮어쓰는
캐시이고 사람이 열어 볼 일이 드물다. 한 폴더에 두면 "이 Excel의 산출물은 여기 전부"가
성립해서 지우기도 쉽다.

### 2. `common.py`에 `resolve_output_path`를 추가한다

```python
def resolve_output_path(output_dir: Path, source_excel: str) -> Path:
    """결과물 pptx 경로를 정한다.

    이름은 원본 Excel 파일명을 그대로 쓴다. 같은 이름이 이미 있으면 뒤에 2, 3…을
    붙여 비어 있는 첫 번호를 쓴다 — 재실행이 이전 산출물을 덮어쓰지 않게 한다.
    """
```

- `source_excel`은 `screens.json`의 `meta.source`에서 온다. `extract.py`가 원본
  Excel 경로를 거기 넣어 두므로 `build.py`가 Excel을 다시 열 필요가 없다.
- `meta.source`가 없거나 비었으면 `화면설계서`로 폴백한다. 손으로 만든
  `screens.json`이나 이 변경 이전에 만들어진 파일에는 없을 수 있고, 그때 크래시가
  나면 안 된다.
- 번호는 구분자 없이 붙인다: `짧은 버전.pptx` → `짧은 버전2.pptx` → `짧은 버전3.pptx`.
- `output_dir`이 아직 없어도 된다. 후보가 하나도 존재하지 않으므로 첫 이름을 돌려주고,
  디렉토리 생성은 저장 직전에 `build()`가 이미 하고 있다. 이 함수는 디렉토리를
  만들지 않는다 — 경로를 정하는 일과 만드는 일을 섞지 않는다.

`resolve_template_path` 옆에 둔다. 둘 다 매핑·중간산출물에 적힌 값을 실제 경로로
푸는 같은 일이고, 순수 함수라 `tests/test_common.py`에서 파일시스템만 놓고 단위
테스트할 수 있다.

### 3. CLI 인자를 개명한다

| 스크립트 | 이전 | 이후 |
|---|---|---|
| `extract.py` | `--work <디렉토리>` (필수) | `--output <디렉토리>` (필수) |
| `build.py` | `--work <디렉토리>` (필수) | `--output <디렉토리>` (필수) |
| `build.py` | `--out <파일>` (필수) | `--out-file <파일>` (**선택**) |
| `analyze.py` | `--out <리포트.json>` | 변경 없음 |

`--out`을 `--out-file`로 개명하는 이유는 `--output`(디렉토리)과 `--out`(파일)이
나란히 있으면 이름만 보고 어느 쪽이 폴더인지 알 수 없기 때문이다. `analyze.py`의
`--out`은 리포트 파일 경로이고 그 스크립트에는 디렉토리 인자가 없어 혼동이
생기지 않으므로 그대로 둔다.

`--out-file`을 주면 그 경로를 그대로 쓰고 **번호를 붙이지 않는다.** 사용자가 파일명을
직접 정한 상황에서는 덮어쓰기가 의도한 동작이고, 명시적 지정을 코드가 뒤집으면
결과를 예측할 수 없다.

### 4. 내부 파라미터명도 `output_dir`로 통일한다

`build()`, `verify_output()`, `resolve_template_path()`, `plan_pages()`,
`fill_slide()`의 `work_dir` → `output_dir`. 기계적 치환이고 호출부가 모두 위치인자라
동작 위험이 없다. 절반만 개명하면 다음에 코드를 읽는 사람이 `--output`과 `work_dir`
두 이름을 다 기억해야 한다.

### 5. 번호가 붙었으면 알린다

```
저장: output/짧은 버전2.pptx (같은 이름이 있어 번호를 붙였습니다)
```

조용히 다른 이름으로 저장하면 사용자가 옛 파일을 열어 보고 "반영이 안 됐다"고
오해한다. 번호가 안 붙었을 때는 지금처럼 `저장: <경로>`만 찍는다.

경고 코드는 늘리지 않는다 — 번호 붙이기는 정상 동작이라 경고 대상이 아니다.
**경고 코드는 아홉 개를 유지한다.**

## 테스트

`tests/test_common.py`에 `resolve_output_path` 단위 테스트를 넣는다.

- 폴더가 비어 있으면 `<stem>.pptx`
- `<stem>.pptx`가 있으면 `<stem>2.pptx`
- `<stem>.pptx`와 `<stem>2.pptx`가 있으면 `<stem>3.pptx`
- `source_excel`이 빈 문자열이면 `화면설계서.pptx`
- `source_excel`이 절대경로여도 stem만 쓴다

`tests/test_build.py`에 CLI 통합 테스트를 넣는다.

- `--out-file` 없이 호출하면 `--output` 안에 Excel 이름으로 저장된다
- 같은 인자로 두 번 호출하면 두 파일이 남는다 (덮어쓰지 않는다)
- `--out-file`을 주면 그 경로에 저장되고 번호가 붙지 않는다

기존 테스트의 `--work` → `--output` 치환: `test_extract.py`(8곳),
`test_extract_build_meta.py`(2곳), `test_sample_e2e.py`(2곳).

## 문서

`SKILL.md`의 `work/` → `output/` (24·32·37·78·98행 및 본문의 `work/` 언급 전부),
5단계 명령에서 `--out` 제거. `README.md`도 같이(29·34·37행).

`docs/superpowers/plans/*.md`는 과거 기록이므로 손대지 않는다.

## 하지 않는 것

- 결과물 파일명에 날짜를 넣지 않는다. Excel 파일명만 쓴다
- 중간 산출물과 결과물을 별도 폴더로 가르지 않는다
- `analyze.py`의 `--out`을 개명하지 않는다
- 번호를 위해 새 경고 코드를 만들지 않는다
