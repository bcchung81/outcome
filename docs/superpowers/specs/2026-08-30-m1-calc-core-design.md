# M1 산출 코어 설계 — AI 성과측정 시스템

> '26.8.30 · 브레인스토밍 결과. 근거 문서: `docs/PRD_AI성과측정시스템.md` v0.2, `docs/61_산출방식.md`(산식 v2),
> `docs/60_산식_고도화안_v2.md`, `docs/20_draft.md`. 참조 구현: `docs/52_측정_자동계산_양식.xlsx`, `docs/53_sessionize.py`,
> 기대값: `docs/51_가상측정_산출대장_v2.xlsx`.
> 이 문서는 산식을 재정의하지 않는다. 산식이 침묵하는 경계와 구현 구조만 정한다.

## 0. 한 장 요약

| 항목 | 결정 |
|---|---|
| 범위 | PRD **M1** = 순수 Python 산출 라이브러리 + 검증 규칙 전부 + 세션 변환 + CLI. 웹(M2~)은 별도 |
| 입력 | 제출 양식 xlsx(`52` 현행본 포함) **와** CSV 4종 둘 다. 어댑터가 정규 CSV로 통일 |
| 산출 단위 | 제출·검증은 (과제, 월). **산출은 연도 누적** Dataset(1월~기준월) |
| 골든 | `52` 원자료 → 3원 검증(라이브러리·독립 스크립트·LibreOffice) 일치 후 `golden/v2.0` 동결 |
| 의존성 | `outcome_core` 0개. `outcome_cli`만 openpyxl |
| 완료 기준 | PRD A1·A2·A3(CLI)·A8. 첫 실용 산출은 #1(9월 시범) |

## 1. 현황 분석에서 확인한 사실

- `52`의 캐시값은 전부 0(재계산본 아님). 기대값은 `51`에만 있다.
- `52` 원자료를 v2 산식으로 재산출하면 **운영형 12건은 `51`과 소수점 일치**, 절감형 5건은 다른 난수분이라 불일치
  (#1 245.17h vs 220.21h). PRD 부록 D가 지정한 골든(`sim2/calc.py`)은 존재하지 않는다.
- `52` 수식과 문서의 불일치 3건: 관측 개월(+1일 없음), ⑤ 이관율 분자에 환경·ERRV 필터 없음, 도입 전 총량 `K5:K9` 고정.
- `52` ①로그 마지막에 「기재요령」 텍스트 행이 A열 데이터로 존재한다.
- `52` ①로그는 세션 단위 13열이라 턴 로그(이용자ID·턴번호·질의문)를 담을 수 없다.
- `52`·`51`의 기대값은 9~12월 **4개월 누적**으로 계산된 것이다(관측 개월 3.99). PRD S1의 월별 제출과 정합하지 않는다.

## 2. 이번 설계에서 새로 정한 규칙 (산식 문서 정정 대상 — D1에 포함)

| # | 규칙 | 내용 | 영향 |
|---|---|---|---|
| R1 | 연누적 산출 | Dataset = (연도, 기준월). 관측 개월·기간 상한·30건 판정 모두 누적 기준. 월별 값은 가산 지표(순절감·처리 건수)에 한해 `Calc(m) − Calc(m−1)`로 파생. 비율 지표는 누적만 공시 | PRD §6·§10, E06 범위 → 연도 내 |
| R2 | 시각 절사 | 모든 시각은 파서에서 초 단위로 절사(소수 초 버림). `52_v2.1` 양식도 동일 절사 | A6 오차 |
| R3 | 관측 개월 | PRD §10: `((MAX(start) − MIN(start)).days + 1) / 30.4375`, 최소 `1/30.4375`, 운영 행만 | `52` 수식 정정 |
| R4 | 정수 초 합산 | 기계·개입 총합은 정수 초로 더한 뒤 마지막에 ÷60. 부동소수 합은 `math.fsum` | 결정론 |
| R5 | 등급 재판정 | #11·#16은 `AI 출력등급 == 현장 정답등급`으로 「정확」을 재판정. 인적확인결과와 다르면 W19 | `61` §3② |
| R6 | 이관율 분자 | `담당자이관` ∧ 운영 ∧ ≠ERRV. 분모 운영 ∧ ≠ERRV 세션 | `52` ⑤ 수식 정정 |
| R7 | 읽기 종료 | 과제번호 열이 정수가 아닌 첫 행에서 읽기 종료. 그 뒤 행 수를 W17로 표기 | 파서 |
| R8 | 파일 범위 | 선택 과제 외 과제번호 행 → **E17** 오류. 양식 버전·시트명·헤더 불일치 → **E18** 오류. 구양식(`52` 현행) → W18 경고 | 규칙 카탈로그 |
| R9 | 파일의 마스터 무시 | 제출 파일의 과제마스터 시트는 읽지 않는다. 도입 전 값은 시스템 `master.csv`만 | G1·조작 차단 |

## 3. 아키텍처

```
outcome/
├─ docs/                     기존 문서(불변) + superpowers/specs
├─ src/outcome_core/         ★ 의존성 0. 스냅샷 재산출에 그대로 동봉
│   ├─ model.py              LogRow·TurnRow·VerifyRow·MonthlyRow·TaskMaster·TaskSetting·Dataset·Result·Aggregate
│   ├─ parse.py              시각 8형식·엑셀 시리얼·절사(R2)·별칭 정규화(§7.7)·ParseIssue
│   ├─ validate/             규칙 레지스트리. 규칙 1개 = 함수 1개. row→file→cross 3단계
│   ├─ sessionize.py         53 승계 함수. 규칙 7개 + 50턴 상한 + 질의문 폐기
│   ├─ calc/                 savings · quality · response · monthly · flags · aggregate · engine
│   └─ version.py            FORMULA_VERSION="2.0" · RULESET_VERSION(레지스트리 해시) · ALIAS_VERSION
├─ src/outcome_cli/          ★ openpyxl 허용
│   ├─ adapters/             xlsx52.py(구양식·v3 양식) · csv.py → 공통 RawTable
│   ├─ commands/             validate · submit · calc · master · compare · golden · recompute
│   └─ report/               intermediates.json · ledger.csv/.xlsx · aggregate.json/.md · flags.md · report.csv/.md
├─ templates/                제출양식_v3.xlsx · 52_v2.1(참조용 산출본, 수식 정정)
├─ golden/v2.0/              inputs/ · expected/ · manifest.json
├─ store/                    (런타임) {연}/{과제}/{월}/v{n}/ 정규 CSV + meta.json, 추가 전용
└─ tests/                    unit · independent · golden · session · adapter
```

경계: `outcome_core`는 파일 경로가 아니라 **행 iterable**을 받는다. 어댑터만 파일을 안다.

## 4. 데이터 모델

| 타입 | 필드 |
|---|---|
| `LogRow` | task, proc_id, start, end, result, env, q=1, review_open, review_save, session_id, first_resp, escal, requery, src_row |
| `TurnRow` | task, session_id, user_id, turn_no, start, end, first_resp, result, env, escal, query, src_row |
| `VerifyRow` | task, proc_id, sample_kind(탐지·비탐지·판정), outcome, ai_grade, truth_grade, src_row |
| `MonthlyRow` | task, ym(`YYYY-MM`\|`도입전`), item, value, src_row |
| `TaskMaster` | task, name, type, category, unit, pre_per_unit, pre_annual, evidence, approved, approved_at, applies_from, requires{log,verify,monthly,session}, metrics[] |
| `TaskSetting` | task, idle_min=30, requery_thr=0.60, max_turns=50, review_rate_thr=0.80 |
| `Dataset` | year, upto_month, logs, verifies, monthlies, master, settings |
| `Result` | task, metrics{11종→값\|None}, intermediates(PRD 부록 C + capped·clip 등), flags[{id,reason}], warnings[], versions{}, inputs[{path,sha256,rows}] |
| `Aggregate` | year, upto_month, axes{직원시간·대국민·운영품질·참고}, 각 축 {numer, denom, value, n_tasks, excluded[], notes[]} |

파서는 정규화만 하고 판단하지 않는다. 해석 실패는 `ParseIssue(row, col, raw, rule_id)`로 넘겨 규칙 엔진이 E02·E04 등으로 판정한다.

## 5. 검증 규칙 레지스트리

```python
@rule("E12", scope="log", stage="row", since="1.0")
def rej_without_review(row: LogRow, ctx) -> Issue | None: ...
```

- 단계: `row`(스트리밍) → `file`(E01·E06·E16·E17·E18) → `cross`(E08·E10·E11·W15 등 타 자료 참조)
- `RULESET_VERSION` = 레지스트리 내용 해시. 규칙이 바뀌면 자동 변경(FR-2.4)
- 산출 단계 플래그(W04~W13)는 `calc/flags.py`가 같은 `Issue` 타입으로 낸다
- 규칙 목록: PRD §9 E01~E16·W01~W16 + E17·E18·W17·W18·W19(본 문서 §2)
- 오류 ≥1 → Submission `검증실패`, 산출로 넘어가지 않음

## 6. 산출 엔진

### 6.1 실행 순서
`Dataset` → 과제별 분기(master.type·category) → `savings | quality | response(+monthly) | count-only` → `flags` → `aggregate`.
각 모듈은 `(rows, master, setting) → Result` 순수 함수. 공통 전처리 `filter_live` = 운영 ∧ ≠ERRV.

### 6.2 절감 계열 (`61` §4-1 + PRD §10 + R1~R4)
- 처리 건수 = `len(live)`; Σq(열 없으면 1); 기계 = Σ(end−start) 정수 초; 개입 = review_open 있는 행만 `min(save−open, pre×3×q)`, 클립 행·초 기록
- 관측 개월 R3; 기간 상한 = `pre_annual × months / 12`; 절감 적용 단위 = `min(Σq − Σq[REJ], 상한)`, 상한 적용 시 `capped=True`
- 순절감(분) 원값 저장. 반올림은 표시 계층에서만. 과제 단축률 = `(pre−post)/pre`(참고값)

### 6.3 품질·응답 계열
- 표본구분별 분자·분모 따로. 표본 0 → `None`. 등급 과제는 R5 재판정
- 이관율 R6. 대기시간 = 중앙값(짝수는 가운데 두 값 평균), first_resp 없는 행 제외, 전무하면 W13
- #14 문의감소율: `도입전` 행 필수(E11), 도입 후 = 기준월까지 동기 합. #15 대기시간 = 월별집계 값

### 6.4 플래그
| 플래그 | 조건 | 값 | 전사 집계 |
|---|---|---|---|
| 잠정 | 절감형 처리 건수 < 30 (W04) | 유지 | 포함 + 건수 병기 |
| 미확정 | 순절감 < 0 (W10) | 음수 유지 | 포함 + 플래그 |
| T2미승인 | master.approved=False | 유지 | 포함 + 건수 병기 |
| 미공시 | 표본구분별 < 30 (W05) / 세션 < 30 (W06) | `None` | 분자·분모 양쪽 제외 |

### 6.5 전사 집계
- 축별 `numer, denom, value, excluded[]` 원값. 단축률 = Σ순절감 ÷ Σ(pre×applied), 잠정·미확정 포함
- 최대 과제 비중 50% 초과 시 note. `Target` 없으면 목표 대비 `None`
- 월별 차분은 가산 지표만(R1)

### 6.6 결정론
정렬 키 명시 `(task, proc_id)` / 세션 `(task, session_id, start, turn_no)`; `math.fsum`; 정수 초; `Result` JSON 키 정렬 직렬화.

## 7. 세션 변환
`sessionize(turns, setting) -> (list[LogRow], SessionReport)`. `53`의 규칙 7개 그대로. 추가: 재질의 비교 최근 `max_turns`턴, `query`는 판정 직후 폐기하고 폐기 건수만 리포트. A8 대조 테스트는 50턴 미만 데이터로 `53_sessionize.py`와 세션 수·이관·재질의·시각 일치 확인.

## 8. CLI

| 명령 | 입력 → 출력 |
|---|---|
| `validate <파일> --task N --month YYYY-MM` | 어댑터 → 규칙. `report.csv`·`report.md`. 오류 시 종료코드 2 |
| `submit <파일> --task N --month YYYY-MM --store DIR` | validate 통과 시 `store/{연}/{과제}/{월}/v{n}/` 정규 CSV + `meta.json`(SHA-256·업로드자·시각·버전 3종). 추가 전용 |
| `calc --store DIR --year Y --upto M` | Dataset(CLOSED.json 있으면 그 버전, 없으면 최신 확정본=가집계) → `results/{연}-{월}/` |
| `master …` | `master.csv` 등록·승인 편집(새 버전) |
| `compare --xlsx 52_v2.1.xlsx --results DIR` | `soffice --headless` 재계산 후 ④⑤ 값과 A6 대조 |
| `golden freeze \| check` | `golden/v2.0` 동결·회귀 |
| `recompute <dir>` | `outcome_core`만으로 재산출(M1에선 calc 별칭) |

## 9. 산출물 규격
- `intermediates.json`: PRD 부록 C + versions·inputs·flags. 키 정렬·UTF-8·소수 원값
- `ledger.csv`: 20행 × 지표 11종 + 플래그. `None`=빈칸, 「미적용」은 문자열
- `ledger.xlsx`: `51` 열 순서, 값만
- `aggregate.md`: `20_draft` 보고 축 표 형태. 병기 항목(과제 수·도입 전 총량·증거 등급·최대 비중) 자동 채움
- 제출 양식 v3: `①로그`·`①-턴로그`·`②검증대장`·`③월별집계`·`기재요령`. 3행 헤더·4행부터 데이터. 수식·마스터·산출 시트 없음. 데이터 유효성만
- `52_v2.1`: 수식 3건 정정 + 절사. `compare` 전용

## 10. 테스트

| 층 | 방법 | 기준 |
|---|---|---|
| 단위 | 규칙 39개·산식 항·세션 규칙 7·파서. 규칙당 정상·경계·오류 케이스 (`tests/fixtures/` 손 작성) | 통과 |
| 독립 대조 | `tests/independent/calc_ref.py` — `outcome_core` 미import 별도 구현 | A6: 시간·건수 ≤0.01, 비율 ≤0.01%p |
| 골든 회귀 | `golden check` | 완전 일치 |
| 교차 | `compare`(LibreOffice) | A6 |
| 세션 | `53` 실행 결과 대조 | A8 |

## 11. 골든 v2.0 동결 절차
1. `52` 현행본 → 구양식 어댑터 → 17과제 정규 CSV. `master.csv`는 `52` 과제마스터 값으로 작성
2. R2 절사 규칙 확정 후 정규 CSV·`52_v2.1` 양쪽 적용 — 이 전엔 동결하지 않음
3. 라이브러리·독립 스크립트·LibreOffice 3원 일치(A6). 불일치는 원인 문서화 후 수정 대상 결정
4. 운영형 12건 `51` 소수점 일치 재확인
5. `golden freeze` → `manifest.json`(SHA-256·버전 3종·일시) 커밋. 산식 변경 시 `golden/v2.1/` 추가(교체 아님)

## 12. 구현 순서

| 주차 | 산출물 | 완료 판정 |
|---|---|---|
| 1 | 저장소 · model · parse · csv 어댑터 · 레지스트리 뼈대 + row 규칙(E02~E05·E12·W03·W14) | 단위 테스트 |
| 2 | savings · flags(절감) · xlsx52(구양식) · validate/submit/calc 최소형 | #1을 `52`에서 읽어 산출한 값이 독립 스크립트(`tests/independent/calc_ref.py`)와 A6 일치 |
| 3 | file·cross 규칙 전부 · quality · monthly · aggregate | 17과제 대장·집계, 규칙 테스트 |
| 4 | response · sessionize · A8 · 양식 v3 · 52_v2.1 · compare | 3원 대조 → **골든 동결** |
| 5 | ledger.xlsx · aggregate.md · master · 마감 파일 · README · 버전 노트 · D1 정정 목록 | A1·A2·A3(CLI)·A8 = M1 완료 |

## 13. 비범위
웹 UI(M2~), 가명화 파이프라인(M4 — 단 질의문 폐기는 포함), SSO·알림·대시보드, 스냅샷 반출 절차(D4), 목표값 결정(D2).

## 14. M1 이후로 넘기는 결정
- D1 확장: 산식 v2 + R1~R9 + `52` 수식 3건을 `20_draft`·`61`·PRD에 반영
- D2 목표 재승인 · D5 임계 초기값(설정 기본값으로 착수) · D7 부서 로그 필드 현황

## 15. 화면안(참고)
M2용 목업은 별도 artifact로 확정됨 — 담당자 1화면(엑셀 1파일 → 결과 → 확정), 관리자 2화면(월 마감·과제 마스터), 조회 3화면. 상단 헤더 탭. M1의 `xlsx52` 어댑터·`store` 구조가 그대로 M2 업로드·DB의 원형이다.
