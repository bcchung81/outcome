# M1 Plan 1 — 기반 · 절감 산출 · CLI 최소형 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `52_측정_자동계산_양식.xlsx`(또는 CSV)에서 절감 실적형 과제(#1~#5)의 로그를 읽어 검증(행 단계 규칙)하고 v2 산식으로 순절감·단축률을 산출해, 독립 구현과 허용오차 내 일치하는 결과를 CLI로 낸다.

**Architecture:** `outcome_core`(의존성 0)는 행 iterable만 받는 순수 함수 모음 — 파서(정규화) → 규칙 레지스트리(판정) → 산출(savings) → 플래그. `outcome_cli`는 파일 어댑터(csv·xlsx)와 명령(validate·submit·calc)만 담당한다. 산출은 연도 누적 `Dataset` 단위다.

**Tech Stack:** Python ≥ 3.11 표준 라이브러리. 테스트 pytest. CLI 어댑터만 openpyxl.

**Spec:** `docs/superpowers/specs/2026-08-30-m1-calc-core-design.md` (이하 "스펙"). 산식 원문: `docs/61_산출방식.md` §4, 엣지: `docs/PRD_AI성과측정시스템.md` §10.

## Global Constraints

- `src/outcome_core/**`는 **표준 라이브러리 외 import 금지** (스펙 §0). 테스트로 강제한다(Task 1).
- 모든 시각은 파서에서 **초 단위 절사**(스펙 R2). `datetime`에 microsecond가 남아 있으면 버그.
- 시간 합산은 **정수 초**로 하고 마지막에 ÷60 (R4). 부동소수 합은 `math.fsum`.
- 관측 개월 = `((max(start) − min(start)).days + 1) / 30.4375`, 최소 `1/30.4375`, 운영 행만 (R3).
- 제출 파일의 과제마스터 시트는 읽지 않는다 (R9). 도입 전 값은 `master.csv`만.
- 과제번호 열이 정수가 아닌 첫 행에서 읽기 종료, 이후 행 수는 W17 (R7).
- 정규 CSV 헤더는 `52` ①로그 열 이름 그대로 13개(아래 `LOG_COLS`). UTF-8-BOM.
- 커밋 메시지 끝에 항상:
  ```
  Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01TYQXWxwFwwiTeMqbCSHoWg
  ```
- 결과값 원값은 반올림하지 않는다. 표시 반올림은 이 계획 범위 밖.

## 파일 구조 (이 계획에서 만드는 것)

```
pyproject.toml
src/outcome_core/__init__.py
src/outcome_core/version.py          FORMULA_VERSION, MONTH_DAYS
src/outcome_core/model.py            dataclass들 + Issue + LOG_COLS
src/outcome_core/parse.py            parse_ts, canon_env, canon_result, canon_escal, to_log_rows
src/outcome_core/validate/__init__.py  rule 데코레이터, REGISTRY, run_rules, ruleset_version
src/outcome_core/validate/log_rows.py  E02 E03 E04 E05 E12 W03 W14
src/outcome_core/calc/__init__.py
src/outcome_core/calc/common.py      filter_live, observation_months
src/outcome_core/calc/savings.py     calc_savings
src/outcome_core/calc/flags.py       apply_savings_flags
src/outcome_core/calc/engine.py      run(dataset) → dict[task, Result]
src/outcome_cli/__init__.py
src/outcome_cli/adapters/__init__.py
src/outcome_cli/adapters/rawtable.py RawTable
src/outcome_cli/adapters/csvfile.py  read_csv
src/outcome_cli/adapters/xlsx52.py   read_52_log
src/outcome_cli/store.py             submit_to_store, load_dataset
src/outcome_cli/__main__.py          validate / submit / calc
tests/conftest.py
tests/test_no_third_party.py
tests/unit/test_parse.py
tests/unit/test_rules_log.py
tests/unit/test_savings.py
tests/unit/test_flags.py
tests/adapter/test_csvfile.py
tests/adapter/test_xlsx52.py
tests/independent/calc_ref.py        outcome_core를 import하지 않는 참조 구현
tests/independent/test_against_ref.py
tests/cli/test_commands.py
```

---

### Task 1: 프로젝트 뼈대 · 버전 상수 · 의존성 0 강제 테스트

**Files:**
- Create: `pyproject.toml`, `src/outcome_core/__init__.py`, `src/outcome_core/version.py`, `src/outcome_cli/__init__.py`, `tests/conftest.py`, `tests/test_no_third_party.py`, `.gitignore`

**Interfaces:**
- Produces: `outcome_core.version.FORMULA_VERSION: str = "2.0"`, `MONTH_DAYS: float = 30.4375`, `MIN_N: int = 30`

- [ ] **Step 1: 뼈대 파일 작성**

`pyproject.toml`:
```toml
[project]
name = "outcome"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = ["openpyxl>=3.1"]

[project.optional-dependencies]
dev = ["pytest>=8"]

[project.scripts]
outcome = "outcome_cli.__main__:main"

[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[tool.setuptools.packages.find]
where = ["src"]

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["src"]
```

`src/outcome_core/version.py`:
```python
"""버전·상수. 산식 버전은 61_산출방식.md v2 기준."""
FORMULA_VERSION = "2.0"
MONTH_DAYS = 30.4375          # PRD §10 관측 개월 분모
MIN_N = 30                    # 잠정·미공시 판정 기준 건수
```

`src/outcome_core/__init__.py`, `src/outcome_cli/__init__.py`: 빈 파일.

`.gitignore`:
```
__pycache__/
*.egg-info/
.pytest_cache/
store/
results/
```

`tests/conftest.py`:
```python
import pathlib
import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent

@pytest.fixture
def docs_dir() -> pathlib.Path:
    return ROOT / "docs"
```

- [ ] **Step 2: 의존성 0 테스트 작성**

`tests/test_no_third_party.py`:
```python
"""outcome_core는 표준 라이브러리만 import한다 (스펙 §0, PRD §12)."""
import ast, pathlib, sys

STDLIB = set(sys.stdlib_module_names)
CORE = pathlib.Path(__file__).resolve().parent.parent / "src" / "outcome_core"

def _imports(path: pathlib.Path):
    tree = ast.parse(path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                yield a.name.split(".")[0]
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            yield node.module.split(".")[0]

def test_core_uses_only_stdlib():
    bad = []
    for py in CORE.rglob("*.py"):
        for mod in _imports(py):
            if mod not in STDLIB and mod != "outcome_core":
                bad.append((py.name, mod))
    assert bad == []
```

- [ ] **Step 3: 설치·실행**

Run: `python3 -m venv .venv && . .venv/bin/activate && pip install -e '.[dev]' && pytest -q`
Expected: `1 passed`

- [ ] **Step 4: Commit**

```bash
git add pyproject.toml .gitignore src tests
git commit -m "chore: 프로젝트 뼈대, 버전 상수, outcome_core 의존성 0 강제 테스트"
```

---

### Task 2: 데이터 모델

**Files:**
- Create: `src/outcome_core/model.py`
- Test: `tests/unit/test_model.py`

**Interfaces:**
- Produces:
  - `LOG_COLS: list[str]` = 13개 정규 열 이름
  - `@dataclass LogRow(task:int, proc_id:str, start:datetime, end:datetime, result:str, env:str, q:int=1, review_open:datetime|None=None, review_save:datetime|None=None, session_id:str="", first_resp:datetime|None=None, escal:str="", requery:str="", src_row:int=0)`
  - `@dataclass TaskMaster(task:int, name:str, type:str, category:str, unit:str, pre_per_unit:float|None, pre_annual:float|None, evidence:str, approved:bool, approved_at:str="", applies_from:str="")`
  - `@dataclass TaskSetting(task:int, idle_min:int=30, requery_thr:float=0.60, max_turns:int=50, review_rate_thr:float=0.80)`
  - `@dataclass Issue(rule:str, level:str, row:int, col:str, value:str, reason:str)` — `level ∈ {"E","W"}`
  - `@dataclass Dataset(year:int, upto_month:int, logs:list[LogRow], master:dict[int,TaskMaster], settings:dict[int,TaskSetting])`
  - `@dataclass Result(task:int, metrics:dict[str,float|None], intermediates:dict, flags:list[dict], warnings:list[Issue], versions:dict, inputs:list[dict])`
  - `Result.to_json() -> str` (키 정렬, ensure_ascii=False)

- [ ] **Step 1: 테스트 작성**

`tests/unit/test_model.py`:
```python
from datetime import datetime
from outcome_core.model import LOG_COLS, LogRow, Result, Issue

def test_log_cols_match_52_header_order():
    assert LOG_COLS == ["과제번호","처리번호","시작시각","종료시각","처리결과","환경구분","처리수량",
                        "검수열림시각","저장제출시각","세션ID","첫응답시각","이관사유","재질의여부"]

def test_logrow_defaults():
    r = LogRow(task=1, proc_id="P1-1", start=datetime(2026,9,1), end=datetime(2026,9,1,0,1), result="OK", env="운영")
    assert r.q == 1 and r.review_open is None and r.escal == ""

def test_result_json_is_sorted_and_utf8():
    res = Result(task=1, metrics={"net_h": 1.5, "rate_pct": None}, intermediates={"b": 1, "a": 2},
                 flags=[], warnings=[Issue("W04","W",0,"","","n<30")], versions={"formula":"2.0"}, inputs=[])
    s = res.to_json()
    assert s.index('"a"') < s.index('"b"')
    assert "n<30" in s and "\\u" not in s
```

- [ ] **Step 2: 실패 확인**

Run: `pytest tests/unit/test_model.py -q`
Expected: FAIL — `ModuleNotFoundError: outcome_core.model`

- [ ] **Step 3: 구현**

`src/outcome_core/model.py`:
```python
"""정규화된 행·마스터·결과 타입. 파일 형식을 모른다."""
from __future__ import annotations
import json
from dataclasses import dataclass, field, asdict
from datetime import datetime

LOG_COLS = ["과제번호","처리번호","시작시각","종료시각","처리결과","환경구분","처리수량",
            "검수열림시각","저장제출시각","세션ID","첫응답시각","이관사유","재질의여부"]

RESULT_CODES = ("OK", "REV", "REJ", "ERRV")
ENV_CODES = ("운영", "시험")
ESCAL_CODES = ("담당자이관", "미해결종료", "검수반려")

@dataclass
class LogRow:
    task: int
    proc_id: str
    start: datetime
    end: datetime
    result: str
    env: str
    q: int = 1
    review_open: datetime | None = None
    review_save: datetime | None = None
    session_id: str = ""
    first_resp: datetime | None = None
    escal: str = ""
    requery: str = ""
    src_row: int = 0

@dataclass
class TaskMaster:
    task: int
    name: str
    type: str            # 절감 실적형 · 운영 실적형 · 측정 제외
    category: str        # 문서·자료 처리 · 감시·경보 · 판독·진단 · 응답·안내 · 예측·배정
    unit: str
    pre_per_unit: float | None
    pre_annual: float | None
    evidence: str = ""
    approved: bool = False
    approved_at: str = ""
    applies_from: str = ""

@dataclass
class TaskSetting:
    task: int
    idle_min: int = 30
    requery_thr: float = 0.60
    max_turns: int = 50
    review_rate_thr: float = 0.80

@dataclass
class Issue:
    rule: str
    level: str           # "E" | "W"
    row: int
    col: str
    value: str
    reason: str

@dataclass
class Dataset:
    year: int
    upto_month: int
    logs: list[LogRow]
    master: dict[int, TaskMaster]
    settings: dict[int, TaskSetting] = field(default_factory=dict)

@dataclass
class Result:
    task: int
    metrics: dict[str, float | None]
    intermediates: dict
    flags: list[dict]
    warnings: list[Issue]
    versions: dict
    inputs: list[dict]

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, sort_keys=True, indent=2, default=str)
```

- [ ] **Step 4: 통과 확인**

Run: `pytest tests/unit/test_model.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outcome_core/model.py tests/unit/test_model.py
git commit -m "feat(core): 데이터 모델 (LogRow·TaskMaster·Issue·Dataset·Result)"
```

---

### Task 3: 파서 — 시각 절사·별칭 정규화·행 변환

**Files:**
- Create: `src/outcome_core/parse.py`
- Test: `tests/unit/test_parse.py`

**Interfaces:**
- Consumes: `model.LogRow`, `model.Issue`, `model.LOG_COLS`
- Produces:
  - `parse_ts(v: str|float|datetime|None) -> datetime|None` — 8형식+엑셀 시리얼, **초 절사**, 타임존 접미 무시. 해석 불가 → `None` (빈값도 `None`)
  - `has_tz_suffix(v: str) -> bool`
  - `canon_env(v) -> str|None` — "운영"/"시험"/None(미인식)
  - `canon_result(v) -> str|None` — OK·REV·REJ·ERRV 또는 None
  - `canon_escal(v) -> str` — 표준 3종 또는 ""(빈값·미인식 모두 "")
  - 처리수량: 열이 **없으면** 전 행 1. 열이 **있는데** 빈값·0·비정수면 `P_Q`(→E05) — PRD §7.1 그대로. 골든(`52`)은 빈값 0건 확인됨
  - `to_log_rows(header: list[str], rows: list[list], first_src_row: int) -> tuple[list[LogRow], list[Issue], int]`
    반환: (행, 파싱 이슈, 읽기 종료 이후 남은 행 수). 파싱 이슈의 rule은 `"P_TS"`(시각)·`"P_INT"`(정수)·`"P_ENV"`·`"P_RESULT"`·`"P_ESCAL"`·`"P_TZ"`·`"P_Q"` — 규칙 엔진(Task 4)이 E/W로 매핑한다. 파싱 실패 행은 결과 목록에서 **제외**하되 이슈는 남긴다.

- [ ] **Step 1: 테스트 작성**

`tests/unit/test_parse.py`:
```python
from datetime import datetime
from outcome_core.parse import parse_ts, canon_env, canon_result, canon_escal, to_log_rows, has_tz_suffix
from outcome_core.model import LOG_COLS

def test_parse_ts_formats_and_truncation():
    assert parse_ts("2026-09-11 14:23:05.987") == datetime(2026,9,11,14,23,5)
    assert parse_ts("2026/09/11 14:23") == datetime(2026,9,11,14,23)
    assert parse_ts("2026-09-11T14:23:05") == datetime(2026,9,11,14,23,5)
    assert parse_ts("20260911142305") == datetime(2026,9,11,14,23,5)
    assert parse_ts(datetime(2026,9,11,14,23,5,435000)) == datetime(2026,9,11,14,23,5)
    assert parse_ts("") is None and parse_ts(None) is None and parse_ts("abc") is None

def test_parse_ts_excel_serial_and_tz():
    assert parse_ts(46276.5) == datetime(2026,9,11,12,0,0)          # 1899-12-30 기준
    assert parse_ts("2026-09-11T14:23:05+09:00") == datetime(2026,9,11,14,23,5)
    assert has_tz_suffix("2026-09-11T14:23:05+09:00") and has_tz_suffix("2026-09-11T14:23:05Z")
    assert not has_tz_suffix("2026-09-11 14:23:05")

def test_canon_values():
    assert canon_env("prod") == "운영" and canon_env("stg") == "시험" and canon_env("uat") is None
    assert canon_result(" ok ") == "OK" and canon_result("errv") == "ERRV" and canon_result("fail") is None
    assert canon_escal("상담원 연결") == "담당자이관" and canon_escal("") == "" and canon_escal("기타") == ""

def _row(**kw):
    base = {"과제번호":"1","처리번호":"P1-1","시작시각":"2026-09-01 09:00:00","종료시각":"2026-09-01 09:02:00",
            "처리결과":"OK","환경구분":"운영","처리수량":"1","검수열림시각":"","저장제출시각":"",
            "세션ID":"","첫응답시각":"","이관사유":"","재질의여부":""}
    base.update(kw)
    return [base[c] for c in LOG_COLS]

def test_to_log_rows_basic_and_q_default():
    rows, issues, rest = to_log_rows(LOG_COLS, [_row(), _row(처리번호="P1-2", 처리수량="3")], first_src_row=4)
    assert [r.proc_id for r in rows] == ["P1-1","P1-2"]
    assert rows[0].q == 1 and rows[1].q == 3 and rows[0].src_row == 4 and issues == [] and rest == 0

def test_to_log_rows_stops_at_non_integer_task():
    rows, issues, rest = to_log_rows(LOG_COLS, [_row(), ["기재요령 — …"]+[""]*12, _row(처리번호="P1-9")], first_src_row=4)
    assert len(rows) == 1 and rest == 2

def test_to_log_rows_collects_parse_issues():
    rows, issues, _ = to_log_rows(LOG_COLS, [_row(종료시각="abc"), _row(처리번호="P1-2", 처리결과="fail"),
                                             _row(처리번호="P1-3", 처리수량="0")], first_src_row=4)
    assert len(rows) == 0
    assert {(i.rule, i.row, i.col) for i in issues} == {("P_TS",4,"종료시각"),("P_RESULT",5,"처리결과"),("P_Q",6,"처리수량")}
```

- [ ] **Step 2: 실패 확인**

Run: `pytest tests/unit/test_parse.py -q`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: 구현**

`src/outcome_core/parse.py`:
```python
"""원문 값 → 정규 값. 판단하지 않고 정규화와 해석 실패 기록만 한다."""
from __future__ import annotations
import re
from datetime import datetime, timedelta
from .model import LogRow, Issue, LOG_COLS, RESULT_CODES

TS_FORMATS = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M",
              "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y%m%d%H%M%S", "%Y-%m-%d %H:%M:%S.%f",
              "%Y-%m-%dT%H:%M:%S.%f"]
_TZ_RE = re.compile(r"(Z|[+-]\d{2}:?\d{2})$")
_EXCEL_EPOCH = datetime(1899, 12, 30)

# 53_sessionize.py ALIAS 승계 (PRD §7.7)
TEST_ENV = ["시험","테스트","개발","검증","test","testing","stg","stage","staging","dev","development","qa","sandbox"]
PROD_ENV = ["운영","prod","production","live","real","본번","운영계"]
ESCAL_ALIAS = {
    "담당자이관": ["담당자이관","담당자 이관","이관","escalate","escalated","transfer","상담원연결","상담원 연결"],
    "미해결종료": ["미해결종료","미해결 종료","미해결","unresolved","abandoned","이탈","중도이탈"],
    "검수반려":   ["검수반려","검수 반려","반려","reject","rejected"],
}

def norm(s) -> str:
    return re.sub(r"[\s_\-]", "", str(s or "")).lower()

def has_tz_suffix(v) -> bool:
    return isinstance(v, str) and bool(_TZ_RE.search(v.strip()))

def parse_ts(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.replace(microsecond=0)
    if isinstance(v, (int, float)):
        d = _EXCEL_EPOCH + timedelta(days=float(v))
        return d.replace(microsecond=0)
    s = str(v).strip()
    if not s:
        return None
    s = _TZ_RE.sub("", s)
    for f in TS_FORMATS:
        try:
            return datetime.strptime(s, f).replace(microsecond=0)
        except ValueError:
            pass
    try:
        return (_EXCEL_EPOCH + timedelta(days=float(s))).replace(microsecond=0)
    except ValueError:
        return None

def canon_env(v):
    n = norm(v)
    if n in {norm(x) for x in TEST_ENV}:
        return "시험"
    if n in {norm(x) for x in PROD_ENV}:
        return "운영"
    return None

def canon_result(v):
    u = str(v or "").strip().upper()
    return u if u in RESULT_CODES else None

def canon_escal(v) -> str:
    n = norm(v)
    if not n:
        return ""
    for std, cands in ESCAL_ALIAS.items():
        if n in {norm(c) for c in cands}:
            return std
    return ""

def _int_or_none(v):
    if v is None or str(v).strip() == "":
        return None
    try:
        f = float(str(v).strip())
    except ValueError:
        return None
    return int(f) if f == int(f) else None

def to_log_rows(header: list[str], rows: list[list], first_src_row: int):
    idx = {h: i for i, h in enumerate(header)}
    has_q = "처리수량" in idx
    def g(r, col):
        i = idx.get(col)
        return r[i] if i is not None and i < len(r) else ""
    out, issues = [], []
    stopped_at = None
    for k, r in enumerate(rows):
        src = first_src_row + k
        task = _int_or_none(g(r, "과제번호"))
        if task is None:
            stopped_at = k
            break
        bad = False
        def ts(col, required):
            nonlocal bad
            raw = g(r, col)
            if raw is None or str(raw).strip() == "":
                if required:
                    issues.append(Issue("P_TS", "E", src, col, "", "필수 시각이 비어 있다")); bad = True
                return None
            if has_tz_suffix(raw):
                issues.append(Issue("P_TZ", "W", src, col, str(raw), "타임존 접미를 무시하고 KST로 읽었다"))
            d = parse_ts(raw)
            if d is None:
                issues.append(Issue("P_TS", "E", src, col, str(raw), "시각 형식을 해석할 수 없다")); bad = True
            return d
        start = ts("시작시각", True); end = ts("종료시각", True)
        res = canon_result(g(r, "처리결과"))
        if res is None:
            issues.append(Issue("P_RESULT", "E", src, "처리결과", str(g(r, "처리결과")), "OK·REV·REJ·ERRV 외의 값")); bad = True
        env = canon_env(g(r, "환경구분"))
        if env is None:
            issues.append(Issue("P_ENV", "W", src, "환경구분", str(g(r, "환경구분")), "운영·시험으로 인식되지 않아 운영으로 처리"))
            env = "운영"
        q = 1
        if has_q:
            raw_q = g(r, "처리수량")
            q = _int_or_none(raw_q)
            if q is None or q < 1:
                issues.append(Issue("P_Q", "E", src, "처리수량", str(raw_q), "처리수량 열이 있는데 빈값·0 이하·비정수")); bad = True
        esc_raw = g(r, "이관사유"); esc = canon_escal(esc_raw)
        if str(esc_raw or "").strip() and not esc:
            issues.append(Issue("P_ESCAL", "W", src, "이관사유", str(esc_raw), "표준 3종으로 매핑되지 않음"))
        ro = ts("검수열림시각", False); rs = ts("저장제출시각", False); fr = ts("첫응답시각", False)
        if bad:
            continue
        out.append(LogRow(task=task, proc_id=str(g(r, "처리번호")).strip(), start=start, end=end, result=res, env=env,
                          q=q, review_open=ro, review_save=rs, session_id=str(g(r, "세션ID") or "").strip(),
                          first_resp=fr, escal=esc, requery=str(g(r, "재질의여부") or "").strip().upper(), src_row=src))
    rest = 0 if stopped_at is None else len(rows) - stopped_at
    return out, issues, rest
```

- [ ] **Step 4: 통과 확인**

Run: `pytest tests/unit/test_parse.py -q`
Expected: `6 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outcome_core/parse.py tests/unit/test_parse.py
git commit -m "feat(core): 파서 — 시각 8형식·초 절사·별칭 정규화·행 변환(R2·R7)"
```

---

### Task 4: 검증 규칙 레지스트리 + 행 단계 규칙 7개

**Files:**
- Create: `src/outcome_core/validate/__init__.py`, `src/outcome_core/validate/log_rows.py`
- Test: `tests/unit/test_rules_log.py`

**Interfaces:**
- Consumes: `model.LogRow`, `model.Issue`
- Produces:
  - `rule(id: str, scope: str, stage: str, since: str)` 데코레이터. `REGISTRY: dict[str, RuleSpec]`, `RuleSpec(id, scope, stage, since, fn)`
  - `map_parse_issues(issues: list[Issue]) -> list[Issue]` — `P_TS→E02`, `P_RESULT→E04`, `P_Q→E05`, `P_ENV→W03`, `P_TZ→W14`, `P_ESCAL→W02`
  - `run_row_rules(rows: list[LogRow], scope="log") -> list[Issue]`
  - `ruleset_version() -> str` — 등록된 규칙 소스의 SHA-256 앞 12자
  - 규칙 함수 시그니처: `fn(row: LogRow) -> Issue | None`
  - 이 태스크의 규칙: **E03**(종료<시작, 저장<열림, 첫응답<시작), **E12**(REJ·REV인데 검수열림 없음 / 열림 있는데 저장 없음; ERRV 행은 검사 안 함). E02·E04·E05·W03·W14는 파서 이슈 매핑으로 성립.

- [ ] **Step 1: 테스트 작성**

`tests/unit/test_rules_log.py`:
```python
from datetime import datetime as D
from outcome_core.model import LogRow, Issue
from outcome_core.validate import REGISTRY, run_row_rules, map_parse_issues, ruleset_version
import outcome_core.validate.log_rows  # noqa: F401  규칙 등록

def R(**kw):
    base = dict(task=1, proc_id="P", start=D(2026,9,1,9,0), end=D(2026,9,1,9,5), result="OK", env="운영", src_row=4)
    base.update(kw); return LogRow(**base)

def test_registry_has_row_rules():
    assert {"E03","E12"} <= set(REGISTRY)
    assert REGISTRY["E03"].stage == "row" and REGISTRY["E03"].scope == "log"

def test_e03_end_before_start():
    out = run_row_rules([R(end=D(2026,9,1,8,59))])
    assert [(i.rule, i.col) for i in out] == [("E03","종료시각")]

def test_e03_save_before_open_and_first_resp_before_start():
    out = run_row_rules([R(review_open=D(2026,9,1,9,6), review_save=D(2026,9,1,9,5)),
                         R(first_resp=D(2026,9,1,8,0))])
    assert [(i.rule, i.col) for i in out] == [("E03","저장제출시각"),("E03","첫응답시각")]

def test_e12_rej_without_review_and_open_without_save():
    out = run_row_rules([R(result="REJ"), R(result="REV"), R(review_open=D(2026,9,1,9,6)),
                         R(result="ERRV"), R(result="REJ", review_open=D(2026,9,1,9,6), review_save=D(2026,9,1,9,9))])
    assert [(i.rule, i.col) for i in out] == [("E12","검수열림시각"),("E12","검수열림시각"),("E12","저장제출시각")]

def test_map_parse_issues():
    src = [Issue("P_TS","E",4,"종료시각","x","..."), Issue("P_ENV","W",5,"환경구분","uat","..."),
           Issue("P_TZ","W",6,"시작시각","..Z","..."), Issue("P_Q","E",7,"처리수량","0","..."),
           Issue("P_RESULT","E",8,"처리결과","f","..."), Issue("P_ESCAL","W",9,"이관사유","기타","...")]
    assert [i.rule for i in map_parse_issues(src)] == ["E02","W03","W14","E05","E04","W02"]

def test_ruleset_version_is_stable_hex():
    v = ruleset_version()
    assert len(v) == 12 and v == ruleset_version()
```

- [ ] **Step 2: 실패 확인**

Run: `pytest tests/unit/test_rules_log.py -q`
Expected: FAIL — `ModuleNotFoundError: outcome_core.validate`

- [ ] **Step 3: 구현**

`src/outcome_core/validate/__init__.py`:
```python
"""검증 규칙 레지스트리. 규칙 1개 = 함수 1개. RULESET_VERSION은 규칙 소스 해시."""
from __future__ import annotations
import hashlib, inspect
from dataclasses import dataclass
from typing import Callable
from ..model import Issue, LogRow

@dataclass(frozen=True)
class RuleSpec:
    id: str
    scope: str      # log | verify | monthly | master
    stage: str      # row | file | cross
    since: str
    fn: Callable

REGISTRY: dict[str, RuleSpec] = {}

def rule(id: str, scope: str, stage: str, since: str = "1.0"):
    def deco(fn):
        REGISTRY[id] = RuleSpec(id, scope, stage, since, fn)
        return fn
    return deco

PARSE_MAP = {"P_TS": "E02", "P_RESULT": "E04", "P_Q": "E05", "P_ENV": "W03", "P_TZ": "W14", "P_ESCAL": "W02"}

def map_parse_issues(issues: list[Issue]) -> list[Issue]:
    return [Issue(PARSE_MAP[i.rule], i.level, i.row, i.col, i.value, i.reason) for i in issues]

def run_row_rules(rows: list[LogRow], scope: str = "log") -> list[Issue]:
    specs = [s for s in REGISTRY.values() if s.scope == scope and s.stage == "row"]
    specs.sort(key=lambda s: s.id)
    out: list[Issue] = []
    for r in rows:
        for s in specs:
            got = s.fn(r)
            if got is None:
                continue
            out.extend(got if isinstance(got, list) else [got])
    return out

def ruleset_version() -> str:
    h = hashlib.sha256()
    for rid in sorted(REGISTRY):
        h.update(rid.encode()); h.update(inspect.getsource(REGISTRY[rid].fn).encode())
    return h.hexdigest()[:12]
```

`src/outcome_core/validate/log_rows.py`:
```python
"""로그 행 단계 규칙 (PRD §9). E02·E04·E05·W03·W14는 파서 이슈 매핑으로 성립한다."""
from ..model import LogRow, Issue
from . import rule

def _ts(d):
    return d.strftime("%Y-%m-%d %H:%M:%S") if d else ""

@rule("E03", scope="log", stage="row")
def time_order(row: LogRow):
    out = []
    if row.end < row.start:
        out.append(Issue("E03", "E", row.src_row, "종료시각", _ts(row.end), f"종료시각이 시작시각({_ts(row.start)})보다 앞선다"))
    if row.review_open and row.review_save and row.review_save < row.review_open:
        out.append(Issue("E03", "E", row.src_row, "저장제출시각", _ts(row.review_save), "저장제출시각이 검수열림시각보다 앞선다"))
    if row.first_resp and row.first_resp < row.start:
        out.append(Issue("E03", "E", row.src_row, "첫응답시각", _ts(row.first_resp), "첫응답시각이 시작시각보다 앞선다"))
    return out or None

@rule("E12", scope="log", stage="row")
def review_consistency(row: LogRow):
    if row.result == "ERRV":
        return None
    if row.result in ("REJ", "REV") and row.review_open is None:
        return Issue("E12", "E", row.src_row, "검수열림시각", "", "반려·수정은 검수의 결과라 검수열림시각 없이 존재할 수 없다")
    if row.review_open is not None and row.review_save is None:
        return Issue("E12", "E", row.src_row, "저장제출시각", "", "검수열림시각이 있는데 저장제출시각이 없다")
    return None
```

- [ ] **Step 4: 통과 확인**

Run: `pytest tests/unit/test_rules_log.py -q`
Expected: `6 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outcome_core/validate tests/unit/test_rules_log.py
git commit -m "feat(core): 규칙 레지스트리 + 행 단계 규칙 E03·E12, 파서 이슈 매핑 E02/E04/E05/W02/W03/W14"
```

---

### Task 5: 절감 산출 `calc_savings`

**Files:**
- Create: `src/outcome_core/calc/__init__.py`(빈 파일), `src/outcome_core/calc/common.py`, `src/outcome_core/calc/savings.py`
- Test: `tests/unit/test_savings.py`

**Interfaces:**
- Consumes: `model.LogRow`, `model.TaskMaster`, `version.MONTH_DAYS`
- Produces:
  - `common.filter_live(rows) -> list[LogRow]` — 운영 ∧ ≠ERRV
  - `common.observation_months(rows_ops: list[LogRow]) -> float` — R3. 빈 목록 → `1/MONTH_DAYS`
  - `common.sec(a: datetime, b: datetime) -> int` — `int((b-a).total_seconds())`
  - `savings.calc_savings(rows: list[LogRow], master: TaskMaster) -> tuple[dict, dict]` = (metrics, intermediates)
    metrics 키: `n_proc, net_h, rate_pct, rejrev_pct` ; intermediates 키는 PRD 부록 C + `capped: bool, clip_rows, clip_min` (`iv_clipped_rows`/`iv_clipped_min` 유지)
    `rows`는 **해당 과제 행만** 넘긴다(분기는 engine 책임).

- [ ] **Step 1: 테스트 작성** — 손으로 계산한 소형 케이스

`tests/unit/test_savings.py`:
```python
from datetime import datetime as D, timedelta as T
import math
from outcome_core.model import LogRow, TaskMaster
from outcome_core.calc.savings import calc_savings
from outcome_core.calc.common import observation_months, filter_live

M = TaskMaster(task=1, name="t", type="절감 실적형", category="문서·자료 처리", unit="건", pre_per_unit=80.0, pre_annual=1000, approved=True)
S = D(2026, 9, 1, 9, 0, 0)

def row(i, mach_min, iv_min=None, result="OK", env="운영", q=1, day=0):
    st = S + T(days=day); en = st + T(minutes=mach_min)
    ro = rs = None
    if iv_min is not None:
        ro = en + T(minutes=1); rs = ro + T(minutes=iv_min)
    return LogRow(task=1, proc_id=f"P{i}", start=st, end=en, result=result, env=env, q=q, review_open=ro, review_save=rs, src_row=i)

def test_filter_live_and_months():
    rows = [row(1, 2), row(2, 2, env="시험"), row(3, 2, result="ERRV"), row(4, 2, day=29)]
    live = filter_live(rows)
    assert [r.proc_id for r in live] == ["P1", "P4"]
    assert observation_months([r for r in rows if r.env == "운영"]) == (29 + 1) / 30.4375
    assert observation_months([]) == 1 / 30.4375

def test_savings_hand_computed():
    # 4 live rows: mach 2,2,4,2 = 10분; 개입 30,30,None,900(→상한 240) = 300분; q=1 → post = 2.5 + 75 = 77.5
    rows = [row(1, 2, 30), row(2, 2, 30, result="REV"), row(3, 4), row(4, 2, 900, result="REJ"), row(5, 2, env="시험")]
    m, iv = calc_savings(rows, M)
    assert m["n_proc"] == 4 and iv["sum_q"] == 4 and iv["n_reviewed"] == 3
    assert math.isclose(iv["mach_per_unit"], 2.5) and math.isclose(iv["iv_per_unit"], 75.0)
    assert iv["iv_clipped_rows"] == 1 and math.isclose(iv["iv_clipped_min"], 660.0)
    months = 1 / 30.4375                       # 하루 관측 → 최소값
    cap = 1000 * months / 12
    assert math.isclose(iv["cap_period"], cap) and iv["capped"] is True
    applied = min(4 - 1, cap)                  # REJ 1건 차감, 상한 적용
    assert math.isclose(iv["applied_units"], applied)
    assert math.isclose(iv["net_min"], (80 - 77.5) * applied)
    assert math.isclose(m["net_h"], iv["net_min"] / 60)
    assert math.isclose(m["rate_pct"], (80 - 77.5) / 80 * 100)
    assert math.isclose(m["rejrev_pct"], (1 + 1) / 4 * 100)

def test_savings_uses_integer_seconds_and_q():
    r = row(1, 2, 10, q=3)                     # q=3: 상한 = 80*3*3 = 720분
    r.end = r.end + T(microseconds=999)        # 파서가 절사하지만 방어: 초 단위 정수 합
    m, iv = calc_savings([r], M)
    assert iv["sum_q"] == 3 and math.isclose(iv["mach_per_unit"], 2 / 3) and math.isclose(iv["iv_per_unit"], 10 / 3)

def test_savings_negative_kept():
    rows = [row(i, 100, 100) for i in range(1, 4)]
    m, iv = calc_savings(rows, M)
    assert m["net_h"] < 0
```

- [ ] **Step 2: 실패 확인**

Run: `pytest tests/unit/test_savings.py -q`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: 구현**

`src/outcome_core/calc/common.py`:
```python
from __future__ import annotations
from datetime import datetime
from ..model import LogRow
from ..version import MONTH_DAYS

def filter_live(rows: list[LogRow]) -> list[LogRow]:
    return [r for r in rows if r.env == "운영" and r.result != "ERRV"]

def sec(a: datetime, b: datetime) -> int:
    return int((b - a).total_seconds())

def observation_months(rows_ops: list[LogRow]) -> float:
    """PRD §10: (MAX(시작)−MIN(시작)+1일) ÷ 30.4375, 최소 1÷30.4375. 운영 행만 넘길 것."""
    if not rows_ops:
        return 1 / MONTH_DAYS
    lo = min(r.start for r in rows_ops); hi = max(r.start for r in rows_ops)
    return max(((hi - lo).days + 1) / MONTH_DAYS, 1 / MONTH_DAYS)
```

`src/outcome_core/calc/savings.py`:
```python
"""절감 실적형 산식 v2 (61_산출방식.md §4-1). 반올림 없음, 원값 그대로."""
from __future__ import annotations
import math
from ..model import LogRow, TaskMaster
from .common import filter_live, observation_months, sec

def calc_savings(rows: list[LogRow], master: TaskMaster) -> tuple[dict, dict]:
    pre = float(master.pre_per_unit); annual = float(master.pre_annual)
    ops = [r for r in rows if r.env == "운영"]
    live = filter_live(rows)
    n_rows = len(rows); n_test = sum(1 for r in rows if r.env == "시험")
    n_errv = sum(1 for r in ops if r.result == "ERRV")
    n_proc = len(live)
    sum_q = sum(r.q for r in live)
    mach_sec = sum(sec(r.start, r.end) for r in live)
    iv_sec = 0; clip_rows = 0; clip_sec = 0; n_reviewed = 0
    for r in live:
        if r.review_open is None or r.review_save is None:
            continue
        n_reviewed += 1
        raw = sec(r.review_open, r.review_save)
        cap = pre * 3 * r.q * 60
        if raw > cap:
            clip_rows += 1; clip_sec += raw - cap; iv_sec += cap
        else:
            iv_sec += raw
    mach_pu = (mach_sec / 60) / sum_q if sum_q else 0.0
    iv_pu = (iv_sec / 60) / sum_q if sum_q else 0.0
    post = mach_pu + iv_pu
    months = observation_months(ops)
    cap_period = annual * months / 12
    rej_q = sum(r.q for r in live if r.result == "REJ")
    rev_q = sum(r.q for r in live if r.result == "REV")
    applied = min(sum_q - rej_q, cap_period)
    net_min = (pre - post) * applied
    inter = {
        "n_rows": n_rows, "n_test": n_test, "n_errv": n_errv, "n_proc": n_proc,
        "sum_q": sum_q, "avg_q": (sum_q / n_proc) if n_proc else 0.0,
        "n_reviewed": n_reviewed, "review_rate": (n_reviewed / n_proc * 100) if n_proc else 0.0,
        "mach_total_min": mach_sec / 60, "mach_per_unit": mach_pu,
        "iv_total_min": iv_sec / 60, "iv_clipped_rows": clip_rows, "iv_clipped_min": clip_sec / 60, "iv_per_unit": iv_pu,
        "post_per_unit": post, "pre_per_unit": pre, "pre_annual": annual,
        "months": months, "cap_period": cap_period, "capped": (sum_q - rej_q) > cap_period,
        "rej_q": rej_q, "rev_q": rev_q, "applied_units": applied,
        "net_min": net_min, "net_h": net_min / 60,
        "pre_total_h": pre * applied / 60, "rate_pct": ((pre - post) / pre * 100) if pre else None,
    }
    metrics = {"n_proc": n_proc, "net_h": inter["net_h"], "rate_pct": inter["rate_pct"],
               "rejrev_pct": ((rej_q + rev_q) / sum_q * 100) if sum_q else None}
    return metrics, inter
```

- [ ] **Step 4: 통과 확인**

Run: `pytest tests/unit/test_savings.py -q`
Expected: `4 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outcome_core/calc tests/unit/test_savings.py
git commit -m "feat(core): 절감 산식 v2 calc_savings (정수 초 합산·행 단위 상한·기간 안분)"
```

---

### Task 6: 플래그(W04·W07·W08·W09·W10·T2미승인) + 엔진

**Files:**
- Create: `src/outcome_core/calc/flags.py`, `src/outcome_core/calc/engine.py`
- Test: `tests/unit/test_flags.py`

**Interfaces:**
- Consumes: `calc_savings`, `model.*`, `version.*`, `validate.ruleset_version`
- Produces:
  - `flags.apply_savings_flags(metrics, inter, master, setting) -> tuple[list[dict], list[Issue]]` — flags `{"id": "잠정"|"미확정"|"T2미승인", "reason": str}`, warnings W04·W07·W08·W09·W10
  - `engine.run(ds: Dataset, inputs: list[dict] | None = None) -> dict[int, Result]` — 이 계획에서는 `master.type == "절감 실적형"`인 과제만 산출, 나머지 과제는 `metrics={"n_proc": …}`만(처리 건수). `Result.versions = {"formula": FORMULA_VERSION, "ruleset": ruleset_version()}`

- [ ] **Step 1: 테스트 작성**

`tests/unit/test_flags.py`:
```python
from datetime import datetime as D, timedelta as T
from outcome_core.model import LogRow, TaskMaster, TaskSetting, Dataset
from outcome_core.calc.flags import apply_savings_flags
from outcome_core.calc.engine import run
from outcome_core.calc.savings import calc_savings

M = TaskMaster(task=1, name="t", type="절감 실적형", category="문서·자료 처리", unit="건", pre_per_unit=80.0, pre_annual=1000, approved=True)
SET = TaskSetting(task=1)

def rows(n, mach=2, iv=30, result="OK"):
    out = []
    for i in range(n):
        st = D(2026, 9, 1, 9) + T(days=i % 60); en = st + T(minutes=mach)
        out.append(LogRow(task=1, proc_id=f"P{i}", start=st, end=en, result=result, env="운영",
                          review_open=en, review_save=en + T(minutes=iv), src_row=i + 4))
    return out

def _flags(rs, master=M, setting=SET):
    m, iv = calc_savings(rs, master); return apply_savings_flags(m, iv, master, setting)

def test_w04_provisional_under_30():
    flags, warns = _flags(rows(10))
    assert {f["id"] for f in flags} == {"잠정"} and "W04" in {w.rule for w in warns}

def test_w10_reversal_flag():
    flags, warns = _flags(rows(40, mach=100, iv=100))
    assert "미확정" in {f["id"] for f in flags} and "W10" in {w.rule for w in warns}

def test_t2_unapproved_flag():
    m2 = TaskMaster(**{**M.__dict__, "approved": False})
    flags, _ = _flags(rows(40), master=m2)
    assert "T2미승인" in {f["id"] for f in flags}

def test_w07_clip_ratio_and_w09_low_review():
    rs = rows(40); 
    for r in rs[:3]: r.review_save = r.review_open + T(minutes=1000)   # 3/40 = 7.5% > 5%
    for r in rs[3:20]: r.review_open = r.review_save = None            # 검수율 20/40 = 50% < 80%
    _, warns = _flags(rs)
    assert {"W07", "W09"} <= {w.rule for w in warns}

def test_w08_volume_over_annual():
    m2 = TaskMaster(**{**M.__dict__, "pre_annual": 10})
    _, warns = _flags(rows(40), master=m2)          # 40건/약 2개월 → 연환산 ≫ 10×1.3
    assert "W08" in {w.rule for w in warns}

def test_engine_runs_savings_and_count_only():
    m7 = TaskMaster(task=7, name="신호등", type="운영 실적형", category="감시·경보", unit="알림", pre_per_unit=None, pre_annual=None)
    logs = rows(40) + [LogRow(task=7, proc_id="S1", start=D(2026,9,1), end=D(2026,9,1,0,1), result="OK", env="운영")]
    ds = Dataset(year=2026, upto_month=12, logs=logs, master={1: M, 7: m7}, settings={1: SET})
    res = run(ds)
    assert res[1].metrics["net_h"] > 0 and res[7].metrics == {"n_proc": 1}
    assert res[1].versions["formula"] == "2.0" and len(res[1].versions["ruleset"]) == 12
```

- [ ] **Step 2: 실패 확인**

Run: `pytest tests/unit/test_flags.py -q`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: 구현**

`src/outcome_core/calc/flags.py`:
```python
"""산출 단계 플래그·경고 (PRD §9 W04·W07~W10, FR-4.5)."""
from __future__ import annotations
from ..model import TaskMaster, TaskSetting, Issue
from ..version import MIN_N

def apply_savings_flags(metrics: dict, inter: dict, master: TaskMaster, setting: TaskSetting):
    flags: list[dict] = []; warns: list[Issue] = []
    n = inter["n_proc"]
    if n < MIN_N:
        flags.append({"id": "잠정", "reason": f"처리 건수 {n}건 < {MIN_N}"})
        warns.append(Issue("W04", "W", 0, "", str(n), "처리 건수 30건 미만 — 순절감·단축률 잠정, 다음 분기 재산출"))
    if inter["net_min"] < 0:
        flags.append({"id": "미확정", "reason": "도입 후 단위당이 도입 전보다 크다(역전)"})
        warns.append(Issue("W10", "W", 0, "", f"{inter['post_per_unit']:.3f}", "역전(도입 후 > 도입 전) — 미확정 분리, 원인 규명"))
    if not master.approved:
        flags.append({"id": "T2미승인", "reason": "도입 전 값이 승인되지 않았다"})
    if n and inter["iv_clipped_rows"] / n > 0.05:
        warns.append(Issue("W07", "W", 0, "개입시간", str(inter["iv_clipped_rows"]), "개입시간 상한 클립 행 비율 5% 초과 — 무입력 구간 과다"))
    if master.pre_annual and inter["months"] > 0:
        annualized = inter["sum_q"] / inter["months"] * 12
        if annualized > master.pre_annual * 1.3:
            warns.append(Issue("W08", "W", 0, "처리수량", f"{annualized:.1f}", "관측 물량 연환산이 기준 처리량의 1.3배 초과"))
    if n and inter["review_rate"] / 100 < setting.review_rate_thr:
        warns.append(Issue("W09", "W", 0, "검수열림시각", f"{inter['review_rate']:.2f}%", f"검수 발생률이 임계 {setting.review_rate_thr:.0%} 미만 — 단축률 과대계상 위험"))
    return flags, warns
```

`src/outcome_core/calc/engine.py`:
```python
"""Dataset → 과제별 Result. 이 단계(Plan 1)는 절감 실적형 + 처리 건수만."""
from __future__ import annotations
from collections import defaultdict
from ..model import Dataset, Result, TaskSetting
from ..version import FORMULA_VERSION
from ..validate import ruleset_version
from .common import filter_live
from .savings import calc_savings
from .flags import apply_savings_flags

def run(ds: Dataset, inputs: list[dict] | None = None) -> dict[int, Result]:
    by_task: dict[int, list] = defaultdict(list)
    for r in ds.logs:
        by_task[r.task].append(r)
    versions = {"formula": FORMULA_VERSION, "ruleset": ruleset_version()}
    out: dict[int, Result] = {}
    for task, master in sorted(ds.master.items()):
        rows = sorted(by_task.get(task, []), key=lambda r: (r.task, r.proc_id))
        setting = ds.settings.get(task, TaskSetting(task=task))
        if master.type == "절감 실적형" and master.pre_per_unit is not None and rows:
            metrics, inter = calc_savings(rows, master)
            flags, warns = apply_savings_flags(metrics, inter, master, setting)
        else:
            metrics, inter, flags, warns = {"n_proc": len(filter_live(rows))}, {}, [], []
        out[task] = Result(task=task, metrics=metrics, intermediates=inter, flags=flags, warnings=warns,
                           versions=versions, inputs=list(inputs or []))
    return out
```

- [ ] **Step 4: 통과 확인**

Run: `pytest tests/unit/test_flags.py -q`
Expected: `6 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outcome_core/calc/flags.py src/outcome_core/calc/engine.py tests/unit/test_flags.py
git commit -m "feat(core): 절감 플래그(잠정·미확정·T2미승인, W04·W07~W10)와 엔진"
```

---

### Task 7: CSV 어댑터 + RawTable + master.csv 로더

**Files:**
- Create: `src/outcome_cli/adapters/__init__.py`(빈 파일), `src/outcome_cli/adapters/rawtable.py`, `src/outcome_cli/adapters/csvfile.py`, `src/outcome_cli/masterfile.py`
- Test: `tests/adapter/test_csvfile.py`

**Interfaces:**
- Produces:
  - `@dataclass RawTable(kind: str, header: list[str], rows: list[list], first_src_row: int, source: str)` — `kind ∈ {"log","turn","verify","monthly"}`
  - `read_csv(path, kind="log") -> RawTable` — 인코딩 utf-8-sig → cp949 순으로 시도, 구분자 `,`/`\t`/`;` 자동. `first_src_row=2`
  - `write_csv(path, header, rows)` — utf-8-sig
  - `load_master(path) -> dict[int, TaskMaster]` — 열: `과제번호,과제명,유형,구분,단위정의,도입전단위당분,도입전기준처리량연,증거등급,승인,승인일,적용시작월` (`승인`은 `Y`/빈값, 숫자 빈값·`-`는 None)

- [ ] **Step 1: 테스트 작성**

`tests/adapter/test_csvfile.py`:
```python
import pathlib
from outcome_cli.adapters.csvfile import read_csv, write_csv
from outcome_cli.masterfile import load_master
from outcome_core.model import LOG_COLS

def test_read_csv_utf8_and_cp949(tmp_path: pathlib.Path):
    p1 = tmp_path / "a.csv"; p1.write_text(",".join(LOG_COLS) + "\n1,P1,2026-09-01 09:00,2026-09-01 09:01,OK,운영,,,,,,,\n", encoding="utf-8-sig")
    p2 = tmp_path / "b.csv"; p2.write_bytes((";".join(LOG_COLS) + "\n1;P1;2026-09-01 09:00;2026-09-01 09:01;OK;운영;;;;;;;\n").encode("cp949"))
    for p in (p1, p2):
        t = read_csv(p)
        assert t.kind == "log" and t.header == LOG_COLS and t.rows[0][1] == "P1" and t.first_src_row == 2

def test_write_then_read_roundtrip(tmp_path):
    p = tmp_path / "c.csv"
    write_csv(p, LOG_COLS, [["1","P1","2026-09-01 09:00:00","2026-09-01 09:01:00","OK","운영","1","","","","","",""]])
    assert read_csv(p).rows[0][0] == "1"

def test_load_master(tmp_path):
    p = tmp_path / "master.csv"
    p.write_text("과제번호,과제명,유형,구분,단위정의,도입전단위당분,도입전기준처리량연,증거등급,승인,승인일,적용시작월\n"
                 "1,방송 심사자료 제작,절감 실적형,문서·자료 처리,기업 1건,80,1000,T2,Y,2026-09-20,2026-09\n"
                 "7,안전 신호등,운영 실적형,감시·경보,알림 1건,-,-,T1,,,\n", encoding="utf-8-sig")
    m = load_master(p)
    assert m[1].pre_per_unit == 80.0 and m[1].approved is True and m[1].applies_from == "2026-09"
    assert m[7].pre_per_unit is None and m[7].approved is False
```

- [ ] **Step 2: 실패 확인**

Run: `pytest tests/adapter/test_csvfile.py -q`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: 구현**

`src/outcome_cli/adapters/rawtable.py`:
```python
from dataclasses import dataclass

@dataclass
class RawTable:
    kind: str               # log | turn | verify | monthly
    header: list[str]
    rows: list[list]
    first_src_row: int      # rows[0]의 원본 행 번호(1-based)
    source: str
```

`src/outcome_cli/adapters/csvfile.py`:
```python
import csv, io, pathlib
from .rawtable import RawTable

def _decode(path: pathlib.Path) -> str:
    b = path.read_bytes()
    for enc in ("utf-8-sig", "cp949"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass
    raise ValueError(f"{path}: UTF-8·CP949 어느 쪽으로도 읽을 수 없다")

def read_csv(path, kind: str = "log") -> RawTable:
    path = pathlib.Path(path)
    text = _decode(path)
    first = text.splitlines()[0] if text else ""
    delim = max([",", "\t", ";"], key=first.count)
    rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    header = [h.strip() for h in rows[0]] if rows else []
    return RawTable(kind=kind, header=header, rows=rows[1:], first_src_row=2, source=str(path))

def write_csv(path, header: list[str], rows: list[list]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(header); w.writerows(rows)
```

`src/outcome_cli/masterfile.py`:
```python
"""master.csv ↔ TaskMaster. 제출 파일의 마스터 시트는 절대 읽지 않는다(스펙 R9)."""
from outcome_core.model import TaskMaster
from .adapters.csvfile import read_csv

MASTER_COLS = ["과제번호","과제명","유형","구분","단위정의","도입전단위당분","도입전기준처리량연","증거등급","승인","승인일","적용시작월"]

def _num(v):
    v = (v or "").strip()
    if v in ("", "-"):
        return None
    return float(v)

def load_master(path) -> dict[int, TaskMaster]:
    t = read_csv(path, kind="master")
    idx = {h: i for i, h in enumerate(t.header)}
    missing = [c for c in MASTER_COLS if c not in idx]
    if missing:
        raise ValueError(f"master.csv 열 누락: {missing}")
    out = {}
    for r in t.rows:
        if not r or not r[idx["과제번호"]].strip():
            continue
        g = lambda c: r[idx[c]] if idx[c] < len(r) else ""
        task = int(g("과제번호"))
        out[task] = TaskMaster(task=task, name=g("과제명"), type=g("유형"), category=g("구분"), unit=g("단위정의"),
                               pre_per_unit=_num(g("도입전단위당분")), pre_annual=_num(g("도입전기준처리량연")),
                               evidence=g("증거등급"), approved=g("승인").strip().upper() == "Y",
                               approved_at=g("승인일"), applies_from=g("적용시작월"))
    return out
```

- [ ] **Step 4: 통과 확인**

Run: `pytest tests/adapter/test_csvfile.py -q`
Expected: `3 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outcome_cli tests/adapter/test_csvfile.py
git commit -m "feat(cli): CSV 어댑터(RawTable)와 master.csv 로더"
```

---

### Task 8: `52` 통합문서 어댑터 (구양식, read_only) + master.csv 추출

**Files:**
- Create: `src/outcome_cli/adapters/xlsx52.py`
- Test: `tests/adapter/test_xlsx52.py`

**Interfaces:**
- Consumes: `RawTable`, `docs/52_측정_자동계산_양식.xlsx`
- Produces:
  - `read_52_log(path) -> RawTable` — 시트 `①로그`, 3행 헤더, 4행부터, **13열만**(N~Q 수식열 버림), `first_src_row=4`, `kind="log"`. 셀 `datetime`은 `strftime("%Y-%m-%d %H:%M:%S")` 문자열로(마이크로초는 절사 — R2와 동일 효과), 숫자는 `str`, None은 `""`
  - `detect_template(path) -> str` — 시트명에 `④산출대장` 있으면 `"52-legacy"`, 아니면 `"v3"`(Plan 3에서 사용). 이 계획에선 legacy만
  - `export_master_from_52(path, out_csv)` — 과제마스터 시트 → `master.csv`(Task 7 열 규격, 승인=`Y` 일괄, 승인일 빈값). **골든 작성용 1회 도구**이며 제출 처리에서는 호출하지 않는다

- [ ] **Step 1: 테스트 작성**

`tests/adapter/test_xlsx52.py`:
```python
import pytest
from outcome_cli.adapters.xlsx52 import read_52_log, detect_template, export_master_from_52
from outcome_cli.masterfile import load_master
from outcome_core.model import LOG_COLS
from outcome_core.parse import to_log_rows

@pytest.fixture(scope="module")
def x52(docs_dir):
    p = docs_dir / "52_측정_자동계산_양식.xlsx"
    if not p.exists():
        pytest.skip("52 없음")
    return p

def test_detect_legacy(x52):
    assert detect_template(x52) == "52-legacy"

def test_read_52_log_shape(x52):
    t = read_52_log(x52)
    assert t.header == LOG_COLS and t.first_src_row == 4 and t.kind == "log"
    assert len(t.rows) >= 30903 and t.rows[0][1] == "P1-00001"
    assert t.rows[0][2] == "2026-09-11 14:23:00" and t.rows[0][3] == "2026-09-11 14:25:02"   # ms 절사

def test_read_52_log_stops_at_footer_via_parser(x52):
    t = read_52_log(x52)
    rows, issues, rest = to_log_rows(t.header, t.rows, t.first_src_row)
    assert rest >= 1                                # 「기재요령」 행
    assert sum(1 for r in rows if r.task == 1) == 365

def test_export_master(x52, tmp_path):
    out = tmp_path / "master.csv"
    export_master_from_52(x52, out)
    m = load_master(out)
    assert m[1].pre_per_unit == 80 and m[1].pre_annual == 1000 and m[3].pre_per_unit == 480
    assert m[7].pre_per_unit is None and m[18].type == "측정 제외"
```

- [ ] **Step 2: 실패 확인**

Run: `pytest tests/adapter/test_xlsx52.py -q`
Expected: FAIL — `ModuleNotFoundError`

- [ ] **Step 3: 구현**

`src/outcome_cli/adapters/xlsx52.py`:
```python
"""52_측정_자동계산_양식.xlsx(구양식) 읽기. openpyxl read_only 스트리밍."""
from __future__ import annotations
from datetime import datetime
import openpyxl
from outcome_core.model import LOG_COLS
from .rawtable import RawTable
from .csvfile import write_csv

LOG_SHEET = "①로그"; MASTER_SHEET = "과제마스터"; HEADER_ROW = 3

def _cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def detect_template(path) -> str:
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return "52-legacy" if "④산출대장" in wb.sheetnames else "v3"
    finally:
        wb.close()

def read_52_log(path) -> RawTable:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[LOG_SHEET]
        rows = []
        for i, r in enumerate(ws.iter_rows(min_row=HEADER_ROW, max_col=len(LOG_COLS), values_only=True)):
            if i == 0:
                header = [_cell(h) for h in r]
                if header != LOG_COLS:
                    raise ValueError(f"①로그 헤더가 양식과 다르다: {header}")
                continue
            if all(v is None or str(v).strip() == "" for v in r):
                continue
            rows.append([_cell(v) for v in r])
        return RawTable(kind="log", header=list(LOG_COLS), rows=rows, first_src_row=HEADER_ROW + 1, source=str(path))
    finally:
        wb.close()

MASTER_OUT = ["과제번호","과제명","유형","구분","단위정의","도입전단위당분","도입전기준처리량연","증거등급","승인","승인일","적용시작월"]
CATEGORY = {1:"문서·자료 처리",2:"문서·자료 처리",3:"문서·자료 처리",4:"문서·자료 처리",5:"문서·자료 처리",
            6:"감시·경보",7:"감시·경보",8:"감시·경보",9:"감시·경보",10:"판독·진단",11:"판독·진단",
            12:"응답·안내",13:"응답·안내",14:"응답·안내",15:"응답·안내",16:"예측·배정",17:"예측·배정",18:"",19:"",20:""}

def export_master_from_52(path, out_csv) -> None:
    """골든 작성용 1회 도구. 52 과제마스터 시트(열: 과제번호·과제명·유형·도입전단위당·기준처리량·단위정의·증거등급) → master.csv"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[MASTER_SHEET]
        rows = []
        for r in ws.iter_rows(min_row=HEADER_ROW + 1, max_col=7, values_only=True):
            if r[0] is None:
                continue
            task = int(r[0])
            num = lambda v: "" if v in (None, "-") else _cell(v)
            approved = "Y" if r[2] == "절감 실적형" else ""
            rows.append([task, _cell(r[1]), _cell(r[2]), CATEGORY.get(task, ""), _cell(r[5]), num(r[3]), num(r[4]), _cell(r[6]), approved, "", ""])
        write_csv(out_csv, MASTER_OUT, rows)
    finally:
        wb.close()
```

- [ ] **Step 4: 통과 확인**

Run: `pytest tests/adapter/test_xlsx52.py -q`
Expected: `4 passed` (약 10~20초, 30,903행 스트리밍)

- [ ] **Step 5: Commit**

```bash
git add src/outcome_cli/adapters/xlsx52.py tests/adapter/test_xlsx52.py
git commit -m "feat(cli): 52 통합문서 어댑터(구양식 read_only)와 master.csv 추출"
```

---

### Task 9: 독립 참조 구현 + 엔진 대조 (A6)

**Files:**
- Create: `tests/independent/__init__.py`(빈 파일), `tests/independent/calc_ref.py`, `tests/independent/test_against_ref.py`

**Interfaces:**
- `calc_ref.py`는 **`outcome_core`를 import하지 않는다**(테스트로 강제). 정규 CSV 문자열 행과 마스터 dict를 받아 절감 5건의 `net_h, rate_pct, n_proc, rejrev_pct, months, applied`를 낸다. 산식은 스펙 §6.2와 R2·R3·R4를 **따로 읽고** 구현한다 — 아래 코드는 그 결과다.

- [ ] **Step 1: 참조 구현 작성**

`tests/independent/calc_ref.py`:
```python
"""독립 참조 구현. outcome_core를 import하지 않는다. 61 §4-1 + PRD §10 + 스펙 R2~R4를 직접 옮김."""
from datetime import datetime, timedelta
from collections import defaultdict
import re

_FMT = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y%m%d%H%M%S", "%Y-%m-%d %H:%M:%S.%f"]
_TZ = re.compile(r"(Z|[+-]\d{2}:?\d{2})$")

def ts(s):
    s = (s or "").strip()
    if not s:
        return None
    s = _TZ.sub("", s)
    for f in _FMT:
        try:
            return datetime.strptime(s, f).replace(microsecond=0)
        except ValueError:
            pass
    return datetime(1899, 12, 30) + timedelta(days=float(s))

TEST = {"시험","테스트","개발","검증","test","testing","stg","stage","staging","dev","development","qa","sandbox"}

def env_of(v):
    return "시험" if re.sub(r"[\s_\-]", "", (v or "")).lower() in TEST else "운영"

def compute(header, rows, master):
    """rows: 정규 CSV 문자열 행. master: {task: (pre_per_unit, pre_annual)}. 반환 {task: dict}"""
    ix = {h: i for i, h in enumerate(header)}
    by = defaultdict(list)
    for r in rows:
        try:
            t = int(float(r[ix["과제번호"]]))
        except (ValueError, IndexError):
            break
        by[t].append(r)
    out = {}
    for t, (pre, annual) in master.items():
        rs = by.get(t, [])
        ops = [r for r in rs if env_of(r[ix["환경구분"]]) == "운영"]
        live = [r for r in ops if r[ix["처리결과"]].strip().upper() != "ERRV"]
        q = lambda r: int(float(r[ix["처리수량"]])) if "처리수량" in ix and r[ix["처리수량"]].strip() else 1
        sq = sum(q(r) for r in live)
        mach = sum(int((ts(r[ix["종료시각"]]) - ts(r[ix["시작시각"]])).total_seconds()) for r in live)
        iv = 0
        for r in live:
            o, s = ts(r[ix["검수열림시각"]]), ts(r[ix["저장제출시각"]])
            if o and s:
                iv += min(int((s - o).total_seconds()), pre * 3 * q(r) * 60)
        starts = [ts(r[ix["시작시각"]]) for r in ops]
        months = max(((max(starts) - min(starts)).days + 1) / 30.4375, 1 / 30.4375) if starts else 1 / 30.4375
        cap = annual * months / 12
        rej = sum(q(r) for r in live if r[ix["처리결과"]].strip().upper() == "REJ")
        rev = sum(q(r) for r in live if r[ix["처리결과"]].strip().upper() == "REV")
        post = (mach / 60 + iv / 60) / sq if sq else 0
        applied = min(sq - rej, cap)
        net_min = (pre - post) * applied
        out[t] = {"n_proc": len(live), "net_h": net_min / 60, "rate_pct": (pre - post) / pre * 100,
                  "rejrev_pct": (rej + rev) / sq * 100 if sq else None, "months": months, "applied": applied}
    return out
```

- [ ] **Step 2: 대조 테스트 작성**

`tests/independent/test_against_ref.py`:
```python
import ast, pathlib, math, pytest
from outcome_core.parse import to_log_rows
from outcome_core.model import Dataset
from outcome_core.calc.engine import run
from outcome_cli.adapters.xlsx52 import read_52_log, export_master_from_52
from outcome_cli.masterfile import load_master
from . import calc_ref

def test_ref_does_not_import_core():
    src = (pathlib.Path(__file__).parent / "calc_ref.py").read_text(encoding="utf-8")
    names = {n.names[0].name.split(".")[0] if isinstance(n, ast.Import) else (n.module or "").split(".")[0]
             for n in ast.walk(ast.parse(src)) if isinstance(n, (ast.Import, ast.ImportFrom))}
    assert "outcome_core" not in names and "outcome_cli" not in names

@pytest.fixture(scope="module")
def data(docs_dir, tmp_path_factory):
    p = docs_dir / "52_측정_자동계산_양식.xlsx"
    if not p.exists():
        pytest.skip("52 없음")
    t = read_52_log(p)
    mp = tmp_path_factory.mktemp("m") / "master.csv"; export_master_from_52(p, mp)
    return t, load_master(mp)

def test_engine_matches_independent_ref_within_a6(data):
    t, master = data
    rows, issues, _ = to_log_rows(t.header, t.rows, t.first_src_row)
    assert not [i for i in issues if i.level == "E"], issues[:5]
    res = run(Dataset(year=2026, upto_month=12, logs=rows, master=master))
    ref = calc_ref.compute(t.header, t.rows, {k: (m.pre_per_unit, m.pre_annual) for k, m in master.items() if m.type == "절감 실적형"})
    for task in range(1, 6):
        a, b = res[task], ref[task]
        assert a.metrics["n_proc"] == b["n_proc"]
        assert abs(a.metrics["net_h"] - b["net_h"]) <= 0.01, (task, a.metrics["net_h"], b["net_h"])
        assert abs(a.metrics["rate_pct"] - b["rate_pct"]) <= 0.01
        assert abs(a.metrics["rejrev_pct"] - b["rejrev_pct"]) <= 0.01
        assert math.isclose(a.intermediates["months"], b["months"])
        assert math.isclose(a.intermediates["applied_units"], b["applied"])
    # 운영형 처리 건수도 확인 (51 값)
    assert res[6].metrics["n_proc"] == 2400 and res[12].metrics["n_proc"] == 4800
```

- [ ] **Step 3: 실행**

Run: `pytest tests/independent -q`
Expected: `2 passed`. 실패하면 **어느 쪽이 스펙과 다른지** 먼저 판단한다(라이브러리가 옳다고 전제하지 않는다). 대표 확인점: #1 `n_proc == 342`, #2 `sum_q == 485`, #4 `sum_q == 113`(현황분석 probe 값 — 절사 전이므로 `net_h`는 probe와 소폭 달라도 정상).

- [ ] **Step 4: Commit**

```bash
git add tests/independent
git commit -m "test: 독립 참조 구현과 엔진 대조 (A6 허용오차, 52 원자료 절감 5건)"
```

---

### Task 10: CLI `validate` · `submit` · `calc` + store

**Files:**
- Create: `src/outcome_cli/store.py`, `src/outcome_cli/__main__.py`
- Test: `tests/cli/test_commands.py`

**Interfaces:**
- Consumes: 어댑터, `to_log_rows`, `map_parse_issues`, `run_row_rules`, `engine.run`, `load_master`
- Produces:
  - `store.submit_to_store(table: RawTable, task: int, month: str, store_dir) -> pathlib.Path` — `store/{YYYY}/{task}/{MM}/v{n}/log.csv` + `meta.json{sha256, rows, source, submitted_at, versions}`. 기존 v가 있으면 n+1(추가 전용)
  - `store.load_dataset(store_dir, year: int, upto: int, master) -> tuple[Dataset, list[dict]]` — 과제·월별 **최신 v**의 `log.csv`를 이어 붙임(마감 파일은 Plan 3). 두 번째 반환은 `inputs` 메타 목록
  - `validate_table(table: RawTable, task: int) -> tuple[list[LogRow], list[Issue]]` — 파싱 + 매핑 + 행 규칙 + **E17**(선택 과제 외 과제번호) + **W17**(읽기 종료 이후 행 수). E17·W17은 여기서 함수로 구현하고 Plan 2에서 레지스트리로 옮긴다
  - `main(argv=None) -> int` — 서브커맨드 `validate`(종료코드: 오류 있으면 2), `submit`, `calc`
  - `calc` 출력: `results/{year}-{upto:02d}/task_{N}.json`(= `Result.to_json()`), `ledger.csv`(열: `과제번호,처리건수,순절감h,단축률pct,반려수정률pct,플래그`), `flags.md`
  - 리포트: `validate`는 `--report DIR`에 `report.csv`(열: `구분,규칙,행,열,값,사유`)와 `report.md`(오류·경고 건수 + 상위 20건)

- [ ] **Step 1: 테스트 작성**

`tests/cli/test_commands.py`:
```python
import json, pathlib, csv
from outcome_cli.__main__ import main
from outcome_core.model import LOG_COLS

MASTER = ("과제번호,과제명,유형,구분,단위정의,도입전단위당분,도입전기준처리량연,증거등급,승인,승인일,적용시작월\n"
          "1,방송,절감 실적형,문서·자료 처리,기업,80,1000,T2,Y,,\n")

def _log(rows):
    return ",".join(LOG_COLS) + "\n" + "\n".join(",".join(r) for r in rows) + "\n"

def good(i, day):
    return [f"1", f"P{i}", f"2026-09-{day:02d} 09:00:00", f"2026-09-{day:02d} 09:02:00", "OK", "운영", "1", f"2026-09-{day:02d} 09:03:00", f"2026-09-{day:02d} 09:33:00", "", "", "", ""]

def test_validate_reports_errors_and_exit_code(tmp_path):
    f = tmp_path / "log.csv"
    f.write_text(_log([good(1, 1), ["1","P2","2026-09-01 09:00","2026-09-01 08:00","OK","운영","1","","","","","",""],
                       ["2","P3","2026-09-01 09:00","2026-09-01 09:01","OK","운영","1","","","","","",""],
                       ["기재요령","","","","","","","","","","","",""]]), encoding="utf-8-sig")
    rc = main(["validate", str(f), "--task", "1", "--month", "2026-09", "--report", str(tmp_path / "rep")])
    assert rc == 2
    rep = list(csv.DictReader(open(tmp_path / "rep" / "report.csv", encoding="utf-8-sig")))
    rules = {r["규칙"] for r in rep}
    assert {"E03", "E17", "W17"} <= rules

def test_submit_then_calc(tmp_path):
    (tmp_path / "master.csv").write_text(MASTER, encoding="utf-8-sig")
    f = tmp_path / "log.csv"; f.write_text(_log([good(i, (i % 28) + 1) for i in range(1, 41)]), encoding="utf-8-sig")
    store = tmp_path / "store"
    assert main(["submit", str(f), "--task", "1", "--month", "2026-09", "--store", str(store)]) == 0
    assert main(["submit", str(f), "--task", "1", "--month", "2026-09", "--store", str(store)]) == 0
    assert (store / "2026" / "1" / "09" / "v2" / "log.csv").exists()
    meta = json.loads((store / "2026" / "1" / "09" / "v2" / "meta.json").read_text(encoding="utf-8"))
    assert len(meta["sha256"]) == 64 and meta["rows"] == 40
    out = tmp_path / "results"
    assert main(["calc", "--store", str(store), "--master", str(tmp_path / "master.csv"), "--year", "2026", "--upto", "9", "--out", str(out)]) == 0
    res = json.loads((out / "2026-09" / "task_1.json").read_text(encoding="utf-8"))
    assert res["metrics"]["n_proc"] == 40 and res["metrics"]["net_h"] > 0 and res["inputs"][0]["sha256"] == meta["sha256"]
    ledger = (out / "2026-09" / "ledger.csv").read_text(encoding="utf-8-sig")
    assert ledger.splitlines()[0] == "과제번호,처리건수,순절감h,단축률pct,반려수정률pct,플래그"

def test_submit_refuses_on_error(tmp_path):
    f = tmp_path / "log.csv"
    f.write_text(_log([["1","P1","2026-09-01 09:00","2026-09-01 08:00","OK","운영","1","","","","","",""]]), encoding="utf-8-sig")
    assert main(["submit", str(f), "--task", "1", "--month", "2026-09", "--store", str(tmp_path / "s")]) == 2
    assert not (tmp_path / "s").exists()
```

- [ ] **Step 2: 실패 확인**

Run: `pytest tests/cli -q`
Expected: FAIL — `ModuleNotFoundError: outcome_cli.__main__`

- [ ] **Step 3: 구현**

`src/outcome_cli/store.py`:
```python
"""추가 전용 파일 저장소. store/{YYYY}/{task}/{MM}/v{n}/log.csv + meta.json"""
from __future__ import annotations
import hashlib, json, pathlib
from datetime import datetime, timezone
from outcome_core.model import Dataset, Issue, LogRow, LOG_COLS
from outcome_core.parse import to_log_rows
from outcome_core.validate import map_parse_issues, run_row_rules, ruleset_version
from outcome_core.version import FORMULA_VERSION
import outcome_core.validate.log_rows  # noqa: F401
from .adapters.rawtable import RawTable
from .adapters.csvfile import write_csv, read_csv

def validate_table(table: RawTable, task: int) -> tuple[list[LogRow], list[Issue]]:
    rows, pissues, rest = to_log_rows(table.header, table.rows, table.first_src_row)
    issues = map_parse_issues(pissues) + run_row_rules(rows)
    for r in rows:
        if r.task != task:
            issues.append(Issue("E17", "E", r.src_row, "과제번호", str(r.task), f"선택한 과제 #{task} 외의 행"))
    if rest:
        issues.append(Issue("W17", "W", table.first_src_row + len(table.rows) - rest, "과제번호", "", f"과제번호가 정수가 아닌 행에서 읽기를 종료했다. 이후 {rest}행 무시"))
    issues.sort(key=lambda i: (i.level != "E", i.row, i.rule))
    return [r for r in rows if r.task == task], issues

def _sha(path: pathlib.Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()

def submit_to_store(table: RawTable, task: int, month: str, store_dir) -> pathlib.Path:
    y, m = month.split("-")
    base = pathlib.Path(store_dir) / y / str(task) / m
    base.mkdir(parents=True, exist_ok=True)
    n = 1 + max([int(p.name[1:]) for p in base.glob("v*") if p.name[1:].isdigit()] or [0])
    d = base / f"v{n}"; d.mkdir()
    write_csv(d / "log.csv", table.header, table.rows)
    meta = {"sha256": _sha(d / "log.csv"), "rows": len(table.rows), "source": table.source, "task": task, "month": month,
            "version": n, "submitted_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "versions": {"formula": FORMULA_VERSION, "ruleset": ruleset_version()}}
    (d / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return d

def load_dataset(store_dir, year: int, upto: int, master) -> tuple[Dataset, list[dict]]:
    logs: list[LogRow] = []; inputs: list[dict] = []
    ydir = pathlib.Path(store_dir) / str(year)
    for tdir in sorted(ydir.glob("*"), key=lambda p: int(p.name)) if ydir.exists() else []:
        for mdir in sorted(tdir.glob("*")):
            if int(mdir.name) > upto:
                continue
            vs = sorted([p for p in mdir.glob("v*") if p.name[1:].isdigit()], key=lambda p: int(p.name[1:]))
            if not vs:
                continue
            latest = vs[-1]
            t = read_csv(latest / "log.csv")
            rows, _, _ = to_log_rows(t.header, t.rows, 2)
            logs.extend(rows)
            meta = json.loads((latest / "meta.json").read_text(encoding="utf-8"))
            inputs.append({"path": str(latest / "log.csv"), "sha256": meta["sha256"], "rows": meta["rows"], "task": int(tdir.name), "month": f"{year}-{mdir.name}", "version": meta["version"]})
    return Dataset(year=year, upto_month=upto, logs=logs, master=master), inputs
```

`src/outcome_cli/__main__.py`:
```python
"""outcome CLI: validate · submit · calc (Plan 1 최소형)."""
from __future__ import annotations
import argparse, csv, pathlib, sys
from outcome_core.calc.engine import run
from .adapters.csvfile import read_csv, write_csv
from .adapters.xlsx52 import read_52_log
from .masterfile import load_master
from .store import validate_table, submit_to_store, load_dataset

def _read(path: str):
    p = pathlib.Path(path)
    return read_52_log(p) if p.suffix.lower() == ".xlsx" else read_csv(p)

def _write_report(issues, rep_dir):
    rep = pathlib.Path(rep_dir); rep.mkdir(parents=True, exist_ok=True)
    write_csv(rep / "report.csv", ["구분","규칙","행","열","값","사유"],
              [["오류" if i.level == "E" else "경고", i.rule, i.row, i.col, i.value, i.reason] for i in issues])
    ne = sum(1 for i in issues if i.level == "E"); nw = len(issues) - ne
    lines = [f"# 검증 결과", "", f"- 오류 {ne}건 · 경고 {nw}건", ""]
    lines += [f"- [{i.rule}] 행 {i.row} {i.col} `{i.value}` — {i.reason}" for i in issues[:20]]
    (rep / "report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

def cmd_validate(a) -> int:
    t = _read(a.file); rows, issues = validate_table(t, a.task)
    if a.report:
        _write_report(issues, a.report)
    ne = sum(1 for i in issues if i.level == "E")
    print(f"행 {len(rows)} · 오류 {ne} · 경고 {len(issues) - ne}")
    for i in issues[:20]:
        print(f"  [{i.rule}] 행 {i.row} {i.col} {i.value!r} — {i.reason}")
    return 2 if ne else 0

def cmd_submit(a) -> int:
    t = _read(a.file); rows, issues = validate_table(t, a.task)
    if any(i.level == "E" for i in issues):
        print("오류가 있어 저장하지 않는다. validate로 확인할 것"); return 2
    t.rows = [t.rows[r.src_row - t.first_src_row] for r in rows]     # 선택 과제 행만 저장
    d = submit_to_store(t, a.task, a.month, a.store)
    print(f"저장: {d}"); return 0

def cmd_calc(a) -> int:
    master = load_master(a.master)
    ds, inputs = load_dataset(a.store, a.year, a.upto, master)
    res = run(ds, inputs)
    out = pathlib.Path(a.out) / f"{a.year}-{a.upto:02d}"; out.mkdir(parents=True, exist_ok=True)
    ledger, flags_md = [], ["# 플래그·경고", ""]
    for task, r in res.items():
        (out / f"task_{task}.json").write_text(r.to_json(), encoding="utf-8")
        m = r.metrics
        ledger.append([task, m.get("n_proc", ""), m.get("net_h", ""), m.get("rate_pct", ""), m.get("rejrev_pct", ""),
                       " ".join(f["id"] for f in r.flags)])
        for f in r.flags:
            flags_md.append(f"- #{task} {f['id']}: {f['reason']}")
        for w in r.warnings:
            flags_md.append(f"- #{task} {w.rule}: {w.reason}")
    write_csv(out / "ledger.csv", ["과제번호","처리건수","순절감h","단축률pct","반려수정률pct","플래그"], ledger)
    (out / "flags.md").write_text("\n".join(flags_md) + "\n", encoding="utf-8")
    print(f"산출: {out}"); return 0

def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="outcome")
    sp = p.add_subparsers(dest="cmd", required=True)
    v = sp.add_parser("validate"); v.add_argument("file"); v.add_argument("--task", type=int, required=True); v.add_argument("--month", required=True); v.add_argument("--report"); v.set_defaults(fn=cmd_validate)
    s = sp.add_parser("submit"); s.add_argument("file"); s.add_argument("--task", type=int, required=True); s.add_argument("--month", required=True); s.add_argument("--store", required=True); s.set_defaults(fn=cmd_submit)
    c = sp.add_parser("calc"); c.add_argument("--store", required=True); c.add_argument("--master", required=True); c.add_argument("--year", type=int, required=True); c.add_argument("--upto", type=int, required=True); c.add_argument("--out", default="results"); c.set_defaults(fn=cmd_calc)
    a = p.parse_args(argv)
    return a.fn(a)

if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: 통과 확인**

Run: `pytest tests/cli -q && pytest -q`
Expected: `3 passed`, 전체 통과

- [ ] **Step 5: 52로 #1 실제 산출 (수동 확인)**

```bash
python -c "from outcome_cli.adapters.xlsx52 import export_master_from_52; export_master_from_52('docs/52_측정_자동계산_양식.xlsx','master.csv')"
outcome submit docs/52_측정_자동계산_양식.xlsx --task 1 --month 2026-12 --store store
outcome calc --store store --master master.csv --year 2026 --upto 12 --out results
cat results/2026-12/ledger.csv
```
Expected: 과제 1 행에 처리건수 342, 순절감 h가 Task 9의 참조값과 같음. (`master.csv`·`store/`·`results/`는 `.gitignore` 대상 — 커밋하지 않는다.)

- [ ] **Step 6: Commit**

```bash
git add src/outcome_cli/store.py src/outcome_cli/__main__.py tests/cli
git commit -m "feat(cli): validate·submit·calc 명령과 추가 전용 store (E17·W17 포함)"
```

---

## 완료 판정 (Plan 1)

- `pytest -q` 전체 통과, `tests/test_no_third_party.py` 통과
- `outcome validate docs/52_측정_자동계산_양식.xlsx --task 1 --month 2026-12` → 오류 0, W17 1건
- `calc` 결과 #1~#5가 `tests/independent/calc_ref.py`와 A6 이내 일치 (Task 9)
- 스펙 §12 주차 1~2 산출물 충족. 다음: Plan 2(file·cross 규칙, quality·monthly·aggregate).
