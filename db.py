"""SQLite 저장소 — 범용 실패비용(COPQ) 관리 시스템

데이터셋(사업장/제품) 하나 = data/ 폴더의 .db 파일 하나.
기본 구조(원본: ECV 실패비용 SYSTEM VER_9.4 엑셀):
  부품단가DB        -> parts (+ price_history: 단가 변경 이력)
  유형별BOM         -> bom (long format: 불량유형 x 부품 x 사용수량)
  유형별초기원인    -> defect_types (불량유형 마스터, BOM 공정 포함)
  일일불량입력(매일) -> daily
  귀책입력          -> judgments (클레임 협의 결과 포함)
  KPI관리 목표      -> targets
  월간 조치내역     -> actions
"""
from __future__ import annotations

import glob
import json
import os
import sqlite3
from contextlib import contextmanager
from datetime import datetime

import pandas as pd

DATA_DIR = os.environ.get("COPQ_DATA", os.path.join(os.path.dirname(os.path.abspath(__file__)), "data"))
DB_PATH = os.environ.get("COPQ_DB", os.path.join(DATA_DIR, "default.db"))

# 계산 로직에 쓰이는 고정 값 (변경 불가)
COST_CLASSES = ["실패", "예방", "제외"]          # 제외 = 집계 제외, 예방 = 예방비용으로 분리
RESPONSIBILITY = ["", "사내", "공급"]            # 공급 = 협력사 귀책(클레임 대상)
REUSE_FLAGS = ["", "재사용", "보류"]             # 재사용 = REWORK, 보류 = 손실(잠재)

# 사업장별로 바꿀 수 있는 설정 기본값
DEFAULT_SETTINGS = {
    "company": "",
    "system_name": "실패비용(COPQ) 관리",
    "currency": "원",
    "week_start": "sun",                      # sun = 엑셀 WEEKNUM(,1) / mon = WEEKNUM(,2)
    "opt_guess_cause": ["협력사", "공정", "경계치", "NTF", "설계", "설비"],
    "opt_disposition": ["폐기", "재검", "AS", "격리후사용", "REWORK"],
    "opt_state": ["SUB", "단품", "완성품"],
    "opt_confirmed_cause": ["공정", "협력사", "NTF", "경계치", "설계", "설비", "작업자"],
    "opt_action_status": ["원인분석 중", "개선 중", "조치 예정", "클레임 예정", "완료"],
    "judge_causes": ["협력사", "NTF", "경계치"],   # 이 원인분류는 귀책 판정 대상
    "judge_dispositions": ["재검"],               # 이 처리방안은 귀책 판정 대상
}
SETTING_LABELS = {
    "opt_guess_cause": "원인분류(추측원인) 목록", "opt_disposition": "처리방안 목록", "opt_state": "불량상태 목록",
    "opt_confirmed_cause": "확정원인 예시 목록", "opt_action_status": "조치 진행상태 목록",
    "judge_causes": "귀책 판정 대상 원인분류", "judge_dispositions": "귀책 판정 대상 처리방안",
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY, value TEXT
);
CREATE TABLE IF NOT EXISTS processes (
    name TEXT PRIMARY KEY, sort INTEGER DEFAULT 0, active INTEGER DEFAULT 1, note TEXT
);
CREATE TABLE IF NOT EXISTS parts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    price REAL NOT NULL DEFAULT 0,
    changed_at TEXT,
    note TEXT
);
CREATE TABLE IF NOT EXISTS price_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    part_name TEXT, old_price REAL, new_price REAL,
    changed_at TEXT, memo TEXT
);
CREATE TABLE IF NOT EXISTS defect_types (
    defect_type TEXT PRIMARY KEY,
    process TEXT,               -- BOM 기준 공정
    cost_class TEXT DEFAULT '실패',   -- 비용구분: 실패/예방/제외
    initial_cause TEXT,         -- 초기원인 (성능/압입력/외관/하중/조립 ...)
    detection TEXT,             -- 검출기준
    state TEXT,                 -- 불량상태 (SUB/단품/완성품)
    initial_resp TEXT,          -- 초기귀책 = 추측원인 (협력사/공정/경계치/NTF)
    factor1 TEXT, factor2 TEXT, factor3 TEXT,
    key_proc1 TEXT, key_proc2 TEXT,
    disposition TEXT DEFAULT '폐기',  -- 처리방안
    active INTEGER DEFAULT 1
);
CREATE TABLE IF NOT EXISTS bom (
    defect_type TEXT NOT NULL,
    part_name TEXT NOT NULL,
    qty REAL NOT NULL DEFAULT 1,
    PRIMARY KEY (defect_type, part_name)
);
CREATE TABLE IF NOT EXISTS daily (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,
    process TEXT NOT NULL,
    defect_type TEXT NOT NULL,
    qty INTEGER NOT NULL,
    remark TEXT,
    exclude INTEGER DEFAULT 0,   -- 개별 건 실패비용 제외 (테스트 등)
    created_at TEXT
);
CREATE INDEX IF NOT EXISTS ix_daily_date ON daily(date);
CREATE TABLE IF NOT EXISTS judgments (
    key TEXT PRIMARY KEY,        -- 연|월|주차|공정|불량유형
    year INTEGER, month INTEGER, week INTEGER,
    process TEXT, defect_type TEXT,
    confirmed_cause TEXT,        -- 확정원인(귀책)
    responsibility TEXT,         -- 귀책: 사내/공급
    retest_ng INTEGER DEFAULT 0, -- 재검불량
    retest_ok INTEGER DEFAULT 0, -- 재검합격
    reuse_flag TEXT,             -- 재사용여부: 재사용/보류
    ncr_no TEXT,
    claim_yn TEXT,
    claim_vendor TEXT,
    resp_qty INTEGER DEFAULT 0,  -- 귀책수량(클레임 수량)
    claim_no TEXT,
    actual_amount REAL,          -- 실청구액
    agreed_date TEXT,            -- 협의완료일
    memo TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS targets (
    year INTEGER, process TEXT,
    annual REAL DEFAULT 0, monthly REAL DEFAULT 0,
    PRIMARY KEY (year, process)
);
CREATE TABLE IF NOT EXISTS actions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    year INTEGER, month INTEGER,
    defect_type TEXT, analysis TEXT, result TEXT,
    status TEXT, owner TEXT, due_date TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS change_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    at TEXT, target TEXT, content TEXT, reason TEXT, who TEXT
);
"""


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@contextmanager
def connect(path: str | None = None):
    p = path or DB_PATH
    os.makedirs(os.path.dirname(p), exist_ok=True)
    con = sqlite3.connect(p)
    con.execute("PRAGMA foreign_keys=ON")
    try:
        yield con
        con.commit()
    finally:
        con.close()


def init_db(path: str | None = None) -> None:
    with connect(path) as con:
        con.executescript(SCHEMA)


def read(table: str, where: str = "", params: tuple = (), path: str | None = None) -> pd.DataFrame:
    sql = f"SELECT * FROM {table}" + (f" WHERE {where}" if where else "")
    with connect(path) as con:
        return pd.read_sql_query(sql, con, params=params)


def query(sql: str, params: tuple = (), path: str | None = None) -> pd.DataFrame:
    with connect(path) as con:
        return pd.read_sql_query(sql, con, params=params)


def execute(sql: str, params: tuple = (), path: str | None = None) -> int:
    with connect(path) as con:
        cur = con.execute(sql, params)
        return cur.lastrowid


def executemany(sql: str, rows: list, path: str | None = None) -> None:
    with connect(path) as con:
        con.executemany(sql, [tuple(py(v) for v in r) for r in rows])


def py(v):
    """numpy/pandas 스칼라 → 파이썬 기본형 (SQLite BLOB 저장 방지)."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(v, "item"):
        return v.item()
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    return v


def replace_table(table: str, df: pd.DataFrame, path: str | None = None) -> None:
    """테이블 내용을 df 로 통째로 교체 (스키마 유지)."""
    with connect(path) as con:
        con.execute(f"DELETE FROM {table}")
        if len(df):
            cols = list(df.columns)
            ph = ",".join("?" * len(cols))
            con.executemany(
                f"INSERT INTO {table} ({','.join(cols)}) VALUES ({ph})",
                [tuple(py(v) for v in r) for r in df.itertuples(index=False)],
            )


def log_change(target: str, content: str, reason: str = "", who: str = "", path: str | None = None) -> None:
    execute("INSERT INTO change_log(at,target,content,reason,who) VALUES (?,?,?,?,?)",
            (now(), target, content, reason, who), path)


def is_empty(path: str | None = None) -> bool:
    """공정·불량유형이 하나도 없으면 초기 설정 전."""
    with connect(path) as con:
        n = con.execute("SELECT (SELECT COUNT(*) FROM processes) + (SELECT COUNT(*) FROM defect_types)").fetchone()[0]
        return n == 0


# ---------------------------------------------------------------- 설정
def get_setting(key: str, path: str | None = None):
    with connect(path) as con:
        r = con.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    if r is None:
        return DEFAULT_SETTINGS.get(key)
    try:
        return json.loads(r[0])
    except (TypeError, ValueError):
        return r[0]


def set_setting(key: str, value, path: str | None = None) -> None:
    execute("INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, json.dumps(value, ensure_ascii=False)), path)


def options(name: str) -> list[str]:
    v = get_setting(name)
    return [x for x in (v or []) if str(x).strip()]


# ---------------------------------------------------------------- 공정
def get_processes(active_only: bool = True, path: str | None = None) -> list[str]:
    sql = "SELECT name FROM processes" + (" WHERE active=1" if active_only else "") + " ORDER BY sort, name"
    with connect(path) as con:
        return [r[0] for r in con.execute(sql)]


def ensure_processes(names, path: str | None = None) -> None:
    """없는 공정만 뒤에 추가."""
    have = set(get_processes(False, path))
    with connect(path) as con:
        n = con.execute("SELECT COALESCE(MAX(sort),0) FROM processes").fetchone()[0]
        for x in dict.fromkeys(str(v).strip() for v in names if v is not None and str(v).strip()):
            if x not in have:
                n += 1
                con.execute("INSERT INTO processes(name,sort,active) VALUES (?,?,1)", (x, n))


# ---------------------------------------------------------------- 데이터셋
def list_datasets() -> list[str]:
    os.makedirs(DATA_DIR, exist_ok=True)
    return sorted(os.path.splitext(os.path.basename(f))[0] for f in glob.glob(os.path.join(DATA_DIR, "*.db")))


def dataset_path(name: str) -> str:
    return os.path.join(DATA_DIR, f"{name}.db")


def backup_dir() -> str:
    d = os.path.join(DATA_DIR, "backups")
    os.makedirs(d, exist_ok=True)
    return d


def use_dataset(name: str) -> None:
    global DB_PATH
    DB_PATH = dataset_path(name)
    init_db()
