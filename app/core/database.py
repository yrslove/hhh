from __future__ import annotations
import json, sqlite3
from datetime import datetime, timezone
from typing import Any
from .paths import DB_PATH

def utc_now():
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")

def connect():
    conn=sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory=sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings(
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS audit_events(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at TEXT NOT NULL,
  event_type TEXT NOT NULL,
  level TEXT NOT NULL,
  action_id TEXT,
  message TEXT NOT NULL,
  data_json TEXT NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS workflow_states(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  function_id TEXT NOT NULL,
  patient_key TEXT NOT NULL DEFAULT '',
  state TEXT NOT NULL,
  data_json TEXT NOT NULL DEFAULT '{}',
  updated_at TEXT NOT NULL,
  UNIQUE(function_id, patient_key)
);
CREATE TABLE IF NOT EXISTS extracted_documents(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  patient_key TEXT NOT NULL,
  document_type TEXT NOT NULL,
  content TEXT NOT NULL,
  source TEXT,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS diagnosis_drafts(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  patient_key TEXT NOT NULL,
  payload_json TEXT NOT NULL,
  approved INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS check_runs(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  function_id TEXT NOT NULL,
  workbook_path TEXT NOT NULL,
  created_at TEXT NOT NULL,
  summary_json TEXT NOT NULL,
  result_json TEXT NOT NULL
);
"""

def init_db():
    with connect() as c:
        c.executescript(SCHEMA)

def set_setting(key: str, value: Any):
    with connect() as c:
        c.execute("""INSERT INTO settings(key,value,updated_at) VALUES(?,?,?)
          ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at""",
          (key,json.dumps(value,ensure_ascii=False),utc_now()))

def get_setting(key: str, default=None):
    with connect() as c:
        row=c.execute("SELECT value FROM settings WHERE key=?",(key,)).fetchone()
    if not row: return default
    try: return json.loads(row["value"])
    except Exception: return row["value"]

def set_workflow(function_id: str, patient_key: str, state: str, data: dict|None=None):
    with connect() as c:
        c.execute("""INSERT INTO workflow_states(function_id,patient_key,state,data_json,updated_at)
          VALUES(?,?,?,?,?) ON CONFLICT(function_id,patient_key)
          DO UPDATE SET state=excluded.state,data_json=excluded.data_json,updated_at=excluded.updated_at""",
          (function_id,patient_key or "",state,json.dumps(data or {},ensure_ascii=False),utc_now()))

def get_workflow(function_id: str, patient_key: str):
    with connect() as c:
        row=c.execute("SELECT * FROM workflow_states WHERE function_id=? AND patient_key=?",
                      (function_id,patient_key or "")).fetchone()
    if not row: return None
    d=dict(row)
    try: d["data"]=json.loads(d.pop("data_json"))
    except Exception: d["data"]={}
    return d

def clear_workflow(function_id: str, patient_key: str):
    with connect() as c:
        c.execute("DELETE FROM workflow_states WHERE function_id=? AND patient_key=?",
                  (function_id,patient_key or ""))

def save_check_run(function_id: str, workbook_path: str, summary: dict, result: dict)->int:
    with connect() as c:
        cur=c.execute("""INSERT INTO check_runs(function_id,workbook_path,created_at,summary_json,result_json)
          VALUES(?,?,?,?,?)""",(function_id,workbook_path,utc_now(),
          json.dumps(summary,ensure_ascii=False),json.dumps(result,ensure_ascii=False)))
        return int(cur.lastrowid)

def get_check_run(run_id:int):
    with connect() as c:
        row=c.execute("SELECT * FROM check_runs WHERE id=?",(run_id,)).fetchone()
    if not row:return None
    d=dict(row); d["summary"]=json.loads(d.pop("summary_json")); d["result"]=json.loads(d.pop("result_json"))
    return d
