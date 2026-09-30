from __future__ import annotations
import json, traceback
from datetime import date, datetime, time, timezone
from pathlib import Path
from typing import Any
from .database import connect
from .paths import AUDIT_JSONL, APP_LOG

SENSITIVE={"password","kep_password","otp","otp_code","sms_code","token","api_key","document_text","content"}

def _safe(v:Any,key:str|None=None):
    if key and key.lower() in SENSITIVE: return "[REDACTED]"
    if isinstance(v,(datetime,date,time)): return v.isoformat()
    if isinstance(v,dict): return {str(k):_safe(x,str(k)) for k,x in v.items()}
    if isinstance(v,(list,tuple)): return [_safe(x) for x in v]
    if isinstance(v,Path): return str(v)
    if isinstance(v,str) and len(v)>1500:return v[:1500]+"…"
    return v

def log_event(event_type:str,message:str,*,level="INFO",action_id=None,data=None):
    ts=datetime.now(timezone.utc).isoformat(timespec="milliseconds")
    safe=_safe(data or {})
    payload={"created_at":ts,"event_type":event_type,"level":level,"action_id":action_id,
             "message":message,"data":safe}
    AUDIT_JSONL.parent.mkdir(parents=True,exist_ok=True)
    with AUDIT_JSONL.open("a",encoding="utf-8") as f:
        f.write(json.dumps(payload,ensure_ascii=False)+"\n")
    with APP_LOG.open("a",encoding="utf-8") as f:
        f.write(f"{ts} [{level}] {event_type} {action_id or '-'} {message}\n")
    with connect() as c:
        cur=c.execute("""INSERT INTO audit_events(created_at,event_type,level,action_id,message,data_json)
          VALUES(?,?,?,?,?,?)""",(ts,event_type,level,action_id,message,json.dumps(safe,ensure_ascii=False)))
        return int(cur.lastrowid)

def log_exception(action_id:str, exc:Exception, data=None):
    return log_event("action.error",f"{type(exc).__name__}: {exc}",level="ERROR",action_id=action_id,
                     data={**(data or {}),"traceback":traceback.format_exc(limit=12)})

def get_events(limit=300):
    limit=max(1,min(int(limit),1000))
    with connect() as c:
        rows=c.execute("""SELECT id,created_at,event_type,level,action_id,message,data_json
          FROM audit_events ORDER BY id DESC LIMIT ?""",(limit,)).fetchall()
    out=[]
    for r in rows:
        d=dict(r)
        try:d["data"]=json.loads(d.pop("data_json"))
        except Exception:d["data"]={}
        out.append(d)
    return out
