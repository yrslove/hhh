from __future__ import annotations
import json, os, shutil, uuid
from datetime import datetime
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .core.paths import WEB_DIR,STATIC_DIR,IMPORT_DIR
from .core.database import init_db,set_setting,get_setting,get_check_run
from .core.audit import get_events,log_event
from .core.function_registry import registry
from .core.config import load_json,save_json
from .core.ui_layout import load_layout,save_layout,create_group,rename_group
from .core.action_control import active_actions,cancel_all
from .automation.browser import stop_active_playwright_drivers

app=FastAPI(title="Doctor Eleks Automation",version="1.0.0")
app.mount("/static",StaticFiles(directory=STATIC_DIR),name="static")

@app.middleware("http")
async def disable_ui_cache(request:Request,call_next):
    response=await call_next(request)
    response.headers["Cache-Control"]="no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"]="no-cache"
    response.headers["Expires"]="0"
    return response

class ExecuteBody(BaseModel):
    payload: dict = {}
class LayoutBody(BaseModel):
    groups:list
    functions:dict
class GroupBody(BaseModel):
    title:str
class RenameGroupBody(BaseModel):
    title:str
class CustomizeBody(BaseModel):
    title:str|None=None
    note:str|None=None
    group_id:str|None=None
    order:int|None=None
class SettingsBody(BaseModel):
    values:dict

def _lifecycle_log(state:str,message:str):
    stamp=datetime.now().astimezone().isoformat(timespec="milliseconds")
    pid=os.getpid()
    line=f"[{stamp}] [APP {state}] pid={pid} — {message}"
    print(line,flush=True)
    try:
        log_event("application.lifecycle",message,action_id="application",
                  data={"state":state.casefold(),"pid":pid,"timestamp":stamp})
    except Exception:
        # Console output must still confirm reload even if audit storage is unavailable.
        pass

@app.on_event("startup")
def startup():
    init_db()
    registry.reload()
    _lifecycle_log("STARTED","Застосунок запущено; початкове завантаження або reload завершено.")

@app.on_event("shutdown")
def shutdown():
    _lifecycle_log("STOPPED","Застосунок зупинено; якщо увімкнено --reload, очікується новий запуск.")

@app.get("/")
def index():
    return FileResponse(WEB_DIR/"index.html")

@app.get("/api/health")
def health():
    return {"ok":True,"version":"1.0.0"}

@app.get("/api/functions")
def functions():
    return registry.merged()

@app.post("/api/functions/{function_id}/execute")
def execute(function_id:str, body:ExecuteBody):
    return registry.execute(function_id,body.payload)

@app.get("/api/actions/status")
def actions_status():
    return {"actions":active_actions()}

@app.post("/api/actions/cancel")
def cancel_actions():
    count=cancel_all()
    drivers=stop_active_playwright_drivers()
    log_event("action.cancel.request","Запитано зупинку активних функцій",level="WARNING",
              data={"count":count,"drivers_stopped":drivers})
    return {"ok":True,"count":count,"drivers_stopped":drivers}

@app.post("/api/upload")
async def upload(file:UploadFile=File(...)):
    name=Path(file.filename or "workbook.xlsx").name
    if Path(name).suffix.lower() not in {".xlsx",".xlsm"}:
        raise HTTPException(400,"Підтримуються тільки XLSX/XLSM.")
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dst=IMPORT_DIR/f"{stamp}_{name}"
    with dst.open("wb") as out:
        while chunk:=await file.read(1024*1024):
            out.write(chunk)
    set_setting("current_workbook",str(dst))
    log_event("file.upload","Завантажено робочу копію Excel",data={"path":str(dst),"size":dst.stat().st_size})
    # sheet names for UI
    try:
        from openpyxl import load_workbook
        wb=load_workbook(dst,read_only=True,data_only=True)
        sheets=wb.sheetnames
        wb.close()
    except Exception:sheets=[]
    return {"ok":True,"path":str(dst),"name":name,"sheets":sheets}

@app.get("/api/current-workbook")
def current_workbook():
    return {"path":get_setting("current_workbook","")}

@app.get("/api/logs")
def logs(limit:int=300):
    return {"events":get_events(limit)}

@app.get("/api/settings")
def get_settings():
    return load_json("settings.json",{})

@app.post("/api/settings")
def set_settings(body:SettingsBody):
    old=load_json("settings.json",{})
    allowed={
      "graph1_path","coordinators_path","desktop_eleks_path","ehealth_base_url","browser_executable",
      "remote_debug_port","frm_sheet","etft_sheet","distribution_sheet","coordinator_sheet",
      "safe_mode","schedule_year","specialists","default_place"
    }
    for k,v in body.values.items():
        if k in allowed:old[k]=v
    save_json("settings.json",old)
    log_event("settings.update","Оновлено налаштування",data={"keys":sorted([k for k in body.values if k in allowed])})
    return {"ok":True,"settings":old}

@app.get("/api/layout")
def get_layout():
    defs=load_json("functions.json",{})
    return load_layout(defs)

@app.post("/api/layout")
def set_layout(body:LayoutBody):
    layout=save_layout({"groups":body.groups,"functions":body.functions})
    return {"ok":True,"layout":layout}

@app.post("/api/groups")
def add_group(body:GroupBody):
    if not body.title.strip():raise HTTPException(400,"Назва групи порожня.")
    defs=load_json("functions.json",{})
    return {"ok":True,"layout":create_group(defs,body.title.strip())}

@app.post("/api/groups/{group_id}/rename")
def group_rename(group_id:str,body:RenameGroupBody):
    defs=load_json("functions.json",{})
    return {"ok":True,"layout":rename_group(defs,group_id,body.title.strip())}

@app.post("/api/functions/{function_id}/customize")
def customize(function_id:str,body:CustomizeBody):
    defs=load_json("functions.json",{})
    layout=load_layout(defs)
    if function_id not in layout.get("functions",{}):
        raise HTTPException(404,"Функцію не знайдено.")
    item=layout["functions"][function_id]
    if body.title is not None:item["title"]=body.title.strip()[:120] or item["title"]
    if body.note is not None:item["note"]=body.note[:300]
    if body.group_id is not None:item["group_id"]=body.group_id
    if body.order is not None:item["order"]=int(body.order)
    layout=save_layout(layout)
    return {"ok":True,"layout":layout}

@app.get("/api/check-runs/{run_id}")
def check_run(run_id:int):
    data=get_check_run(run_id)
    if not data:raise HTTPException(404,"Результат не знайдено.")
    return data
