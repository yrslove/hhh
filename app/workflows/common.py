from __future__ import annotations
from pathlib import Path
from datetime import datetime,date,timedelta
from ..core.config import load_json
from ..core.database import get_setting
from ..core.text import normalize_patient_name
from ..excel.workbook import get_graph_patient

def settings():
    cfg=load_json("settings.json",{})
    # runtime overrides stored in DB
    overrides=get_setting("settings_overrides",{}) or {}
    cfg.update(overrides)
    return cfg

def graph_path(payload=None):
    payload=payload or {}
    s=settings()
    p=payload.get("workbook_path") or s.get("graph1_path") or get_setting("current_workbook","")
    if not p or not Path(p).exists():raise ValueError("Не підключено Графік 1. Завантажте Excel або вкажіть шлях у Налаштуваннях.")
    return str(Path(p))

def coordinator_path():
    p=settings().get("coordinators_path","")
    if not p or not Path(p).exists():raise ValueError("Не налаштовано шлях до таблиці координаторів.")
    return str(Path(p))

def graph_sheet():
    return settings().get("distribution_sheet","Розподіл AR2")

def patient_data(name:str,payload=None,workbook=None):
    payload=payload or {}
    return get_graph_patient(graph_path(payload),graph_sheet(),name,workbook=workbook,
                             month=payload.get("month"))

def patient_key(name:str):
    return normalize_patient_name(name)

def fmt_datetime(v):
    if v is None:return None
    if isinstance(v,datetime):
        v=(v+timedelta(seconds=30)).replace(second=0,microsecond=0)
        return v.strftime("%d.%m.%Y %H:%M")
    if isinstance(v,date):return v.strftime("%d.%m.%Y 00:00")
    return str(v)

def require(payload,name,label=None):
    value=(payload or {}).get(name)
    if value is None or (isinstance(value,str) and not value.strip()):
        raise ValueError(f"Потрібно заповнити: {label or name}.")
    return value
