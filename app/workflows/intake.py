from __future__ import annotations
import json,re
from ..core.result import ActionResult,waiting
from ..core.audit import log_event
from ..core.database import connect,utc_now
from ..core.text import normalize_patient_name
from ..automation.desktop import DesktopEleks
from ..automation.browser import ManualCheckpoint
from ..excel.workbook import sync_coordinator_to_graph,find_patient_row,backup
from .common import settings,graph_path,coordinator_path,graph_sheet,require,patient_key

SECTION_ALIASES={
 "complaints":["скарги","скарги та запит","жалобы"],
 "anamnesis":["анамнез","анамнез захворювання"],
 "neurological":["неврологічний статус","неврологический статус"],
 "objective":["об'єктивно","обʼєктивно","объективно"],
 "diagnosis":["діагноз","диагноз"]
}

def coordinator_to_graph(payload):
    name=require(payload,"patient_name","ПІБ пацієнта"); cfg=settings()
    result=sync_coordinator_to_graph(coordinator_path(),graph_path(payload),name,
                                     cfg.get("coordinator_sheet",""),graph_sheet())
    return ActionResult(True,f"Дані «{name}» перенесено/оновлено у Графік 1.",result)

def neurologist_report(payload):
    action="desktop_neurologist_report"; name=require(payload,"patient_name","ПІБ пацієнта")
    cfg=settings()
    try:
        content=DesktopEleks(action,cfg).find_neurologist_report(name)
    except ManualCheckpoint as cp:
        return waiting(str(cp),cp.code,cp.data)
    with connect() as c:
        c.execute("""INSERT INTO extracted_documents(patient_key,document_type,content,source,created_at)
          VALUES(?,?,?,?,?)""",(patient_key(name),"neurologist_report",content,"Doctor Eleks Desktop",utc_now()))
    log_event("document.extracted","Збережено консультативний висновок невролога",action_id=action,
              data={"patient_key":patient_key(name),"length":len(content)})
    sections=_split_sections(content)
    return ActionResult(True,"Висновок невролога знайдено та збережено локально.",
                        {"sections":{k:v[:1200] for k,v in sections.items()},"length":len(content)})

def _split_sections(text:str):
    txt=text or ""; lower=txt.casefold()
    hits=[]
    for key,aliases in SECTION_ALIASES.items():
        best=None
        for a in aliases:
            m=re.search(rf"(?im)^\s*{re.escape(a)}\s*:?\s*",txt)
            if m and (best is None or m.start()<best.start()):best=m
        if best:hits.append((best.start(),best.end(),key))
    hits.sort()
    out={}
    for i,(start,end,key) in enumerate(hits):
        stop=hits[i+1][0] if i+1<len(hits) else len(txt)
        out[key]=txt[end:stop].strip()
    return out

def _latest_doc(name):
    with connect() as c:
        row=c.execute("""SELECT * FROM extracted_documents WHERE patient_key=? AND document_type='neurologist_report'
          ORDER BY id DESC LIMIT 1""",(patient_key(name),)).fetchone()
    return dict(row) if row else None

def diagnosis_prepare(payload):
    action="diagnosis_prepare"; name=require(payload,"patient_name","ПІБ пацієнта")
    doc=_latest_doc(name)
    if not doc:
        return ActionResult(False,"Спочатку виконайте «Знайти висновок невролога».",{})
    sections=_split_sections(doc["content"])
    approved=(payload.get("approved_codes") or "").strip()
    if not approved:
        # No autonomous clinical decision: prepare a structured packet for the user's existing formula.
        with connect() as c:
            now=utc_now()
            c.execute("""INSERT INTO diagnosis_drafts(patient_key,payload_json,approved,created_at,updated_at)
              VALUES(?,?,?,?,?)""",(patient_key(name),json.dumps({"sections":sections},ensure_ascii=False),0,now,now))
        return waiting(
          "Матеріал підготовлено. Внесіть коди, підтверджені вашим діагностичним правилом/ChatGPT, у поле «Підтверджені коди» і повторіть функцію.",
          "diagnosis_confirmation",
          {"sections":{k:v[:1800] for k,v in sections.items()},
           "required":["Скарги/запит","Анамнез","Неврологічний статус","Об'єктивно","Діагноз лікаря"]})
    codes=[x.strip() for x in re.split(r"[,;\n]+",approved) if x.strip()]
    if not codes: return ActionResult(False,"Не розпізнано жодного підтвердженого коду.",{})
    # write into N for the unique patient row; preserve source with backup
    path=graph_path(payload); sheet=graph_sheet()
    rows=find_patient_row(path,sheet,name,3)
    if len(rows)!=1:return ActionResult(False,f"У Графік 1 знайдено {len(rows)} рядків пацієнта; автоматичний запис зупинено.",{})
    from openpyxl import load_workbook
    b=backup(path); wb=load_workbook(path); ws=wb[sheet]
    ws.cell(rows[0],14).value=", ".join(codes); wb.save(path); wb.close()
    with connect() as c:
        now=utc_now()
        c.execute("""INSERT INTO diagnosis_drafts(patient_key,payload_json,approved,created_at,updated_at)
          VALUES(?,?,?,?,?)""",(patient_key(name),json.dumps({"sections":sections,"codes":codes},ensure_ascii=False),1,now,now))
    log_event("diagnosis.approved","Підтверджені коди записано в колонку N",action_id=action,
              data={"patient_key":patient_key(name),"count":len(codes),"backup":b})
    return ActionResult(True,f"Записано {len(codes)} підтверджених кодів у колонку N.",{"codes":codes,"backup":b})
