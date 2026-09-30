from __future__ import annotations
import re
from datetime import datetime, timezone, timedelta
from ..core.result import ActionResult,waiting
from ..core.audit import log_event
from ..core.database import get_workflow,set_workflow
from ..automation.browser import WebEleks,ManualCheckpoint
from ..excel.workbook import get_coordinator_status,update_coordinator_status,write_cell
from .common import settings,graph_path,graph_sheet,patient_data,patient_key,require,coordinator_path,fmt_datetime

def _codes(text):
    found=[]
    s=str(text or "")
    for c in re.findall(r"\b96[.\s]?\d{3}\b",s):
        code=c.replace(" ","")
        if "." not in code and len(code)==5:code=code[:2]+"."+code[2:]
        if code not in found:found.append(code)
    return found

def _manual_sign_step(web,action,pkey,data):
    try:
        web.click("save_sign")
        set_workflow(action,pkey,"awaiting_kep",data)
        return waiting("Накладіть КЕП у відкритому вікні. Після успішного підпису повторіть цю функцію.","kep",data)
    except ManualCheckpoint as cp:
        if cp.code=="safe_mode":
            set_workflow(action,pkey,"safe_mode_final",data)
            return waiting("Безпечний режим не натиснув фінальну кнопку. Перевірте форму, натисніть «Зберегти та підписати» вручну, накладіть КЕП і повторіть функцію.","safe_mode_final",data)
        return waiting(str(cp),cp.code,{**data,**cp.data})

def _resume_simple_signature(action,pkey,name,cfg):
    st=get_workflow(action,pkey)
    if not st or st["state"] not in {"awaiting_kep","safe_mode_final"}:return None
    # A rerun is treated as explicit user confirmation that the manual KEP step was completed.
    # We still reconnect and verify the patient page can be opened, avoiding silent advancement while logged out.
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name)
            set_workflow(action,pkey,"complete",{"manual_signature_confirmed":True})
            return ActionResult(True,"Етап КЕП прийнято як завершений після повторного запуску.",{"resumed":True})
    except ManualCheckpoint as cp:
        return waiting(str(cp),cp.code,cp.data)

def submit_care_plan(payload):
    action="care_plan_submit";name=require(payload,"patient_name","ПІБ пацієнта");cfg=settings();pkey=patient_key(name)
    resumed=_resume_simple_signature(action,pkey,name,cfg)
    if resumed:return resumed
    try:status=get_coordinator_status(coordinator_path(),name,cfg.get("coordinator_sheet",""))
    except Exception as e:return ActionResult(False,f"Не вдалося перевірити статус координатора: {e}",{})
    if "кидайте" not in str(status or "").casefold():
        return ActionResult(False,f"Статус координатора зараз «{status}». Для подання очікується «Кидайте».",{"status":status})
    pdata=patient_data(name,payload)
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name,pdata.get("dob"));web.click("treatment_plans")
            if web.page_has("чернет"):web.choose_text("чернет",optional=True)
            web.click("actions");web.click("edit");web.click("create_request")
            return _manual_sign_step(web,action,pkey,{"coordinator_status":status})
    except ManualCheckpoint as cp:return waiting(str(cp),cp.code,cp.data)

def _request_otp(web,action,pkey,name,cfg,codes,attempts):
    web.click("primary_prescription")
    web.click("auth_method",optional=True);web.click("actions",optional=True);web.click("auth_access",optional=True);web.click("create_request")
    attempts+=1
    requested=datetime.now(timezone.utc).isoformat()
    set_workflow(action,pkey,"awaiting_otp",{"requested_at":requested,"attempts":attempts,"codes":codes})
    try:update_coordinator_status(coordinator_path(),name,"Код надіслано",cfg.get("coordinator_sheet",""))
    except Exception as e:log_event("coordinator.status.warning",str(e),level="WARNING",action_id=action)
    return waiting("Код автентифікації надіслано. Отримайте 4 цифри у координатора і повторіть функцію з OTP.","otp",
                   {"expires_minutes":10,"max_requests":2,"attempts":attempts,"codes":codes})

def _otp_expired(st):
    try:
        requested=datetime.fromisoformat((st.get("data") or {}).get("requested_at",""))
        if requested.tzinfo is None:requested=requested.replace(tzinfo=timezone.utc)
        return datetime.now(timezone.utc)-requested > timedelta(minutes=10)
    except:return False

def primary_prescriptions(payload):
    action="primary_prescriptions";name=require(payload,"patient_name","ПІБ пацієнта");cfg=settings();pkey=patient_key(name)
    pdata=patient_data(name,payload);codes=_codes(pdata.get("codes_specialists"))
    if not codes:return ActionResult(False,"У колонці R не знайдено кодів первинних призначень.",{"raw":str(pdata.get("codes_specialists") or "")})
    otp=str(payload.get("otp") or "").strip();request_new=bool(payload.get("request_new_code",False))
    st=get_workflow(action,pkey)
    # On rerun after a KEP checkpoint, the current code is considered signed by explicit user action.
    if st and st["state"] in {"awaiting_kep","safe_mode_final"}:
        data=st.get("data") or {};completed=list(data.get("completed",[]));current=data.get("current")
        if current and current not in completed:completed.append(current)
        set_workflow(action,pkey,"creating_prescriptions",{"codes":codes,"completed":completed})
        st=get_workflow(action,pkey)
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name,pdata.get("dob"));web.click("treatment_plans")
            st=get_workflow(action,pkey)
            if not st or st["state"] not in {"awaiting_otp","creating_prescriptions","complete"}:
                return _request_otp(web,action,pkey,name,cfg,codes,0)
            if st["state"]=="awaiting_otp":
                attempts=int((st.get("data") or {}).get("attempts",1))
                expired=_otp_expired(st)
                if request_new:
                    if attempts>=2:return ActionResult(False,"OTP уже запитувався двічі. Автоматичне третє отримання заблоковано.",{"attempts":attempts})
                    return _request_otp(web,action,pkey,name,cfg,codes,attempts)
                if not otp:
                    msg="Введіть 4-значний OTP."
                    if expired:msg+=" Попередній код старший за 10 хвилин; увімкніть «Повторно надіслати OTP», якщо потрібен другий код."
                    return waiting(msg,"otp",{"attempts":attempts,"expired":expired,"codes":codes})
                if expired:
                    return ActionResult(False,"Введений OTP відноситься до заявки старшої за 10 хвилин. Запитайте новий код (якщо ще не використано 2 спроби).",{"attempts":attempts})
                if not re.fullmatch(r"\d{4}",otp):return ActionResult(False,"OTP повинен містити рівно 4 цифри.",{})
                web.fill("otp",otp);web.click("save",optional=True);web.click("save_exit",optional=True)
                set_workflow(action,pkey,"creating_prescriptions",{"codes":codes,"completed":[]})
            st=get_workflow(action,pkey) or {"data":{"codes":codes,"completed":[]}}
            if st["state"]=="complete":
                return ActionResult(True,"Первинні призначення вже завершені.",{"codes":codes})
            completed=list((st.get("data") or {}).get("completed",[]))
            first_dt=fmt_datetime(pdata.get("first_interaction"));start_date=first_dt.split(" ")[0] if first_dt else ""
            for code in codes:
                if code in completed:continue
                web.click("actions",optional=True);web.click("primary_prescription",optional=True)
                web.fill("prescription_code",code,optional=True)
                web.fill("prescription_minutes","420",optional=True)
                if start_date:web.fill("prescription_start",start_date,optional=True)
                web.fill("prescription_duration","1",optional=True)
                if web.click("prescription_duration_unit",optional=True):web.choose_text("місяць",optional=True)
                if web.click("goal",optional=True):web.click("support_goal",optional=True)
                data={"codes":codes,"completed":completed,"current":code}
                result=_manual_sign_step(web,action,pkey,data)
                if result:return result
            set_workflow(action,pkey,"complete",{"codes":codes,"completed":codes})
            try:write_cell(graph_path(payload),graph_sheet(),pdata["row"],17,"E + призначення")
            except Exception as e:log_event("excel.status.warning",str(e),level="WARNING",action_id=action)
            return ActionResult(True,f"Первинні призначення створено: {len(codes)}. Q оновлено на «E + призначення».",{"codes":codes})
    except ManualCheckpoint as cp:return waiting(str(cp),cp.code,cp.data)

def referrals_from_prescriptions(payload):
    action="referrals_from_prescriptions";name=require(payload,"patient_name","ПІБ пацієнта");cfg=settings();pkey=patient_key(name)
    pdata=patient_data(name,payload);codes=_codes(pdata.get("codes_specialists"))
    if not codes:return ActionResult(False,"У колонці R не знайдено кодів призначень.",{})
    st=get_workflow(action,pkey)
    if st and st["state"] in {"awaiting_kep","safe_mode_final"}:
        data=st.get("data") or {};completed=list(data.get("completed",[]));current=data.get("current")
        if current and current not in completed:completed.append(current)
        set_workflow(action,pkey,"creating_referrals",{"completed":completed})
        st=get_workflow(action,pkey)
    completed=list((st.get("data") or {}).get("completed",[])) if st else []
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name,pdata.get("dob"));web.click("treatment_plans");web.click("primary_prescriptions_section",optional=True)
            for code in codes:
                if code in completed:continue
                if not web.choose_text(code,optional=True):
                    raise ManualCheckpoint("prescription_select",f"Не вдалося знайти первинне призначення {code}. Відкрийте його вручну і повторіть.")
                web.click("actions");web.click("create_referral")
                # Important: do NOT click "take_to_work"; other specialists take these referrals.
                result=_manual_sign_step(web,action,pkey,{"completed":completed,"current":code})
                if result:return result
            write_cell(graph_path(payload),graph_sheet(),pdata["row"],16,"повний")
            write_cell(graph_path(payload),graph_sheet(),pdata["row"],17,"E + направлення")
            set_workflow(action,pkey,"complete",{"completed":codes})
            return ActionResult(True,"Усі направлення створено без взяття в роботу. P = «повний», Q = «E + направлення».",{"codes":codes})
    except ManualCheckpoint as cp:return waiting(str(cp),cp.code,cp.data)
