from __future__ import annotations
from time import perf_counter
from ..core.result import ActionResult
from ..core.audit import log_event
from ..core.database import save_check_run
from ..excel.schedule import (parse_schedule,build_schedule_index,check_frm_before_etft,
                              check_etft_overlaps,serialize)
from .common import graph_path,settings

def _log_violations(action_id,errors):
    for e in errors:
        log_event("schedule.violation",e.get("message",e.get("type","Порушення")),level="ERROR",
                  action_id=action_id,data=e)

def _log_ar3_monoservice(action_id,items):
    for item in items:
        log_event("schedule.ar3_monoservice",item.type,level="INFO",action_id=action_id,data=item.as_dict())

def _log_technical_warnings(action_id,warnings):
    """Keep parser diagnostics in the audit trail, never in user results."""
    for warning in warnings:
        log_event("schedule.technical",warning.message,level="DEBUG",action_id=action_id,data=warning.as_dict())

def _summary(appts):
    patients={a.patient_name_normalized for a in appts}
    dates={a.date for a in appts}
    return len(appts),len(patients),len(dates)

def _schedule_human_log(title, errors):
    """Build the visible schedule result, grouped by the first found month."""
    if not errors:
        return f"{title}: 0\nПомилок не знайдено."

    ordered=sorted(errors,key=lambda item:(str(item.get("date") or ""),str(item.get("message") or "")))
    lines=[f"{title}: {len(ordered)}"]
    current_month=None
    for item in ordered:
        date_value=str(item.get("date") or "")
        month=date_value[5:7] if len(date_value)>=7 and date_value[4]=="-" else ""
        month_title=f"{month} МЕСЯЦ" if month else "МЕСЯЦ НЕ УКАЗАН"
        if month_title!=current_month:
            current_month=month_title
            lines.extend(["",month_title])
        lines.append(str(item.get("message") or item.get("type") or "Помилка"))
    return "\n".join(lines)

def run_frm_before(payload):
    action="check_frm_before_etft"; path=graph_path(payload); cfg=settings()
    frm_sheet=payload.get("frm_sheet") or cfg.get("frm_sheet","ФРМ")
    etft_sheet=payload.get("etft_sheet") or cfg.get("etft_sheet","ЕТФТ")
    log_event("schedule.start","Запущено перевірку ФРМ → ЕТ/ФТ",action_id=action,data={"file":path,"frm":frm_sheet,"etft":etft_sheet})
    schedule=build_schedule_index(path,frm_sheet,etft_sheet,int(cfg.get("schedule_year",2026)))
    frm,et=schedule.frm_appointments,schedule.etft_appointments
    log_event("schedule.parsed",f"Прочитано записів ФРМ: {len(frm)}; ЕТ/ФТ: {len(et)}",action_id=action,
              data={"frm_records":len(frm),"etft_records":len(et),"frm_sheet":frm_sheet,"etft_sheet":etft_sheet})
    errors,hourly=check_frm_before_etft(frm,et,schedule.etft_by_patient_date)
    log_event("schedule.hourly",f"Годинних записів ФРМ: {hourly}",action_id=action,data={"hourly":hourly})
    _log_violations(action,errors)
    _log_ar3_monoservice(action,schedule.ar3_monoservice)
    _log_technical_warnings(action,schedule.structure_warnings)
    fa,fp,fd=_summary(frm); ea,ep,ed=_summary(et)
    summary={"frm_records":fa,"frm_hourly":hourly,"frm_primary":len(schedule.frm_primary),
             "frm_repeat_ignored":len(schedule.frm_repeat),"ar3_monoservice":len(schedule.ar3_monoservice),
             "etft_records":ea,"etft_patients":ep,
             "dates_checked":len({a.date for a in frm}|{a.date for a in et}),"errors":len(errors)}
    result={"errors":errors,"ar3_monoservice":serialize(schedule.ar3_monoservice),"summary":summary,
            "human_log":_schedule_human_log("Перевірка ФРМ → ЕТ/ФТ",errors)}
    rid=save_check_run(action,path,summary,result)
    log_event("schedule.finish",f"Перевірка завершена. Знайдено помилок: {len(errors)}",action_id=action,data=summary)
    return ActionResult(True,f"Перевірено. Помилок порядку ФРМ → ЕТ/ФТ: {len(errors)}.",{**result,"run_id":rid})

def run_overlaps(payload):
    action="check_etft_overlaps"; path=graph_path(payload); cfg=settings()
    sheet=payload.get("etft_sheet") or cfg.get("etft_sheet","ЕТФТ")
    log_event("schedule.start","Запущено перевірку перетинів ЕТ/ФТ",action_id=action,data={"file":path,"sheet":sheet})
    et,w=parse_schedule(path,sheet,int(cfg.get("schedule_year",2026)))
    log_event("schedule.parsed",f"Прочитано записів ЕТ/ФТ: {len(et)}",action_id=action,
              data={"etft_records":len(et),"sheet":sheet})
    errors=check_etft_overlaps(et); _log_violations(action,errors)
    _log_technical_warnings(action,w)
    ea,ep,ed=_summary(et)
    summary={"etft_records":ea,"etft_patients":ep,"dates_checked":ed,"errors":len(errors)}
    result={"errors":errors,"summary":summary,
            "human_log":_schedule_human_log("Перетини ЕТ/ФТ",errors)}
    rid=save_check_run(action,path,summary,result)
    log_event("schedule.finish",f"Перевірка завершена. Перетинів: {len(errors)}",action_id=action,data=summary)
    return ActionResult(True,f"Перевірено. Перетинів ЕТ/ФТ: {len(errors)}.",{**result,"run_id":rid})

def run_all(payload):
    action="important_000"; path=graph_path(payload); cfg=settings()
    started=perf_counter()
    frm_sheet=payload.get("frm_sheet") or cfg.get("frm_sheet","ФРМ")
    etft_sheet=payload.get("etft_sheet") or cfg.get("etft_sheet","ЕТФТ")
    log_event("schedule.start","IMPORTANT_000: запущено всі перевірки розкладу",action_id=action,data={"file":path})
    # One XLSX/XLSM import; both checks consume this immutable set of parsed appointments.
    schedule=build_schedule_index(path,frm_sheet,etft_sheet,int(cfg.get("schedule_year",2026)))
    frm,et=schedule.frm_appointments,schedule.etft_appointments
    log_event("schedule.parsed",f"IMPORTANT_000: ФРМ {len(frm)}, ЕТ/ФТ {len(et)}",action_id=action,
              data={"frm_records":len(frm),"etft_records":len(et),"frm_sheet":frm_sheet,"etft_sheet":etft_sheet})
    order,hourly=check_frm_before_etft(frm,et,schedule.etft_by_patient_date)
    log_event("schedule.hourly",f"Годинних записів ФРМ: {hourly}",action_id=action,data={"hourly":hourly})
    log_event("schedule.frm_occurrences",
              f"PRIMARY ФРМ: {len(schedule.frm_primary)}; REPEAT ФРМ пропущено: {len(schedule.frm_repeat)}",
              action_id=action,data={"frm_primary":len(schedule.frm_primary),"frm_repeat_ignored":len(schedule.frm_repeat)})
    if schedule.ar3_monoservice:
        log_event("schedule.ar3_monoservice_count",f"54 пакет група АР3-1 (монопослуга): {len(schedule.ar3_monoservice)}",
                  level="INFO",action_id=action,data={"ar3_monoservice":len(schedule.ar3_monoservice)})
    for primary in schedule.frm_primary:
        repeat_cell=primary.paired_with_cell or ""
        log_event("schedule.frm_pair",
                  f"PRIMARY {primary.date} {primary.cell_address}; REPEAT {primary.paired_with_date} {repeat_cell}; gap {primary.pair_gap_days}",
                  level="DEBUG",action_id=action,
                  data={"patient_key":primary.patient_name_normalized,"primary_date":primary.date,
                        "primary_cell":primary.cell_address,"repeat_date":primary.paired_with_date,
                        "repeat_cell":repeat_cell,"gap_days":primary.pair_gap_days})
    overlaps=check_etft_overlaps(et)
    all_errors=order+overlaps; _log_violations(action,all_errors)
    _log_ar3_monoservice(action,schedule.ar3_monoservice)
    _log_technical_warnings(action,schedule.structure_warnings)
    fa,fp,fd=_summary(frm); ea,ep,ed=_summary(et)
    summary={"frm_records":fa,"frm_hourly":hourly,"frm_primary":len(schedule.frm_primary),
             "frm_repeat_ignored":len(schedule.frm_repeat),"ar3_monoservice":len(schedule.ar3_monoservice),
             "etft_records":ea,"etft_patients":ep,
             "dates_checked":len({a.date for a in frm}|{a.date for a in et}),
             "order_errors":len(order),"overlap_errors":len(overlaps),"errors":len(all_errors),
             "duration_ms":round((perf_counter()-started)*1000,1)}
    result={"order_errors":order,"overlap_errors":overlaps,
            "ar3_monoservice":serialize(schedule.ar3_monoservice),"summary":summary,
            "human_log":_schedule_human_log("ЕТ_ФТ_ФРМ_Пересечение времени",all_errors)}
    rid=save_check_run(action,path,summary,result)
    log_event("schedule.finish",f"IMPORTANT_000 завершено. Помилок: {len(all_errors)}",action_id=action,data=summary)
    return ActionResult(True,f"IMPORTANT_000 завершено: {len(order)} помилок порядку, {len(overlaps)} перетинів.",{**result,"run_id":rid})
