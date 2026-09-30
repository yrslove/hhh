from __future__ import annotations
import time
from typing import Callable,Any
from .config import load_json
from .result import ActionResult
from .audit import log_event,log_exception
from .ui_layout import load_layout
from .action_control import action_scope, set_progress, raise_if_cancelled, ActionCancelled

from ..workflows.procedures import create_procedures,batch_sign
from ..workflows.schedule_checks import run_all,run_frm_before,run_overlaps
from ..workflows.procedure_counts import count_procedures
from ..workflows.intake import coordinator_to_graph,neurologist_report,diagnosis_prepare
from ..workflows.first_interaction import (
    diagnostic_interaction,rehab_referral_from_diagnostic,rehab_interaction,
    rehab_referral_after_interaction,care_plan_draft,first_rehab_interaction
)
from ..workflows.care_plan import submit_care_plan,primary_prescriptions,referrals_from_prescriptions

HANDLERS={
 "procedures_create":create_procedures,
 "procedures_batch_sign":batch_sign,
 "important_000":run_all,
 "check_frm_before_etft":run_frm_before,
 "check_etft_overlaps":run_overlaps,
 "procedure_counts":count_procedures,
 "coordinator_to_graph":coordinator_to_graph,
 "desktop_neurologist_report":neurologist_report,
 "diagnosis_prepare":diagnosis_prepare,
 "diagnostic_interaction":diagnostic_interaction,
 "rehab_referral_from_diagnostic":rehab_referral_from_diagnostic,
 "rehab_interaction":rehab_interaction,
 "first_rehab_interaction":first_rehab_interaction,
 "rehab_referral_after_interaction":rehab_referral_after_interaction,
 "care_plan_draft":care_plan_draft,
 "care_plan_submit":submit_care_plan,
 "primary_prescriptions":primary_prescriptions,
 "referrals_from_prescriptions":referrals_from_prescriptions,
}

class FunctionRegistry:
    def __init__(self):
        self.reload()

    def reload(self):
        self._defs=load_json("functions.json",{})
        self._flat={}
        for g in self._defs.get("groups",[]):
            for fn in g.get("functions",[]):
                item=dict(fn); item["default_group_id"]=g["id"]; item["default_group_title"]=g["title"]
                self._flat[item["id"]]=item

    def merged(self):
        self.reload()
        layout=load_layout(self._defs)
        groups={g["id"]:{**g,"functions":[]} for g in layout["groups"]}
        # retain defaults if a saved group disappeared
        for g in self._defs.get("groups",[]):
            groups.setdefault(g["id"],{"id":g["id"],"title":g["title"],"order":g.get("order",99),"functions":[]})
        for fid,fn in self._flat.items():
            ui=layout.get("functions",{}).get(fid,{})
            gid=ui.get("group_id") or fn["default_group_id"]
            if gid not in groups:
                groups[gid]={"id":gid,"title":"Інше","order":999,"functions":[]}
            x=dict(fn)
            x["title"]=ui.get("title") or fn["title"]
            x["note"]=ui.get("note","")
            x["group_id"]=gid
            x["order"]=ui.get("order",999)
            x["status"]="ready" if fid in HANDLERS else "unavailable"
            groups[gid]["functions"].append(x)
        out=[]
        for g in sorted(groups.values(),key=lambda x:x.get("order",999)):
            g=dict(g); g["functions"]=sorted(g["functions"],key=lambda x:x.get("order",999))
            out.append(g)
        return {"groups":out}

    def definition(self,fid): self.reload(); return self._flat.get(fid)

    def execute(self,fid,payload=None):
        payload=payload or {}
        fn=self.definition(fid)
        if not fn:return ActionResult(False,"Невідома функція.",{}).to_dict()
        handler=HANDLERS.get(fid)
        if not handler:return ActionResult(False,"Для цієї функції немає handler.",{}).to_dict()
        started=time.perf_counter()
        log_event("action.start",f"Запущено: {fn['title']}",action_id=fid,
                  data={"payload_keys":sorted(payload.keys())})
        try:
            with action_scope(fid):
                set_progress(8,"Функція виконується")
                raise_if_cancelled()
                try:
                    result=handler(payload)
                except Exception:
                    raise_if_cancelled()
                    raise
                # A handler may translate a browser transport interruption
                # into its ordinary checkpoint result. Honor Stop before that
                # result can be treated as a normal completion.
                raise_if_cancelled()
                set_progress(100,"Завершення")
            duration=round((time.perf_counter()-started)*1000,1)
            log_event("action.finish",result.message,level="INFO" if result.ok else "WARNING",
                      action_id=fid,data={"duration_ms":duration,"status":result.status,
                      "result_keys":sorted((result.data or {}).keys())})
            d=result.to_dict(); d["duration_ms"]=duration; return d
        except ActionCancelled as exc:
            duration=round((time.perf_counter()-started)*1000,1)
            log_event("action.cancelled",str(exc),level="WARNING",action_id=fid,
                      data={"duration_ms":duration})
            return ActionResult(False,str(exc),{"duration_ms":duration},"cancelled").to_dict()
        except Exception as exc:
            duration=round((time.perf_counter()-started)*1000,1)
            log_exception(fid,exc,{"duration_ms":duration})
            return ActionResult(False,f"Помилка: {exc}",{"duration_ms":duration},"failed").to_dict()

registry=FunctionRegistry()
