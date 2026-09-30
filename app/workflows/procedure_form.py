"""State and request handling for the Doctor Eleks procedure dxForm.

Keep the interaction with DevExtreme in one place.  In particular, do not keep
widget instances or DOM nodes across a form update: every read resolves the
current form and its editors again.
"""
from __future__ import annotations

import json
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from ..automation.browser import ManualCheckpoint, wait_visible_recheck
from ..core.audit import log_event
from ..core.paths import RUNTIME_DIR
from ..core.action_control import raise_if_cancelled


SAVE_NAME = "spiprocedureSave"
AUTH_401_MESSAGE = (
    "Сервер Doctor Eleks відхилив запит процедури (HTTP 401). "
    "Відкрита сторінка не підтверджує чинність токена API. "
    "Перевірте доступ і натисніть «Продовжити»; перед повтором програма перевірить наявний запис."
)
PAYLOAD_FIELDS = (
    "id", "patient", "patientRef", "based_on", "basedOn", "code", "performer", "performerRef", "recorded_byRef",
    "paperReferralRef", "encounterRef", "pre_primary_source",
    "report_originsCode", "performed_date_time", "outcome", "division",
    "category", "selectedReferralRef", "procedurePeriodStart",
    "procedurePeriodEnd", "procedurePeriod", "is_date_or_period", "isDateOrPeriod",
    "date", "dateStart", "dateEnd",
    "showErrorMessage", "toValidate",
)
FORM_FIELDS = (
    "procedureRef", "patientRef", "selectedReferralCode", "serviceCode",
    "performerRef", "division", "procedureCategoryRef", "procedureStatus",
    "procedureOutcome", "report_originsCode", "procedure_primary_source",
    "isDateOrPeriod", "procedurePeriod", "date", "dateStart", "dateEnd",
)


def _json_safe(value: Any):
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()
                if str(k).casefold() not in {"authorization", "token", "password", "otp"}}
    if isinstance(value, (list, tuple)):
        return [_json_safe(v) for v in value]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _walk_named(value, names):
    if isinstance(value, dict):
        for key, item in value.items():
            if key in names:
                yield key, item
            yield from _walk_named(item, names)
    elif isinstance(value, list):
        for item in value:
            yield from _walk_named(item, names)


def sanitize_payload(body):
    """Return only the requested semantic fields; never include headers."""
    if not isinstance(body, dict):
        return {}
    # Portal serializes procedure arguments as [{"name": ..., "value": ...}].
    # The values are strings, including the literal string "null".
    params=body.get("params")
    if isinstance(params,list):
        return {entry["name"]:_json_safe(entry.get("value")) for entry in params
                if isinstance(entry,dict) and entry.get("name") in PAYLOAD_FIELDS}
    names=set(PAYLOAD_FIELDS)
    found={key: _json_safe(value) for key, value in _walk_named(body, names)}
    # Keep wrapper flags even when they are adjacent to the procedure object.
    for key in ("showErrorMessage", "toValidate"):
        if key in body:
            found[key]=_json_safe(body[key])
    return found


def _present(value):
    return value is not None and str(value).strip().casefold() not in {"", "null", "undefined"}


def _sanitized_response(value):
    """Keep problem details and response shape without arbitrary server data."""
    if not isinstance(value,dict):
        return {"type":type(value).__name__}
    allowed=("type","title","status","detail","traceId","error","error_description")
    return {key:re.sub(r"Bearer\s+\S+","Bearer [REDACTED]",str(value[key]))[:2000]
            for key in allowed if key in value} or {"keys":list(value)[:30]}


def structured_payload_diff(failed, successful):
    """Diff the requested semantic payload fields, ignoring generated IDs."""
    left=failed.get("payload",failed); right=successful.get("payload",successful)
    ignored={"id"}
    case_insensitive={"patient","patientRef","based_on","basedOn","code",
                      "performer","performerRef","recorded_byRef","paperReferralRef",
                      "encounterRef","division"}
    def different(key):
        a,b=left.get(key),right.get(key)
        if key in case_insensitive and a is not None and b is not None:
            return str(a).casefold()!=str(b).casefold()
        return a!=b
    return {key:{"failed":left.get(key),"manual_success":right.get(key)}
            for key in PAYLOAD_FIELDS if key not in ignored and different(key)}


class DoctorEleksProcedureForm:
    """Fresh-reference procedure form adapter with save capture and hard stops."""

    def __init__(self, web, target=None):
        self.web=web
        self.page=web.page
        # The application lives in a legacy iframe; network events are emitted
        # by its owning top-level Page, while form evaluation stays in Frame.
        self.network_page=getattr(web,"browser_page",None) or web.page
        self.target=target or {}
        self.last_action="FORM_OPEN"
        self._mutations_started=False
        self._console_pending=[]
        self._active_record_id=None
        self._final_save_confirmed=False
        self.saves=[]
        self.failure=None
        self._sequence=0
        self._out=RUNTIME_DIR / "procedure-save-capture.jsonl"
        self._on_request=self._capture_request
        self._on_response=self._capture_response
        self._on_console=self._capture_console
        self.network_page.on("request",self._on_request)
        self.network_page.on("response",self._on_response)
        self.network_page.on("console",self._on_console)
        self._log("PROCEDURE_FORM_OPEN",{})

    def close(self):
        for event, callback in (("request",self._on_request),("response",self._on_response),
                                ("console",self._on_console)):
            try: self.network_page.remove_listener(event,callback)
            except Exception: pass

    def recover_after_failure(self):
        """Take the failure artifacts and force a fresh manual form open."""
        status,record=self.failure or ("UNKNOWN",{})
        state=self._try_read_state()
        try:
            folder=Path(self.web.snapshot_error("procedure_reopen_required"))
            (folder/"procedure-form-state.json").write_text(
                json.dumps(_json_safe(state),ensure_ascii=False,indent=2),encoding="utf-8")
        except Exception: pass
        self._log("FORM_REOPEN",{"status":status,"form_state":state,
                                  "instruction":"Close the damaged procedure and open a new form."})
        self.close()
        if status==401:
            raise ManualCheckpoint("login",AUTH_401_MESSAGE,
                                   {"status":401,"form_state":state})
        raise ManualCheckpoint("procedure_form_reopen",
          "Поточна форма зупинена. Закрийте її та відкрийте процедуру заново, потім продовжуйте з новим станом.",
          {"status":status,"payload":record.get("payload"),"response":record.get("response"),
           "form_state":state,"last_action":self.last_action})

    def _log(self, event, data):
        payload={"event":event,"sequence":self._sequence,"last_action":self.last_action,
                 "target":_json_safe(self.target),**_json_safe(data)}
        log_event("procedure.form."+event.casefold(),event,action_id=self.web.action_id,data=payload)
        try:
            self._out.parent.mkdir(parents=True,exist_ok=True)
            with self._out.open("a",encoding="utf-8") as stream:
                stream.write(json.dumps({"timestamp":datetime.now().astimezone().isoformat(),**payload},
                                        ensure_ascii=False)+"\n")
        except Exception: pass

    def begin_mutation(self, action):
        self._mutations_started=True
        self.last_action=str(action)
        self._log("FIELD_ACTION_BEGIN",{"action":self.last_action})

    @staticmethod
    def _request_json(request):
        try: return request.post_data_json
        except Exception:
            try: return json.loads(request.post_data or "null")
            except Exception: return None

    @staticmethod
    def _is_save(request):
        if request.method != "POST" or "/api/run/portal" not in request.url:
            return False
        body=DoctorEleksProcedureForm._request_json(request)
        return isinstance(body,dict) and body.get("name")==SAVE_NAME

    def _capture_request(self, request):
        try:
            if not self._is_save(request): return
            self._sequence+=1
            body=self._request_json(request)
            record={"sequence":self._sequence,"request":request,"payload":sanitize_payload(body),
                    "status":None,"response":None,"requested_at":time.monotonic()}
            self.saves.append(record)
            self._log("AUTOSAVE_REQUEST",{"save_payload":record["payload"]})
        except Exception: pass

    def _capture_response(self, response):
        try:
            request=response.request
            record=next((item for item in reversed(self.saves) if item.get("request") is request),None)
            if record is None:
                if response.status==401 and "/api/run/portal" in response.url:
                    self.failure=(401,{"payload":None,"response":{"status":401}})
                    self._log("AUTH_EXPIRED",{"status":401,"endpoint":"/api/run/portal"})
                return
            try: result=response.json()
            except Exception:
                try: result=response.text()[:4000]
                except Exception: result=""
            record["status"]=response.status
            record["response"]=_sanitized_response(result)
            record.pop("request",None)
            if response.status == 401:
                self.failure=(401,record)
                self._log("AUTH_EXPIRED",{"status":401,"save_payload":record["payload"],
                                          "response":record["response"]})
            elif response.status >= 400:
                if self.failure is None:
                    self.failure=(response.status,record)
                self._log("AUTOSAVE_FAILURE",{"status":response.status,"save_payload":record["payload"],
                                               "response":record["response"],"form_state":self._try_read_state()})
            else:
                self._log("AUTOSAVE_SUCCESS",{"status":response.status,"save_payload":record["payload"],
                                                "response":record["response"]})
        except Exception: pass

    def _capture_console(self, message):
        try:
            text=message.text or ""
            lowered=text.casefold()
            if not ("e1035" in lowered or "cannot read properties of undefined" in lowered or
                    "node cannot be found in the current page" in lowered):
                return
            if not self._mutations_started:
                # Doctor Eleks/DevExtreme can emit E1035 while the procedure
                # dialog's dependent editors are still initializing. Record
                # it for diagnosis, but do not claim automation corrupted a
                # form before it has changed any field.
                self._log("FORM_INITIALIZATION_CONSOLE",{"console":text[:1000],
                                                          "form_state":self._try_read_state()})
                return
            # This installation emits E1035 during a successful referral and
            # after a 200 division save, while all required widgets remain
            # attached and functional. A console line alone is therefore a
            # diagnostic signal; confirm the live form before stopping it.
            self._console_pending=(getattr(self,"_console_pending",[])+[text[:1000]])[-5:]
            self._log("FORM_CONSOLE_WARNING",{"console":text[:1000],
                                               "last_action":self.last_action})
        except Exception: pass

    def guard(self):
        if not self.failure and getattr(self,"_console_pending",[]):
            if getattr(self,"_final_save_confirmed",False):
                warnings=self._console_pending
                self._console_pending=[]
                self._log("FORM_CLOSED_AFTER_FINAL",{"warnings":warnings,
                    "reason":"final save already confirmed; verify persisted grid row"})
            else:
                self._confirm_console_state()
        if self.failure:
            status,record=self.failure
            if status=="CORRUPTED":
                self._log("FORM_CORRUPTED",{"status":status,"save_payload":record.get("payload"),
                                             "response":record.get("response")})
            if status==401:
                raise ManualCheckpoint("login",AUTH_401_MESSAGE,
                                       {"status":401,"payload":record.get("payload"),
                                        "response":record.get("response")})
            snap=self._try_read_state()
            try:
                folder=Path(self.web.snapshot_error("procedure_autosave_failure"))
                (folder/"procedure-form-state.json").write_text(
                    json.dumps(_json_safe(snap),ensure_ascii=False,indent=2),encoding="utf-8")
                (folder/"spiprocedureSave.json").write_text(
                    json.dumps({"payload":record.get("payload"),"response":record.get("response"),
                                "status":status,"last_action":self.last_action},
                               ensure_ascii=False,indent=2),encoding="utf-8")
            except Exception: pass
            checkpoint="procedure_form_corrupted" if status=="CORRUPTED" else "procedure_autosave_failed"
            message=("DevExtreme/DOM повідомив про пошкоджену або замінену форму. Відкрийте її заново."
                     if status=="CORRUPTED" else
                     f"spiprocedureSave зупинив форму: HTTP {status}. Відкрийте процедуру заново.")
            raise ManualCheckpoint(checkpoint,message,
                                   {"status":status,"console":record.get("console"),"payload":record.get("payload"),
                                    "response":record.get("response"),"form_state":snap,
                                    "last_action":self.last_action})

    def _confirm_console_state(self, timeout_ms=5000):
        """Classify a console warning using fresh editor and route state."""
        warnings=list(self._console_pending)
        deadline=time.monotonic()+timeout_ms/1000
        last_state={}
        while time.monotonic()<deadline:
            raise_if_cancelled()
            if self.failure:return
            if hasattr(self.web,"require_authenticated"):
                self.web.require_authenticated()
            last_state=self._try_read_state()
            data=last_state.get("formData",{})
            editors=last_state.get("editors",{})
            period_mode=str(data.get("isDateOrPeriod")).casefold()=="true"
            required=("division","isDateOrPeriod","dateStart","dateEnd","procedurePeriod")\
                if period_mode else ("division","isDateOrPeriod","date")
            same_record=(not self._active_record_id or
                         str(data.get("procedureRef","")).casefold()==
                         str(self._active_record_id).casefold())
            if (last_state.get("ready") and same_record and
                    all(editors.get(key,{}).get("dataSourceReady") for key in required)):
                self._console_pending=[]
                self._log("FORM_STABLE_AFTER_CONSOLE",{"warnings":warnings,
                    "procedureRef":data.get("procedureRef"),"editors":list(required)})
                return
            self.page.wait_for_timeout(100)
        self.failure=("CORRUPTED",{"console":"\n".join(warnings),
                                    "payload":None,"response":None,
                                    "form_state":last_state})
        self._log("FORM_CORRUPTED",{"reason":"required editors did not recover",
                                    "warnings":warnings,"form_state":last_state})

    def _try_read_state(self):
        try: return self.read_state()
        except Exception as exc: return {"read_error":str(exc)[:200]}

    def wait_ready(self, timeout_ms=20000):
        self.web.require_authenticated()
        try:
            self.page.wait_for_function("""() => {
          const anchor=document.querySelector('input[id$="_isDateOrPeriod"]');
          if(!anchor?.isConnected || !window.angular || document.readyState==='loading')return false;
          const scope=window.angular.element(anchor).scope(), form=scope?.formInstance;
          if(!form || typeof form.option!=='function' || !scope.formVm?.formData)return false;
          const root=anchor.closest('.dx-form');
          if(!root?.isConnected)return false;
          const periodMode=String(scope.formVm.formData.isDateOrPeriod)==='true';
          const required=['isDateOrPeriod','division','performerRef','serviceCode',
            ...(periodMode?['dateStart','dateEnd','procedurePeriod']:['date'])];
          if(required.some(k=>!root.querySelector('input[id$="_'+k+'"]')))return false;
          for(const field of ['isDateOrPeriod','division']) {
            const host=root.querySelector('input[id$="_'+field+'"]').closest('.dx-selectbox');
            const widget=host && window.DevExpress?.ui?.dxSelectBox?.getInstance(host);
            const source=widget?.getDataSource?.();
            const items=widget?.option('items')||
              ((source?.isLoaded?.() && source.items?.())||[]);
            if(!widget || !items.length)return false;
          }
          return true;
            }""",timeout=timeout_ms)
        except Exception as exc:
            diagnostic=self._try_read_state()
            self._log("FORM_READY_TIMEOUT",{"form_state":diagnostic,"reason":str(exc)[:300]})
            raise ManualCheckpoint("procedure_form_ready",
              "Форма процедури не завершила ініціалізацію. Перевірте авторизацію і дочекайтеся видимих полів форми.",
              {"form_state":diagnostic,"reason":str(exc)[:300]}) from exc
        state=self.read_state()
        if not state.get("ready"):
            raise ManualCheckpoint("procedure_form_ready","Форма процедури не пройшла перевірку готовності.",state)
        self._log("FORM_READY",{"form_state":state})
        self._log("FORM_STATE_INITIAL",{"form_state":state})
        return state

    def read_state(self):
        # Entire read happens in one evaluation, and returns values only. No DOM
        # or widget references cross the Playwright boundary.
        return self.page.evaluate("""() => {
          const anchor=document.querySelector('input[id$="_isDateOrPeriod"]');
          const node=anchor?.closest('.dx-form');
          const scope=anchor && window.angular?.element(anchor).scope();
          const form=scope?.formInstance;
          if(!node?.isConnected || !form)return {ready:false,reason:'procedure dxForm missing'};
          const data=form.option('formData')||{}, angularData=scope.formVm?.formData||{};
          const fields=['procedureRef','patientRef','selectedReferralCode','serviceCode',
            'performerRef','division','procedureCategoryRef','procedureStatus',
            'procedureOutcome','report_originsCode','procedure_primary_source',
            'isDateOrPeriod','procedurePeriod','date','dateStart','dateEnd'];
          const editors={};
          for(const host of node.querySelectorAll('.dx-selectbox,.dx-datebox,.dx-textbox')) {
            if(!host.isConnected)continue;
            const input=host.querySelector('input[id]'); if(!input)continue;
            const field=input.id.split('_').pop();
            const apiName=host.classList.contains('dx-selectbox')?'dxSelectBox':
              host.classList.contains('dx-datebox')?'dxDateBox':'dxTextBox';
            const widget=window.DevExpress?.ui?.[apiName]?.getInstance(host);
            if(widget) {
              const source=widget.getDataSource?.();
              const ready=!host.classList.contains('dx-selectbox') ||
                !!(widget.option('items')?.length ||
                   (source?.isLoaded?.() && source.items?.().length));
              editors[field]={display:input.value,value:widget.option('value')??null,
                dataSourceReady:ready};
            }
          }
          const clean=v=>v instanceof Date?v.toISOString():v??null;
          const pick=x=>Object.fromEntries(fields.map(k=>[k,clean(x[k])]));
          return {ready:node.isConnected,formData:pick(angularData),dxFormData:pick(data),editors};
        }""")

    def set_complete_state(self, updates):
        """Fill through the same stock controls as a successful manual run.

        The live Doctor Eleks form saves after individual editor actions. A
        manual run showed that its intermediate null period saves are accepted
        when the date fields retain the application's ISO string format. Direct
        dxForm/Angular mutation instead sent localized JavaScript Date strings
        and produced HTTP 500. Each action below waits for its own response.
        """
        self.guard()
        before=self.read_state()
        if before.get("formData",{}).get("selectedReferralCode")!="digitalReferral":
            raise ManualCheckpoint("procedure_preflight","Електронне направлення ще не вибрано.",before)
        self._active_record_id=before.get("formData",{}).get("procedureRef")
        if str(updates.get("isDateOrPeriod")).casefold()!="true":
            raise ValueError("procedure period mode is required")
        start=datetime.fromisoformat(str(updates["dateStart"]))
        end=datetime.fromisoformat(str(updates["dateEnd"]))
        if end-start!=timedelta(minutes=30):
            raise ValueError("procedure end must be start + 30 minutes")
        if str(updates.get("date"))[:16]!=start.isoformat()[:16]:
            raise ValueError("performed date must match the procedure start")
        self._log("FIELD_BATCH_BEGIN",{"expected":_json_safe(updates),"before":before})
        self.choose_division(str(updates["division"]))
        self.choose_calendar_datetime("dateStart",start)
        # At this point dateEnd is still the form's default value. Doctor
        # Eleks may derive a large temporary duration from dateStart to that
        # default end, so only the server payload is meaningful for this
        # intermediate selection. Validate all three model layers after the
        # expected end time has been set below.
        self.choose_period_30(verify_state=False)
        # Selecting a period can rebuild the date editors. A successful HTTP
        # save may arrive while Angular formData is temporarily empty. Wait
        # for the already selected start date to return in all three layers
        # before opening dateEnd's calendar.
        state=self._wait_datetime_state("dateStart",start,timeout_ms=15000)
        if not self._state_datetime_matches(state,"dateEnd",end):
            self.choose_calendar_datetime("dateEnd",end)
        # The dateEnd item's XML hardcodes a null period. The period item's
        # own click XML includes {{procedurePeriod}}. Re-selecting it captures
        # a complete server payload after dateEnd has settled.
        payload=(self.saves[-1].get("payload",{}) if self.saves else {})
        period_state=self.read_state()
        if (payload.get("procedurePeriod")!="30" or
                not self._payload_datetime_matches(payload.get("procedurePeriodEnd"),end) or
                not self._period_state_matches(period_state)):
            self.choose_period_30(verify_state=True)
        after=self.read_state()
        self.last_action="FORM_STATE_AFTER_BATCH"
        self._log("FORM_STATE_AFTER_BATCH",{"expected":_json_safe(updates),"form_state":after})
        return after

    @staticmethod
    def _payload_datetime_matches(value,expected):
        if not _present(value): return False
        try:
            parsed=datetime.fromisoformat(str(value).replace("Z","+00:00"))
            return parsed.tzinfo is None and parsed.replace(second=0,microsecond=0)==expected
        except ValueError:
            return False

    @staticmethod
    def _state_datetime_matches(state,field,expected):
        values=(state.get("formData",{}).get(field),
                state.get("dxFormData",{}).get(field),
                state.get("editors",{}).get(field,{}).get("value"))
        return all(DoctorEleksProcedureForm._payload_datetime_matches(value,expected)
                   for value in values)

    def _wait_datetime_state(self,field,expected,timeout_ms=10000):
        """Wait for the editor, dxForm and Angular to reflect a saved date."""
        deadline=time.monotonic()+timeout_ms/1000
        state={}
        while time.monotonic()<deadline:
            raise_if_cancelled()
            self.guard()
            state=self.read_state()
            if self._state_datetime_matches(state,field,expected):
                return state
            self.page.wait_for_timeout(100)
        raise ManualCheckpoint("procedure_calendar",
          "Дата/час не збігається між календарем, dxForm і Angular після очікування оновлення форми.",
          {"field":field,"expected":expected.isoformat(),"form_state":state})

    def choose_division(self,code):
        """Select a loaded division item and verify its underlying UUID."""
        self.guard()
        options=self.selectbox_options("division")
        matches=[item for item in options if str(item.get("code")).casefold()==code.casefold()]
        if len(matches)!=1:
            raise ManualCheckpoint("procedure_division_options",
              "Потрібний підрозділ відсутній у завантаженому списку.",
              {"code":code,"options":options})
        state=self.read_state()
        if (str(state.get("formData",{}).get("division")).casefold()==code.casefold() and
                str(state.get("dxFormData",{}).get("division")).casefold()==code.casefold() and
                str(state.get("editors",{}).get("division",{}).get("value")).casefold()==code.casefold()):
            return state
        self.begin_mutation("SELECT_DIVISION")
        division=wait_visible_recheck(self.page,
            lambda:self.page.locator('input[id$="_division"]').first)
        division.click(timeout=5000)
        self.guard()
        expected_name=str(matches[0].get("name") or "").strip()
        selected=self._find_visible_option(expected_name)
        if selected is None:
            raise ManualCheckpoint("procedure_division_options",
              "Не знайдено пункт підрозділу у відкритому списку.",
              {"code":code,"name":expected_name})
        sequence=self.latest_sequence
        selected.click(timeout=5000)
        self.wait_autosave(sequence,timeout_ms=20000)
        state=self._wait_field_state("division",code,timeout_ms=5000)
        values=(state.get("formData",{}).get("division"),
                state.get("dxFormData",{}).get("division"),
                state.get("editors",{}).get("division",{}).get("value"))
        if not all(str(value).casefold()==code.casefold() for value in values):
            raise ManualCheckpoint("procedure_division",
              "Підрозділ не збігається між editor, dxForm і Angular.",
              {"expected":code,"form_state":state})
        self._log("FORM_STATE_AFTER_FIELD",{"field":"division","form_state":state,
                                             "save_payload":self.saves[-1].get("payload")})
        return state

    def choose_calendar_datetime(self,field,expected):
        """Use the readonly dxDateBox calendar, clock controls and OK button."""
        if field not in {"dateStart","dateEnd"}:
            raise ValueError("unsupported procedure calendar field")
        self.guard()
        state=self.read_state()
        if self._state_datetime_matches(state,field,expected):return state
        self.begin_mutation("CALENDAR_"+field.upper())
        input_box=wait_visible_recheck(self.page,
            lambda:self.page.locator(f'input[id$="_{field}"]').first,
            initial_ms=1800,retry_ms=3000)
        if input_box.get_attribute("readonly") is None:
            raise ManualCheckpoint("procedure_calendar",
              "Редактор дати змінив тип; календар не підтверджено.",{"field":field})
        # Doctor Eleks may re-render the button while the previous editor's
        # autosave settles. The DateBox API opens the same calendar without a
        # coordinate click on a moving DOM element.
        opened=False;last_reason=""
        for attempt in range(2):
            self.guard()
            try:
                self.page.wait_for_function("""([field, expectedId]) => {
                  const input=document.querySelector('input[id$="_'+field+'"]');
                  const host=input?.closest('.dx-datebox');
                  const widget=host&&window.DevExpress?.ui?.dxDateBox?.getInstance(host);
                  const scope=input&&window.angular?.element(input).scope();
                  const record=scope?.formVm?.formData?.procedureRef;
                  return !!(input?.isConnected && widget && typeof widget.open==='function' &&
                    (!expectedId || String(record).toLowerCase()===String(expectedId).toLowerCase()));
                }""",arg=[field,self._active_record_id],timeout=7000)
                picker=self.page.evaluate("""field => {
                  const input=document.querySelector('input[id$="_'+field+'"]');
                  const host=input?.closest('.dx-datebox');
                  const widget=host&&window.DevExpress?.ui?.dxDateBox?.getInstance(host);
                  if(!input?.isConnected || !widget || typeof widget.open!=='function')
                    return {opened:false,reason:'dxDateBox was replaced'};
                  if(widget.option('pickerType')==='native')
                    return {opened:false,reason:'native picker'};
                  widget.open();
                  return {opened:!!widget.option('opened'),reason:''};
                }""",field)
                opened=bool(picker.get("opened"))
                last_reason=str(picker.get("reason") or "")
                if opened:
                    self.page.wait_for_function("""field => {
                  const input=document.querySelector('input[id$="_'+field+'"]');
                  const host=input?.closest('.dx-datebox');
                  const widget=host&&window.DevExpress?.ui?.dxDateBox?.getInstance(host);
                  return !!(input?.isConnected && widget?.option('opened') &&
                    [...document.querySelectorAll('.dx-datebox-wrapper')].some(node =>
                      node.getClientRects().length && getComputedStyle(node).visibility!=='hidden'));
                    }""",arg=field,timeout=7000)
                    break
            except Exception as exc:
                opened=False
                last_reason=str(exc)[:300]
            self.page.wait_for_timeout(300)
        if not opened:
            raise ManualCheckpoint("procedure_calendar",
              "Календар дати не відкрився після оновлення форми Doctor Eleks.",
              {"field":field,"reason":last_reason,"form_state":self._try_read_state()})
        self.guard()
        self._log("CALENDAR_OPENED",{"field":field,"method":"dxDateBox.open"})
        overlay=wait_visible_recheck(self.page,
            lambda:self.page.locator('.dx-datebox-wrapper:visible').first,
            initial_ms=1800,retry_ms=3000)
        for _ in range(36):
            year,month=overlay.evaluate("""el => {
              const node=el.querySelector('.dx-calendar');
              const widget=node&&window.DevExpress?.ui?.dxCalendar?.getInstance(node);
              const date=widget?.option('currentDate');
              return date?[date.getFullYear(),date.getMonth()+1]:[null,null];
            }""")
            if year is None:
                raise ManualCheckpoint("procedure_calendar","Календар DevExtreme ще не ініціалізовано.")
            delta=(expected.year-year)*12+expected.month-month
            if delta==0:break
            selector=('.dx-calendar-navigator-next-month' if delta>0 else
                      '.dx-calendar-navigator-previous-month')
            month_button=wait_visible_recheck(self.page,
                lambda:self.page.locator('.dx-datebox-wrapper:visible').first
                  .locator(selector).first)
            month_button.click(timeout=3000)
            self.guard()
        else:
            raise ManualCheckpoint("procedure_calendar","Дата графіка поза доступним діапазоном календаря.")
        day=expected.strftime("%Y/%m/%d")
        day_cell=wait_visible_recheck(self.page,
            lambda:self.page.locator('.dx-datebox-wrapper:visible').first
              .locator(f'.dx-calendar-cell:not(.dx-calendar-other-month)[data-value="{day}"]').first)
        day_cell.click(timeout=5000)
        hours=wait_visible_recheck(self.page,
            lambda:self.page.locator('.dx-datebox-wrapper:visible').first
              .locator('input[role="spinbutton"][aria-label="hours"]'))
        # Do not press Tab inside the popup. On this portal, leaving the
        # minutes spinbutton can dismiss the datebox before its OK button is
        # clicked (the observed failure was the page receiving the click).
        # Filling the next control already blurs the hours field; clicking OK
        # below blurs and commits the minutes field.
        hours.fill(str(expected.hour))
        minutes=wait_visible_recheck(self.page,
            lambda:self.page.locator('.dx-datebox-wrapper:visible').first
              .locator('input[role="spinbutton"][aria-label="minutes"]'))
        minutes.fill(str(expected.minute))
        sequence=self.latest_sequence
        done=wait_visible_recheck(self.page,
            lambda:self.page.locator('.dx-datebox-wrapper:visible').first
              .locator('.dx-popup-done[role="button"]').first)
        try:
            done.click(timeout=5000)
        except Exception:
            # Some form refreshes close the picker as the value is committed.
            # Accept that only when all three independent date sources agree.
            committed=self._try_read_state()
            if not self._state_datetime_matches(committed,field,expected):
                raise
            self._log("CALENDAR_CLOSED_AFTER_COMMIT",{"field":field,
                                                         "form_state":committed})
        self.wait_optional_autosave(sequence,timeout_ms=5000)
        state=self._wait_datetime_state(field,expected)
        self._log("FORM_STATE_AFTER_FIELD",{"field":field,"form_state":state,
                                             "save_payload":self.saves[-1].get("payload") if self.saves else None})
        return state

    def selectbox_options(self, field):
        """Read loaded options from a newly resolved editor instance."""
        self.guard()
        self.page.wait_for_function("""field => {
          const anchor=document.querySelector('input[id$="_'+field+'"]');
          const host=anchor?.closest('.dx-selectbox');
          const widget=host&&window.DevExpress?.ui?.dxSelectBox?.getInstance(host);
          const source=widget?.getDataSource?.();
          return !!(anchor?.isConnected && widget &&
            (widget.option('items')?.length ||
             (source?.isLoaded?.() && source.items?.().length)));
        }""",arg=field,timeout=10000)
        self.guard()
        return self.page.evaluate("""field => {
          const anchor=document.querySelector('input[id$="_'+field+'"]');
          const host=anchor?.closest('.dx-selectbox');
          const widget=host&&window.DevExpress?.ui?.dxSelectBox?.getInstance(host);
          const source=widget?.getDataSource?.();
          if(!anchor?.isConnected || !widget)return null;
          const items=widget.option('items')?.length ? widget.option('items') :
            ((source?.isLoaded?.() && source.items?.())||[]);
          const flat=[];
          const walk=values=>(values||[]).forEach(item=>{
            if(Array.isArray(item?.items))walk(item.items);
            else flat.push(item);
          });
          walk(items);
          return flat.map(item=>({
            code:item.code??item.id??item.value??null,
            name:item.name??item.text??item.title??String(item)
          }));
        }""",field)

    def _find_visible_option(self,label):
        """Resolve current dropdown rows twice before treating a value as absent."""
        for attempt in range(2):
            self.guard()
            candidates=self.page.locator('.dx-list-item-content:visible')
            for index in range(candidates.count()):
                item=candidates.nth(index)
                try:
                    if item.inner_text(timeout=500).strip()==label:
                        return item
                except Exception:continue
            if attempt==0:self.page.wait_for_timeout(2000)
        return None

    @staticmethod
    def _field_state_matches(state,field,expected):
        values=(state.get("formData",{}).get(field),
                state.get("dxFormData",{}).get(field),
                state.get("editors",{}).get(field,{}).get("value"))
        if field=="division":
            return all(value is not None and str(value).casefold()==str(expected).casefold()
                       for value in values)
        return all(value is not None and str(value)==str(expected) for value in values)

    @staticmethod
    def _period_state_matches(state):
        return DoctorEleksProcedureForm._field_state_matches(state,"procedurePeriod","30")

    def _wait_field_state(self,field,expected,timeout_ms=5000):
        """Wait for Angular, dxForm and the freshly resolved editor to agree."""
        deadline=time.monotonic()+timeout_ms/1000
        state={}
        while time.monotonic()<deadline:
            raise_if_cancelled();self.guard()
            state=self.read_state()
            if self._field_state_matches(state,field,expected):
                return state
            self.page.wait_for_timeout(100)
        return state

    def _wait_period_state(self,timeout_ms=5000):
        """Wait for the Angular model, dxForm and current editor to converge."""
        return self._wait_field_state("procedurePeriod","30",timeout_ms)

    def choose_period_30(self,verify_state=True):
        """The period item's own Doctor Eleks action is the valid save path."""
        self.guard()
        action=self.page.evaluate("""() => {
          const anchor=document.querySelector('input[id$="_isDateOrPeriod"]');
          const scope=anchor&&window.angular?.element(anchor).scope();
          const root=scope?.formInstance?.option('items')||[];
          let found=null;
          const walk=items=>(items||[]).forEach(item=>{
            if(item.dataField==='procedurePeriod')found=item;
            walk(item.items);
            (item.tabs||[]).forEach(tab=>walk(tab.items));
          });
          walk(root);
          const xml=found?.editorOptions?.onItemClickActionXml||'';
          const match=xml.match(/<param\\s+name="procedurePeriod">([^<]*)<\\/param>/);
          return {found:!!found,source:match?.[1]??null};
        }""")
        if action.get("source")!="{{procedurePeriod}}":
            raise ManualCheckpoint("procedure_period_action",
              "Поточна форма не підтверджує штатне збереження коду тривалості.",action)
        options=self.selectbox_options("procedurePeriod")
        if not options or not any(str(item.get("code"))=="30" for item in options):
            raise ManualCheckpoint("procedure_period_options",
              "Список тривалості не завантажив код 30.",{"options":options})
        self.begin_mutation("SELECT_PROCEDURE_PERIOD_30")
        period=wait_visible_recheck(self.page,
            lambda:self.page.locator('input[id$="_procedurePeriod"]').first)
        period.click(timeout=5000)
        self.guard()
        selected=self._find_visible_option("30 хвилин")
        if selected is None:
            raise ManualCheckpoint("procedure_period_options","Не знайдено видимий пункт «30 хвилин».")
        sequence=self.latest_sequence
        selected.click(timeout=5000)
        save=self.wait_autosave(sequence,timeout_ms=20000)
        if save.get("payload",{}).get("procedurePeriod")!="30":
            raise ManualCheckpoint("procedure_period_payload",
              "Doctor Eleks не відправив procedurePeriod=30 після вибору пункту.",
              {"save":save,"form_state":self._try_read_state()})
        state=self._wait_period_state() if verify_state else self.read_state()
        if verify_state and not self._period_state_matches(state):
            raise ManualCheckpoint("procedure_period_state",
              "Після збереження код 30 не збігається між editor, dxForm і Angular.",
              {"save":save,"form_state":state})
        self._log("FORM_STATE_AFTER_FIELD",{"field":"procedurePeriod",
                                             "form_state":state,"save_payload":save.get("payload")})
        return save

    def wait_autosave(self, after_sequence=0, timeout_ms=15000):
        deadline=time.monotonic()+timeout_ms/1000
        while time.monotonic()<deadline:
            raise_if_cancelled();self.guard()
            pending=[x for x in self.saves if x["sequence"]>after_sequence]
            if (pending and all(x.get("status") is not None for x in pending)
                    and time.monotonic()-pending[-1].get("requested_at",0)>=0.5):
                self.guard()
                return pending[-1]
            self.page.wait_for_timeout(100)
        pending=[x for x in self.saves if x["sequence"]>after_sequence]
        if not pending:
            raise ManualCheckpoint("procedure_autosave_missing","Зміна не породила spiprocedureSave.",
                                   {"last_action":self.last_action,"form_state":self._try_read_state()})
        raise ManualCheckpoint("procedure_autosave_timeout","spiprocedureSave не завершився вчасно.",
                               {"payload":pending[-1].get("payload"),"last_action":self.last_action})

    @property
    def latest_sequence(self):
        return max((entry["sequence"] for entry in self.saves),default=0)

    def wait_optional_autosave(self, after_sequence, timeout_ms=5000):
        """Wait for a save caused by a UI action, if that action starts one."""
        deadline=time.monotonic()+timeout_ms/1000
        while time.monotonic()<deadline:
            raise_if_cancelled();self.guard()
            if any(entry["sequence"]>after_sequence for entry in self.saves):
                return self.wait_autosave(after_sequence,
                                          timeout_ms=max(100,int((deadline-time.monotonic())*1000)))
            self.page.wait_for_timeout(100)
        self.guard()
        return None

    def wait_final_save(self, after_sequence, timeout_ms=25000, expected_id=None):
        """Require the validate=1 save caused by «Підписати пізніше»."""
        deadline=time.monotonic()+timeout_ms/1000
        while time.monotonic()<deadline:
            raise_if_cancelled();self.guard()
            entries=[item for item in self.saves if item["sequence"]>after_sequence]
            validated=[item for item in entries
                       if str(item.get("payload",{}).get("toValidate"))=="1" and
                       (expected_id is None or
                        str(item.get("payload",{}).get("id","")).casefold()==str(expected_id).casefold())]
            if (validated and all(item.get("status") is not None for item in entries)
                    and time.monotonic()-entries[-1].get("requested_at",0)>=0.5):
                self.guard()
                result=validated[-1]
                if 200<=(result.get("status") or 0)<300:
                    return result
            self.page.wait_for_timeout(100)
        raise ManualCheckpoint("procedure_final_save",
          "Фінальне spiprocedureSave з toValidate=1 не підтверджено сервером.",
          {"saves":[{k:v for k,v in item.items() if k!="request"}
                    for item in self.saves if item["sequence"]>after_sequence]})

    def capture_manual_success(self, timeout_ms=900000):
        """Keep listeners live while the operator completes one procedure by hand."""
        self.last_action="MANUAL_CAPTURE_WAIT"
        self._log("MANUAL_CAPTURE_READY",{"form_state":self.read_state(),"timeout_ms":timeout_ms})
        deadline=time.monotonic()+timeout_ms/1000
        while time.monotonic()<deadline:
            raise_if_cancelled(); self.guard()
            success=next((entry for entry in reversed(self.saves)
                          if entry.get("status") is not None and 200<=entry["status"]<300),None)
            if success:
                record={"payload":success["payload"],"status":success["status"],
                        "response":success["response"],"timestamp":datetime.now().astimezone().isoformat()}
                path=RUNTIME_DIR/"procedure-manual-success.json"
                path.write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding="utf-8")
                failed=next((json.loads(line) for line in reversed(self._out.read_text(encoding="utf-8").splitlines())
                             if line.strip() and '"event": "AUTOSAVE_FAILURE"' in line),None) if self._out.exists() else None
                diff=structured_payload_diff(failed or {},record)
                (RUNTIME_DIR/"procedure-manual-diff.json").write_text(
                    json.dumps(diff,ensure_ascii=False,indent=2),encoding="utf-8")
                self._log("MANUAL_CAPTURE_SUCCESS",{"status":record["status"],"payload":record["payload"],
                                                     "diff":diff})
                return record
            self.page.wait_for_timeout(150)
        raise ManualCheckpoint("procedure_manual_capture_timeout",
                               "Не отримано успішний ручний spiprocedureSave за час очікування.")

    def verify_payload(self, sequence, expected):
        record=next((item for item in self.saves if item["sequence"]==sequence),None)
        if not record:
            raise ManualCheckpoint("procedure_payload_missing","Немає captured spiprocedureSave payload.")
        for key,value in expected.items():
            if record.get("payload",{}).get(key)!=_json_safe(value):
                raise ManualCheckpoint("procedure_payload_mismatch",
                  f"Payload поля {key} не відповідає очікуваному значенню.",
                  {"expected":_json_safe(expected),"payload":record.get("payload")})
        return record
