from __future__ import annotations
import re
from datetime import datetime,date,timedelta
from collections import Counter
from ..core.result import ActionResult,waiting
from ..core.text import normalize_patient_name
from ..core.action_control import raise_if_cancelled,ActionCancelled
from ..core.audit import log_event
from ..core.database import get_workflow,set_workflow
from ..automation.browser import WebEleks,ManualCheckpoint,AutomationError,wait_visible_recheck
from ..excel.workbook import parse_diagnoses_with_primary,write_cell,update_coordinator_status
from .common import settings,patient_data,patient_key,require,graph_path,graph_sheet,coordinator_path,fmt_datetime
from .rehab_first import (age_on as _rehab_age_on, format_interaction_datetime,
                          parse_interaction_datetime,
                          newest_referral_index, normalize_icd, parse_icd_codes,
                          diagnosis_time_matches,validate_age_specific_reports,
                          plan_diagnosis_reconciliation as _plan_diagnosis_reconciliation,
                          barthel_rows_to_remove)
from .procedures import _datetime_text_matches

CODE_TOKEN=re.compile(r"\b[A-ZА-ЯІЇЄҐ]\d{1,3}(?:[.\-]\d{1,3})*\b|\b\d{2,3}(?:[.\s]\d{2,3})+\b",re.I)
_FI_DATETIME_ALIASES={"episode_start","start_datetime","diagnosis_installed_at","report_datetime",
                      "interaction_start"}
_FI_CALENDAR_ONLY={"episode_start","interaction_start"}

def _is_self_referral(v):
    s=str(v or "").casefold().replace(" ","")
    return "самозвер" in s or "самообращ" in s

def _referral_digits(v):
    return "".join(re.findall(r"\d",str(v or "")))

def _green_sync(web):
    body=web.current_text().casefold()
    if any(x in body for x in ("синхронізовано","синхронизировано","успішно синхрон","успешно синхрон")):
        return True
    try:
        loc=web.page.locator(".text-success,.status-success,[class*='success'],[class*='synch']")
        return loc.count()>0
    except:
        return False

def _resume_after_signature(action,pkey,name,cfg,require_sync=False):
    st=get_workflow(action,pkey)
    if not st or st["state"] not in {"awaiting_kep","safe_mode_final"}:
        return None
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name)
            if require_sync and not _green_sync(web):
                return waiting("Підпис уже очікується. Завершіть КЕП і дочекайтеся зеленої синхронізації, потім повторіть функцію.","kep")
            set_workflow(action,pkey,"complete",{"verified_after_signature":True})
            return ActionResult(True,"Ручний підпис/синхронізацію прийнято як завершений етап.",{"resumed":True})
    except ManualCheckpoint as cp:
        return waiting(str(cp),cp.code,cp.data)

def _click_save_sign_wait(web,action,pkey):
    try:
        web.click("save_sign")
        set_workflow(action,pkey,"awaiting_kep",{})
        return waiting("Форма відправлена на підпис. Накладіть КЕП і після завершення повторіть функцію.","kep")
    except ManualCheckpoint as cp:
        if cp.code=="safe_mode":
            set_workflow(action,pkey,"safe_mode_final",{})
            return waiting("Безпечний режим не натиснув фінальну кнопку. Перевірте форму, натисніть «Зберегти та підписати» вручну, накладіть КЕП і повторіть функцію.","safe_mode_final")
        return waiting(str(cp),cp.code,cp.data)

def _append_episode_suffix(web,suffix):
    try:
        loc=web.locator("episode_name",timeout=500)
        current=""
        try: current=loc.input_value(timeout=500)
        except:
            try: current=loc.inner_text(timeout=500)
            except: pass
        current=(current or "").strip()
        if current and not current.casefold().endswith(suffix.casefold()):
            loc.fill(current+" "+suffix)
            log_event("episode.name",f"До назви епізоду додано {suffix}",action_id=web.action_id)
    except Exception:
        log_event("episode.name.warning","Поле назви епізоду не знайдено; суфікс треба перевірити вручну.",level="WARNING",action_id=web.action_id)

def _row_code(text):
    hits=CODE_TOKEN.findall(text or "")
    return hits[0].replace(" ","") if hits else ""

def _set_diag_metadata(web,is_primary):
    if web.click("diagnosis_role",optional=True):
        web.choose_text("Основний" if is_primary else "Супутній",optional=True)
    if web.click("diagnosis_certainty",optional=True):web.choose_text("Заключний",optional=True)
    if web.click("diagnosis_clinical",optional=True):web.choose_text("Активний",optional=True)

def _reconcile_diagnoses(web,desired,primary):
    web.click("diagnoses",optional=True)
    desired=[x.upper().replace(" ","") for x in desired]
    primary=primary.upper().replace(" ","")
    rows=web.page.locator("tr")
    current={}
    row_objs={}
    for i in range(rows.count()):
        try:
            txt=rows.nth(i).inner_text(timeout=250)
            code=_row_code(txt).upper().replace(" ","")
            if code:
                current.setdefault(code,[]).append(txt)
                row_objs.setdefault(code,rows.nth(i))
        except:pass
    missing=[c for c in desired if c not in current]
    extras=[c for c in current if c not in desired]
    # Replace template extras with missing desired codes before adding new rows.
    while extras and missing:
        old=extras.pop(0);new=missing.pop(0);row=row_objs.get(old)
        try:
            edit=row.get_by_text("Редагувати",exact=False)
            if not edit.count():edit=row.get_by_text("Редактировать",exact=False)
            if not edit.count():raise RuntimeError("edit missing")
            edit.first.click()
            web.click("change",optional=True);web.fill("code_search",new,optional=True);web.click("choose",optional=True)
            _set_diag_metadata(web,new==primary);web.click("save",optional=True)
            log_event("diagnosis.replace",f"{old} → {new}",action_id=web.action_id)
        except Exception:
            missing.insert(0,new)
            break
    for code in missing:
        web.click("add_diagnosis",optional=True)
        if not web.fill("code_search",code,optional=True):
            raise ManualCheckpoint("diagnosis_add",f"Не вдалося автоматично додати діагноз {code}. Додайте його вручну.")
        web.click("choose",optional=True);_set_diag_metadata(web,code==primary);web.click("save",optional=True)
    # Configure roles on rows we can identify.
    rows=web.page.locator("tr")
    for code in desired:
        for i in range(rows.count()):
            try:
                row=rows.nth(i);txt=row.inner_text(timeout=200)
                if code.casefold() not in txt.casefold():continue
                edit=row.get_by_text("Редагувати",exact=False)
                if not edit.count():edit=row.get_by_text("Редактировать",exact=False)
                if edit.count():
                    edit.first.click();_set_diag_metadata(web,code==primary);web.click("save",optional=True)
                break
            except:pass
    # Remaining extras must not be silently kept.
    if extras:
        raise ManualCheckpoint("diagnosis_extra",
            "У шаблоні залишилися діагнози, яких немає в колонці N: "+", ".join(extras)+". Видаліть/замініть їх вручну і повторіть.")

def _edit_report_96022(web,dt,doctor,cfg):
    web.click("diagnostic_reports",optional=True)
    row=None
    rows=web.page.locator("tr")
    for i in range(rows.count()):
        try:
            txt=rows.nth(i).inner_text(timeout=250).replace(" ","")
            if "9602200" in txt or "96.0.22.00" in txt:
                row=rows.nth(i);break
        except:pass
    if row is None:
        if not web.page_has("96 0 22") and not web.page_has("96.0.22"):
            raise ManualCheckpoint("report_96022","Не знайдено діагностичний звіт 96 0 22 00. Додайте/перевірте його вручну.")
        web.click("diagnostic_report_edit",optional=True)
    else:
        try:
            edit=row.get_by_text("Редагувати",exact=False)
            if not edit.count():edit=row.get_by_text("Редактировать",exact=False)
            if edit.count():edit.first.click()
        except:pass
    web.fill("report_datetime",dt,optional=True)
    web.choose_text(cfg.get("default_place","ТОВ 2 Міжнародний центр реабілітації та нейрофізіології"),optional=True)
    if doctor:
        for alias in ("interpreting_worker","diagnostic_executor","report_submitter"):
            if web.click(alias,optional=True):web.choose_text(doctor,optional=True)
    web.fill("conclusion","Оцінка підтримання здоров'я або одужання: задовільне, пацієнт здатний отримати реабілітаційні послуги.",optional=True)
    web.click("save",optional=True)

def _validate_and_prepare_icf(web,dt):
    web.click("examinations",optional=True)
    required=("D599","B799","D4509","E325")
    for code in required:
        if web.page_has(code):continue
        web.click("add_examination",optional=True)
        if web.fill("code_search",code,optional=True):
            web.click("choose",optional=True);web.fill("measurement_datetime",dt,optional=True);web.click("save",optional=True)
        else:
            raise ManualCheckpoint("icf_code",f"У МКФ немає {code}, і автоматичне додавання недоступне. Додайте код вручну.")
    body=web.current_text().casefold()
    # Body Structures code was not specified by the user: validate category instead of inventing a code.
    if "структур" not in body:
        raise ManualCheckpoint("icf_structure","Не підтверджено жодного обстеження категорії «Структури організму». Додайте мінімум одне вручну.")
    # Update dates where edit rows are available.
    rows=web.page.locator("tr")
    for i in range(rows.count()):
        try:
            txt=rows.nth(i).inner_text(timeout=200).replace(" ","").upper()
            if any(code in txt for code in required):
                edit=rows.nth(i).get_by_text("Редагувати",exact=False)
                if not edit.count():edit=rows.nth(i).get_by_text("Редактировать",exact=False)
                if edit.count():
                    edit.first.click();web.fill("measurement_datetime",dt,optional=True);web.click("save",optional=True)
        except:pass

def _age_on(dob,at_value):
    try:
        if isinstance(dob,datetime):d=dob.date()
        elif isinstance(dob,date):d=dob
        else:d=datetime.fromisoformat(str(dob)).date()
        if isinstance(at_value,datetime):at=at_value.date()
        elif isinstance(at_value,date):at=at_value
        else:
            s=str(at_value).split(" ")[0]
            try:at=datetime.strptime(s,"%d.%m.%Y").date()
            except:at=datetime.fromisoformat(s).date()
        return at.year-d.year-((at.month,at.day)<(d.month,d.day))
    except:return None

def diagnostic_interaction(payload):
    action="diagnostic_interaction";name=require(payload,"patient_name","ПІБ пацієнта");cfg=settings();pkey=patient_key(name)
    resumed=_resume_after_signature(action,pkey,name,cfg,require_sync=True)
    if resumed:return resumed
    pdata=patient_data(name,payload);dt=fmt_datetime(pdata.get("first_interaction"))
    if not dt:return ActionResult(False,"У колонці M не задано дату/час першої взаємодії.",{})
    dx=parse_diagnoses_with_primary(graph_path(payload),graph_sheet(),name)
    if not dx["codes"]:return ActionResult(False,"У колонці N немає кодів діагнозів.",{})
    if dx["ambiguous_primary"]:return ActionResult(False,"Не вдалося однозначно визначити жирний основний діагноз у N.",dx)
    primary=(dx["primary"] or dx["codes"][:1])[0];referral=pdata.get("referral")
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name,pdata.get("dob"))
            if not _is_self_referral(referral):
                digits=_referral_digits(referral)
                if len(digits)!=16:return ActionResult(False,"Е-направлення повинно містити 16 цифр.",{"value":str(referral)})
                web.click("referrals",optional=True);web.click("referral_search_number",optional=True);web.fill("referral_number",digits);web.click("search_button")
                if not web.choose_text("фізичної та реабілітаційної медицини",optional=True):
                    raise ManualCheckpoint("referral_validation","Підтвердіть, що направлення саме на консультацію лікаря ФРМ, і візьміть його в роботу.")
                web.click("take_to_work",optional=True)
            web.click("interactions",optional=True);web.click("new_interaction");web.click("create_new_episode")
            web.fill("episode_start",dt,optional=True);web.fill("start_datetime",dt,optional=True);web.fill("duration","30",optional=True)
            web.click("episode_type",optional=True);web.choose_text("Діагностика",optional=True)
            web.click("priority",optional=True);web.choose_text("Ургентний",optional=True)
            web.choose_text(cfg.get("default_place","ТОВ 2 Міжнародний центр реабілітації та нейрофізіології"),optional=True)
            if _is_self_referral(referral):web.click("direction_self",optional=True)
            else:
                web.click("electronic_referral",optional=True);web.click("choose_referral",optional=True);web.choose_text(digits,optional=True)
            web.click("add_action",optional=True);web.fill("code_search","A67-02-22",optional=True);web.choose_text("A67-02-22",optional=True);web.click("add_selected",optional=True)
            web.click("add_diagnosis",optional=True);web.fill("code_search",primary,optional=True);web.click("choose",optional=True)
            _set_diag_metadata(web,True);web.click("save",optional=True)
            _append_episode_suffix(web,"D")
            if not _is_self_referral(referral):web.click("redeem_referral",optional=True)
            return _click_save_sign_wait(web,action,pkey)
    except ManualCheckpoint as cp:return waiting(str(cp),cp.code,cp.data)

def _create_rehab_referral(payload,action):
    name=require(payload,"patient_name","ПІБ пацієнта");cfg=settings();pkey=patient_key(name)
    resumed=_resume_after_signature(action,pkey,name,cfg,require_sync=False)
    if resumed:return resumed
    pdata=patient_data(name,payload)
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name,pdata.get("dob"));web.click("interactions",optional=True);web.click("actions");web.click("create_referral")
            web.click("referral_category",optional=True);web.click("nonstationary_rehab",optional=True)
            web.click("choose_service",optional=True);web.click("medical_rehab",optional=True)
            web.click("take_to_work",optional=True)
            return _click_save_sign_wait(web,action,pkey)
    except ManualCheckpoint as cp:return waiting(str(cp),cp.code,cp.data)

def rehab_referral_from_diagnostic(payload):return _create_rehab_referral(payload,"rehab_referral_from_diagnostic")
def rehab_referral_after_interaction(payload):return _create_rehab_referral(payload,"rehab_referral_after_interaction")

def rehab_interaction(payload):
    action="rehab_interaction";name=require(payload,"patient_name","ПІБ пацієнта");cfg=settings();pkey=patient_key(name)
    resumed=_resume_after_signature(action,pkey,name,cfg,require_sync=True)
    if resumed:
        if resumed.status=="done":
            pdata=patient_data(name,payload);write_cell(graph_path(payload),graph_sheet(),pdata["row"],15,"E PV Reab")
            resumed.message="Реабілітаційна взаємодія синхронізована; O оновлено на E PV Reab."
        return resumed
    pdata=patient_data(name,payload);dt=fmt_datetime(pdata.get("first_interaction"))
    if not dt:return ActionResult(False,"У колонці M немає дати/часу першої взаємодії.",{})
    dx=parse_diagnoses_with_primary(graph_path(payload),graph_sheet(),name)
    if not dx["codes"]:return ActionResult(False,"У колонці N немає діагнозів.",{})
    if dx["ambiguous_primary"]:return ActionResult(False,"Не визначено один жирний основний діагноз у N.",dx)
    primary=(dx["primary"] or dx["codes"][:1])[0]
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name,pdata.get("dob"));web.click("interactions",optional=True);web.click("new_interaction");web.click("create_new_episode")
            web.fill("episode_start",dt,optional=True);web.fill("start_datetime",dt,optional=True);web.fill("duration","60",optional=True)
            web.click("episode_type",optional=True);web.choose_text("Реабілітація",optional=True)
            web.click("priority",optional=True);web.choose_text("Плановий",optional=True)
            web.choose_text(cfg.get("default_place","ТОВ 2 Міжнародний центр реабілітації та нейрофізіології"),optional=True)
            web.click("electronic_referral",optional=True);web.click("choose_referral",optional=True);web.choose_text("Медична реабілітація",optional=True)
            web.click("templates",optional=True);web.click("organization_templates",optional=True);web.choose_text("AR1/2",optional=True)
            _reconcile_diagnoses(web,dx["codes"],primary)
            doctor=payload.get("doctor_profile","")
            _edit_report_96022(web,dt,doctor,cfg)
            _validate_and_prepare_icf(web,dt)
            age=_age_on(pdata.get("dob"),pdata.get("first_interaction"))
            if age is not None and age>=18:
                web.click("diagnostic_reports",optional=True)
                rows=web.page.locator("tr")
                report_rows=0
                for i in range(rows.count()):
                    try:
                        txt=rows.nth(i).inner_text(timeout=200).casefold()
                        if "звіт" in txt or "отчет" in txt or "96" in txt:report_rows+=1
                    except:pass
                if report_rows<2:
                    raise ManualCheckpoint("adult_reports","Для дорослого не підтверджено мінімум два діагностичні звіти. Додайте другий; за можливості Barthel Extended 60–70.")
            _append_episode_suffix(web,"R")
            web.click("redeem_referral",optional=True)
            return _click_save_sign_wait(web,action,pkey)
    except ManualCheckpoint as cp:return waiting(str(cp),cp.code,cp.data)


def _fi_click(web, alias, *, locator=None):
    try:
        loc=locator if locator is not None else web.wait_for_alias(alias, enabled=True)
        return web.click(alias,locator=loc)
    except (ManualCheckpoint,ActionCancelled):raise
    except Exception as exc:
        raise AutomationError(f"Крок click/{alias}: очікувався видимий і доступний елемент; фактично {exc}.") from exc


def _fi_fill(web, alias, value):
    try:
        loc=(_fi_diagnosis_control(web,alias) if alias in _FI_DIAGNOSIS_FIELDS
             else web.wait_for_alias(alias, enabled=True))
        if alias in _FI_DATETIME_ALIASES:
            _fi_set_datetime(web,loc,alias,value)
        else:
            web.fill(alias,value,locator=loc)
            actual=loc.input_value(timeout=1500)
            if str(actual).strip()!=str(value).strip():
                raise AutomationError(f"Поле «{alias}»: очікувалося {value!r}, отримано {actual!r}.")
    except (ManualCheckpoint,ActionCancelled):raise
    except Exception as exc:
        raise AutomationError(f"Крок fill/{alias}: очікувалося {value!r}; фактично {exc}.") from exc


def _fi_datetime_state(loc, *, timeout=None):
    return loc.evaluate("""el => {
      const host=el.closest('.dx-datebox');
      const widget=host && window.DevExpress?.ui?.dxDateBox?.getInstance(host);
      const date=widget?.option('value');
      const pad=n=>String(n).padStart(2,'0');
      const local=date instanceof Date && !isNaN(date)
        ? `${pad(date.getDate())}.${pad(date.getMonth()+1)}.${date.getFullYear()} ${pad(date.getHours())}:${pad(date.getMinutes())}`
        : (typeof date==='string' ? date : null);
      return {display:el.value, widget:local, isDatebox:!!host};
    }""",timeout=timeout)


def _fi_set_datetime(web, loc, alias, value):
    expected=datetime.strptime(str(value),"%d.%m.%Y %H:%M")
    settle=lambda: web.page.wait_for_timeout(2000) if alias in _FI_CALENDAR_ONLY else None
    state=_fi_datetime_state(loc)
    if _datetime_text_matches(state.get("display"),expected) and (
        not state.get("isDatebox") or _datetime_text_matches(state.get("widget"),expected)):
        return
    if loc.get_attribute("readonly") is None and alias not in _FI_CALENDAR_ONLY:
        web.fill(alias,value,locator=loc)
        loc.press("Tab",timeout=3000)
    else:
        if not state.get("isDatebox"):
            raise AutomationError(f"Дата {alias}: поле readonly без перевіреного календаря dxDateBox.")
        dropdown=wait_visible_recheck(web.page,lambda:loc.locator(
            'xpath=ancestor::*[contains(@class,"dx-datebox")][1]'
            ).locator('.dx-dropdowneditor-button').first)
        dropdown.click(timeout=10000)
        settle()
        overlay=wait_visible_recheck(web.page,lambda:web.page.locator('.dx-datebox-wrapper:visible').first)
        for _ in range(60):
            year,month=overlay.evaluate("""el => {
              const node=el.querySelector('.dx-calendar');
              const widget=node&&window.DevExpress?.ui?.dxCalendar?.getInstance(node);
              const date=widget?.option('currentDate');
              return date?[date.getFullYear(),date.getMonth()+1]:[null,null];
            }""")
            if year is None:raise AutomationError(f"Дата {alias}: календар DevExtreme не ініціалізовано.")
            delta=(expected.year-year)*12+expected.month-month
            if delta==0:break
            selector=('.dx-calendar-navigator-next-month' if delta>0 else
                      '.dx-calendar-navigator-previous-month')
            wait_visible_recheck(web.page,lambda:overlay.locator(selector).first).click(timeout=3000)
            settle()
        else:raise AutomationError(f"Дата {alias}: потрібний місяць поза межами календаря.")
        day=expected.strftime("%Y/%m/%d")
        wait_visible_recheck(web.page,lambda:overlay.locator(
            f'.dx-calendar-cell:not(.dx-calendar-other-month)[data-value="{day}"]').first).click(timeout=10000)
        settle()
        hours=wait_visible_recheck(web.page,lambda:overlay.locator('input[role="spinbutton"][aria-label="hours"]').first)
        # Tabbing out of the last time spinbutton can close this DevExtreme
        # popup before its OK button is clicked. Filling the next field blurs
        # hours; clicking OK below blurs and commits minutes.
        hours.fill(str(expected.hour))
        minutes=wait_visible_recheck(web.page,lambda:overlay.locator('input[role="spinbutton"][aria-label="minutes"]').first)
        minutes.fill(str(expected.minute))
        try:
            wait_visible_recheck(web.page,lambda:overlay.locator('.dx-popup-done[role="button"]').first).click(timeout=10000)
        except Exception:
            observed=_fi_datetime_state(loc)
            if not (_datetime_text_matches(observed.get("display"),expected) and
                    _datetime_text_matches(observed.get("widget"),expected)):
                raise
        settle()
    observed={}
    def accepted():
        nonlocal observed
        observed=_fi_datetime_state(loc)
        return (_datetime_text_matches(observed.get("display"),expected) and
                (not observed.get("isDatebox") or _datetime_text_matches(observed.get("widget"),expected)))
    try:_fi_wait_condition(web,accepted,f"дата/час {alias}={value}")
    except AutomationError as exc:
        raise AutomationError(f"Дата {alias}: очікувалося {value!r}, стан поля {observed!r}.") from exc


def _fi_template_date_locator(web, alias):
    for candidate in web.selector_map.get(alias,[alias]):
        try:
            loc=web._candidate_locator(candidate)
            for index in range(loc.count()):
                current=loc.nth(index)
                if current.is_visible():return current
        except ActionCancelled:
            raise
        except Exception:
            continue
    return None


def _fi_type_template_date(loc, value):
    """Type through the date box mask so DevExtreme commits the value on Tab."""
    expected=datetime.strptime(value,"%d.%m.%Y %H:%M")
    loc.press("Control+A",timeout=3000)
    loc.type(expected.strftime("%d.%m.%Y, %H:%M"),delay=35,timeout=10000)


def _fi_template_dates_ready(web):
    for alias in ("template_episode_start","template_interaction_start"):
        loc=_fi_template_date_locator(web,alias)
        if loc is None:
            return False
        try:
            ready=loc.evaluate("""el => {
              const host=el.closest('.dx-datebox');
              const rect=el.getBoundingClientRect();
              return !!host && rect.width>0 && rect.height>0 &&
                !!window.DevExpress?.ui?.dxDateBox?.getInstance(host);
            }""",timeout=500)
            if not ready:
                return False
        except ActionCancelled:
            raise
        except Exception:
            return False
    try:
        return bool(web.page.evaluate("""() => {
          const visible=el => {
            const rect=el.getBoundingClientRect(), style=getComputedStyle(el);
            return rect.width>0 && rect.height>0 && style.display!=='none' &&
              style.visibility!=='hidden' && Number(style.opacity||1)>0;
          };
          const layers=[...document.querySelectorAll(
            '.dx-overlay-content,.popup.-extended,.popup.-default,[role="dialog"],.modal'
          )].filter(visible);
          const episodeSelector='input[id$="_templateOptionEpisodeDate"],input[id$="_copyOptionEpisodeDate"]';
          const encounterSelector='input[id$="_templateOptionEncounterDate"],input[id$="_copyOptionEncounterDate"]';
          if (!layers.length) {
            return [...document.querySelectorAll(episodeSelector)].some(episode =>
              visible(episode) && [...document.querySelectorAll(encounterSelector)].some(encounter => visible(encounter))
            );
          }
          return layers.some(layer => {
            const episode=[...layer.querySelectorAll(episodeSelector)].find(visible);
            const encounter=[...layer.querySelectorAll(encounterSelector)].find(visible);
            const rect=layer.getBoundingClientRect();
            return rect.height>=120 && !!episode && !!encounter;
          });
        }"""))
    except ActionCancelled:
        raise
    except Exception:
        return False


def _fi_wait_template_date_dialog(web):
    """Wait 1.5 seconds for the template dialog, then poll its date fields."""
    clock=__import__("time")
    raise_if_cancelled()
    web.page.wait_for_timeout(1500)
    deadline=clock.monotonic()+20
    previous_state=None
    stable=0
    while clock.monotonic()<deadline:
        raise_if_cancelled()
        fields_visible=_fi_template_dates_visible(web)
        ready=fields_visible and _fi_template_dates_ready(web)
        state=(_fi_modal_visible(web),fields_visible,ready)
        if state!=previous_state:
            log_event("template.dialog.state","Стан вікна параметрів шаблону змінився.",
                      action_id=web.action_id,
                      data={"modal_visible":state[0],"dates_visible":state[1],"dates_ready":state[2]})
            previous_state=state
        stable=stable+1 if ready else 0
        if stable>=3:
            return
        web.page.wait_for_timeout(500)
    raise AutomationError("Після очікування 1,5 секунди вікно шаблонів не показало обидва готові поля дат за 20 секунд.")


def _fi_wait_template_dates_saved(web, expected):
    """Confirm both template modal dates reached their DevExtreme widgets."""
    expected_date=datetime.strptime(expected,"%d.%m.%Y %H:%M")
    clock=__import__("time")
    deadline=clock.monotonic()+8
    observed={}
    stable=0
    while clock.monotonic()<deadline:
        raise_if_cancelled()
        observed={}
        ready=_fi_template_dates_ready(web)
        if ready:
            for alias in ("template_episode_start","template_interaction_start"):
                loc=_fi_template_date_locator(web,alias)
                try:
                    state=_fi_datetime_state(loc,timeout=500) if loc is not None else {}
                except ActionCancelled:
                    raise
                except Exception:
                    state={}
                observed[alias]=state
                if not (_datetime_text_matches(state.get("display"),expected_date) and
                        _datetime_text_matches(state.get("widget"),expected_date)):
                    ready=False
        stable=stable+1 if ready else 0
        if stable>=2:
            return
        web.page.wait_for_timeout(500)
    raise AutomationError(f"Дати у вікні шаблонів не збереглися: очікувалося {expected!r} в обох полях; стан {observed!r}.")


def _fi_template_dates_visible(web):
    return any(_fi_template_date_locator(web,alias) is not None
               for alias in ("template_episode_start","template_interaction_start"))


def _fi_interaction_form_visible(web):
    return any(_fi_template_date_locator(web,alias) is not None
               for alias in ("episode_start","interaction_start"))


def _fi_wait_before_template_button(web):
    """Let the new episode form settle before opening its template dialog."""
    clock=__import__("time")
    for stage,seconds in enumerate((5,5,10)):
        deadline=clock.monotonic()+seconds
        stable=0
        while clock.monotonic()<deadline:
            raise_if_cancelled()
            ready=not _fi_modal_visible(web) and _fi_interaction_form_visible(web)
            stable=stable+1 if ready else 0
            if stage and stable>=3:
                return
            web.page.wait_for_timeout(500)
    raise AutomationError("Форма нового епізоду не стабілізувалася перед відкриттям шаблонів за 5+5+10 секунд.")


def _fi_post_template_form_ready(web):
    """Return true only after the selected template has rendered its form."""
    # The form can contain other visible DevExtreme overlays, so detect the
    # template picker itself rather than treating every overlay as a blocker.
    if _fi_template_dates_visible(web) or _fi_organization_template_picker_visible(web):
        return False
    aliases=("episode_start","interaction_start","interaction_end","duration")
    controls={}
    for alias in aliases:
        try:
            # Poll without the staged 12-second wait_for_alias timeout; this
            # predicate is called repeatedly by the bounded readiness wait.
            for candidate in web.selector_map.get(alias,[alias]):
                found=web._candidate_locator(candidate).first
                if found.count() and found.is_visible() and found.is_enabled():
                    controls[alias]=found
                    break
            if alias not in controls:
                return False
        except ActionCancelled:
            raise
        except Exception:
            return False
    # A matching input can briefly exist before DevExtreme has initialized
    # the date widget that the workflow uses to read and validate its value.
    for alias in ("episode_start","interaction_start","interaction_end"):
        try:
            if not controls[alias].evaluate("""el => {
              const host=el.closest('.dx-datebox');
              return !host || !!window.DevExpress?.ui?.dxDateBox?.getInstance(host);
            }""",timeout=500):
                return False
        except ActionCancelled:
            raise
        except Exception:
            return False
    return True


def _fi_wait_post_template_form(web):
    """Leave the selected template untouched until its form stays ready."""
    clock=__import__("time")
    for stage,seconds in enumerate((5,5,10)):
        deadline=clock.monotonic()+seconds
        stable=0
        while clock.monotonic()<deadline:
            raise_if_cancelled()
            stable=stable+1 if _fi_post_template_form_ready(web) else 0
            if stage and stable>=3:
                return
            web.page.wait_for_timeout(500)
    raise AutomationError("Після вибору шаблону форма взаємодії не завантажила всі поля та календарі за 5+5+10 секунд.")


def _fi_visible_picker_contains(web, pattern):
    rows=web.page.locator(
        ".popup.-extended:visible tr, .popup.-default:visible tr, "
        ".dx-overlay-content:visible tr, [role='dialog']:visible tr"
    )
    for index in range(rows.count()):
        try:
            if pattern.search(rows.nth(index).inner_text(timeout=250)):
                return True
        except ActionCancelled:
            raise
        except Exception:
            continue
    return False


def _fi_visible_text(web, text):
    loc=web.page.get_by_text(text,exact=True)
    for index in range(loc.count()):
        try:
            if loc.nth(index).is_visible():return True
        except ActionCancelled:
            raise
        except Exception:
            continue
    return False


def _fi_modal_visible(web):
    return web.page.locator(
        ".popup.-extended:visible, .popup.-default:visible, "
        ".dx-overlay-content:visible, [role='dialog']:visible, .modal:visible"
    ).count()>0


def _fi_episode_picker_visible(web):
    return _fi_visible_picker_contains(web,re.compile(r"Створити новий епізод|Создать новый эпизод",re.I))


def _fi_organization_template_picker_visible(web):
    return _fi_visible_picker_contains(web,re.compile(r"^для\s+F84\.0\s+Діти\b",re.I))


def _fi_close_template_options(web):
    web.page.keyboard.press("Escape")
    deadline=__import__("time").monotonic()+3
    while __import__("time").monotonic()<deadline:
        raise_if_cancelled()
        if not _fi_template_dates_visible(web):
            return
        web.page.wait_for_timeout(200)
    # Some Eleks dialogs do not bind Escape. Close the visible overlay by its
    # own close control, then verify that the template date fields disappeared.
    close=web.page.locator(
        ".dx-overlay-content:visible .dx-closebutton:visible, "
        ".dx-overlay-content:visible [aria-label='Close']:visible, "
        ".dx-overlay-content:visible [aria-label='Закрити']:visible, "
        ".dx-overlay-content:visible [aria-label='Закрыть']:visible"
    )
    for index in range(close.count()-1,-1,-1):
        button=close.nth(index)
        try:
            if button.is_enabled():
                button.click(timeout=2000)
                break
        except ActionCancelled:
            raise
        except Exception:
            continue
    deadline=__import__("time").monotonic()+3
    while __import__("time").monotonic()<deadline:
        raise_if_cancelled()
        if not _fi_template_dates_visible(web):
            return
        web.page.wait_for_timeout(200)
    try:
        _fi_click(web,"templates")
    except ActionCancelled:
        raise
    except Exception:
        pass
    deadline=__import__("time").monotonic()+3
    while __import__("time").monotonic()<deadline:
        raise_if_cancelled()
        if not _fi_template_dates_visible(web):
            return
        web.page.wait_for_timeout(200)
    raise AutomationError("Не вдалося закрити неповністю відкрите вікно шаблонів взаємодій.")


def _fi_open_template_options(web):
    if _fi_template_dates_ready(web):
        return
    if _fi_template_dates_visible(web):
        _fi_close_template_options(web)
    for attempt in range(2):
        _fi_click(web,"templates")
        # Give the overlay time to mount before checking that its date controls
        # are inside the fully-rendered window rather than behind a blank modal.
        web.page.wait_for_timeout(1200)
        deadline=__import__("time").monotonic()+15
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            if _fi_template_dates_ready(web):
                return
            web.page.wait_for_timeout(250)
        if attempt==0:
            log_event("template.options.retry",
                      "Вікно параметрів шаблону відкрилося не повністю; закриваємо і відкриваємо повторно.",
                      action_id=web.action_id)
            _fi_close_template_options(web)
    raise AutomationError("Вікно шаблонів взаємодій не відобразило обидва готові поля дат після повторного відкриття.")


def _fi_row_code(text):
    codes=parse_icd_codes(text)
    return codes[0] if codes else ""


def _fi_table_rows(web, section_alias=None):
    grid_selectors={
        "diagnoses":"extendedgrid[id^='conditiongrid'] tr.dx-data-row:visible",
        "examinations":"#observationgrid tr.dx-data-row:visible",
    }
    row_selector=grid_selectors.get(section_alias,"tr.dx-data-row:visible")
    if section_alias:
        labels={
            "interactions":("Взаємодії","Взаимодействия"),
            "diagnostic_reports":("Діагностичні звіти","Диагностические отчеты"),
            "diagnoses":None,
            "examinations":None,
        }
        if section_alias not in labels:
            raise AutomationError(f"Невідома вкладка таблиці: {section_alias}.")
        if labels[section_alias]:
            web.click_content_text(labels[section_alias],retry_click=True)
    rows=web.page.locator(row_selector)
    out=[]
    for i in range(rows.count()):
        row=rows.nth(i)
        try:
            if row.locator("xpath=ancestor::*[contains(@class,'dx-overlay-content') or contains(@class,'popup')][1]").count():
                continue
            out.append((row,row.inner_text(timeout=500)))
        except Exception:pass
    return out


def _fi_barthel_exam_rows(web, *, wait=True):
    """Find Barthel_Extended by its grid cell, independent of generated column IDs."""
    clock=__import__("time")
    for seconds in ((2,5,5) if wait else (0.5,)):
        deadline=clock.monotonic()+seconds
        while clock.monotonic()<deadline:
            raise_if_cancelled()
            found=[];seen=set()
            containers=web.page.locator(".dx-scrollable-content:visible")
            for ci in range(containers.count()):
                # Match the named data cell directly; generated dx-col-* IDs
                # change between sessions and scanning every cell is noisy.
                cells=containers.nth(ci).locator(
                    'td[aria-describedby]:visible:has-text("Barthel_Extended")')
                for cell_index in range(cells.count()):
                    cell=cells.nth(cell_index)
                    try:
                        cell_text=" ".join(cell.inner_text(timeout=300).split())
                        if "barthel_extended" not in cell_text.casefold():continue
                        row=cell.locator("xpath=ancestor::tr[1]")
                        if row.count()!=1:continue
                        if row.locator("xpath=ancestor::*[contains(@class,'dx-overlay-content') or contains(@class,'popup')][1]").count():
                            continue
                        row_text=" ".join(row.inner_text(timeout=300).split())
                        key=row_text.casefold()
                        if key not in seen:
                            seen.add(key);found.append((row,row_text))
                    except ActionCancelled:raise
                    except Exception:pass
            if found:return found
            web.page.wait_for_timeout(250)
    return []


def _fi_diagnostic_report_rows(web, *, activate=True):
    """Read 96022/96037 report rows from the visible DevExtreme scroll table."""
    if activate:
        web.click_content_text(("Діагностичні звіти","Диагностические отчеты"),retry_click=True)
    observed={"scrollables":0,"rows":0}
    for seconds in (2,5,5):
        deadline=__import__("time").monotonic()+seconds
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            matches=[]
            containers=web.page.locator(".dx-scrollable-content:visible")
            observed["scrollables"]=containers.count()
            for ci in range(containers.count()):
                container=containers.nth(ci)
                rows=container.locator("tr:visible, [role='row']:visible")
                for ri in range(rows.count()):
                    row=rows.nth(ri)
                    try:
                        if row.locator("xpath=ancestor::*[contains(@class,'dx-overlay-content') or contains(@class,'popup')][1]").count():
                            continue
                        text=row.inner_text(timeout=400)
                        digits=re.sub(r"\D","",text)
                        if "9602200" in digits or "9603700" in digits:
                            matches.append((row,text))
                    except ActionCancelled:raise
                    except Exception:pass
            # Nested scrollable wrappers can expose the same row twice.
            unique=[];seen=set()
            for row,text in matches:
                key=(re.sub(r"\D","",text)," ".join(text.split()))
                if key in seen:continue
                seen.add(key);unique.append((row,text))
            observed["rows"]=len(unique)
            if unique:return unique
            web.page.wait_for_timeout(250)
    return []


def _fi_popup_with_text(web, predicate, description):
    """Wait for one visible popup whose text proves the expected transition."""
    selectors=(".popup.-notification:visible", ".popup.-default:visible",
               ".popup.-extended:visible", ".dx-overlay-content:visible",
               "[role='dialog']:visible", ".modal:visible")
    clock=__import__("time")
    for seconds in (2,5,5):
        deadline=clock.monotonic()+seconds
        while clock.monotonic()<deadline:
            raise_if_cancelled()
            for selector in selectors:
                candidates=web.page.locator(selector)
                matched=[]
                for index in range(candidates.count()):
                    popup=candidates.nth(index)
                    try:
                        text=" ".join(popup.inner_text(timeout=300).split())
                        if popup.is_visible() and predicate(text):matched.append((popup,text))
                    except ActionCancelled:raise
                    except Exception:pass
                if len(matched)==1:return matched[0]
                if len(matched)>1:
                    raise AutomationError(f"{description}: знайдено кілька відповідних модальних вікон.")
            web.page.wait_for_timeout(250)
    raise AutomationError(f"Не з'явилося підтвердження: {description} після очікування 2+5+5 секунд.")


def _fi_popup_text_visible(web, expected_text):
    expected=normalize_patient_name(expected_text)
    for selector in (".popup.-notification:visible", ".popup.-default:visible",
                     ".popup.-extended:visible", ".dx-overlay-content:visible",
                     "[role='dialog']:visible", ".modal:visible"):
        candidates=web.page.locator(selector)
        for index in range(candidates.count()):
            try:
                if normalize_patient_name(candidates.nth(index).inner_text(timeout=300))==expected:
                    return True
            except ActionCancelled:raise
            except Exception:pass
    return False


def _fi_confirm_popup_ok(web, popup, description, popup_text=None):
    # Doctor Eleks notifications use this directive as the actual close/next
    # action. Clicking a nested dx-button label can leave the popup open.
    popup_text=popup_text or popup.inner_text(timeout=500).strip()
    web.page.wait_for_timeout(2000)
    matches=[]
    primary=popup.locator('.popup_buttons button.btn.btn-primary')
    for index in range(primary.count()):
        button=primary.nth(index)
        try:
            label=normalize_patient_name(button.inner_text(timeout=300))
            if button.is_visible() and button.is_enabled() and label in (
                    normalize_patient_name("ОК"),"ok"):
                matches.append(button)
        except ActionCancelled:raise
        except Exception:pass
    proven=popup.locator('.popup_buttons button[click-next-action="true"]')
    if not matches:
        for index in range(proven.count()):
            button=proven.nth(index)
            try:
                label=normalize_patient_name(button.inner_text(timeout=300))
                if button.is_visible() and button.is_enabled() and label in (
                        normalize_patient_name("ОК"),"ok"):
                    matches.append(button)
            except ActionCancelled:raise
            except Exception:pass
    if not matches:
        controls=popup.locator('button, [role="button"], .dx-button')
        for index in range(controls.count()):
            control=controls.nth(index)
            try:
                label=normalize_patient_name(control.inner_text(timeout=300))
                if control.is_visible() and control.is_enabled() and label in (
                        normalize_patient_name("ОК"),"ok"):
                    matches.append(control)
            except ActionCancelled:raise
            except Exception:pass
    if len(matches)!=1:
        raise AutomationError(f"{description}: очікувалася одна доступна кнопка «ОК», знайдено {len(matches)}.")
    matches[0].click(timeout=10000)
    web.page.wait_for_timeout(2000)
    _fi_wait_condition(web,lambda:not _fi_popup_text_visible(web,popup_text),
                       f"закриття/зміна вікна: {description}")


def _fi_confirm_diagnostic_report_delete(web):
    """Validate and acknowledge both Doctor Eleks delete dialogs."""
    def is_delete_prompt(text):
        normalized=normalize_patient_name(text)
        return "видалити введений діагностичний звіт" in normalized or (
            "видалити" in normalized and "діагностичний звіт" in normalized)
    popup,text=_fi_popup_with_text(web,is_delete_prompt,"підтвердження видалення звіту")
    if not any(code in re.sub(r"\D","",text) for code in ("9603700","960370")):
        # The prompt can omit the report code; its exact delete wording is
        # still required before the destructive confirmation is acknowledged.
        if "видалити введений діагностичний звіт" not in normalize_patient_name(text):
            raise AutomationError("Діалог видалення не підтверджує видалення саме діагностичного звіту.")
    _fi_confirm_popup_ok(web,popup,"підтвердження видалення звіту",text)

    def is_success(text):
        normalized=normalize_patient_name(text)
        deleted=("видал" in normalized or "удал" in normalized)
        return deleted and (("успіш" in normalized or "успеш" in normalized) or
                            "діагностичний звіт видалено" in normalized or
                            "диагностический отчет удален" in normalized)
    popup,text=_fi_popup_with_text(web,is_success,"успішне видалення діагностичного звіту")
    _fi_confirm_popup_ok(web,popup,"повідомлення про успішне видалення звіту",text)


def _fi_confirm_barthel_delete(web):
    """Acknowledge the pre-delete question; row disappearance proves success."""
    def is_prompt(text):
        normalized=normalize_patient_name(text)
        return (("видал" in normalized or "удал" in normalized) and
                any(token in normalized for token in ("обстеж", "наблюден", "спостереж", "введен")))
    popup,text=_fi_popup_with_text(web,is_prompt,"підтвердження видалення обстеження Barthel_Extended")
    _fi_confirm_popup_ok(web,popup,"підтвердження видалення обстеження Barthel_Extended",text)


def _fi_report_code(row_text):
    """Extract the report code token without folding the row's date into it."""
    text=str(row_text or "")
    if re.search(r"(?<!\d)(?:96022|96[.]?0[.]?22)(?:[ .-]?00)?(?!\d)",text):
        return "9602200"
    if re.search(r"(?<!\d)(?:96037|96[.]?0[.]?37)(?:[ .-]?00)?(?!\d)",text):
        return "9603700"
    return ""


def _fi_diagnosis_page_state(web):
    grids=web.page.locator("extendedgrid[id^='conditiongrid']:visible")
    if grids.count()!=1:
        raise AutomationError(f"Таблиця діагнозів: очікувалася одна видима таблиця, знайдено {grids.count()}.")
    grid=grids.first
    # In this screen DevExtreme places the pager next to the custom
    # <extendedgrid> element, not always inside it. Walk up only to the nearest
    # wrapper that contains the diagnosis grid and its pager.
    root=grid;pager=None
    for _depth in range(8):
        pagers=root.locator(".dx-pager:visible")
        if pagers.count()==1:
            pager=pagers.first;break
        if pagers.count()>1:break
        try:root=root.locator("xpath=..")
        except Exception:break
    if pager is None:
        global_pagers=web.page.locator(".dx-pager:visible")
        if global_pagers.count()==1:pager=global_pagers.first
        elif global_pagers.count()>1:
            raise AutomationError("Таблиця діагнозів: пагінатор є, але не вдалося однозначно прив'язати його до таблиці.")
    if pager is None:
        external_pages=web.page.locator(".dx-pages:visible .dx-page[role='button']")
        if external_pages.count():
            raise AutomationError("Таблиця діагнозів: кнопки сторінок знайдено, але пагінатор не прив'язаний до таблиці.")
        return grid,None,[1],1
    buttons=pager.locator(".dx-pages .dx-page[role='button']")
    if not buttons.count():buttons=pager.locator(".dx-page[role='button']")
    def page_number(button):
        label=button.get_attribute("aria-label") or ""
        match=re.search(r"(?:page\s*)?(\d+)",label,re.I)
        if not match:
            match=re.search(r"^\s*(\d+)\s*$",button.inner_text(timeout=300))
        return int(match.group(1)) if match else None
    pages=[]
    for index in range(buttons.count()):
        number=page_number(buttons.nth(index))
        if number is not None:pages.append(number)
    pages=sorted(set(pages))
    if not pages:raise AutomationError("Таблиця діагнозів: не вдалося прочитати номери сторінок.")
    active=pager.locator(".dx-pages .dx-page.dx-selection[role='button']")
    if not active.count():active=pager.locator(".dx-page.dx-selection[role='button']")
    active_number=page_number(active.first) if active.count()==1 else None
    current=active_number if active_number in pages else pages[0]
    return grid,pager,pages,current


def _fi_switch_diagnosis_page(web, page_number):
    pager=None;current=None;button=None
    for seconds in (2,5,5):
        deadline=__import__("time").monotonic()+seconds
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            try:
                _grid,pager,pages,current=_fi_diagnosis_page_state(web)
                if page_number in pages:
                    if current==page_number:return
                    buttons=(pager.locator(".dx-pages .dx-page[role='button']") if pager else None)
                    if buttons is not None and not buttons.count():buttons=pager.locator(".dx-page[role='button']")
                    if buttons:
                        for index in range(buttons.count()):
                            candidate=buttons.nth(index)
                            label=candidate.get_attribute("aria-label") or ""
                            match=re.search(r"(?:page\s*)?(\d+)",label,re.I)
                            if not match:
                                match=re.search(r"^\s*(\d+)\s*$",candidate.inner_text(timeout=300))
                            if match and int(match.group(1))==page_number:
                                button=candidate;break
                    if button is not None:break
            except ActionCancelled:raise
            except Exception:pass
            web.page.wait_for_timeout(250)
        if button is not None:break
    if button is None:
        raise AutomationError(f"Таблиця діагнозів: сторінка {page_number} не з'явилася після очікування 2+5+5 секунд.")
    button.click(timeout=10000)
    def page_active():
        _grid,_pager,_pages,current=_fi_diagnosis_page_state(web)
        return current==page_number
    _fi_wait_condition(web,page_active,f"відкриття сторінки діагнозів {page_number}")


def _fi_all_diagnosis_rows(web):
    _grid,_pager,pages,original_page=_fi_diagnosis_page_state(web)
    records=[]
    for page_number in pages:
        _fi_switch_diagnosis_page(web,page_number)
        records.extend((row,text,page_number) for row,text in _fi_table_rows(web,"diagnoses"))
    _fi_switch_diagnosis_page(web,original_page)
    return records


def _fi_find_diagnosis_row(web, code, occurrence=0):
    _grid,_pager,pages,original_page=_fi_diagnosis_page_state(web)
    seen=0
    for page_number in pages:
        _fi_switch_diagnosis_page(web,page_number)
        matches=[(row,text) for row,text in _fi_table_rows(web,"diagnoses")
                 if _fi_row_code(text)==code]
        if seen+len(matches)>occurrence:return matches[occurrence-seen]
        seen+=len(matches)
    _fi_switch_diagnosis_page(web,original_page)
    return None


def _fi_saved_interaction_matches(row_text, start_text):
    start=datetime.strptime(start_text,"%d.%m.%Y %H:%M")
    end=start+timedelta(minutes=60)
    formats=("%d.%m.%Y %H:%M","%Y-%m-%d %H:%M")
    row=" ".join(str(row_text).split())
    return (any(start.strftime(fmt) in row for fmt in formats) and
            any(end.strftime(fmt) in row for fmt in formats))


def _fi_wait_condition(web, predicate, description):
    for seconds in (2,5,5):
        deadline=__import__("time").monotonic()+seconds
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            if predicate():return True
            web.page.wait_for_timeout(250)
    raise AutomationError(f"Після дії не настало очікуване DOM-стан: {description}.")


def _fi_control_text(locator):
    try:return locator.input_value(timeout=500).strip()
    except Exception:return locator.inner_text(timeout=500).strip()


def _fi_place_matches(expected, actual):
    """Compare the configured facility with its Doctor Eleks display name."""
    words=lambda value:re.findall(r"[^\W_]+",normalize_patient_name(value),re.UNICODE)
    wanted=words(expected)
    shown=words(actual)
    if not wanted:return False
    if shown==wanted:return True
    # Eleks appends the same numbered facility suffix after the quoted name:
    # ТОВ 2 Назва -> ТОВ 2 \"Назва\" 2.  No other extra word is accepted.
    return (len(wanted)>=3 and wanted[1].isdigit() and
            shown==wanted+[wanted[1]])


_FI_DIAGNOSIS_FIELDS={
    "diagnosis_role":"diagnosesrole",
    "diagnosis_certainty":"conditionVerification_status",
    "diagnosis_clinical":"conditionClinical_status",
    "diagnosis_installed_at":"conditionOnset_date",
}
_FI_DIAGNOSIS_FIELDS=frozenset(_FI_DIAGNOSIS_FIELDS)
_FI_DIAGNOSIS_FIELD_SUFFIXES={
    "diagnosis_role":"_diagnosesrole",
    "diagnosis_certainty":"_conditionVerification_status",
    "diagnosis_clinical":"_conditionClinical_status",
    "diagnosis_installed_at":"_conditionOnset_date",
}


def _fi_diagnosis_modal(web):
    modal=web.page.locator(".popup.-extended:visible")
    count=modal.count()
    if count!=1:
        raise AutomationError(f"Форма діагнозу: очікувалася одна відкрита модалка, знайдено {count}.")
    return modal.first


def _fi_diagnosis_control(web, alias):
    if alias not in _FI_DIAGNOSIS_FIELD_SUFFIXES:
        raise AutomationError(f"Невідоме поле модалки діагнозу: {alias}.")
    observed={"count":0}
    def resolve():
        modal=_fi_diagnosis_modal(web)
        field=modal.locator(
            f'input.dx-texteditor-input[role="combobox"][id$="{_FI_DIAGNOSIS_FIELD_SUFFIXES[alias]}"]')
        observed["count"]=field.count()
        if observed["count"]==1 and field.first.is_visible() and field.first.is_enabled():
            return field.first
        return None
    for seconds in (2,5,5):
        deadline=__import__("time").monotonic()+seconds
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            try:
                control=resolve()
                if control is not None:return control
            except (ActionCancelled,ManualCheckpoint):raise
            except Exception:pass
            web.page.wait_for_timeout(250)
    raise AutomationError(f"У відкритій модалці діагнозу не знайдено доступне поле {alias}; збігів {observed['count']}.")


def _fi_diagnosis_button(web, label):
    observed={"count":0}
    def resolve():
        modal=_fi_diagnosis_modal(web)
        button=modal.locator(f'[role="button"][aria-label="{label}"]:visible')
        observed["count"]=button.count()
        if observed["count"]==1 and button.first.is_enabled():return button.first
        return None
    for seconds in (2,5,5):
        deadline=__import__("time").monotonic()+seconds
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            try:
                button=resolve()
                if button is not None:return button
            except (ActionCancelled,ManualCheckpoint):raise
            except Exception:pass
            web.page.wait_for_timeout(250)
    raise AutomationError(f"У відкритій модалці діагнозу не знайдено кнопку «{label}»; збігів {observed['count']}.")


def _fi_wait_extended_popup_closed(web, description):
    _fi_wait_condition(web,lambda:web.page.locator(".popup.-extended:visible").count()==0,
                       description)
    # DevExtreme removes the popup before its grid finishes re-layout.
    web.page.wait_for_timeout(2000)


def _fi_wait_diagnosis_modal_closed(web):
    _fi_wait_extended_popup_closed(web,"закриття модалки реєстрації діагнозу")


def _fi_select_field(web, alias, expected):
    control=(_fi_diagnosis_control(web,alias) if alias in _FI_DIAGNOSIS_FIELDS
             else web.wait_for_alias(alias,enabled=True))
    if normalize_patient_name(expected)==normalize_patient_name(_fi_control_text(control)):
        return False
    _fi_click(web,alias,locator=control)
    for seconds in (2,5,5):
        deadline=__import__("time").monotonic()+seconds
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            matches=[]
            for selector in (".dx-list-item-content:visible", "[role='option']:visible"):
                options=web.page.locator(selector)
                for index in range(options.count()):
                    option=options.nth(index)
                    try:
                        if normalize_patient_name(option.inner_text(timeout=300))==normalize_patient_name(expected):
                            matches.append(option)
                    except ActionCancelled:raise
                    except Exception:pass
                if matches:break
            if len(matches)>1:raise AutomationError(f"Поле {alias}: кілька пунктів «{expected}» у списку.")
            if len(matches)==1:
                _fi_click(web,alias+".option",locator=matches[0])
                _fi_wait_condition(web,lambda:normalize_patient_name(_fi_control_text(control))==normalize_patient_name(expected),
                                   f"встановлення {alias}={expected}")
                return True
            web.page.wait_for_timeout(250)
    raise AutomationError(f"Поле {alias}: не знайдено точний пункт «{expected}» після 2+5+5 секунд.")


def _fi_row_action(web, row, labels):
    """Resolve the interactive control inside one exact data row."""
    clock=__import__("time")
    wanted={normalize_patient_name(label) for label in labels}
    previous_boxes={}
    for seconds in (2,5,5):
        deadline=clock.monotonic()+seconds
        while clock.monotonic()<deadline:
            raise_if_cancelled()
            matches=[]
            controls=row.locator('button, [role="button"], a, [ng-click], [data-ng-click], [onclick]')
            for index in range(controls.count()):
                control=controls.nth(index)
                try:
                    if not control.is_visible() or not control.is_enabled():continue
                    title=control.get_attribute("aria-label") or control.inner_text(timeout=300)
                    if normalize_patient_name(title) in wanted:
                        box=control.bounding_box()
                        if not box:continue
                        geometry=tuple(round(box[key],1) for key in ("x","y","width","height"))
                        previous=previous_boxes.get(title)
                        previous_boxes[title]=geometry
                        if previous is not None and all(abs(a-b)<=0.5 for a,b in zip(previous,geometry)):
                            matches.append(control)
                except ActionCancelled:raise
                except Exception:pass
            if len(matches)==1:
                web._log("row.action.ready","/".join(labels))
                return matches[0]
            if len(matches)>1:
                raise AutomationError(f"У рядку знайдено {len(matches)} елементів дії «{labels[0]}»; вибір неоднозначний.")
            web.page.wait_for_timeout(250)
    return None


def _fi_picker_rows(web):
    """Read only the currently visible picker, using the proven DevExtreme rows."""
    for seconds in (2,5,5):
        deadline=__import__("time").monotonic()+seconds
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            for selector in (".popup.-extended", ".popup.-default", ".dx-overlay-content", "[role='dialog']", ".modal"):
                popups=web.page.locator(selector)
                for index in range(popups.count()):
                    popup=popups.nth(index)
                    try:
                        if not popup.is_visible():continue
                        rows=popup.locator("tr.dx-data-row")
                        if not rows.count():rows=popup.locator("tr")
                        visible=[rows.nth(i) for i in range(rows.count()) if rows.nth(i).is_visible()]
                        visible=[row for row in visible if
                                 row.get_by_text("Вибрати",exact=True).count() or
                                 row.get_by_text("Выбрать",exact=True).count() or
                                 row.locator('[aria-label="Вибрати"], [aria-label="Выбрать"]').count()]
                        if visible:return popup,visible
                    except ActionCancelled:raise
                    except Exception:pass
            web.page.wait_for_timeout(250)
    raise AutomationError("Модальне вікно не показало табличні рядки після очікування 2+5+5 секунд.")


def _fi_create_new_episode(web):
    popup,rows=_fi_picker_rows(web);matches=[]
    for row in rows:
        try:
            if re.search(r"Створити новий епізод|Создать новый эпизод",row.inner_text(timeout=400),re.I):
                matches.append(row)
        except Exception:pass
    if len(matches)!=1:
        raise AutomationError(f"У списку епізодів очікувався один рядок «Створити новий епізод», знайдено {len(matches)}.")
    choose=_fi_row_action(web,matches[0],["Вибрати","Выбрать"])
    if not choose:raise AutomationError("У рядку створення епізоду не знайдено кнопку «Вибрати».")
    choose.click(timeout=10000)
    matches[0].wait_for(state="hidden",timeout=12000)


def _fi_click_new_interaction(web):
    """Open the interactions tab, then click its own add button."""
    try:
        web.click_content_text(("Взаємодії","Взаимодействия"),retry_click=True)
        web.page.wait_for_timeout(2000)
        labels=web.page.get_by_text(re.compile(r"^(Нова взаємодія|Новое взаимодействие):?$"))
        last_state={"labels":0,"buttons":0}
        for seconds in (2,5,5):
            deadline=__import__("time").monotonic()+seconds
            while __import__("time").monotonic()<deadline:
                raise_if_cancelled()
                last_state["labels"]=labels.count()
                for index in range(last_state["labels"]):
                    label=labels.nth(index)
                    try:
                        if not label.is_visible():continue
                        text=normalize_patient_name(label.inner_text(timeout=300)).rstrip(":")
                        if text not in {"нова взаємодія","новое взаимодействие"}:continue
                        group=label.locator(
                            "xpath=ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' dx-box-flex ')][1]")
                        buttons=group.locator(':scope > .dx-box-item [role="button"][aria-label="add"]')
                        last_state["buttons"]=buttons.count()
                        ready=[buttons.nth(i) for i in range(buttons.count())
                               if buttons.nth(i).is_visible() and buttons.nth(i).is_enabled()]
                        if len(ready)==1:
                            _fi_click(web,"new_interaction.add",locator=ready[0])
                            return
                    except (ActionCancelled,ManualCheckpoint):raise
                    except Exception:continue
                web.page.wait_for_timeout(250)
        raise AutomationError(f"Не знайдено доступний «+» біля «Нова взаємодія»; стан {last_state!r}.")
    except (ActionCancelled,ManualCheckpoint):raise
    except Exception as exc:
        raise AutomationError(
            f"Не вдалося відкрити вкладку «Взаємодії» або натиснути «+» біля «Нова взаємодія»: {exc}"
        ) from exc


def _fi_choose_exact_code(web, code, *, change):
    if change:_fi_click(web,"change")
    _fi_fill_diagnosis_search(web,code)
    popup,rows=_fi_picker_rows(web)
    matches=[]
    for row in rows:
        if _fi_row_has_exact_code(row,code):matches.append(row)
    if len(matches)!=1:
        raise AutomationError(f"Пошук МКХ {code}: очікувався один рядок з точним кодом, знайдено {len(matches)}.")
    choose=_fi_row_action(web,matches[0],["Вибрати","Выбрать"])
    if not choose:raise AutomationError(f"У рядку точного коду МКХ {code} немає «Вибрати».")
    choose.click(timeout=10000)
    matches[0].wait_for(state="hidden",timeout=12000)


def _fi_row_has_exact_code(row, code):
    cells=row.locator('td[role="gridcell"], td')
    for index in range(cells.count()):
        try:
            cell_text=cells.nth(index).inner_text(timeout=400).strip()
            if re.match(rf"^{re.escape(code)}(?![\w.])",cell_text,re.I):return True
        except ActionCancelled:raise
        except Exception:pass
    return False


def _fi_diagnosis_search_input(web):
    """Resolve the code filter input by its own column header inside the picker."""
    popup=web.page.locator(".popup.-extended:visible")
    if popup.count()!=1:
        raise AutomationError(f"Пошук МКХ: очікувалася одна відкрита модалка коду, знайдено {popup.count()}.")
    headers=popup.first.locator('tr.dx-header-row td[role="columnheader"], [role="columnheader"]')
    code_headers=[]
    for index in range(headers.count()):
        header=headers.nth(index)
        try:
            label=normalize_patient_name(header.inner_text(timeout=300))
            aria=normalize_patient_name(header.get_attribute("aria-label") or "")
            if label=="код" or aria in {"стовпець код","столбец код"}:
                code_headers.append(header)
        except ActionCancelled:raise
        except Exception:pass
    if len(code_headers)!=1:
        raise AutomationError(f"Пошук МКХ: заголовок колонки «Код» знайдено {len(code_headers)} разів.")
    header_id=code_headers[0].get_attribute("id")
    if not header_id:
        raise AutomationError("Пошук МКХ: у заголовка колонки «Код» немає id для прив'язки поля фільтра.")
    filters=popup.first.locator(f'input.dx-texteditor-input[aria-label="Фільтр"][aria-describedby="{header_id}"]')
    if filters.count()!=1 or not filters.first.is_visible() or not filters.first.is_enabled():
        raise AutomationError(f"Пошук МКХ: для колонки «Код» очікувалося одне доступне поле dx-texteditor-input, знайдено {filters.count()}.")
    return filters.first


def _fi_fill_diagnosis_search(web, code):
    observed={"value":"","header":"Код"}
    def resolve():
        try:
            return _fi_diagnosis_search_input(web)
        except ActionCancelled:raise
        except Exception:return None
    field=None
    for seconds in (2,5,5):
        deadline=__import__("time").monotonic()+seconds
        while __import__("time").monotonic()<deadline:
            raise_if_cancelled()
            field=resolve()
            if field is not None:break
            web.page.wait_for_timeout(250)
        if field is not None:break
    if field is None:
        raise AutomationError("Пошук МКХ: не знайдено доступне поле фільтра під заголовком «Код» у відкритій модалці.")
    web.fill("code_search",code,locator=field)
    def value_accepted():
        observed["value"]=field.input_value(timeout=500).strip()
        return observed["value"]==str(code)
    _fi_wait_condition(web,value_accepted,f"введення коду МКХ {code} у фільтр колонки «Код»")
    def result_visible():
        # This picker is rendered by datagriddx in some builds and extendedgrid
        # in others; bind to its visible data rows instead of the wrapper tag.
        rows=web.page.locator(".popup.-extended:visible tr.dx-data-row:visible")
        return any(_fi_row_has_exact_code(rows.nth(i),code) for i in range(rows.count()))
    _fi_wait_condition(web,result_visible,f"поява точного коду МКХ {code} у результатах фільтра")


def _fi_diagnosis_row_on_page(web, code, page_number):
    """Resolve one diagnosis on its known page without sweeping all pages."""
    _fi_switch_diagnosis_page(web,page_number)
    matches=[(row,text) for row,text in _fi_table_rows(web,"diagnoses")
             if _fi_row_code(text)==code]
    if len(matches)>1:
        raise AutomationError(f"На сторінці {page_number} код {code} дублюється; рядок не змінено.")
    return matches[0] if matches else None


def _fi_reconcile_diagnoses(web, source_codes, target_dt):
    source=[normalize_icd(c) for c in source_codes]
    if not source:raise AutomationError("У джерелі немає кодів діагнозів для звірки.")
    primary_code=source[0]
    source_mutable=[code for code in source if not code.startswith("Z")]
    initial_rows=_fi_all_diagnosis_rows(web)
    template=[]
    for row,text,page_number in initial_rows:
        code=_fi_row_code(text)
        if code:
            template.append({"row":row,"code":normalize_icd(code),"text":text,
                             "page":page_number})
    current_primary=[record for record in template if "основний" in record["text"].casefold()]
    if len(current_primary)>1:
        raise AutomationError("У шаблоні більше одного основного діагнозу; автоматичне редагування зупинено.")
    if current_primary:
        primary_row=current_primary[0]
        if primary_row["code"]!=primary_code:
            if primary_code.startswith("Z"):
                raise AutomationError(f"Основний діагноз {primary_code} належить до незмінних Z-кодів; автоматичне редагування зупинено.")
            existing_primary=[record for record in template
                              if record["code"]==primary_code and record is not primary_row]
            if len(existing_primary)>1:
                raise AutomationError(f"Код основного діагнозу {primary_code} уже дублюється серед супутніх; потрібна ручна перевірка.")
            if existing_primary:
                duplicate=existing_primary[0]
                current=_fi_diagnosis_row_on_page(web,primary_code,duplicate["page"])
                if current is None:raise AutomationError(f"Супутній діагноз {primary_code} зник до перенесення в основний.")
                delete=_fi_row_action(web,current[0],["Видалити","Удалить"])
                if not delete:raise AutomationError(f"Не знайдено видалення супутнього діагнозу {primary_code} перед перенесенням у основний.")
                delete.click(timeout=10000)
                _fi_wait_condition(web,lambda:not any(
                    _fi_row_code(text)==primary_code for _,text in _fi_table_rows(web,"diagnoses")),
                    f"видалення попереднього супутнього рядка {primary_code}")
            fresh=[(row,text,page) for row,text,page in _fi_all_diagnosis_rows(web)
                   if _fi_row_code(text)==primary_row["code"] and "основний" in text.casefold()]
            if len(fresh)!=1:
                raise AutomationError("Не вдалося повторно знайти єдиний рядок шаблону з роллю «Основний».")
            current=fresh[0]
            edit=_fi_row_action(web,current[0],["Редагувати","Редактировать"])
            if not edit:raise AutomationError(f"Не знайдено редагування поточного основного діагнозу {primary_row['code']}.")
            edit.click(timeout=10000)
            _fi_choose_exact_code(web,primary_code,change=True)
            _fi_set_diagnosis_time(web,target_dt,primary=True)
            log_event("diagnoses.primary_row_replaced",
                      "Основний діагноз шаблону замінено кодом першого діагнозу з джерела.",
                      action_id=web.action_id,data={"source_primary":primary_code})
            initial_rows=_fi_all_diagnosis_rows(web)
            template=[]
            for row,text,page_number in initial_rows:
                code=_fi_row_code(text)
                if code:
                    template.append({"row":row,"code":normalize_icd(code),"text":text,
                                     "page":page_number})
            current_primary=[record for record in template if "основний" in record["text"].casefold()]
            if len(current_primary)!=1 or current_primary[0]["code"]!=primary_code:
                raise AutomationError(f"Після заміни код {primary_code} не підтверджено в ролі «Основний».")
        # When the primary code from the source is missing, assign it to the
        # template row that is already primary. Doctor Eleks may hide the
        # «Основний» option while a different row still owns that role.
        primary_row=current_primary[0]
        template.sort(key=lambda record:record is not primary_row)
    plan=_plan_diagnosis_reconciliation(source_mutable,template)
    replacements=plan["replace"]
    surplus=plan["delete"]

    # Replace only a unique, wrong code, preserving its existing table row.
    for record,code in sorted(replacements,key=lambda pair:pair[0]["page"]):
        current=_fi_diagnosis_row_on_page(web,record["code"],record["page"])
        if current is None:raise AutomationError(f"Діагноз {record['code']} зник до заміни.")
        action=_fi_row_action(web,current[0],["Редагувати","Редактировать"])
        if not action:raise AutomationError(f"Не знайдено «Редагувати» для діагнозу {record['code']}.")
        action.click(timeout=10000);_fi_choose_exact_code(web,code,change=True)
        _fi_set_diagnosis_time(web,target_dt,primary=(code==primary_code))

    # Delete only a proven unique surplus, last page first to avoid pager shifts.
    for record in sorted(surplus,key=lambda item:item.get("page",1),reverse=True):
        current=_fi_diagnosis_row_on_page(web,record["code"],record.get("page",1))
        if current is None:raise AutomationError(f"Зайвий діагноз {record['code']} зник до видалення.")
        action=_fi_row_action(web,current[0],["Видалити","Удалить"])
        if not action:raise AutomationError(f"Не знайдено «Видалити» для зайвого діагнозу {record['code']}.")
        action.click(timeout=10000)
        _fi_wait_condition(web,lambda:not any(
            _fi_row_code(text)==record["code"] for _,text in _fi_table_rows(web,"diagnoses")),
            f"видалення діагнозу {record['code']}")

    for code in plan["add"]:
        _fi_click(web,"add_diagnosis");_fi_choose_exact_code(web,code,change=False)
        _fi_set_diagnosis_time(web,target_dt,primary=(code==primary_code))

    # Keep an already-correct code and role completely untouched. Correct only
    # a unique retained row whose role differs from the source order.
    duplicate_codes={record["code"] for record in plan["protected_duplicates"]}
    for record in sorted(plan["keep"],key=lambda item:item["code"]==primary_code):
        code=record["code"]
        if code in duplicate_codes:continue
        expected_role="основний" if code==primary_code else "супутній"
        if expected_role in record.get("text","").casefold():continue
        current=_fi_diagnosis_row_on_page(web,code,record.get("page",1))
        if current is None:raise AutomationError(f"Діагноз {code} зник до перевірки ролі.")
        action=_fi_row_action(web,current[0],["Редагувати","Редактировать"])
        if not action:raise AutomationError(f"Не знайдено «Редагувати» для діагнозу {code}.")
        action.click(timeout=10000)
        _fi_set_diagnosis_time(web,target_dt,primary=(code==primary_code))

    # One final sweep verifies the complete result after all planned changes.
    rows=_fi_all_diagnosis_rows(web)
    final=[]
    for row,text,_page_number in rows:
        code=_fi_row_code(text)
        if code:final.append((row,normalize_icd(code),text))
    final_mutable=Counter(code for _,code,_ in final if not code.startswith("Z"))
    expected_mutable=Counter(plan["expected_mutable_counts"])
    if final_mutable!=expected_mutable:
        raise AutomationError(f"Після звірки діагнозів очікувалося {dict(expected_mutable)}, отримано {dict(final_mutable)}.")
    final_z=Counter(code for _,code,_ in final if code.startswith("Z"))
    expected_z=Counter(record["code"] for record in plan["immutable_z"])
    if final_z!=expected_z:
        raise AutomationError(f"Змінився незмінний список Z-діагнозів: очікувалося {dict(expected_z)}, отримано {dict(final_z)}.")
    primary=[text for _,code,text in final if code==primary_code and "основний" in text.casefold()]
    if len(primary)!=1:
        raise AutomationError(f"Очікувався один перший діагноз {primary_code} з роллю «Основний», знайдено {len(primary)}.")
    protected_codes={record["code"] for record in plan["protected_duplicates"]}
    wrong_roles=[]
    for _row,code,text in final:
        if code in protected_codes:continue
        expected_role="основний" if code==primary_code else "супутній"
        if expected_role not in text.casefold():wrong_roles.append(code)
    if wrong_roles:
        raise AutomationError("Після звірки неправильна роль у діагнозів: "+", ".join(wrong_roles)+".")
    main_rows=[code for _row,code,text in final if "основний" in text.casefold()]
    if main_rows!=[primary_code]:
        raise AutomationError(f"Основним має бути лише перший код {primary_code}; фактичні основні коди: {main_rows}.")
    log_event("diagnoses.reconciled","Діагнози звірено з джерелом.",action_id=web.action_id,
              data={"source_count":len(source),"template_z_count":sum(expected_z.values()),
                    "replaced":len(replacements),"deleted":len(surplus),"added":len(plan["add"]),
                    "protected_duplicates":len(plan["protected_duplicates"])})


def _fi_set_diagnosis_time(web, target_dt, primary=False):
    # Let the diagnosis editor settle after the code picker closes.
    web.page.wait_for_timeout(2000)
    _fi_select_field(web,"diagnosis_role","Основний" if primary else "Супутній")
    _fi_select_field(web,"diagnosis_certainty","Заключний")
    _fi_select_field(web,"diagnosis_clinical","Активний")
    _fi_fill(web,"diagnosis_installed_at",target_dt)
    _fi_click(web,"save",locator=_fi_diagnosis_button(web,"Зберегти"))
    _fi_wait_diagnosis_modal_closed(web)


def _fi_referral_date(value):
    parsed=parse_interaction_datetime(value)
    if parsed:return parsed.date()
    for fmt in ("%d.%m.%Y","%Y-%m-%d"):
        try:return datetime.strptime(str(value).strip(),fmt).date()
        except ValueError:pass
    return None


def _fi_read_named_input(web, name):
    locator=web.page.locator(f'input[name="{name}"]')
    values=[]
    for index in range(locator.count()):
        field=locator.nth(index)
        try:
            if not field.is_visible():continue
            value=field.input_value(timeout=2000).strip()
            if value:values.append(value)
        except ActionCancelled:raise
        except Exception:continue
    distinct=set(values)
    return next(iter(distinct)) if len(distinct)==1 else ""


def _fi_selected_referral_state(web):
    references=web.page.locator('input[name="encounterIncoming_referrals"]')
    radio=web.page.get_by_role("radio",name="Електронне направлення",exact=True)
    try:radio_checked=radio.count()==1 and radio.first.is_checked()
    except ActionCancelled:raise
    except Exception:radio_checked=False
    return {
        "number":_fi_read_named_input(web,"serviceRequestNumber"),
        "created":_fi_read_named_input(web,"serviceRequestCreationDate"),
        "has_reference":any(bool(references.nth(i).input_value(timeout=2000).strip())
                            for i in range(references.count())),
        "radio_checked":radio_checked,
    }


def _fi_referral_matches(state, expected_number, expected_created):
    actual_number=re.sub(r"\D","",state.get("number") or "")
    wanted_number=re.sub(r"\D","",expected_number or "")
    return (bool(wanted_number) and actual_number==wanted_number and
            _fi_referral_date(state.get("created"))==_fi_referral_date(expected_created) and
            bool(state.get("has_reference")) and bool(state.get("radio_checked")))


def _fi_select_latest_referral(web):
    radio=web.wait_for_alias("electronic_referral")
    if not radio.is_checked():
        _fi_click(web,"electronic_referral")
        _fi_wait_condition(web,lambda:radio.is_checked(),"обрано radio «Електронне направлення»")
        web.page.wait_for_timeout(2000)
    _fi_click(web,"choose_referral")
    popup,_rows=_fi_picker_rows(web)
    headers=popup.locator('tr.dx-header-row td, thead th, [role="columnheader"]')
    names=[normalize_patient_name(headers.nth(i).inner_text(timeout=400)) for i in range(headers.count())]
    dates=[i for i,label in enumerate(names) if "дата створення" in label or "дата создания" in label]
    if len(dates)!=1:
        raise AutomationError(f"Направлення: очікувалася одна колонка «Дата створення», знайдено {len(dates)}; заголовки {names!r}.")
    date_column=dates[0]
    number_column=next((i for i,label in enumerate(names) if "номер" in label or "№" in label),None)
    if number_column is None:
        raise AutomationError("Направлення: немає колонки номера для перевірки обраного запису.")
    pager=popup.get_by_text("Сторінки:",exact=True)
    pager_root=pager.first.locator("xpath=ancestor::div[contains(@class,'dx-box-flex')][1]") if pager.count() else None
    page_numbers=[]
    if pager_root is not None:
        controls=pager_root.locator('[role="button"][aria-label]')
        page_numbers=sorted({int(label) for button in controls.all()
                             if (label:=button.get_attribute("aria-label") or "").isdigit()})
        if not page_numbers:raise AutomationError("Направлення: пагінатор є, але номери його сторінок не визначені.")
    def switch_page(number):
        if number is None:return
        active=pager_root.locator('[role="button"].dx-button-success[aria-label]')
        if active.count() and active.first.get_attribute("aria-label")==str(number):return
        button=pager_root.locator(f'[role="button"][aria-label="{number}"]')
        if button.count()!=1:raise AutomationError(f"Направлення: не знайдено сторінку {number}.")
        before=popup.locator("tr.dx-data-row").first
        before_text=before.inner_text(timeout=500) if before.count() else ""
        button.first.click(timeout=10000)
        _fi_wait_condition(web,lambda:active.count() and active.first.get_attribute("aria-label")==str(number)
                           and (not before_text or (popup.locator("tr.dx-data-row").count() and
                                popup.locator("tr.dx-data-row").first.inner_text(timeout=500)!=before_text)),
                           f"відкриття сторінки направлень {number}")
    entries=[]
    for page_number in (page_numbers or [None]):
        switch_page(page_number)
        _popup,rows=_fi_picker_rows(web)
        for index,row in enumerate(rows):
            cells=row.locator('td[role="gridcell"], td')
            if cells.count()<=date_column:
                raise AutomationError("Направлення: рядок не містить колонку «Дата створення».")
            created=cells.nth(date_column).inner_text(timeout=500).strip()
            number=(cells.nth(number_column).inner_text(timeout=500).strip()
                    if number_column is not None and cells.count()>number_column else "")
            entries.append((page_number,index,created,number,row.inner_text(timeout=500)))
    idx=newest_referral_index([entry[2] for entry in entries])
    chosen=entries[idx]
    switch_page(chosen[0])
    popup,rows=_fi_picker_rows(web)
    if chosen[1]>=len(rows) or rows[chosen[1]].inner_text(timeout=500)!=chosen[4]:
        raise AutomationError("Направлення: вибраний рядок змінився після переходу між сторінками.")
    action=_fi_row_action(web,rows[chosen[1]],["Вибрати","Выбрать"])
    if not action:raise AutomationError("Найновіше направлення не має доступної кнопки «Вибрати».")
    action.click(timeout=10000)
    rows[chosen[1]].wait_for(state="hidden",timeout=12000)
    state={}
    def selected_matches():
        nonlocal state
        try:state=_fi_selected_referral_state(web)
        except ActionCancelled:raise
        except Exception as exc:
            state={"read_error":str(exc)[:250]}
            return False
        return _fi_referral_matches(state,chosen[3],chosen[2])
    try:_fi_wait_condition(web,selected_matches,"номер і дата обраного електронного направлення у формі")
    except AutomationError as exc:
        raise AutomationError(f"Направлення не підтверджено у формі: очікувалося номер {chosen[3]!r}, "
                              f"дата {chosen[2]!r}; фактично {state!r}.") from exc
    log_event("referral.selected","Вибрано найновіше електронне направлення.",action_id=web.action_id)


def _fi_validate_template(web, start_dt, cfg):
    required={
      "interaction_start":start_dt,"episode_start":start_dt,"duration":"60",
      "interaction_class":"Амбулаторна медична допомога",
      "interaction_type":"Взаємодія в закладі охорони здоров'я",
      "priority":"Плановий","episode_type":"Реабілітація",
      "paid_service":"false","health40":"false",
    }
    expected_start=datetime.strptime(start_dt,"%d.%m.%Y %H:%M")
    for alias,expected in required.items():
        if alias=="paid_service":
            loc=None
            for candidate in web.selector_map.get(alias,[]):
                try:
                    found=web._candidate_locator(candidate)
                    if found.count() and found.first.is_visible():
                        loc=found.first
                        break
                except ActionCancelled:raise
                except Exception:continue
            if loc is None:
                # Some interaction templates omit this optional checkbox; an
                # absent control cannot represent a paid service selection.
                log_event("template.optional_field_absent",
                          "У шаблоні немає прапорця «Платна послуга»; відсутній елемент вважається вимкненим.",
                          action_id=web.action_id,data={"field":alias})
                continue
        else:
            loc=web.wait_for_alias(alias)
        actual=_fi_control_text(loc)
        if alias in {"interaction_start","episode_start"}:
            state=_fi_datetime_state(loc)
            valid=_datetime_text_matches(state.get("display"),expected_start) and (
                not state.get("isDatebox") or _datetime_text_matches(state.get("widget"),expected_start))
            actual=str(state)
        elif alias=="duration":
            number=re.search(r"\d+(?:[.,]\d+)?",actual)
            valid=bool(number and float(number.group().replace(",","."))==60)
        elif alias in {"paid_service","health40"}:
            valid=(loc.is_checked()==(expected=="true"))
        else:valid=normalize_patient_name(expected)==normalize_patient_name(actual)
        if not valid:raise AutomationError(f"Перевірка шаблону: {alias}, очікувалося {expected!r}, отримано {actual!r}.")
    reasons=[]
    for row in web.page.locator("tr.dx-data-row:visible").all():
        cells=row.locator('td[role="gridcell"]')
        if not cells.count():continue
        text=cells.first.inner_text(timeout=500).strip()
        if re.match(r"^\(?A57\)?\s*/\s*Фізична медицина\s*/\s*реабілітація\s*$",text,re.I):
            reasons.append(row)
    if len(reasons)!=1:
        raise AutomationError(f"Причина звернення: очікувався один рядок A57 / Фізична медицина / реабілітація, знайдено {len(reasons)}.")
    for alias in ("service_place","interaction_participant","treating_physician"):
        loc=web.wait_for_alias(alias);actual=_fi_control_text(loc)
        if alias=="interaction_participant":
            tags=loc.locator('xpath=ancestor::*[contains(@class,"dx-tagbox")][1]').locator('.dx-tag-content')
            actual="; ".join(tags.nth(i).inner_text(timeout=500).strip() for i in range(tags.count()))
        if not actual:raise AutomationError(f"Шаблон залишив обов'язкове поле {alias} порожнім.")
        if alias=="service_place" and cfg.get("default_place") and not _fi_place_matches(
            cfg["default_place"],actual):
            raise AutomationError(f"Шаблон: місце надання послуги, очікувалося {cfg['default_place']!r}, отримано {actual!r}.")
    expected_end=(datetime.strptime(start_dt,"%d.%m.%Y %H:%M")+timedelta(minutes=60)).strftime("%d.%m.%Y %H:%M")
    loc=web.wait_for_alias("interaction_end")
    state=_fi_datetime_state(loc)
    expected_end_dt=datetime.strptime(expected_end,"%d.%m.%Y %H:%M")
    if not _datetime_text_matches(state.get("display"),expected_end_dt) or (
        state.get("isDatebox") and not _datetime_text_matches(state.get("widget"),expected_end_dt)):
        raise AutomationError(f"Кінець взаємодії: очікувалося {expected_end!r}, отримано {state!r}.")
    log_event("template.validated","Шаблон взаємодії пройшов перевірку.",action_id=web.action_id)


def _fi_verify_patient_card(web, name, dob):
    """Check the card's patient fields; page text may contain only initials."""
    if not web._patient_card_matches(name,dob):
        raise AutomationError(f"Поля відкритої картки не підтверджують пацієнта «{name}».")
    if not dob:return
    if isinstance(dob,(datetime,date)):
        dob_date=dob.date() if isinstance(dob,datetime) else dob
    else:
        dob_date=None
        for fmt in ("%Y-%m-%d","%d.%m.%Y","%d/%m/%Y","%d-%m-%Y"):
            try:dob_date=datetime.strptime(str(dob).strip(),fmt).date();break
            except ValueError:pass
    if dob_date is None:raise AutomationError(f"Не вдалося перевірити дату народження пацієнта: {dob!r}.")
    # This field is absent for many patients. Inspect it only if it is already
    # present; never wait for it or let a card re-render abort the workflow.
    card_birth=""
    try:
        birth_fields=web.page.locator('input[name="patientBirthDate"]')
        for i in range(birth_fields.count()):
            field=birth_fields.nth(i)
            if field.is_visible():
                card_birth=field.input_value(timeout=0).strip()
                if card_birth:break
    except ActionCancelled:
        raise
    except Exception:
        card_birth=""
    if not card_birth:
        log_event("patient.birth_date_unavailable",
                  "Дата народження в картці відсутня або недоступна без очікування; пацієнта підтверджено за полями ПІБ.",
                  action_id=web.action_id)
        return
    dob_variants=(dob_date.strftime("%d.%m.%Y"),dob_date.strftime("%d/%m/%Y"),
                  dob_date.strftime("%d-%m-%Y"),dob_date.isoformat())
    if not any(value in card_birth for value in dob_variants):
        raise AutomationError(f"Дата народження відкритої картки {card_birth!r} не збігається з джерелом {dob_date.strftime('%d.%m.%Y')}.")


def first_rehab_interaction(payload):
    """Create the first rehab interaction using the selected schedule patient row."""
    action="first_rehab_interaction"
    name=require(payload,"patient_name","ПІБ пацієнта")
    cfg=settings();pkey=patient_key(name)
    prior=get_workflow(action,pkey)
    explicit_resume=bool(payload.get("_resume_from_checkpoint"))
    resume_requested=explicit_resume or bool(prior and prior.get("state")=="waiting_user")
    resume_data=(prior or {}).get("data") or {}
    if explicit_resume:
        if str(payload.get("_resume_key") or "").strip()!=pkey:
            return ActionResult(False,"Контрольна точка належить іншому пацієнту або відсутня. Запустіть функцію заново.",{},"failed")
    if resume_requested:
        if not prior or prior.get("state")!="waiting_user" or not resume_data.get("resume_stage"):
            return ActionResult(False,"Контрольна точка вже неактуальна. Запустіть функцію заново.",{},"failed")
    if prior and prior.get("state")=="complete" and not resume_requested:
        # A completed checkpoint alone is not proof that the interaction is
        # still present in Doctor Eleks. Reopen the patient and let the exact
        # date+duration check against the live interaction list decide whether
        # a retry is safe; it will stop before creating a duplicate.
        log_event("interaction.retry_after_checkpoint",
                  "Попередній запуск позначений завершеним; повторно звіряємо цільову взаємодію у Doctor Eleks.",
                  action_id=action,data={"previous_updated_at":prior.get("updated_at")})
    try:
        pdata=patient_data(name,payload)
        if normalize_patient_name(pdata.get("name"))!=normalize_patient_name(name):
            raise ValueError("ПІБ вибраного рядка таблиці не збігається повністю.")
        dt_value=pdata.get("first_interaction")
        dt=format_interaction_datetime(dt_value)
        if resume_requested and resume_data.get("target_datetime") not in (None,dt):
            raise ValueError("Дата взаємодії змінилася після зупинки; стару контрольну точку не можна продовжити.")
        diagnoses=parse_icd_codes(pdata.get("system_diagnoses"))
        if not diagnoses:raise ValueError("У колонці N не знайдено кодів МКХ.")
        status=str(pdata.get("first_interaction_status") or "").strip().casefold()
        if any(token in status for token in ("e pv reab","створено","створена","підписано","завершено")):
            raise ValueError(f"У колонці O вже вказано стан першої взаємодії ({pdata.get('first_interaction_status')}); повторний запуск зупинено.")
        log_event("source.parsed","Дані вибраного рядка перевірено.",action_id=action,
                  data={"patient":normalize_patient_name(name),"diagnosis_count":len(diagnoses)})
    except Exception as exc:
        return ActionResult(False,f"{name} | source parsed: {exc}",{},"failed")
    checkpoint_data={"source_row":pdata.get("row"),"target_datetime":dt,
                     "baseline_interactions":list(resume_data.get("baseline_interactions") or []),
                     "resume_stage":resume_data.get("resume_stage") or "initial"}
    resume_stage=checkpoint_data["resume_stage"]
    template_dialog_ready=False
    def mark_resume_stage(stage):
        nonlocal resume_stage
        resume_stage=stage
        checkpoint_data["resume_stage"]=stage
        set_workflow(action,pkey,"in_progress",dict(checkpoint_data))
    try:
        with WebEleks(action,cfg,confirmed_authenticated=resume_requested) as web:
            if resume_requested and web._patient_card_matches(name,pdata.get("dob")):
                log_event("interaction.resume.patient_present","Продовжуємо у відкритій картці пацієнта.",action_id=action)
            else:
                web.select_patient(name,pdata.get("dob"))
            if resume_requested and resume_stage=="post_template":
                # Older paused runs changed dates after selecting a template.
                # Reopen template options and redo the required preselection
                # date sequence instead of continuing that invalid state.
                resume_stage="template_button"
                mark_resume_stage(resume_stage)
            _fi_verify_patient_card(web,name,pdata.get("dob"))
            log_event("patient.verified","Пацієнта перевірено за полями ПІБ; дату народження звірено за наявності поля в картці.",action_id=action)
            baseline_interactions=checkpoint_data["baseline_interactions"]
            if not resume_requested:
                baseline_interactions=[text for _,text in _fi_table_rows(web,"interactions")]
                checkpoint_data["baseline_interactions"]=baseline_interactions
                set_workflow(action,pkey,"in_progress",dict(checkpoint_data))
            elif resume_stage=="initial" and not baseline_interactions and not (
                    _fi_template_dates_visible(web) or _fi_interaction_form_visible(web) or _fi_episode_picker_visible(web)):
                baseline_interactions=[text for _,text in _fi_table_rows(web,"interactions")]
                checkpoint_data["baseline_interactions"]=baseline_interactions
                set_workflow(action,pkey,"in_progress",dict(checkpoint_data))
            if resume_requested and resume_stage!="initial" and not (
                    _fi_template_dates_visible(web) or _fi_interaction_form_visible(web) or
                    _fi_episode_picker_visible(web) or _fi_organization_template_picker_visible(web)):
                current_interactions=[text for _,text in _fi_table_rows(web,"interactions")]
                if any(_fi_saved_interaction_matches(text,dt) for text in current_interactions):
                    set_workflow(action,pkey,"complete",{"finalized":True,"resumed_after_save":True})
                    return ActionResult(True,"1 ВЗАЄМОДІЯ вже була збережена до зупинки; повторне створення не виконувалося.",
                                        {"patient":name,"resumed":True})
                baseline_interactions=current_interactions
                checkpoint_data["baseline_interactions"]=current_interactions
                resume_stage="initial"
                checkpoint_data["resume_stage"]=resume_stage
                set_workflow(action,pkey,"in_progress",dict(checkpoint_data))
            if any(_fi_saved_interaction_matches(text,dt) for text in baseline_interactions):
                if resume_requested:
                    set_workflow(action,pkey,"complete",{"finalized":True,"resumed_after_save":True})
                    return ActionResult(True,"1 ВЗАЄМОДІЯ вже була збережена до зупинки; повторне створення не виконувалося.",
                                        {"patient":name,"resumed":True})
                raise AutomationError(f"У Doctor Eleks вже є взаємодія з початком {dt} і тривалістю 60 хвилин; повторне створення зупинено.")
            if prior and prior.get("state")=="in_progress" and not resume_requested:
                log_event("interaction.stale_attempt","Попередню незавершену спробу звірено зі списком; цільової взаємодії немає.",
                          action_id=action,data={"previous_updated_at":prior.get("updated_at")})
            if resume_stage!="post_template":
                form_open=_fi_template_dates_visible(web) or _fi_interaction_form_visible(web)
                if resume_stage=="initial" and not form_open:
                    if not _fi_episode_picker_visible(web):
                        _fi_click_new_interaction(web)
                    _fi_create_new_episode(web)
                    mark_resume_stage("template_button")
                elif resume_stage=="initial" and form_open:
                    # The previous attempt may have opened the episode form
                    # just before timing out; keep it and continue from there.
                    mark_resume_stage("template_button")

                if resume_stage=="template_button":
                    # Open the interaction-template dialog after the episode
                    # form has settled; do not act until its dates are ready.
                    if not _fi_template_dates_visible(web):
                        _fi_wait_before_template_button(web)
                        _fi_click(web,"templates")
                    mark_resume_stage("template_dialog")

                if resume_stage=="organization_picker" and not _fi_organization_template_picker_visible(web):
                    if _fi_template_dates_visible(web):
                        resume_stage="template_dialog"
                        mark_resume_stage(resume_stage)
                    else:
                        # A prior run may have clicked a template before it
                        # timed out waiting for the resulting form.
                        resume_stage="post_template"
                        mark_resume_stage(resume_stage)

                if resume_stage in {"template_dialog","template_options"}:
                    # Set both dates in the template dialog before choosing
                    # the organization template. Pause between fields so the
                    # first DevExtreme date has time to commit.
                    _fi_wait_template_date_dialog(web)
                    first_date=_fi_template_date_locator(web,"template_episode_start")
                    second_date=_fi_template_date_locator(web,"template_interaction_start")
                    if first_date is None or second_date is None:
                        raise AutomationError("Не знайдено обидва поля дати у вікні шаблонів.")
                    first_date.click(timeout=5000)
                    _fi_type_template_date(first_date,dt)
                    first_date.press("Tab",timeout=3000)
                    second_date.focus(timeout=3000)
                    deadline=__import__("time").monotonic()+2
                    while __import__("time").monotonic()<deadline:
                        raise_if_cancelled()
                        web.page.wait_for_timeout(250)
                    expected_date=datetime.strptime(dt,"%d.%m.%Y %H:%M")
                    saved_state=_fi_datetime_state(first_date,timeout=1000)
                    if not (_datetime_text_matches(saved_state.get("display"),expected_date) and
                            _datetime_text_matches(saved_state.get("widget"),expected_date)):
                        raise AutomationError(f"Перша дата не зафіксувалася після переходу до другого поля: {saved_state!r}.")
                    # Type through the second field's mask, then Tab out to
                    # commit that widget before switching template tabs.
                    _fi_type_template_date(second_date,dt)
                    second_date.press("Tab",timeout=3000)
                    _fi_wait_template_dates_saved(web,dt)
                    template_dialog_ready=True
                    mark_resume_stage("template_dates_set")

                if resume_stage=="template_dates_set":
                    if _fi_organization_template_picker_visible(web):
                        # The previous run may have switched tabs just before
                        # it stopped; resume from the visible picker.
                        mark_resume_stage("organization_picker")
                    elif _fi_template_dates_visible(web):
                        if not template_dialog_ready:
                            _fi_wait_template_date_dialog(web)
                        web.click_content_text(("Шаблони установи","Шаблоны учреждения"))
                        _fi_wait_condition(web,lambda:_fi_organization_template_picker_visible(web),
                                           "список шаблонів установи")
                        mark_resume_stage("organization_picker")
                    else:
                        raise AutomationError("Після налаштування дат не знайдено ані діалог шаблону, ані список шаблонів установи.")

                if resume_stage=="organization_picker":
                    _popup,rows=_fi_picker_rows(web)
                    matches=[]
                    for row in rows:
                        cells=row.locator('td[aria-describedby]')
                        for index in range(cells.count()):
                            try:
                                txt=cells.nth(index).inner_text(timeout=400)
                                if re.match(r"^для\s+F84\.0\s+Діти\b",txt.strip(),re.I):
                                    matches.append(row);break
                            except ActionCancelled:raise
                            except Exception:pass
                    if len(matches)!=1:raise AutomationError(f"Очікувався один шаблон «для F84.0 Діти ...», знайдено {len(matches)}.")
                    choose=_fi_row_action(web,matches[0],["Вибрати","Выбрать"])
                    if not choose:raise AutomationError("У потрібного шаблону відсутня кнопка «Вибрати».")
                    web.page.wait_for_timeout(2000)
                    choose.click(timeout=10000)
                    log_event("template.selection.clicked","Кнопку вибору шаблону натиснуто; чекаємо форму.",
                              action_id=action)
                    mark_resume_stage("post_template")
                    _fi_wait_post_template_form(web)
            if resume_stage!="post_template":
                raise AutomationError("Не вдалося визначити безпечне місце продовження після вибору шаблону.")
            _fi_wait_post_template_form(web)
            _fi_validate_template(web,dt,cfg)
            _fi_select_latest_referral(web)
            _fi_reconcile_diagnoses(web,diagnoses,dt)
            age=_rehab_age_on(pdata.get("dob"),dt_value)
            if age is None:raise AutomationError("Не вдалося визначити вік пацієнта за датою народження.")
            rows=_fi_diagnostic_report_rows(web)
            reports=[(row,text) for row,text in rows]
            codes=[_fi_report_code(text) for _,text in reports]
            try:remove_report_indexes=validate_age_specific_reports(age,codes)
            except ValueError as exc:raise AutomationError(str(exc))
            if age<18:
                for index in remove_report_indexes:
                    row,_=reports[index]
                    fresh=_fi_diagnostic_report_rows(web,activate=False)
                    target=next(((r,t) for r,t in fresh if _fi_report_code(t)=="9603700"),None)
                    if target is None:break
                    action_button=_fi_row_action(web,target[0],["Видалити","Удалить"])
                    if not action_button:raise AutomationError("Не знайдено видалення рядка 96037.")
                    old_count=sum(1 for _,t in fresh if _fi_report_code(t)=="9603700")
                    action_button.click(timeout=10000)
                    _fi_confirm_diagnostic_report_delete(web)
                    _fi_wait_condition(web,lambda:sum(1 for _,t in _fi_diagnostic_report_rows(web,activate=False) if _fi_report_code(t)=="9603700")<old_count,
                                       "видалення діагностичного звіту 96037")
                remaining=_fi_diagnostic_report_rows(web,activate=False)
                if any(_fi_report_code(t)=="9603700" for _,t in remaining):
                    raise AutomationError("Звіт 96037 залишився після видалення для дитини.")
            fresh=_fi_diagnostic_report_rows(web,activate=False)
            row,text=next(((r,t) for r,t in fresh if _fi_report_code(t)=="9602200"),(None,""))
            if row is None:raise AutomationError("Звіт 96022 зник під час перевірки.")
            edit=_fi_row_action(web,row,["Редагувати","Редактировать"])
            if not edit:raise AutomationError("Не знайдено редагування рядка 96022.")
            edit.click(timeout=10000)
            report_time=web.wait_for_alias("report_datetime")
            try:actual_report_time=report_time.input_value(timeout=500)
            except Exception:actual_report_time=report_time.inner_text(timeout=500)
            if not diagnosis_time_matches(actual_report_time,dt):
                _fi_fill(web,"report_datetime",dt);_fi_click(web,"save")
            else:_fi_click(web,"report_cancel")
            _fi_wait_extended_popup_closed(web,"закриття модалки діагностичного звіту")
            _fi_wait_condition(web,lambda:any(_fi_report_code(text)=="9602200" for _,text in _fi_diagnostic_report_rows(web,activate=False)),
                               "повернення звіту 96022 після збереження")
            refreshed=_fi_diagnostic_report_rows(web,activate=False)
            report_row=next((row for row,text in refreshed if _fi_report_code(text)=="9602200"),None)
            if report_row is None:raise AutomationError("Звіт 96022 зник після збереження.")
            verify_edit=_fi_row_action(web,report_row,["Редагувати","Редактировать"])
            if not verify_edit:raise AutomationError("Не знайдено повторне редагування 96022 для перевірки часу.")
            verify_edit.click(timeout=10000)
            report_state=_fi_datetime_state(web.wait_for_alias("report_datetime"))
            expected_report=datetime.strptime(dt,"%d.%m.%Y %H:%M")
            if not _datetime_text_matches(report_state.get("display"),expected_report) or (
                report_state.get("isDatebox") and not _datetime_text_matches(report_state.get("widget"),expected_report)):
                raise AutomationError(f"Час звіту 96022 не зберігся: очікувалося {dt}, фактично {report_state!r}.")
            _fi_click(web,"report_cancel")
            _fi_wait_extended_popup_closed(web,"закриття модалки перевірки діагностичного звіту")
            log_event("reports.validated","Діагностичні звіти перевірено.",action_id=action)
            barthel_rows=_fi_barthel_exam_rows(web)
            remove_exam_indexes=barthel_rows_to_remove(age,[text for _,text in barthel_rows])
            for _index in remove_exam_indexes:
                fresh=_fi_barthel_exam_rows(web)
                current=fresh[0] if fresh else None
                if current is None:break
                delete=_fi_row_action(web,current[0],["Видалити","Удалить"])
                if not delete:raise AutomationError("Не знайдено видалення Barthel_Extended.")
                delete.click(timeout=10000)
                _fi_confirm_barthel_delete(web)
                old_count=len(fresh)
                _fi_wait_condition(web,lambda:len(_fi_barthel_exam_rows(web,wait=False))<old_count,
                                   "видалення Barthel_Extended")
            if age<18 and _fi_barthel_exam_rows(web,wait=False):
                raise AutomationError("Barthel_Extended залишився після видалення для дитини.")
            log_event("observations.validated","Обстеження перевірено за віком.",action_id=action)
            # The current interaction form does not expose a redemption checkbox
            # in all EHR versions. The selected referral itself was already
            # verified in _fi_select_latest_referral; only toggle redemption
            # when the form actually provides that control.
            check=None
            for candidate in web.selector_map.get("redeem_referral_checkbox",[]):
                try:
                    loc=web._candidate_locator(candidate)
                    if loc.count() and loc.first.is_visible():
                        check=loc.first
                        break
                except ActionCancelled:raise
                except Exception:continue
            if check is not None:
                if not check.is_enabled():
                    check=web.wait_for_alias("redeem_referral_checkbox",enabled=True)
                if not check.is_checked():check.click(timeout=10000)
                _fi_wait_condition(web,lambda:check.is_checked(),"checkbox «Погасити направлення»")
            else:
                log_event("referral.redemption_control_not_present",
                          "Форма не містить окремого прапорця погашення; вибране електронне направлення вже перевірено.",
                          action_id=action)
            close_episode=web.page.get_by_label(re.compile("Закрити епізод|Закрыть эпизод",re.I))
            if close_episode.count() and close_episode.first.is_checked():
                raise AutomationError("«Закрити епізод» позначено; зупинка перед збереженням.")
            validation=web.page.locator('.dx-invalid-message:visible, .dx-validation-message:visible')
            if validation.count():
                raise AutomationError("Форма має помилки валідації перед збереженням: "+
                    "; ".join(validation.nth(i).inner_text(timeout=300) for i in range(min(validation.count(),5))))
            dialog_seen={"value":False}
            def accept_dialog(dialog):
                dialog.accept();dialog_seen["value"]=True
            save_responses=[]
            def on_response(response):
                try:
                    if response.request.method in {"POST","PUT","PATCH"}:
                        save_responses.append((response.status,response.url))
                except Exception:pass
            web.browser_page.once("dialog",accept_dialog)
            web.browser_page.on("response",on_response)
            try:
                if web.safe_mode:
                    raise ManualCheckpoint("safe_mode","Безпечний режим зупинив фінальне збереження взаємодії.")
                sign_control=web.wait_for_alias("sign_later",enabled=True)
                _fi_click(web,"sign_later")
                _fi_click(web,"confirm_anyway")
                deadline=__import__("time").monotonic()+12
                while __import__("time").monotonic()<deadline:
                    raise_if_cancelled()
                    errors=[status for status,_url in save_responses if status>=400]
                    if errors:raise AutomationError(f"Фінальне збереження повернуло HTTP-помилку: {errors}.")
                    if dialog_seen["value"] and not sign_control.is_visible() and not web.page.get_by_role(
                        "button",name="Ок, все одно підписати").first.is_visible():
                        _fi_table_rows(web,"interactions")
                        def saved_row_visible():
                            current=[text for _,text in _fi_table_rows(web)]
                            added=[text for text in current if text not in baseline_interactions]
                            return any(dt.split(" ")[0] in text for text in current) and (
                                any(dt.split(" ")[0] in text for text in added) or len(current)>len(baseline_interactions))
                        _fi_wait_condition(web,saved_row_visible,"нова взаємодія у списку пацієнта")
                        if not any(200<=status<300 for status,_url in save_responses):
                            raise AutomationError("Нова взаємодія показана у списку, але HTTP-збереження не підтверджено.")
                        log_event("interaction.finalized","Фінальне підтвердження завершено.",action_id=action)
                        set_workflow(action,pkey,"complete",{"finalized":True})
                        return ActionResult(True,"1 ВЗАЄМОДІЯ створена та збережена.",{"patient":name,"age":age})
                    web.page.wait_for_timeout(250)
                raise AutomationError("Після фінального підтвердження форма залишилася відкритою або операція не завершилася.")
            finally:
                web.browser_page.remove_listener("response",on_response)
    except ManualCheckpoint as cp:
        checkpoint_data.update({"checkpoint":cp.code,"last_error":str(cp)[:500]})
        set_workflow(action,pkey,"waiting_user",checkpoint_data)
        return waiting(f"{name} | {cp} Виправте сторінку Doctor Eleks і натисніть «Продовжити». Відновлення продовжиться з етапу «{resume_stage}».",
                       cp.code,{**cp.data,"can_continue":True,"resume_key":pkey,"resume_stage":resume_stage})
    except ActionCancelled as exc:
        checkpoint_data.update({"checkpoint":"cancelled","last_error":str(exc)[:500]})
        set_workflow(action,pkey,"waiting_user",checkpoint_data)
        return waiting(f"{name} | автоматизацію зупинено на етапі «{resume_stage}». Виправте сторінку Doctor Eleks і натисніть «Продовжити».",
                       "cancelled",{"can_continue":True,"resume_key":pkey,"resume_stage":resume_stage})
    except Exception as exc:
        checkpoint_data.update({"checkpoint":"automation_error","last_error":str(exc)[:500]})
        set_workflow(action,pkey,"waiting_user",checkpoint_data)
        log_event("first_interaction.await_continue",f"{name} | сценарій зупинено для ручного виправлення: {exc}",
                  level="WARNING",action_id=action,data={"resume_stage":resume_stage})
        return waiting(f"{name} | автоматизація зупинилася на етапі «{resume_stage}»: {exc}. Виправте сторінку Doctor Eleks і натисніть «Продовжити».",
                       "automation_error",{"can_continue":True,"resume_key":pkey,"resume_stage":resume_stage})

def care_plan_draft(payload):
    action="care_plan_draft";name=require(payload,"patient_name","ПІБ пацієнта");cfg=settings();pdata=patient_data(name,payload)
    try:
        with WebEleks(action,cfg) as web:
            web.select_patient(name,pdata.get("dob"));web.click("interactions",optional=True)
            if "реабіліта" not in web.current_text().casefold():
                raise ManualCheckpoint("rehab_episode","Відкрийте реабілітаційну взаємодію; план лікування можна створювати тільки на ній.")
            web.click("actions");web.click("create_plan");web.click("save")
        try:update_coordinator_status(coordinator_path(),name,"Можемо кидати",cfg.get("coordinator_sheet",""))
        except Exception as e:log_event("coordinator.status.warning",f"План збережено, але статус координатора не оновлено: {e}",level="WARNING",action_id=action)
        return ActionResult(True,"Чернетку плану лікування збережено без заявки; статус «Можемо кидати» виставлено, якщо таблиця координаторів доступна.",{})
    except ManualCheckpoint as cp:return waiting(str(cp),cp.code,cp.data)
