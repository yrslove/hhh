from __future__ import annotations
import re, time
from datetime import datetime, date, timedelta
from collections import Counter
from pathlib import Path
from ..core.result import ActionResult,waiting
from ..core.audit import log_event
from ..core.paths import RUNTIME_DIR
from ..core.database import get_workflow,set_workflow
from ..core.action_control import ActionCancelled, raise_if_cancelled, set_progress, set_progress_count
from ..core.text import normalize_patient_name
from ..automation.browser import WebEleks,ManualCheckpoint,AutomationError,wait_visible_recheck
from ..excel.schedule import parse_schedule
from ..excel.workbook import WorkbookChangedError, get_graph_patient, load_procedure_reference
from .common import graph_path,settings,patient_key,require
from .procedure_form import DoctorEleksProcedureForm

# A code in column R is followed by its service type, e.g.
# ``96128/ФТ`` or ``96.112/аЕТ``.  The type must not be guessed from a
# specialist's name elsewhere in the string.
CODE_RX=re.compile(r"(?<!\d)(\d{2,3}(?:[.\s-]?\d{2,3})?)(?!\d)\s*(?:/|\\)\s*([A-Za-zА-Яа-яІіЇїЄє]+)")
REF_DIGITS=re.compile(r"(?:\d[\s\-]?){16}")
_TYPE_CONFUSABLES=str.maketrans({"a":"а","e":"е","f":"ф","t":"т"})

def _procedure_browser_settings():
    # The shared Chrome may contain Docs/Sheets tabs that Playwright tries to
    # initialize before it can reach Doctor Eleks. Seed an isolated profile
    # from Valeriia's signed-in profile, without its open tabs.
    return {**settings(),"remote_debug_port":0,
            "browser_profile_dir":str(RUNTIME_DIR/"procedure_valeriia_profile"),
            "browser_profile_seed_dir":str(RUNTIME_DIR/"browser_profile")}

def _completion_log(items,skipped=()):
    """Human-readable result for the operator; the structured plan remains in data."""
    if not items and not skipped:
        return "Нових процедур не внесено. Усі заплановані записи вже підтверджені у локальному стані."
    lines=[f"Внесено процедур: {len(items)}"]
    for item in items:
        day=datetime.fromisoformat(item["date"]).strftime("%d.%m.%Y")
        lines.append(f"• {day} {item['time']} — {item['specialist']}")
    if skipped:
        lines.append(f"Пропущено: {len(skipped)} — процедури вже знайдені в Doctor Eleks.")
        for item in skipped:
            day=datetime.fromisoformat(item["date"]).strftime("%d.%m.%Y")
            lines.append(f"• {day} {item['time']} — {item['specialist']}")
    return "\n".join(lines)

def _canonical_service_type(value):
    """Return the workflow key for the only supported trailing type labels."""
    label=str(value or "").strip().casefold().translate(_TYPE_CONFUSABLES)
    if label in {"фт","афт"}:return "ft"
    # аЕТ is an assistant-ET label; it uses the same ET service code.
    if label in {"ет","ает"}:return "et"
    return ""

def _codes_by_type(text):
    out={}
    for code,label in CODE_RX.findall(str(text or "")):
        typ=_canonical_service_type(label)
        if typ:out[typ]=re.sub(r"\s+","",code)
    s=str(text or "")
    if "96.128" in s:out.setdefault("ft","96.128")
    if "96.113" in s:out.setdefault("et","96.113")
    return out

def _profile_meta(specialist,cfg):
    profiles=cfg.get("specialists",{}) or {}
    for name,meta in profiles.items():
        if normalize_patient_name(name)==normalize_patient_name(specialist):
            return meta or {}
    return {}

def _specialist_type(specialist,cfg):
    # The final token is the operator's explicit choice: "Пальчик ФТ",
    # "… ЕТ" or "… аЕТ".  It has priority over a potentially stale profile
    # setting and over words occurring earlier in the name.
    match=re.search(r"([A-Za-zА-Яа-яІіЇїЄє]+)\s*$",str(specialist or ""))
    if match:
        typ=_canonical_service_type(match.group(1))
        if typ:return typ
    return ""

def _service_code5(value):
    digits=re.sub(r"\D","",str(value or ""))
    return digits[:5] if len(digits)>=5 else ""

def _service_codes_in_row(row_text):
    # Service values in Doctor Eleks look like 96112-00 or 96.112-00.
    found=re.findall(r"(?<!\d)(\d{2}[.\s-]?\d{3})(?!\d)",str(row_text or ""))
    return list(dict.fromkeys(filter(None,(_service_code5(value) for value in found))))

def _plan(payload):
    set_progress(10,"Зчитуємо графік процедур і дані пацієнта")
    name=require(payload,"patient_name","ПІБ пацієнта");path=graph_path(payload);cfg=settings()
    sheet=payload.get("sheet") or cfg.get("etft_sheet","ЕТФТ")
    distribution_sheet=cfg.get("distribution_sheet","Розподіл AR2")
    before=Path(path).stat()
    read_started=time.perf_counter()
    wb=load_procedure_reference(path,sheet,distribution_sheet)
    try:
        opened_at=time.perf_counter()
        appts,warns=parse_schedule(wb,sheet,int(cfg.get("schedule_year",2026)))
        parsed_at=time.perf_counter()
        pdata=get_graph_patient(path,distribution_sheet,name,workbook=wb)
        patient_at=time.perf_counter()
        after=Path(path).stat()
        if (before.st_size,before.st_mtime_ns)!=(after.st_size,after.st_mtime_ns):
            raise WorkbookChangedError("Excel-файл змінився під час читання графіка")
        log_event("excel.procedure_plan_timing","Зчитано план процедур з Excel.",action_id="procedures_create",
                  data={"open_seconds":round(opened_at-read_started,3),
                        "schedule_seconds":round(parsed_at-opened_at,3),
                        "patient_seconds":round(patient_at-parsed_at,3),
                        "appointments":len(appts)})
        key=normalize_patient_name(name)
        all_patient=[a for a in appts if a.patient_name_normalized==key]
        if not all_patient:raise ValueError(f"На вкладці {sheet} не знайдено записів пацієнта «{name}».")
        specialist_input=(payload.get("specialist") or "").strip()
        if specialist_input:
            spec_key=normalize_patient_name(specialist_input)
            selected=[a for a in all_patient if normalize_patient_name(a.specialist)==spec_key]
            if not selected:
                variants=sorted({a.specialist for a in all_patient})
                raise ValueError(f"Не знайдено записів спеціаліста «{specialist_input}». Доступні: {', '.join(variants)}")
            specialist=selected[0].specialist
        else:
            variants=sorted({a.specialist for a in all_patient if a.specialist})
            if len(variants)!=1:
                raise ValueError("У пацієнта є записи від кількох спеціалістів. Вкажіть конкретного спеціаліста: "+", ".join(variants))
            specialist=variants[0];selected=all_patient
        selected.sort(key=lambda a:(a.date,a.start_time))
        codes=_codes_by_type(pdata.get("codes_specialists"))
        typ=_specialist_type(specialist,cfg);code=codes.get(typ,"")
        plan=[{"date":a.date,"time":a.start_time,"specialist":specialist,"type":typ,"service_code":code,
               "sheet":a.sheet_name,"cell":a.cell_address} for a in selected]
        return path,pdata,plan,[w.as_dict() for w in warns],specialist
    finally:
        wb.close()

def _row_action_button(row):
    # In the current Doctor Eleks table «Дії» is often a clickable text/span,
    # not a semantic button.  Resolve it by exact text inside the matched row
    # and choose the rightmost visible match.
    exact=[]
    for label in ("Дії","Действия"):
        try:
            matches=row.get_by_text(label,exact=True)
            for index in range(matches.count()):
                item=matches.nth(index)
                if not item.is_visible():continue
                box=item.bounding_box()
                exact.append((box.get("x",0) if box else 0,item))
        except Exception:pass
    if exact:return max(exact,key=lambda item:item[0])[1]
    controls=row.locator("button, a, [role='button'], [ng-click], [data-ng-click], [onclick]")
    for index in range(controls.count()-1,-1,-1):
        try:
            control=controls.nth(index)
            text=control.inner_text(timeout=300).strip().casefold()
            if control.is_visible() and (text.startswith("дії") or text.startswith("действия")):
                return control
        except Exception:pass
    return None

def _confirm_referral_request(web):
    """Acknowledge Doctor Eleks' notification after taking a referral."""
    try:
        popup=wait_visible_recheck(web.page,lambda: web.page.locator(".popup.-notification"),
                             initial_ms=2000,retry_ms=3000)
        message=popup.locator(".popup_text").inner_text(timeout=2000).strip()
        if "Заявка на взяття направлення в роботу успішно створена" not in message:
            raise AutomationError(f"Неочікуване повідомлення після взяття направлення: {message}")
        button=wait_visible_recheck(web.page,lambda: web.page.locator(".popup.-notification").locator(
            '.popup_buttons button[click-next-action="true"]'),
            initial_ms=1800,retry_ms=1800)
        if button.inner_text(timeout=1000).strip().upper() not in ("ОК","OK"):
            raise AutomationError("У повідомленні про направлення немає кнопки «ОК».")
        button.click(timeout=5000)
        popup.wait_for(state="hidden",timeout=5000)
    except AutomationError:
        raise
    except Exception as exc:
        raise AutomationError("Не вдалося підтвердити повідомлення про взяття направлення кнопкою «ОК».") from exc
    log_event("procedure.referral.confirmed","Підтверджено повідомлення про взяття направлення кнопкою «ОК».",action_id=web.action_id)

def _return_to_medical_records(web):
    """Wait until the queue redirect completes before leaving that page."""
    try:
        web.page.wait_for_function("""() =>
          location.hash.toLowerCase().includes('action=spzibwebservicerequestschedulelist') &&
          [...document.querySelectorAll('li.side-list-item.active')].some(
            item => /Черга направлень|Очередь направлений/i.test(item.innerText))
        """,timeout=20000)
    except Exception as exc:
        raise ManualCheckpoint("referral_queue_navigation",
          "Після підтвердження направлення Doctor Eleks не завершив перехід до черги направлень.",
          {"url":web.page.url}) from exc
    # The route can appear before its delayed requests finish. Give the
    # queue a short settling window, then verify the destination as well.
    web.page.wait_for_timeout(1500)
    web.goto_medical_records()
    log_event("procedure.referral.returned_to_records",
              "Після завершення переходу до черги відкрито «Медичні записи» та перевірено сторінку.",
              action_id=web.action_id)

def _referral_service_name(row,code):
    try:
        service=row.locator('td[aria-colindex="4"]').first.inner_text(timeout=1000).strip()
    except Exception as exc:
        raise ManualCheckpoint("referral_service","Не вдалося прочитати назву послуги у вибраному направленні.") from exc
    match=re.match(r"^\s*\d{2}[.\s-]?\d{3}(?:-\d{2})?\s+(.+)$",service,re.S)
    if not match or _service_code5(service)!=code:
        raise ManualCheckpoint("referral_service",f"Код або назву послуги у рядку направлення {code} не вдалося підтвердити.",
                               {"service":service,"code":code})
    return match.group(1).strip()

def _take_referral_to_work(web,code):
    """Open patient referrals, match code in the first four rows and take it to work."""
    target=_service_code5(code)
    if not target:raise ValueError(f"Некоректний код послуги: {code}")
    web.click_content_text(("Направлення","Направления"),timeout=2500,retry_click=True)
    web.page.wait_for_timeout(500)
    service_rows=[];matched=[]
    for attempt in range(2):
        rows=web.page.locator("tr");service_rows=[]
        for index in range(rows.count()):
            raise_if_cancelled()
            try:
                row=rows.nth(index)
                if hasattr(row,"is_visible") and not row.is_visible():continue
                text=row.inner_text(timeout=300)
                codes=_service_codes_in_row(text)
                if not codes:continue
                service_rows.append((row,text,codes))
                if len(service_rows)>=4:break
            except ActionCancelled:raise
            except Exception:pass
        matched=[item for item in service_rows if target in item[2]]
        if len(matched)==1:break
        if attempt==0:web.page.wait_for_timeout(2000)
    if len(matched)!=1:
        seen=[{"codes":codes,"text":text[:240]} for _,text,codes in service_rows]
        if not matched:
            raise ManualCheckpoint("referral_code",
              f"У перших {len(service_rows)} направленнях не знайдено код послуги {target}.",
              {"target_code":target,"rows":seen})
        raise ManualCheckpoint("referral_ambiguous",
          f"У перших чотирьох направленнях знайдено кілька рядків з кодом {target}.",
          {"target_code":target,"matches":len(matched)})
    row,text,_=matched[0]
    service_name=_referral_service_name(row,target)
    numbers=[]
    for match in REF_DIGITS.findall(text):
        digits=re.sub(r"\D","",match)
        if len(digits)==16:numbers.append(digits)
    numbers=list(dict.fromkeys(numbers))
    if len(numbers)!=1:
        raise ManualCheckpoint("referral_number",f"У рядку послуги {target} немає однозначного 16-значного номера направлення.",
                               {"numbers":numbers,"service":service_name})
    if any(marker in normalize_patient_name(text) for marker in
           ("в роботі","в работе","взято в роботу")):
        log_event("procedure.referral.already_taken",
                  f"Направлення з кодом {target} уже в роботі.",action_id=web.action_id,
                  data={"code":target,"referral_number":numbers[0],"service":service_name})
        return {"number":numbers[0],"service":service_name,"code":target}
    def current_referral_action():
        rows=web.page.locator("tr")
        for index in range(rows.count()):
            current=rows.nth(index)
            try:
                current_text=current.inner_text(timeout=300)
                if target in _service_codes_in_row(current_text) and any(
                    re.sub(r"\D","",value)==numbers[0] for value in REF_DIGITS.findall(current_text)):
                    return _row_action_button(current)
            except Exception:pass
        return None
    try:
        action=wait_visible_recheck(web.page,current_referral_action,
                                    initial_ms=1800,retry_ms=1800)
    except Exception as exc:
        raise ManualCheckpoint("referral_actions",
            f"Для направлення з кодом {target} не знайдено кнопку «Дії» у правому краї рядка.") from exc
    action.click(timeout=10000)
    web.page.wait_for_timeout(2000)
    web.click_content_text(("Взяти у роботу","Взяти в роботу","Взять в работу"),timeout=2500)
    _confirm_referral_request(web)
    # The app redirects to the referral queue after OK. Wait for that
    # transition, then return through the medical-records sidebar item.
    _return_to_medical_records(web)
    log_event("procedure.referral.taken",f"Направлення з кодом {target} взято в роботу.",action_id=web.action_id,
              data={"code":target,"referral_number":numbers[0],"service":service_name})
    return {"number":numbers[0],"service":service_name,"code":target}

def _click_create_procedure(web):
    try:
        label=wait_visible_recheck(web.page,
            lambda:web.page.get_by_text("Створити процедуру:",exact=True).first,
            initial_ms=1800,retry_ms=1800)
        add=wait_visible_recheck(web.page,lambda:label.locator(
            "xpath=ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' dx-box-flex ')][1]")
            .locator(':scope > .dx-box-item [role="button"][aria-label="add"]').first,
            initial_ms=1800,retry_ms=1800)
        add.click(timeout=3000)
    except Exception as exc:
        raise AutomationError("Не знайдено кнопку «+» біля «Створити процедуру».") from exc
    web.page.wait_for_timeout(2000)

def _procedure_row_matches(text,item,code,specialist):
    expected_date=item["date"]
    expected_time=item["time"]
    expected_code=_service_code5(code)
    parts=normalize_patient_name(specialist).split()
    if not parts:return False
    surname=parts[0]
    return (expected_date in text and expected_time in text and
            expected_code in _service_codes_in_row(text) and
            surname in normalize_patient_name(text))

def _visible_procedure_rows(web):
    matches=[]
    for row in web.page.locator("tr.dx-data-row").all():
        try:
            if not row.is_visible():continue
            matches.append(row.inner_text(timeout=500)[:500])
        except ActionCancelled:raise
        except Exception:pass
    return matches

def _matching_visible_procedure_rows(web,item,code,specialist):
    return [text for text in _visible_procedure_rows(web)
            if _procedure_row_matches(text,item,code,specialist)]

def _scan_procedure_grid_pages(web,read_visible):
    """Collect a page-specific result once from every visible pager page."""
    matches=list(read_visible(web))
    if not hasattr(web.page,"get_by_text"):
        return matches
    pager_label=web.page.get_by_text("Сторінки:",exact=True)
    if not pager_label.count():
        return matches
    pager=pager_label.first.locator("xpath=ancestor::div[contains(@class,'dx-box-flex')][1]")
    buttons=pager.locator('[role="button"][aria-label]')
    pages=sorted({int(label) for button in buttons.all()
                  if (label:=button.get_attribute("aria-label") or "").isdigit()})
    active=pager.locator('[role="button"].dx-button-success[aria-label]')
    original=active.first.get_attribute("aria-label") if active.count() else None
    if not pages or original is None:
        raise ManualCheckpoint("procedure_grid_pager",
            "Не вдалося визначити поточну сторінку списку процедур перед перевіркою дубля.")
    def switch(number):
        current=active.first.get_attribute("aria-label") if active.count() else None
        if current==str(number):return
        old_row=web.page.locator("tr.dx-data-row").first
        old_text=old_row.inner_text(timeout=1000) if old_row.count() else ""
        pager.locator(f'[role="button"][aria-label="{number}"]').click(timeout=3000)
        web.page.wait_for_function("""([wanted, oldText]) => {
            const active=document.querySelector('[role="button"].dx-button-success[aria-label="'+wanted+'"]');
            const row=document.querySelector('tr.dx-data-row');
            return !!active && !!row && !!row.innerText.trim() &&
                   (!oldText || row.innerText!==oldText);
        }""",arg=[str(number),old_text],timeout=12000)

    try:
        for number in pages:
            raise_if_cancelled()
            if str(number)==original:continue
            switch(number)
            matches.extend(read_visible(web))
    finally:
        if active.count() and active.first.get_attribute("aria-label")!=original:
            switch(original)
    return matches

def _existing_procedure_candidates(web,item,code,specialist):
    """Search all pages only when retrying a possibly interrupted creation."""
    return _scan_procedure_grid_pages(
        web,lambda current:_matching_visible_procedure_rows(current,item,code,specialist))

def _verify_batch_procedure_rows(web,plan,code,specialist):
    """Make one read-only grid pass after the complete batch."""
    rows=_scan_procedure_grid_pages(web,_visible_procedure_rows)
    problems=[]
    for item in plan:
        matches=[row for row in rows if _procedure_row_matches(row,item,code,specialist)]
        if len(matches)!=1:
            problems.append({"cell":item["cell"],"date":item["date"],
                             "time":item["time"],"count":len(matches)})
    if problems:
        raise ManualCheckpoint("procedure_batch_verification",
          "Після завершення пакета не всі процедури однозначно знайдено у списку Doctor Eleks.",
          {"problems":problems,"checked":len(plan),"grid_rows":len(rows)})
    log_event("procedures.batch_verified",
              f"Після завершення пакета у списку Doctor Eleks підтверджено {len(plan)} процедур.",
              action_id=web.action_id,data={"checked":len(plan),"grid_rows":len(rows)})
    return len(plan)

def _needs_retry_duplicate_check(state,item,is_first_pending):
    """A full grid search is needed for a possibly interrupted creation.

    New targets have never been opened by this workflow and can go straight to
    the create button. Older workflow states predate the per-cell attempt flag,
    so check their first unfinished target once when resuming.
    """
    data=(state or {}).get("data") or {}
    if data.get("attempted_cell")==item["cell"]:
        return True
    if data.get("attempt_tracking_version")==1:
        return False
    return bool(is_first_pending and (state or {}).get("state") not in (None,"new","complete"))

def _wait_created_procedure_row(web,item,code,specialist,timeout_ms=15000,expected_id=None):
    deadline=time.monotonic()+timeout_ms/1000
    # A freshly saved row is usually on the current grid page. First wait for
    # that exact server-assigned ID without moving the pager. Only search all
    # pages if the grid has not displayed it after a short refresh window.
    visible_deadline=min(deadline,time.monotonic()+3)
    while time.monotonic()<visible_deadline:
        raise_if_cancelled()
        visible=_matching_visible_procedure_rows(web,item,code,specialist)
        exact=[row for row in visible if not expected_id or
               str(expected_id).casefold() in row.casefold()]
        if len(exact)==1:return exact[0]
        if len(exact)>1:
            raise ManualCheckpoint("procedure_duplicate_rows",
              "Після збереження знайдено кілька процедур з однаковим ID.",{"rows":exact})
        web.page.wait_for_timeout(250)
    while time.monotonic()<deadline:
        raise_if_cancelled()
        matches=_existing_procedure_candidates(web,item,code,specialist)
        if expected_id:
            matches=[row for row in matches if str(expected_id).casefold() in row.casefold()]
        if len(matches)==1:return matches[0]
        if len(matches)>1:
            raise ManualCheckpoint("procedure_duplicate_rows",
              "Після збереження знайдено кілька процедур з однаковими датою, часом, кодом і виконавцем.",
              {"rows":matches})
        web.page.wait_for_timeout(250)
    raise ManualCheckpoint("procedure_final_row",
      "Сервер підтвердив збереження, але нова процедура не з'явилася у списку.",
      {"date":item["date"],"time":item["time"],"code":code})

def _select_matching_referral(web,referral):
    """Open the referral picker, validate its rows and choose the newest exact match."""
    adapter=getattr(web,"_procedure_form_adapter",None)
    popup=web.page.locator(".popup.-extended")
    if not (popup.count() and popup.first.is_visible()):
        electronic=wait_visible_recheck(web.page,
            lambda:web.page.get_by_text("Електронне направлення",exact=True).first,
            initial_ms=1800,retry_ms=1800)
        sequence=adapter.latest_sequence if adapter else 0
        if adapter:adapter.last_action="SELECT_ELECTRONIC_REFERRAL"
        electronic.click(timeout=3000)
        if adapter:adapter.wait_optional_autosave(sequence)
        try:
            choose=wait_visible_recheck(web.page,
                lambda:web.page.get_by_text("Обрати направлення",exact=True).first,
                initial_ms=2000,retry_ms=6000)
            sequence=adapter.latest_sequence if adapter else 0
            choose.click(timeout=3000)
            if adapter:adapter.wait_optional_autosave(sequence)
        except (ManualCheckpoint,ActionCancelled):
            raise
        except Exception as exc:
            raise ManualCheckpoint("referral_picker_open",
              "Після вибору «Електронне направлення» не з'явилася кнопка «Обрати направлення».") from exc
    # Picker implementations vary: do not require a particular popup wrapper.
    # Wait for its own «Вибрати» action inside a visible data row.
    deadline=time.monotonic()+45
    rows=None
    while time.monotonic()<deadline:
        raise_if_cancelled()
        candidates=web.page.locator("tr.dx-row.dx-data-row")
        for index in range(candidates.count()):
            row=candidates.nth(index)
            try:
                if not row.is_visible():continue
                select=row.locator('[role="button"][aria-label="Вибрати"]')
                if select.count() and select.first.is_visible():
                    rows=candidates
                    break
            except ActionCancelled:raise
            except Exception:pass
        if rows is not None:break
        web.page.wait_for_timeout(250)
    if rows is None:
        raise ManualCheckpoint("referral_modal_rows",
          "Вікно вибору направлення не показало табличні рядки з кнопкою «Вибрати» протягом 45 секунд.")
    matches=[];seen=[]
    for index in range(rows.count()):
        raise_if_cancelled()
        row=rows.nth(index)
        if not row.is_visible():continue
        cells=row.locator('td[role="gridcell"]')
        if cells.count()<7:continue
        date_text=cells.nth(0).inner_text(timeout=500).strip()
        service=cells.nth(1).inner_text(timeout=500).strip()
        number=cells.nth(3).inner_text(timeout=500).strip()
        seen.append({"date":date_text,"service":service,"number":number})
        if re.sub(r"\D","",number)!=re.sub(r"\D","",referral["number"]):continue
        if normalize_patient_name(service)!=normalize_patient_name(referral["service"]):continue
        try:date_key=datetime.fromisoformat(date_text.split()[0])
        except ValueError:date_key=datetime.min
        matches.append((date_key,-index,row))
    if not matches:
        raise ManualCheckpoint("referral_match",
          "У вікні вибору не знайдено направлення з тим самим номером і назвою послуги.",
          {"expected":referral,"visible_rows":seen[:20]})
    _,_,row=max(matches,key=lambda item:(item[0],item[1]))
    select=row.locator('[role="button"][aria-label="Вибрати"]').first
    select=wait_visible_recheck(web.page,
        lambda:row.locator('[role="button"][aria-label="Вибрати"]').first,
        initial_ms=2000,retry_ms=8000)
    sequence=adapter.latest_sequence if adapter else 0
    select.click(timeout=10000)
    if adapter:adapter.wait_optional_autosave(sequence)
    try:
        popup.first.wait_for(state="hidden",timeout=10000)
        web.page.wait_for_function("""() =>
          location.hash.toLowerCase().includes('idname=procedureref') &&
          !!document.querySelector('input[id$="_procedurePeriod"]')
        """,timeout=20000)
    except Exception as exc:
        raise ManualCheckpoint("referral_picker_close",
          "Після вибору направлення форма процедури не відкрилася заново.") from exc
    if adapter:
        adapter.guard()
        state=adapter.wait_ready()
        if state.get("formData",{}).get("selectedReferralCode")!="digitalReferral":
            raise ManualCheckpoint("referral_state","Направлення не потрапило в модель процедури.",state)
    log_event("procedure.referral.selected","У процедурі вибрано направлення за номером і назвою послуги.",
              action_id=web.action_id,data={"number":referral["number"],"service":referral["service"]})
    if adapter:adapter.guard()
    return {"number":referral["number"],"service":referral["service"],"verified":True}

def _datetime_text_matches(value,expected):
    text=str(value or "")
    formats=("%d.%m.%Y %H:%M","%d.%m.%Y, %H:%M","%Y-%m-%d %H:%M",
             "%Y-%m-%d, %H:%M","%Y-%m-%dT%H:%M","%d/%m/%Y %H:%M")
    for fmt in formats:
        try:
            if datetime.strptime(text.strip(),fmt)==expected:return True
        except ValueError:pass
    match=re.search(r"(\d{1,4})[./-](\d{1,2})[./-](\d{1,4}).*?(\d{1,2}):(\d{2})",text)
    if not match:return False
    a,b,c,h,m=map(int,match.groups())
    if a>31:year,month,day=a,b,c
    else:day,month,year=a,b,c
    if year<1000:return False
    try:return datetime(year,month,day,h,m)==expected
    except ValueError:return False

def _procedure_field_state(web,suffix):
    """Read what the operator sees and what Doctor Eleks will submit."""
    field=web.page.locator(f'input[id$="_{suffix}"]').first
    return field.evaluate("""(el, name) => {
      const host=el.closest('.dx-datebox,.dx-selectbox,.dx-textbox');
      const ui=window.DevExpress && window.DevExpress.ui;
      const widgetName=host?.classList.contains('dx-datebox')?'dxDateBox':
        host?.classList.contains('dx-selectbox')?'dxSelectBox':'dxTextBox';
      const widget=host && ui?.[widgetName]?.getInstance(host);
      const pad=n=>String(n).padStart(2,'0');
      const value=v=>v instanceof Date
        ? `${v.getFullYear()}-${pad(v.getMonth()+1)}-${pad(v.getDate())} ${pad(v.getHours())}:${pad(v.getMinutes())}`
        : v==null ? '' : String(v);
      const scope=window.angular && window.angular.element(host || el).scope();
      const hidden=host?.querySelector(`input[type="hidden"][name="${name}"]`) ||
        document.querySelector(`input[type="hidden"][name="${name}"]`);
      return {display:el.value,hidden:value(hidden?.value),
        model:value(scope?.formVm?.formData?.[name]),
        widget:value(widget?.option('value'))};
    }""",suffix)

def _procedure_datetime_is_bound(web,suffix,expected):
    state=_procedure_field_state(web,suffix)
    return all(_datetime_text_matches(state.get(part),expected)
               for part in ("display","hidden","model","widget")),state

def _matches_procedure_division(text):
    organization="міжнароднийцентрреабілітаціїтанейрофізіології"
    compact=re.sub(r"[^\w]","",normalize_patient_name(text),flags=re.UNICODE)
    if not compact.endswith("2") or organization not in compact or len(compact)<5:
        return False
    prefix=(compact[0] in "тt" and compact[1] in "оo" and compact[2] in "вbv")
    return prefix and compact[3]=="2"

def _verify_procedure_form_ready(web,item,referral,selected_referral):
    start=datetime.fromisoformat(item["date"]+"T"+item["time"])
    end=start+timedelta(minutes=30)
    start_ok,start_state=_procedure_datetime_is_bound(web,"dateStart",start)
    end_ok,end_state=_procedure_datetime_is_bound(web,"dateEnd",end)
    period=_procedure_field_state(web,"procedurePeriod")
    division=_procedure_field_state(web,"division")
    fields=web.page.locator('input[id$="_dateStart"]').first.evaluate("""el => {
      const data=window.angular?.element(el).scope()?.formVm?.formData || {};
      const keys=['serviceCode','report_originsCode','procedureCategoryRef','procedureStatus',
        'performerRef','procedureOutcome','division','isDateOrPeriod'];
      const result=Object.fromEntries(keys.map(key=>[key,data[key] == null ? '' : String(data[key])]));
      result.performedDate=data.date instanceof Date ?
        `${data.date.getFullYear()}-${String(data.date.getMonth()+1).padStart(2,'0')}-${String(data.date.getDate()).padStart(2,'0')} ${String(data.date.getHours()).padStart(2,'0')}:${String(data.date.getMinutes()).padStart(2,'0')}` :
        String(data.date ?? '');
      return result;
    }""")
    validation=web.page.locator('input[id$="_dateStart"]').first.evaluate("""el => {
      const form=window.angular?.element(el).scope()?.formInstance;
      if(!form || typeof form.validate!=='function')return {isValid:false,errors:['Form validator unavailable']};
      const result=form.validate();
      return {isValid:!!result.isValid,
        errors:(result.brokenRules||[]).map(rule=>rule.message||rule.type||'Invalid').slice(0,10)};
    }""")
    # The CRM form has no serviceRequestNumber/serviceRequestServiceName fields.
    # The picker verifies the number and service before selecting the row;
    # the read-only service field on the form is named serviceCode.
    number_ok=(bool(selected_referral.get("verified")) and
      re.sub(r"\D","",selected_referral.get("number",""))==
      re.sub(r"\D","",referral["number"]))
    service_ok=(normalize_patient_name(selected_referral.get("service",""))==
      normalize_patient_name(referral["service"]) and
      normalize_patient_name(referral["service"]) in
      normalize_patient_name(fields.get("serviceCode","")))
    period_ok=bool(period.get("model")) and all(period.get(part)==period.get("model")
                                                for part in ("widget","hidden"))
    division_ok=bool(division.get("model")) and division.get("model")==fields.get("division")
    performed_ok=_datetime_text_matches(fields.get("performedDate"),start)
    required_ok=all(fields.get(key) for key in
      ("serviceCode","report_originsCode","procedureCategoryRef","procedureStatus",
       "performerRef","isDateOrPeriod"))
    if not all((start_ok,end_ok,period_ok,division_ok,performed_ok,number_ok,service_ok,required_ok,
                validation.get("isValid"))):
        raise ManualCheckpoint("procedure_form_data",
          "Дані процедури не збережені у внутрішній моделі форми; збереження зупинено.",
          {"start":start_state,"end":end_state,"period":period,"division":division,
           "form_data":fields,"validation":validation,
           "checks":{"start":start_ok,"end":end_ok,"period":period_ok,
           "division":division_ok,"performed_date":performed_ok,"referral_number":number_ok,
           "referral_service":service_ok,
           "required":required_ok}})

def _procedure_autosave_preflight(adapter):
    """Compare the stock period save against every live form layer."""
    adapter.guard()
    state=adapter.read_state()
    save=adapter.saves[-1] if adapter.saves else None
    payload=save.get("payload",{}) if save else {}
    if not state.get("ready") or not save or not 200<=(save.get("status") or 0)<300:
        raise ManualCheckpoint("procedure_preflight","Немає підтвердженого успішного autosave.",
                               {"form_state":state,"save":save})
    angular=state.get("formData",{}); dx=state.get("dxFormData",{})
    editors=state.get("editors",{})
    def require(condition,message):
        if not condition:
            raise ManualCheckpoint("procedure_preflight",message,
              {"target":adapter.target,"form_state":state,"save_payload":payload})
    require(all(str(value)=="30" for value in (
        angular.get("procedurePeriod"),dx.get("procedurePeriod"),
        editors.get("procedurePeriod",{}).get("value"),payload.get("procedurePeriod"))),
        "Тривалість 30 хвилин не збігається між editor, dxForm, Angular і payload.")
    require(all(str(value).casefold()=="true" for value in (
        angular.get("isDateOrPeriod"),dx.get("isDateOrPeriod"),payload.get("is_date_or_period"))),
        "Режим періоду не підтверджено в усіх шарах.")
    require(angular.get("selectedReferralCode")==payload.get("selectedReferralRef")=="digitalReferral",
            "Електронне направлення не збігається з payload.")
    for key in ("based_on","code","performer","division"):
        require(_payload_present(payload.get(key)),f"У payload відсутнє {key}.")
    require(payload.get("performer")==angular.get("performerRef"),"Виконавець не збігається.")
    require(payload.get("division")==angular.get("division")==dx.get("division")==
            editors.get("division",{}).get("value"),"Підрозділ не збігається.")
    require(adapter.target["code"] in str(angular.get("serviceCode") or ""),
            "Код послуги не відповідає цільовій процедурі.")
    start=datetime.fromisoformat(adapter.target["date"]+"T"+adapter.target["start_time"])
    end=start+timedelta(minutes=30)
    for key,expected,payload_key in (("dateStart",start,"procedurePeriodStart"),
                                     ("dateEnd",end,"procedurePeriodEnd"),
                                     ("date",start,"performed_date_time")):
        model=angular.get(key); dx_value=dx.get(key)
        require(_payload_present(model) and model==dx_value,
                f"{key} не збігається між Angular і dxForm.")
        if key=="date":
            require(model==angular.get("dateStart"),"Дата виконання не збігається з початком.")
        else:
            editor=editors.get(key,{})
            require(model==editor.get("value") and
                    _datetime_text_matches(editor.get("display"),expected),
                    f"{key} у редакторі не відповідає графіку.")
        require(_procedure_datetime_matches(payload.get(payload_key),expected),
                f"{payload_key} у payload не відповідає графіку.")
    require(str(payload.get("showErrorMessage"))=="0",
            "Autosave тривалості має виконуватися без фінальної валідації.")
    adapter._log("PREFLIGHT_SUCCESS",{"form_state":state,"save_payload":payload,
                                      "target":adapter.target})

def _payload_present(value):
    return value is not None and str(value).strip().casefold() not in {"", "null", "undefined"}

def _procedure_datetime_matches(value,expected):
    if not _payload_present(value):return False
    text=str(value).strip()
    try:
        parsed=datetime.fromisoformat(text.replace("Z","+00:00"))
        if not parsed.tzinfo:
            return parsed.replace(second=0,microsecond=0)==expected
    except ValueError:pass
    try:
        parsed=datetime.strptime(text[:24],"%a %b %d %Y %H:%M:%S")
        return parsed.replace(second=0)==expected
    except ValueError:pass
    return _datetime_text_matches(text,expected)

def _build_procedure_target(name,referral,code,specialist,item):
    start=datetime.fromisoformat(item["date"]+"T"+item["time"])
    return {"patient":name,"referral":referral,"code":code,"specialist":specialist,
            "date":item["date"],"start_time":item["time"],"duration_minutes":30,
            "end_time":(start+timedelta(minutes=30)).strftime("%H:%M"),
            "division":'ТОВ 2 "МІЖНАРОДНИЙ ЦЕНТР РЕАБІЛІТАЦІЇ ТА НЕЙРОФІЗІОЛОГІЇ" 2'}

def _sign_procedure_later(web):
    deadline=time.monotonic()+10000/1000
    # Use the exact interactive DevExtreme element from the form's DOM tree,
    # not its child text span (which can be duplicated or receive no action).
    sign_buttons=web.page.locator(
      'div.backColor.dx-field-item > .dx-field-item-content > '
      'div.dx-button[role="button"][aria-label="Підписати пізніше"]')
    if sign_buttons.count()==0:
        sign_buttons=web.page.locator(
          'div.dx-button[role="button"][aria-label="Підписати пізніше"]')
    sign_button=None
    target_chain=[]
    while time.monotonic()<deadline and sign_button is None:
        raise_if_cancelled()
        for index in range(sign_buttons.count()):
            candidate=sign_buttons.nth(index)
            try:
                if (candidate.is_visible() and candidate.is_enabled() and
                    normalize_patient_name(candidate.get_attribute("aria-label") or "")==
                    normalize_patient_name("Підписати пізніше") and
                    normalize_patient_name(candidate.inner_text(timeout=300))==
                    normalize_patient_name("Підписати пізніше")):
                    sign_button=candidate;break
            except Exception:pass
        if sign_button is None:web.page.wait_for_timeout(250)
    if sign_button is None:
        raise ManualCheckpoint("procedure_save","Не знайдено кнопку «Підписати пізніше».")
    try:
        target_chain=sign_button.evaluate("""el => {
          const out=[];
          for(let node=el; node && out.length<10; node=node.parentElement) {
            out.push({tag:node.tagName.toLowerCase(), classes:String(node.className||''),
              role:node.getAttribute('role'), aria_label:node.getAttribute('aria-label'),
              text:(node.innerText||'').trim().slice(0,120)});
          }
          return out;
        }""")
    except Exception:target_chain=[]
    native_dialogs=[]
    save_responses=[]
    procedure_form=getattr(web,"_procedure_form_adapter",None)
    autosave_sequence=max((entry["sequence"] for entry in procedure_form.saves),default=0) if procedure_form else 0
    record_id=(procedure_form.read_state().get("formData",{}).get("procedureRef")
               if procedure_form else None)
    if procedure_form and not _payload_present(record_id):
        raise ManualCheckpoint("procedure_final_save","У форми немає поточного ID процедури.")
    def on_dialog(dialog):
        detail={"type":dialog.type,"message":dialog.message[:500]}
        native_dialogs.append(detail)
        try:
            dialog.accept()
            detail["accepted"]=True
        except Exception as exc:
            detail["error"]=str(exc)[:200]
    def on_response(response):
        try:
            method=response.request.method
            if method in ("POST","PUT","PATCH") and len(save_responses)<20:
                save_responses.append({"method":method,"status":response.status,
                                       "endpoint":response.url.split("?",1)[0].rsplit("/",1)[-1]})
        except Exception:pass
    web.browser_page.on("dialog",on_dialog)
    web.browser_page.on("response",on_response)
    def stop_observing():
        try:web.browser_page.remove_listener("dialog",on_dialog)
        except Exception:pass
        try:web.browser_page.remove_listener("response",on_response)
        except Exception:pass
    sign_button.scroll_into_view_if_needed(timeout=3000)
    sign_button.evaluate("""el => {
      const probe={pointerdown:0,click:0,trustedPointer:false,trustedClick:false};
      window.__procedureSignClickProbe=probe;
      el.addEventListener('pointerdown',event=>{
        probe.pointerdown++;probe.trustedPointer=event.isTrusted;
      },{capture:true,once:true});
      el.addEventListener('click',event=>{
        probe.click++;probe.trustedClick=event.isTrusted;
      },{capture:true,once:true});
    }""")
    try:
        sign_button.click(timeout=5000)
    except Exception:
        stop_observing()
        raise
    if procedure_form:
        # The final action is allowed only after the complete form passed
        # preflight, and is confirmed by the server response it causes.
        procedure_form.last_action="FINAL_SAVE_REQUEST"
        procedure_form._log("FINAL_SAVE_REQUEST",{"form_state":procedure_form._try_read_state()})
        try:
            saved=procedure_form.wait_final_save(autosave_sequence,timeout_ms=25000,
                                                 expected_id=record_id)
        except Exception:
            stop_observing()
            raise
        if saved.get("status",0)<200 or saved.get("status",0)>=300 or saved.get("response") is None:
            procedure_form.guard()
            raise ManualCheckpoint("procedure_final_save","Фінальне збереження процедури не підтверджено сервером.",
                                   {"save":saved})
        final_payload=saved.get("payload",{})
        start=datetime.fromisoformat(procedure_form.target["date"]+"T"+
                                     procedure_form.target["start_time"])
        end=start+timedelta(minutes=30)
        final_ok=(all(_payload_present(final_payload.get(key)) for key in
                      ("based_on","code","performer","division")) and
                  final_payload.get("selectedReferralRef")=="digitalReferral" and
                  str(final_payload.get("is_date_or_period")).casefold()=="true" and
                  str(final_payload.get("showErrorMessage"))=="1" and
                  _procedure_datetime_matches(final_payload.get("procedurePeriodStart"),start) and
                  _procedure_datetime_matches(final_payload.get("procedurePeriodEnd"),end) and
                  _procedure_datetime_matches(final_payload.get("performed_date_time"),start))
        if not final_ok:
            stop_observing()
            raise ManualCheckpoint("procedure_final_payload",
              "Сервер відповів 2xx, але фінальний payload не збігається з цільовою процедурою.",
              {"target":procedure_form.target,"save_payload":final_payload})
        procedure_form._final_save_confirmed=True
        procedure_form._log("FINAL_SAVE_SUCCESS",{"status":saved["status"],"response":saved["response"]})
    try:click_probe=web.page.evaluate("() => window.__procedureSignClickProbe || null")
    except Exception:click_probe=None
    if click_probe and (not click_probe.get("click") or not click_probe.get("trustedClick")):
        stop_observing()
        raise ManualCheckpoint("procedure_save_click",
          "Клік не дійшов до кнопки «Підписати пізніше».",
          {"click_probe":click_probe,"target_parent_chain":target_chain})
    log_event("procedure.sign_later.clicked","Виконано клік «Підписати пізніше»; очікуємо підтвердження.",
              action_id=web.action_id,data={"click_probe":click_probe,"target_parent_chain":target_chain,
                                          "native_dialogs":native_dialogs})
    web.page.wait_for_timeout(2000)
    if native_dialogs:
        try:
            sign_button.wait_for(state="hidden",timeout=5000)
            stop_observing()
            log_event("procedure.confirmed","Браузерний діалог підтверджено, форму процедури закрито.",
                      action_id=web.action_id,data={"native_dialogs":native_dialogs,
                                                     "save_responses":save_responses})
            return record_id
        except Exception:pass
    modal_selectors=(".popup.-notification", ".dx-overlay-wrapper", ".dx-overlay-content",
                     "[role='dialog']", ".modal")
    observed_modals=[]
    def find_ok_button(timeout_ms):
        deadline=time.monotonic()+timeout_ms/1000
        while time.monotonic()<deadline:
            raise_if_cancelled()
            for root in (web.page,web.browser_page):
                for selector in modal_selectors:
                    modals=root.locator(selector)
                    for modal_index in range(modals.count()):
                        modal=modals.nth(modal_index)
                        try:
                            if not modal.is_visible():continue
                            modal_text=modal.inner_text(timeout=300).strip()
                            if modal_text and modal_text[:500] not in observed_modals:
                                observed_modals.append(modal_text[:500])
                            controls=modal.locator('button, [role="button"], .dx-button, span.dx-button-text')
                            for control_index in range(controls.count()):
                                control=controls.nth(control_index)
                                try:
                                    if (control.is_visible() and normalize_patient_name(control.inner_text(timeout=300))
                                        in (normalize_patient_name("ОК"),"ok")):
                                        return control
                                except Exception:pass
                        except Exception:pass
                # Some builds put the confirmation directly in the page without
                # a modal wrapper. Only accept an exact, visible OK control.
                controls=root.locator('button, [role="button"], .dx-button, span.dx-button-text')
                for index in range(controls.count()):
                    control=controls.nth(index)
                    try:
                        if (control.is_visible() and normalize_patient_name(control.inner_text(timeout=300))
                            in (normalize_patient_name("ОК"),"ok")):
                            return control
                    except Exception:pass
            web.page.wait_for_timeout(250)
        return None

    ok_button=find_ok_button(15000)
    if ok_button is None:
        form_fields={}
        for suffix in ("dateStart","dateEnd","procedurePeriod","division"):
            try:
                form_fields[suffix]=_procedure_field_state(web,suffix)
            except Exception:form_fields[suffix]=None
        try:
            validation=web.page.evaluate("""() => [...document.querySelectorAll('.dx-invalid-message,.dx-validation-message')]
              .filter(el=>el.getBoundingClientRect().width>0).map(el=>(el.innerText||'').trim()).filter(Boolean).slice(0,10)""")
        except Exception:validation=[]
        log_event("procedure.sign_later.unconfirmed",
                  "Після кліку по «Підписати пізніше» не з'явилася кнопка «ОК».",
                  level="WARNING",action_id=web.action_id,
                  data={"form_fields":form_fields,"validation":validation,"click_probe":click_probe,
                        "visible_modals":observed_modals[:10],"native_dialogs":native_dialogs,
                        "save_responses":save_responses,"target_parent_chain":target_chain})
        stop_observing()
        raise ManualCheckpoint("procedure_confirmation",
          "Після натискання «Підписати пізніше» не знайдено видиму кнопку «ОК».",
          {"form_fields":form_fields,"validation":validation,"click_probe":click_probe,
           "visible_modals":observed_modals[:10],"native_dialogs":native_dialogs,
           "save_responses":save_responses,"sign_button_parent_chain":target_chain})
    try:
        ok_button.click(timeout=5000)
        sign_button.wait_for(state="hidden",timeout=10000)
    except Exception as exc:
        stop_observing()
        raise ManualCheckpoint("procedure_confirmation","Знайдено кнопку «ОК», але натискання не спрацювало.",
                               {"visible_modals":observed_modals[:10],
                                "native_dialogs":native_dialogs,"save_responses":save_responses,
                                "sign_button_parent_chain":target_chain}) from exc
    stop_observing()
    log_event("procedure.confirmed","Процедуру підтверджено: «Підписати пізніше» та «ОК»; форму закрито.",
              action_id=web.action_id,data={"native_dialogs":native_dialogs,"save_responses":save_responses})
    return record_id

def _confirm_procedure(web,adapter,item,code,specialist):
    """Confirm the server save and return to the patient's procedure list.

    The final spiprocedureSave is the server write. Its 2xx result, validated
    payload and the subsequent return to the patient's procedure list confirm
    completion. The new row may be on another grid page; searching all pages
    after every save only changes the UI and delays the next creation.
    """
    record_id=_sign_procedure_later(web)
    adapter.guard()
    if not record_id or not getattr(adapter,"_final_save_confirmed",False):
        raise ManualCheckpoint("procedure_final_save",
          "Фінальне збереження процедури не підтверджено сервером.")
    try:
        web.page.wait_for_function("""() =>
          location.hash.toLowerCase().includes('idname=patientref') &&
          !document.querySelector('input[id$="_isDateOrPeriod"]') &&
          !!document.querySelector('extendedgrid[id^="proceduregrid"] tr.dx-data-row')
        """,timeout=15000)
    except Exception as exc:
        raise ManualCheckpoint("procedure_final_navigation",
          "Сервер зберіг процедуру, але список процедур пацієнта не відкрився.",
          {"record_id":record_id}) from exc
    visible=_matching_visible_procedure_rows(web,item,code,specialist)
    row=next((text for text in visible if record_id.casefold() in text.casefold()),None)
    log_event("procedure.final.server_confirmed",
              "Фінальний spiprocedureSave підтверджено сервером; форму закрито, список процедур відкрито.",
              action_id=web.action_id,data={"record_id":record_id,"visible_on_current_page":bool(row)})
    adapter.close()
    return row or record_id

def _check_profile(web,specialist,cfg):
    meta=_profile_meta(specialist,cfg)
    marker=str(meta.get("profile_text","")).strip()
    if marker and marker.casefold() not in web.current_text().casefold():
        raise ManualCheckpoint("specialist_profile",
          f"Потрібен профіль спеціаліста «{specialist}». Поточна сторінка не містить контрольний текст «{marker}». Увійдіть потрібним КЕП і повторіть.")
    if not marker:
        log_event("procedures.profile.unverified","Профіль спеціаліста не перевірено автоматично: profile_text не налаштовано.",
                  level="WARNING",action_id=web.action_id,data={"specialist":specialist})

def _resume_snapshot(path,pdata,plan,warnings,specialist,pkey):
    """Store the expensive Excel result before a browser/login checkpoint."""
    dob=pdata.get("dob")
    if isinstance(dob,(datetime,date)):dob=dob.isoformat()
    elif dob is not None:dob=str(dob)
    return {"patient_key":pkey.split("|",1)[0],"path":path,"patient_dob":dob,
            "plan":plan,"warnings":warnings,"specialist":specialist}

def _resume_procedure_plan(payload,name):
    """Restore an Excel plan saved at the browser-ready checkpoint."""
    if not payload.get("_resume_from_checkpoint"):
        return None
    pkey=str(payload.get("_resume_key") or "").strip()
    if not pkey:
        raise ValueError("Немає збереженої контрольної точки для продовження.")
    state=get_workflow("procedures_create",pkey)
    data=(state or {}).get("data") or {}; snapshot=data.get("resume") or {}
    if not state or state.get("state") not in {"browser_ready","login"}:
        raise ValueError("Контрольна точка вже неактуальна. Запустіть функцію заново.")
    if snapshot.get("patient_key")!=patient_key(name):
        raise ValueError("Контрольна точка належить іншому пацієнту.")
    path=str(snapshot.get("path") or "")
    plan=snapshot.get("plan")
    if not path or not isinstance(plan,list) or not plan:
        raise ValueError("Збережений план процедур неповний. Запустіть функцію заново.")
    return path,{"dob":snapshot.get("patient_dob")},plan,snapshot.get("warnings") or [],str(snapshot.get("specialist") or ""),pkey,state

def create_procedures(payload):
    action="procedures_create";name=require(payload,"patient_name","ПІБ пацієнта")
    restored=_resume_procedure_plan(payload,name);cfg=_procedure_browser_settings()
    if restored:
        path,pdata,plan,warnings,specialist,pkey,st=restored
        log_event("procedures.resume","Продовжено з контрольної точки Doctor Eleks без повторного читання Excel.",
                  action_id=action,data={"specialist":specialist,"remaining":len(plan)})
        set_progress(25,"Продовжуємо з відкритого Doctor Eleks")
    else:
        path,pdata,plan,warnings,specialist=_plan(payload)
        pkey=patient_key(name)+"|"+normalize_patient_name(specialist)
        st=get_workflow(action,pkey) or {"state":"new","data":{}}
    if len(plan)!=14 and not payload.get("allow_nonstandard_count",False):
        return ActionResult(False,
          f"Для «{specialist}» знайдено {len(plan)} записів, а для Розподіл AR2 очікується 14. Перевірте графік або дозвольте нестандартну кількість.",
          {"planned":plan,"warnings":warnings,"count":len(plan),"specialist":specialist})
    code=_service_code5(plan[0]["service_code"] if plan else "")
    if not code:
        return ActionResult(False,"Не вдалося підібрати код з колонки R за фінальним типом спеціаліста. Використовуйте суфікс ФТ, ЕТ або аЕТ та пару на кшталт 96128/ФТ.",
                            {"specialist":specialist,"resolved_type":plan[0].get("type","") if plan else "",
                             "raw":str(pdata.get("codes_specialists") or "")})
    completed=set((st.get("data") or {}).get("completed_cells",[]));todo=[x for x in plan if x["cell"] not in completed]
    attempted_cell=str((st.get("data") or {}).get("attempted_cell") or "")
    if not todo:
        # Local workflow state only records what an earlier run attempted. It
        # cannot prove that the records still exist in Doctor Eleks, or that
        # the intended patient is currently open. Reconcile it against the
        # server-side procedure list before reporting success.
        try:
            with WebEleks(action,cfg) as web:
                web.select_patient(name,pdata.get("dob"));_check_profile(web,specialist,cfg)
                web.click_content_text(("Процедури","Процедуры"),timeout=2500,retry_click=True)
                web.page.locator('extendedgrid[id^="proceduregrid"]').first.wait_for(
                    state="visible",timeout=15000)
                rows=_scan_procedure_grid_pages(web,_visible_procedure_rows)
                problems=[]
                for item in plan:
                    matches=[row for row in rows if _procedure_row_matches(row,item,code,specialist)]
                    if len(matches)!=1:
                        problems.append({"cell":item["cell"],"date":item["date"],
                                         "time":item["time"],"count":len(matches)})
                ambiguous=[problem for problem in problems if problem["count"]>1]
                if ambiguous:
                    raise ManualCheckpoint("procedure_batch_verification",
                      "У Doctor Eleks знайдено неоднозначні дублікати процедур; автоматичне продовження зупинено.",
                      {"problems":ambiguous,"checked":len(plan),"grid_rows":len(rows)})
                missing_cells={problem["cell"] for problem in problems}
                if not missing_cells:
                    log_event("procedures.batch_verified",
                              f"Повторно у Doctor Eleks підтверджено всі {len(plan)} процедур.",
                              action_id=web.action_id,data={"checked":len(plan),"grid_rows":len(rows),
                                                           "patient":name,"specialist":specialist})
                else:
                    completed.difference_update(missing_cells)
                    todo=[item for item in plan if item["cell"] in missing_cells]
                    set_workflow(action,pkey,"in_progress",{
                        **(st.get("data") or {}),"completed_cells":sorted(completed),
                        "total":len(plan),"specialist":specialist})
                    log_event("procedures.local_state_reconciled",
                              f"У Doctor Eleks не знайдено {len(missing_cells)} локально позначених процедур; буде повторено внесення.",
                              level="WARNING",action_id=web.action_id,
                              data={"missing":problems,"patient":name,"specialist":specialist})
        except ManualCheckpoint as cp:
            set_workflow(action,pkey,cp.code,{**(st.get("data") or {}),"checkpoint":cp.code})
            return waiting(str(cp),cp.code,{"completed":len(completed),"total":len(plan),
                                            "specialist":specialist,**cp.data})
        if not todo:
            set_workflow(action,pkey,"complete",{"completed_cells":sorted(completed),
                                                   "total":len(plan),"specialist":specialist})
            return ActionResult(True,f"Усі {len(plan)} процедур від «{specialist}» знайдені у Doctor Eleks.",
                                {"count":len(plan),"specialist":specialist,
                                 "human_log":_completion_log([]),"verified_in_doctor_eleks":True})
    log_event("procedures.plan",f"Підготовлено {len(plan)} процедур; залишилось {len(todo)}.",action_id=action,
              data={"count":len(plan),"remaining":len(todo),"specialist":specialist,"code":code})
    set_progress_count(len(completed),len(plan),f"Підготовлено {len(todo)} процедур")
    resume=_resume_snapshot(path,pdata,plan,warnings,specialist,pkey)
    def state_data(checkpoint=""):
        data={"completed_cells":sorted(completed),"total":len(plan),"specialist":specialist,
              "resume":resume,"attempt_tracking_version":1}
        if attempted_cell:data["attempted_cell"]=attempted_cell
        if checkpoint:data["checkpoint"]=checkpoint
        return data
    try:
        created_now=[];skipped_now=[]
        with WebEleks(action,cfg,confirmed_authenticated=bool(restored)) as web:
            web.select_patient(name,pdata.get("dob"));_check_profile(web,specialist,cfg)
            referral=_take_referral_to_work(web,code)
            for item in todo:
                raise_if_cancelled()
                # Immutable target is complete before opening the form. It is
                # retained in logs/checkpoints throughout this procedure.
                procedure_target=_build_procedure_target(name,referral,code,specialist,item)
                web.click_content_text(("Процедури","Процедуры"),timeout=2500,retry_click=True)
                web.page.wait_for_timeout(2000)
                if _needs_retry_duplicate_check(st,item,item is todo[0]) or attempted_cell==item["cell"]:
                    existing=_existing_procedure_candidates(web,item,code,specialist)
                    if existing:
                        if len(existing)!=1:
                            raise ManualCheckpoint("procedure_duplicate_ambiguous",
                              "Знайдено кілька процедур з датою, часом, кодом і виконавцем; пропуск неоднозначний.",
                              {"target":procedure_target,"existing_rows":existing})
                        # A unique server row for the interrupted target means
                        # it was already saved. Record it as completed and
                        # continue with the next schedule cell instead of
                        # reopening the same creation form.
                        completed.add(item["cell"]);attempted_cell=""
                        skipped_now.append(item)
                        set_workflow(action,pkey,"in_progress",state_data())
                        set_progress_count(len(completed),len(plan),
                                           f"Пропущено наявну процедуру; перевірено {len(completed)} з {len(plan)}")
                        log_event("procedure.skipped_existing",
                          f"Процедуру {item['date']} {item['time']} уже знайдено; перехід до наступної.",
                          action_id=web.action_id,
                          data={"cell":item["cell"],"date":item["date"],"time":item["time"],
                                "code":code,"specialist":specialist,"existing_rows":existing})
                        continue
                attempted_cell=item["cell"]
                set_workflow(action,pkey,"in_progress",state_data())
                _click_create_procedure(web)
                procedure_form=DoctorEleksProcedureForm(web,procedure_target)
                web._procedure_form_adapter=procedure_form
                procedure_form.wait_ready()
                if payload.get("capture_manual_procedure"):
                    procedure_form._log("DIAGNOSTIC_MODE",{
                      "message":"Save capture is active while automation continues through the referral and editor fields."})
                selected_referral=_select_matching_referral(web,referral)
                procedure_form.guard()
                divisions=procedure_form.selectbox_options("division") or []
                matched=[entry for entry in divisions if _matches_procedure_division(entry.get("name",""))]
                if len(matched)!=1 or not matched[0].get("code"):
                    raise ManualCheckpoint("procedure_division_options",
                      "Не знайдено однозначний код підрозділу в завантаженому списку.",
                      {"matches":matched,"option_count":len(divisions)})
                start=datetime.fromisoformat(item["date"]+"T"+item["time"])
                end=start+timedelta(minutes=30)
                procedure_form.set_complete_state({
                    "date":start.isoformat(timespec="seconds"),
                    "dateStart":start.isoformat(timespec="seconds"),
                    "dateEnd":end.isoformat(timespec="seconds"),
                    "division":matched[0]["code"],
                    "isDateOrPeriod":"true",
                })
                procedure_form.guard()
                _verify_procedure_form_ready(web,item,referral,selected_referral)
                _procedure_autosave_preflight(procedure_form)
                _confirm_procedure(web,procedure_form,item,code,specialist)
                completed.add(item["cell"])
                attempted_cell=""
                created_now.append(item)
                set_workflow(action,pkey,"in_progress",state_data())
                set_progress_count(len(completed),len(plan),f"Внесено {len(completed)} з {len(plan)} процедур")
                log_event("procedure.created",f"Створено процедуру {item['date']} {item['time']}",action_id=action,
                          data={"cell":item["cell"],"code":code,"specialist":specialist,"done":len(completed),"total":len(plan)})
            _verify_batch_procedure_rows(web,plan,code,specialist)
            set_workflow(action,pkey,"complete",{"completed_cells":sorted(completed),"total":len(plan),"specialist":specialist})
            human_log=_completion_log(created_now,skipped_now)
            log_event("procedures.summary",human_log,action_id=action,
                      data={"created_count":len(created_now),"skipped_existing_count":len(skipped_now),
                            "specialist":specialist,"procedures":created_now,"skipped_existing":skipped_now})
            return ActionResult(True,f"Оброблено {len(completed)}/{len(plan)} процедур від «{specialist}»: "
                                     f"створено {len(created_now)}, пропущено вже створених {len(skipped_now)}.",
                                {"count":len(plan),"specialist":specialist,"warnings":warnings,
                                 "created_procedures":created_now,"skipped_existing":skipped_now,
                                 "human_log":human_log})
    except ManualCheckpoint as cp:
        if 'procedure_form' in locals():
            try:procedure_form.close()
            except Exception:pass
        set_workflow(action,pkey,cp.code,state_data(cp.code))
        created_now=locals().get("created_now",[])
        can_continue=cp.code in {"browser_ready","login","ehealth_frame"}
        if can_continue:
            log_event("procedures.await_continue","Очікується «Продовжити» після авторизації у Doctor Eleks.",
                      action_id=action,data={"checkpoint":cp.code,"specialist":specialist})
        return waiting(str(cp),cp.code,{"completed":len(completed),"total":len(plan),"specialist":specialist,
                                        "created_procedures":created_now,
                                        "skipped_existing":locals().get("skipped_now",[]),
                                        "human_log":_completion_log(created_now,locals().get("skipped_now",[])),
                                        "can_continue":can_continue,"resume_key":pkey if can_continue else "",**cp.data})
    except Exception:
        if 'procedure_form' in locals():
            try:procedure_form.close()
            except Exception:pass
        raise

def batch_sign(payload):
    action="procedures_batch_sign";name=require(payload,"patient_name","ПІБ пацієнта");cfg=_procedure_browser_settings();pkey=patient_key(name)
    st=get_workflow(action,pkey)
    if st and st["state"] in {"awaiting_kep","safe_mode_final"}:
        set_workflow(action,pkey,"complete",{"manual_signature_confirmed":True})
        return ActionResult(True,"Пакетний підпис прийнято як завершений після повторного запуску.",{"resumed":True})
    try:
        with WebEleks(action,cfg) as web:
            web.require_authenticated();web.click("batch_processing")
            web.click("record_type",optional=True);web.choose_text("Процедура")
            rows=web.page.locator("tr");selected=0
            for i in range(rows.count()):
                raise_if_cancelled()
                try:
                    row=rows.nth(i);txt=row.inner_text(timeout=300)
                    if normalize_patient_name(name) in normalize_patient_name(txt):
                        box=row.locator('input[type="checkbox"]')
                        if box.count():box.first.check();selected+=1
                except:pass
            if selected==0:
                raise ManualCheckpoint("batch_selection","Не вдалося автоматично визначити непідписані процедури пацієнта. Оберіть їх вручну і повторіть.")
            try:
                web.click("sign_selected")
                set_workflow(action,pkey,"awaiting_kep",{"selected":selected})
                return waiting(f"Обрано {selected} записів. Накладіть КЕП і повторіть функцію.","kep",{"selected":selected})
            except ManualCheckpoint as cp:
                if cp.code=="safe_mode":
                    set_workflow(action,pkey,"safe_mode_final",{"selected":selected})
                    return waiting(f"Безпечний режим не натиснув пакетний підпис. Перевірте {selected} вибраних записів, натисніть підпис вручну, накладіть КЕП і повторіть.","safe_mode_final",{"selected":selected})
                return waiting(str(cp),cp.code,cp.data)
    except ManualCheckpoint as cp:return waiting(str(cp),cp.code,cp.data)
