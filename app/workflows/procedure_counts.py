from __future__ import annotations

import re
from collections import defaultdict

from openpyxl import load_workbook

from ..core.action_control import raise_if_cancelled, set_progress
from ..core.audit import log_event
from ..core.database import save_check_run
from ..core.result import ActionResult
from ..core.text import display_name, normalize_patient_name
from ..excel.schedule import _sheet_name_match, parse_schedule
from .common import graph_path, settings

MONTH_PREFIXES = ("січ", "лют", "бер", "квіт", "трав", "черв", "лип",
                  "серп", "вер", "жовт", "лист", "груд")


def _roster_month(value):
    text = str(value or "").strip().casefold()
    for number, prefix in enumerate(MONTH_PREFIXES, 1):
        if text.startswith(prefix):
            return f"{number:02d}"
    match = re.fullmatch(r"\.?\s*(0?[1-9]|1[0-2])\.?", text)
    return f"{int(match.group(1)):02d}" if match else ""


def _name_signature(name: str) -> tuple[str, tuple[str, ...]]:
    """Match abbreviated names in column C only when the result is unique."""
    words = re.findall(r"[^\W\d_]+", normalize_patient_name(name), re.UNICODE)
    return (words[0], tuple(word[0] for word in words[1:3])) if len(words) >= 2 else ("", ())


def _name_pair(name: str) -> tuple[str, ...]:
    words = re.findall(r"[^\W\d_]+", normalize_patient_name(name), re.UNICODE)
    return tuple(words[:2]) if len(words) >= 2 else ()


def _loose_pair(pair: tuple[str, ...]) -> tuple[str, ...]:
    # The graph occasionally uses «е» where the roster uses «є».
    return tuple(part.replace("є", "е") for part in pair)


def _package_schedule_name(name: str) -> str:
    """Return the name part of an ET/FT label ending in package number 25."""
    normalized = normalize_patient_name(name)
    return re.sub(r"\s+25(?:\s+пакет)?$", "", normalized).strip() if re.search(
        r"\s+25(?:\s+пакет)?$", normalized) else ""


def _location(appointment):
    return {
        "sheet": appointment.sheet_name,
        "cell": appointment.cell_address,
        "continuation_cell": appointment.continuation_cell_address or "",
        "time": f"{appointment.start_time}–{appointment.end_time}",
        "specialist": appointment.specialist,
    }


def _status_key(value):
    return re.sub(r"\s+", "", str(value or "").casefold())


def _patient_names(sheet, start_row=3, expected_days=14, daily_target=2):
    names = {}
    header_row = start_row - 1
    status_col = next((cell.column for cell in sheet[header_row]
                       if _status_key(cell.value) == "статуспл"), None)
    if status_col is None:
        raise ValueError(f"На вкладці «{sheet.title}» не знайдено колонку «Статус ПЛ».")
    for row in sheet.iter_rows(min_row=start_row, min_col=2, max_col=max(11, status_col)):
        raise_if_cancelled()
        if _status_key(row[status_col - 2].value) != "є+направлення":
            continue
        value = row[1].value
        if not isinstance(value, str):
            continue
        name = display_name(value)
        key = normalize_patient_name(name)
        if not key or key in {"піп", "піб", "пацієнт"}:
            continue
        group = normalize_patient_name(row[9].value) if row[9].value else ""
        entry = names.setdefault(key, {"patient": name, "source_cell": row[1].coordinate,
                                       "source_sheet": sheet.title, "group": group,
                                       "expected_days": expected_days, "daily_target": daily_target,
                                       "roster_months": set()})
        roster_month = _roster_month(row[0].value)
        if roster_month:
            entry["roster_months"].add(roster_month)
    return names


def count_procedures(payload):
    """Check combined daily counts, optionally within one calendar month."""
    action = "procedure_counts"
    path = graph_path(payload)
    cfg = settings()
    year = int(cfg.get("schedule_year", 2026))
    month = str(payload.get("month") or "").strip()
    if month and (not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", month)):
        raise ValueError("Місяць перевірки потрібно вибрати у форматі РРРР-ММ.")
    distribution = payload.get("distribution_sheet") or cfg.get("distribution_sheet", "Розподіл АР2")
    package = payload.get("package_sheet") or "25 пакет"
    etft = payload.get("etft_sheet") or cfg.get("etft_sheet", "ЕТФТ")
    psych = payload.get("psych_sheet") or "ПСИХЛого"
    set_progress(10, "Зчитуємо пацієнтів і два графіки")
    wb = load_workbook(path, data_only=False, rich_text=True)
    try:
        actual = {label: _sheet_name_match(wb, wanted)
                  for label, wanted in (("distribution", distribution), ("package", package), ("etft", etft), ("psych", psych))}
        missing = [wanted for label, wanted in (("distribution", distribution), ("etft", etft), ("psych", psych))
                   if not actual[label]]
        if missing:
            raise ValueError(f"Не знайдено вкладки: {', '.join(missing)}. Доступні: {', '.join(wb.sheetnames)}")
        patients = _patient_names(wb[actual["distribution"]])
        package_patients = (_patient_names(wb[actual["package"]], start_row=2,
                                           expected_days=21, daily_target=1)
                            if actual["package"] else {})
        package_roster_pairs = defaultdict(set)
        package_roster_loose_pairs = defaultdict(set)
        if actual["package"]:
            for row in wb[actual["package"]].iter_rows(min_row=2, min_col=3, max_col=3):
                if isinstance(row[0].value, str):
                    pair = _name_pair(row[0].value)
                    if pair:
                        roster_name = normalize_patient_name(row[0].value)
                        package_roster_pairs[pair].add(roster_name)
                        package_roster_loose_pairs[_loose_pair(pair)].add(roster_name)
        patients.update(package_patients)
        if not patients:
            raise ValueError("У стовпці C вкладок розподілу не знайдено пацієнтів.")
        et_records, et_warnings = parse_schedule(wb, actual["etft"], year)
        set_progress(45, "Прочитано вкладку ЕТ/ФТ")
        psych_records, psych_warnings = parse_schedule(wb, actual["psych"], year)
        set_progress(70, "Прочитано вкладку Псих-Лого")
    finally:
        wb.close()

    # A broken date/time header would make all patients appear to have no
    # procedures. Stop instead of reporting a misleading mass of zeros.
    fatal = [warning for warning in et_warnings + psych_warnings
             if "Не знайдено рядок" in warning.message or "Не знайдено 30-хвилинну" in warning.message]
    if fatal:
        raise ValueError("Не вдалося прочитати графік: " + "; ".join(w.message for w in fatal))

    by_sheet = []
    all_names = set()
    for records in (et_records, psych_records):
        index = defaultdict(lambda: defaultdict(list))
        for appointment in records:
            index[appointment.patient_name_normalized][appointment.date].append(appointment)
            all_names.add(appointment.patient_name_normalized)
        by_sheet.append(index)
    by_signature = defaultdict(set)
    by_package_pair = defaultdict(set)
    by_package_loose_pair = defaultdict(set)
    for key in all_names:
        by_signature[_name_signature(key)].add(key)
        package_name = _package_schedule_name(key)
        if package_name:
            pair = _name_pair(package_name)
            by_package_pair[pair].add(key)
            by_package_loose_pair[_loose_pair(pair)].add(key)

    resolved = {}
    for source_key, source in patients.items():
        if source["daily_target"] == 1:
            pair = _name_pair(source_key)
            aliases = by_package_pair.get(pair, set()) if pair else set()
            ambiguous = bool(aliases and len(package_roster_pairs[pair]) > 1)
            if pair and not aliases:
                loose_pair = _loose_pair(pair)
                loose_aliases = by_package_loose_pair.get(loose_pair, set())
                schedule_pairs = {_name_pair(_package_schedule_name(alias)) for alias in loose_aliases}
                if (loose_aliases and len(package_roster_loose_pairs[loose_pair]) == 1
                        and len(schedule_pairs) == 1):
                    aliases = loose_aliases
            matches = ({source_key} if source_key in all_names else set())
            if not ambiguous:
                matches.update(aliases)
            if matches or ambiguous:
                resolved[source_key] = (tuple(sorted(matches)), ambiguous)
                continue
        if source_key in all_names:
            resolved[source_key] = ((source_key,), False)
            continue
        candidates = by_signature.get(_name_signature(source_key), set())
        resolved[source_key] = ((next(iter(candidates)),) if len(candidates) == 1 else (),
                                len(candidates) > 1)
    group_dates = defaultdict(set)
    for source_key, source in patients.items():
        matched_keys, ambiguous = resolved[source_key]
        if source["group"] and matched_keys and not ambiguous:
            group_key = (source["source_sheet"], source["group"])
            for matched_key in matched_keys:
                group_dates[group_key].update(by_sheet[0].get(matched_key, {}))
                group_dates[group_key].update(by_sheet[1].get(matched_key, {}))

    issues = []
    summaries = []
    selected_patients = 0
    for source_key, source in patients.items():
        raise_if_cancelled()
        matched_keys, ambiguous = resolved[source_key]
        all_et_days = defaultdict(list)
        all_psych_days = defaultdict(list)
        for matched_key in matched_keys:
            for day, items in by_sheet[0].get(matched_key, {}).items():
                all_et_days[day].extend(items)
            for day, items in by_sheet[1].get(matched_key, {}).items():
                all_psych_days[day].extend(items)
        et_days = {day: items for day, items in all_et_days.items() if not month or day.startswith(month)}
        psych_days = {day: items for day, items in all_psych_days.items() if not month or day.startswith(month)}
        in_month = bool(et_days or psych_days)
        if month and not (month[5:] in source["roster_months"] and month[:4] == str(year) or in_month):
            continue
        selected_patients += 1
        if ambiguous:
            issues.append({"kind": "ambiguous_patient", "patient": source["patient"],
                           "source_cell": source["source_cell"], "source_sheet": source["source_sheet"], "date": "",
                           "message": ("Короткий запис у графіку може відповідати кільком пацієнтам; перевірте прізвище та ім’я."
                                       if source["daily_target"] == 1 else
                                       "У графіках знайдено кілька пацієнтів з таким прізвищем та ініціалами.")})
            continue
        dates = sorted(set(et_days) | set(psych_days))
        et_total = sum(len(items) for items in et_days.values())
        psych_total = sum(len(items) for items in psych_days.values())
        total = et_total + psych_total
        expected_days = source["expected_days"]
        daily_target = source["daily_target"]
        summary = {"patient": source["patient"], "source_cell": source["source_cell"],
                   "source_sheet": source["source_sheet"], "days": len(dates),
                   "daily_target": daily_target,
                   "expected_days": None if month or daily_target == 1 else expected_days,
                   "etft_total": et_total, "psych_total": psych_total,
                   "total": total,
                   "expected_total": (expected_days if daily_target == 1 and month else
                                      expected_days * daily_target if daily_target == 2 and not month else None)}
        summaries.append(summary)
        if daily_target == 1:
            months = ({month} if month else
                      {day[:7] for day in dates} |
                      {f"{year}-{roster_month}" for roster_month in source["roster_months"]})
            for checked_month in sorted(months):
                month_total = sum(len(items) for day, items in et_days.items() if day.startswith(checked_month))
                month_total += sum(len(items) for day, items in psych_days.items() if day.startswith(checked_month))
                if month_total < expected_days:
                    issues.append({"kind": "course_total", **summary, "date": "",
                                   "checked_month": checked_month, "month_total": month_total,
                                   "message": f"За {checked_month}: {month_total} процедур; потрібно щонайменше {expected_days}."})
        if not month and daily_target == 2 and len(dates) != expected_days:
            issues.append({"kind": "course_days", **summary, "date": "",
                           "message": f"Днів із записами: {len(dates)} із {expected_days}; процедур за всі дні: {total} із {expected_days * daily_target}."})
        # A group's common course dates let us identify a wholly empty day. If
        # the group calendar is incomplete or has extras, report only the
        # reliable count of observed days instead of guessing a date.
        group_key = (source["source_sheet"], source["group"])
        cohort_days = group_dates.get(group_key, set()) if source["group"] else set()
        missing_dates = (cohort_days - set(dates)
                         if daily_target == 2 and len(cohort_days) == expected_days else set())
        if month:
            missing_dates = {day for day in missing_dates if day.startswith(month)}
            if daily_target == 2 and not dates and not missing_dates:
                issues.append({"kind": "month_empty", **summary, "date": "",
                               "message": f"За {month} на ЕТ/ФТ та Псих-Лого не знайдено процедур."})
        for day in sorted(set(dates) | missing_dates):
            et_items = et_days.get(day, [])
            psych_items = psych_days.get(day, [])
            day_total = len(et_items) + len(psych_items)
            if day_total == daily_target and len(psych_items) <= 1:
                continue
            issues.append({"kind": "daily_count", "patient": source["patient"],
                           "source_cell": source["source_cell"], "source_sheet": source["source_sheet"], "date": day,
                           "etft_count": len(et_items), "psych_count": len(psych_items),
                           "day_total": day_total, "patient_total": total,
                           "days": len(dates), "daily_target": daily_target,
                           "locations": [_location(a) for a in et_items + psych_items]})
    issue_patients = {(item["source_sheet"], item["source_cell"]) for item in issues}
    summary = {"month": month, "patients": selected_patients,
               "patients_with_issues": len(issue_patients), "issues": len(issues),
               "etft_records": sum(not month or item.date.startswith(month) for item in et_records),
               "psych_records": sum(not month or item.date.startswith(month) for item in psych_records),
               "structure_warnings": len(et_warnings) + len(psych_warnings)}
    scope = f"за {month}" if month else "за весь графік"
    lines = [f"Кол-во процедур {scope}: перевірено {selected_patients} пацієнтів; проблемних: {len(issue_patients)}."]
    for issue in issues:
        if issue["kind"] == "daily_count":
            locations = "; ".join(f"{x['sheet']}!{x['cell']} {x['time']} {x['specialist']}" for x in issue["locations"])
            lines.append(f"{issue['patient']}, {issue['date']}: {issue['day_total']} "
                         f"(ЕТ/ФТ {issue['etft_count']}, Псих-Лого {issue['psych_count']}); "
                         f"всього {scope} {issue['patient_total']}. {locations or 'Записів немає.'}")
        else:
            lines.append(f"{issue['patient']}: {issue['message']}")
    result = {"procedure_count_issues": issues, "patient_summaries": summaries,
              "structure_warnings": [w.as_dict() for w in et_warnings + psych_warnings],
              "summary": summary, "human_log": "\n".join(lines)}
    run_id = save_check_run(action, path, summary, result)
    log_event("procedure_counts.finish", lines[0], action_id=action, data={**summary, "run_id": run_id})
    return ActionResult(True, lines[0], {**result, "run_id": run_id})
