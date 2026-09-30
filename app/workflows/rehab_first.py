"""Pure validation/planning helpers for the first rehabilitation interaction."""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from openpyxl.utils.datetime import from_excel

ICD_TOKEN = re.compile(r"(?<![A-Z0-9])[A-Z][0-9]{2}(?:\.[0-9A-Z]{1,4})?(?![A-Z0-9])", re.I)


def parse_icd_codes(value) -> list[str]:
    """Extract ordered ICD-10 codes, ignoring parenthesized notes and prose."""
    if value is None:
        return []
    text = re.sub(r"\([^)]*\)", " ", str(value))
    result = []
    for match in ICD_TOKEN.finditer(text.upper()):
        code = match.group(0).upper()
        if code not in result:
            result.append(code)
    return result


def normalize_icd(value: str) -> str:
    return re.sub(r"\s+", "", str(value or "")).upper()


def parse_interaction_datetime(value) -> datetime | None:
    """Parse Excel or local UI datetime without converting timezone/clock time."""
    if isinstance(value, datetime):
        return value.replace(tzinfo=None) if value.tzinfo else value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value,(int,float)) and not isinstance(value,bool):
        try:
            parsed=from_excel(value)
            return parsed if isinstance(parsed,datetime) else None
        except (TypeError,ValueError,OverflowError):
            return None
    raw = str(value or "").strip()
    if not raw:
        return None
    for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            pass
    try:
        parsed = datetime.fromisoformat(raw)
        return parsed.replace(tzinfo=None) if parsed.tzinfo else parsed
    except ValueError:
        return None


def format_interaction_datetime(value) -> str:
    parsed = parse_interaction_datetime(value)
    if parsed is None:
        raise ValueError(f"Некоректна дата/час першої взаємодії: {value!r}")
    # Excel formula results can land a fraction of a second before a minute
    # boundary (for example 10:59:59.999). The schedule is minute-granular, so
    # round instead of truncating and accidentally shifting the appointment
    # one minute earlier.
    parsed = (parsed + timedelta(seconds=30)).replace(second=0, microsecond=0)
    return parsed.strftime("%d.%m.%Y %H:%M")


def age_on(dob, at) -> int | None:
    at_dt = parse_interaction_datetime(at)
    dob_dt = parse_interaction_datetime(dob)
    if not dob_dt or not at_dt:
        return None
    years = at_dt.year - dob_dt.year
    return years - ((at_dt.month, at_dt.day) < (dob_dt.month, dob_dt.day))


def diagnosis_time_matches(actual, expected: str) -> bool:
    return str(actual or "").strip() == str(expected or "").strip()


def validate_age_specific_reports(age: int, codes: list[str]) -> list[int]:
    """Return indexes of child-only removals; raise when required reports drift."""
    normalized=[re.sub(r"\D", "", str(code)) for code in codes]
    if "9602200" not in normalized:
        raise ValueError("Не знайдено обов'язковий діагностичний звіт 96022.")
    adults=[i for i,code in enumerate(normalized) if code=="9603700"]
    if age>=18 and not adults:
        raise ValueError("У дорослого відсутній очікуваний звіт 96037.")
    return adults if age<18 else []


def barthel_rows_to_remove(age: int, names: list[str]) -> list[int]:
    if age>=18:return []
    return [i for i,name in enumerate(names) if "barthel_extended" in str(name).casefold()]


def validate_template_drift(expected: dict, actual: dict) -> list[tuple[str, object, object]]:
    return [(key,value,actual.get(key)) for key,value in expected.items()
            if key not in actual or str(actual.get(key)).strip().casefold()!=str(value).strip().casefold()]


def newest_referral_index(dates: list[datetime | date | str]) -> int:
    """Index of the newest creation date; stable UI order resolves ties."""
    parsed = []
    for index, value in enumerate(dates):
        dt = parse_interaction_datetime(value)
        if dt is None:
            for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                try:
                    dt = datetime.strptime(str(value).strip(), fmt)
                    break
                except ValueError:
                    continue
        if dt is None:
            raise ValueError(f"Не вдалося розібрати дату створення направлення: {value!r}")
        parsed.append((dt, -index, index))
    if not parsed:
        raise ValueError("Електронні направлення відсутні.")
    return max(parsed)[2]


def plan_diagnosis_reconciliation(source: list[str], template: list[dict]) -> dict:
    """Plan row-preserving changes. Z and duplicate-code rows are immutable."""
    desired = [normalize_icd(code) for code in source]
    rows = [dict(row, code=normalize_icd(row.get("code"))) for row in template]
    if not desired:
        raise ValueError("У джерелі немає кодів МКХ.")
    z_rows = [row for row in rows if row["code"].startswith("Z")]
    mutable_rows = [row for row in rows if not row["code"].startswith("Z")]
    counts = {}
    for row in mutable_rows:
        counts[row["code"]] = counts.get(row["code"], 0) + 1
    duplicates = [row for row in mutable_rows if counts[row["code"]] > 1]
    mutable = [row for row in mutable_rows if counts[row["code"]] == 1]
    desired_mutable = [code for code in desired if not code.startswith("Z")]
    remaining = list(mutable)
    kept_rows = []
    missing = []
    for code in desired_mutable:
        # Existing duplicated codes are preserved as a group. They satisfy a
        # requested code without editing one arbitrarily selected duplicate.
        if any(row["code"] == code for row in duplicates):
            continue
        index = next((i for i,row in enumerate(remaining) if row["code"] == code), None)
        if index is None:
            missing.append(code)
        else:
            kept_rows.append(remaining.pop(index))
    replacements = list(zip(remaining, missing))
    deletes = remaining[len(replacements):]
    add_codes = missing[len(replacements):]
    expected = {code: max(desired_mutable.count(code), counts.get(code, 0))
                if counts.get(code, 0) > 1 else desired_mutable.count(code)
                for code in set(desired_mutable) | {r["code"] for r in duplicates}}
    return {"keep": kept_rows, "replace": replacements, "delete": deletes, "add": add_codes,
            "immutable_z": z_rows, "protected_duplicates": duplicates,
            "expected_mutable_counts": expected, "desired": desired}
