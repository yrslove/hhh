from __future__ import annotations

import hashlib
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from ...core.config import load_json
from ...core.text import normalize_header, normalize_name, normalize_space


def resolve_workbook_path(payload: dict[str, Any], latest_path: str | None) -> Path:
    raw = payload.get("path") or latest_path
    if not raw:
        raise ValueError("Спочатку завантажте Excel-файл.")
    path = Path(raw).expanduser().resolve()
    if not path.exists() or not path.is_file():
        raise ValueError("Файл не знайдено на локальному диску.")
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Підтримуються .xlsx та .xlsm. Старий .xls потрібно зберегти як .xlsx.")
    return path


def load_aliases() -> tuple[dict[str, list[str]], dict[str, str]]:
    cfg = load_json("column_aliases.json")
    return cfg.get("fields", {}), cfg.get("fallback_columns", {})


def _match_field(cell_value: Any, aliases: dict[str, list[str]]) -> str | None:
    """Спочатку шукає точний заголовок, потім — найбільш специфічний частковий збіг.

    Це не дає загальному тексту на кшталт «Перша взаємодія» перехопити
    колонку «Статус першої взаємодії».
    """
    cell = normalize_header(cell_value)
    if not cell:
        return None

    normalized: list[tuple[str, str]] = []
    for field, values in aliases.items():
        for alias in values:
            candidate = normalize_header(alias)
            if candidate:
                normalized.append((field, candidate))

    for field, candidate in normalized:
        if cell == candidate:
            return field

    partials: list[tuple[int, str]] = []
    for field, candidate in normalized:
        if len(candidate) >= 7 and (candidate in cell or cell in candidate):
            partials.append((len(candidate), field))
    if not partials:
        return None
    partials.sort(reverse=True)
    return partials[0][1]


def inspect_sheet(ws) -> dict[str, Any]:
    aliases, fallback = load_aliases()
    best_row = 1
    best_mapping: dict[str, int] = {}
    best_score = -1

    scan_rows = min(ws.max_row or 1, 35)
    scan_cols = min(ws.max_column or 1, 80)
    for row_idx, values in enumerate(
        ws.iter_rows(min_row=1, max_row=scan_rows, max_col=scan_cols, values_only=True),
        start=1,
    ):
        mapping: dict[str, int] = {}
        for col_idx, cell_value in enumerate(values, start=1):
            field = _match_field(cell_value, aliases)
            if field and field not in mapping:
                mapping[field] = col_idx
        score = len(mapping)
        if score > best_score:
            best_score = score
            best_row = row_idx
            best_mapping = mapping

    mapping_source: dict[str, str] = {key: "header" for key in best_mapping}
    if best_score < 2:
        best_row = 1
        best_mapping = {}
        mapping_source = {}

    for field, letter in fallback.items():
        if field not in best_mapping:
            best_mapping[field] = column_index_from_string(letter)
            mapping_source[field] = f"fallback:{letter}"

    core_fields = ["full_name", "referral", "birth_date", "interaction_at"]
    fallback_core = [field for field in core_fields if mapping_source.get(field, "").startswith("fallback:")]
    return {
        "sheet": ws.title,
        "header_row": best_row,
        "mapping": best_mapping,
        "mapping_source": mapping_source,
        "fallback_core": fallback_core,
        "max_row": ws.max_row,
        "max_column": ws.max_column,
        "header_matches": max(best_score, 0),
    }


def workbook_structure(path: Path) -> dict[str, Any]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = [inspect_sheet(wb[name]) for name in wb.sheetnames]
        return {
            "path": str(path),
            "filename": path.name,
            "sheet_names": wb.sheetnames,
            "sheets": sheets,
        }
    finally:
        wb.close()


def iso_date(value: Any) -> str | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = normalize_space(value)
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return text or None


def iso_datetime(value: Any) -> str | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.isoformat(timespec="minutes")
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time()).isoformat(timespec="minutes")
    text = normalize_space(value)
    for fmt in (
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y %H.%M",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M",
        "%d.%m.%Y",
    ):
        try:
            return datetime.strptime(text, fmt).isoformat(timespec="minutes")
        except ValueError:
            continue
    return text or None


def referral_text(value: Any) -> tuple[str | None, str | None]:
    if value is None or value == "":
        return None, None
    warning = None
    if isinstance(value, bool):
        return str(value), "Некоректний тип номера направлення."
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        warning = (
            "Номер направлення збережений в Excel як число. Для 16-значних номерів "
            "використовуйте текстовий формат, інакше Excel може змінити останні цифри."
        )
    text = normalize_space(value)
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text, warning


def source_key(source_id: str, sheet: str, row: int) -> str:
    raw = f"{source_id}|{sheet}|{row}".encode("utf-8", errors="ignore")
    return hashlib.sha256(raw).hexdigest()


def cell_text(value: Any) -> str | None:
    text = normalize_space(value)
    return text or None


def normalized_patient_name(value: Any) -> str:
    return normalize_name(value)
