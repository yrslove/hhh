from __future__ import annotations

from datetime import datetime, timezone

from openpyxl import load_workbook

from ...core.database import connect, get_setting, set_setting
from ...core.result import ActionResult
from .common import (
    cell_text,
    inspect_sheet,
    iso_date,
    iso_datetime,
    normalized_patient_name,
    referral_text,
    resolve_workbook_path,
    source_key,
)


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def run(payload: dict) -> ActionResult:
    path = resolve_workbook_path(payload, get_setting("latest_workbook"))
    requested_sheet = payload.get("sheet")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet_name = requested_sheet if requested_sheet in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        info = inspect_sheet(ws)
        mapping = info["mapping"]
        name_col = mapping.get("full_name")
        if not name_col:
            return ActionResult(False, "Не знайдено колонку ПІБ.", {"sheet": sheet_name})

        imported = 0
        numeric_referral_warnings = 0
        now = _now()
        source_name = get_setting("latest_workbook_original_name", path.name)
        max_needed_col = max(mapping.values(), default=1)

        def value(row_values: tuple, field: str):
            col = mapping.get(field)
            if not col or col > len(row_values):
                return None
            return row_values[col - 1]

        with connect() as conn:
            # Імпорт одного аркуша — це актуальний знімок цього джерела.
            # Попередня локальна копія того самого filename+sheet замінюється,
            # щоб повторне завантаження не множило дублікати.
            conn.execute(
                "DELETE FROM patients WHERE source_file=? AND source_sheet=?",
                (source_name, sheet_name),
            )
            rows = ws.iter_rows(
                min_row=info["header_row"] + 1,
                max_col=max_needed_col,
                values_only=True,
            )
            for row_idx, row_values in enumerate(rows, start=info["header_row"] + 1):
                full_name = cell_text(value(row_values, "full_name"))
                if not full_name:
                    continue
                if normalized_patient_name(full_name) in {"піб", "піб пацієнта", "пацієнт"}:
                    continue

                referral, referral_warning = referral_text(value(row_values, "referral"))
                if referral_warning:
                    numeric_referral_warnings += 1

                item = {
                    "source_key": source_key(source_name, sheet_name, row_idx),
                    "source_file": source_name,
                    "source_sheet": sheet_name,
                    "source_row": row_idx,
                    "full_name": full_name,
                    "normalized_name": normalized_patient_name(full_name),
                    "referral": referral,
                    "referral_warning": referral_warning,
                    "birth_date": iso_date(value(row_values, "birth_date")),
                    "document_diagnoses": cell_text(value(row_values, "document_diagnoses")),
                    "interaction_at": iso_datetime(value(row_values, "interaction_at")),
                    "system_diagnoses": cell_text(value(row_values, "system_diagnoses")),
                    "first_interaction_status": cell_text(value(row_values, "first_interaction_status")),
                    "treatment_plan": cell_text(value(row_values, "treatment_plan")),
                    "treatment_plan_status": cell_text(value(row_values, "treatment_plan_status")),
                    "procedure_codes": cell_text(value(row_values, "procedure_codes")),
                }

                conn.execute(
                    """
                    INSERT INTO patients(
                        source_key, source_file, source_sheet, source_row,
                        full_name, normalized_name, referral, referral_warning, birth_date,
                        document_diagnoses, interaction_at, system_diagnoses,
                        first_interaction_status, treatment_plan, treatment_plan_status,
                        procedure_codes, created_at, updated_at
                    ) VALUES (
                        :source_key, :source_file, :source_sheet, :source_row,
                        :full_name, :normalized_name, :referral, :referral_warning, :birth_date,
                        :document_diagnoses, :interaction_at, :system_diagnoses,
                        :first_interaction_status, :treatment_plan, :treatment_plan_status,
                        :procedure_codes, :created_at, :updated_at
                    )
                    ON CONFLICT(source_key) DO UPDATE SET
                        full_name=excluded.full_name,
                        normalized_name=excluded.normalized_name,
                        referral=excluded.referral,
                        referral_warning=excluded.referral_warning,
                        birth_date=excluded.birth_date,
                        document_diagnoses=excluded.document_diagnoses,
                        interaction_at=excluded.interaction_at,
                        system_diagnoses=excluded.system_diagnoses,
                        first_interaction_status=excluded.first_interaction_status,
                        treatment_plan=excluded.treatment_plan,
                        treatment_plan_status=excluded.treatment_plan_status,
                        procedure_codes=excluded.procedure_codes,
                        validation_status='not_checked',
                        validation_message=NULL,
                        has_duplicate=0,
                        updated_at=excluded.updated_at
                    """,
                    {**item, "created_at": now, "updated_at": now},
                )
                imported += 1

        set_setting("active_sheet", sheet_name)
        return ActionResult(
            ok=True,
            message=f"Імпортовано/оновлено {imported} пацієнтів.",
            data={
                "sheet": sheet_name,
                "imported": imported,
                "numeric_referral_warnings": numeric_referral_warnings,
                "mapping": mapping,
                "mapping_source": info["mapping_source"],
            },
        )
    finally:
        wb.close()
