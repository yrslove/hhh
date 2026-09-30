from __future__ import annotations

import re
from datetime import datetime, timezone

from ...core.database import connect
from ...core.result import ActionResult

SELF_REFERRAL_WORDS = ("самозвер", "само звер", "самозвернення")


def _looks_iso_date(text: str | None) -> bool:
    if not text:
        return False
    try:
        datetime.fromisoformat(text)
        return True
    except ValueError:
        return False


def run(payload: dict) -> ActionResult:
    counts = {"ok": 0, "warning": 0, "error": 0}
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT id, full_name, referral, referral_warning, birth_date,
                   interaction_at, system_diagnoses
            FROM patients
            """
        ).fetchall()
        for row in rows:
            errors: list[str] = []
            warnings: list[str] = []
            name = (row["full_name"] or "").strip()
            if len(name) < 3:
                errors.append("Не заповнено ПІБ.")

            birth = row["birth_date"]
            if not birth:
                warnings.append("Не вказано дату народження.")
            elif not _looks_iso_date(birth):
                warnings.append("Дата народження має нестандартний формат.")

            referral = (row["referral"] or "").strip()
            lower = referral.casefold()
            is_self = any(word in lower for word in SELF_REFERRAL_WORDS)
            if not referral:
                warnings.append("Не вказано направлення або самозвернення.")
            elif not is_self:
                digits = re.sub(r"\D", "", referral)
                if len(digits) != 16:
                    errors.append("Е-направлення має містити 16 цифр або бути самозверненням.")
            if row["referral_warning"]:
                warnings.append(row["referral_warning"])

            interaction = row["interaction_at"]
            if not interaction:
                warnings.append("Не вказано дату/час першої взаємодії.")
            elif not _looks_iso_date(interaction):
                warnings.append("Дата/час першої взаємодії має нестандартний формат.")

            if not row["system_diagnoses"]:
                warnings.append("Коди діагнозів по системі ще не заповнені.")

            if errors:
                status = "error"
                message = " ".join(errors + warnings)
            elif warnings:
                status = "warning"
                message = " ".join(warnings)
            else:
                status = "ok"
                message = ""
            counts[status] += 1
            conn.execute(
                "UPDATE patients SET validation_status=?, validation_message=?, updated_at=? WHERE id=?",
                (
                    status,
                    message,
                    datetime.now(timezone.utc).isoformat(timespec="seconds"),
                    row["id"],
                ),
            )

    return ActionResult(
        ok=True,
        message=f"Перевірено {sum(counts.values())} пацієнтів.",
        data=counts,
    )
