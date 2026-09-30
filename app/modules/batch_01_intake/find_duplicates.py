from __future__ import annotations

from ...core.database import connect
from ...core.result import ActionResult


def run(payload: dict) -> ActionResult:
    with connect() as conn:
        conn.execute("UPDATE patients SET has_duplicate=0")
        groups = conn.execute(
            """
            SELECT normalized_name, COALESCE(birth_date, '') AS bd, COUNT(*) AS cnt
            FROM patients
            WHERE normalized_name <> ''
            GROUP BY normalized_name, COALESCE(birth_date, '')
            HAVING COUNT(*) > 1
            """
        ).fetchall()

        affected = 0
        details = []
        for group in groups:
            cur = conn.execute(
                """
                UPDATE patients
                SET has_duplicate=1
                WHERE normalized_name=? AND COALESCE(birth_date, '')=?
                """,
                (group["normalized_name"], group["bd"]),
            )
            affected += cur.rowcount
            details.append({"key": group["normalized_name"], "birth_date": group["bd"], "count": group["cnt"]})

    return ActionResult(
        ok=True,
        message=f"Знайдено {len(groups)} груп дубльованих записів.",
        data={"groups": len(groups), "affected_rows": affected, "duplicates": details[:50]},
    )
