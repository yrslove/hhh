from __future__ import annotations

from ...core.database import get_setting, set_setting
from ...core.result import ActionResult
from .common import resolve_workbook_path, workbook_structure


def run(payload: dict) -> ActionResult:
    path = resolve_workbook_path(payload, get_setting("latest_workbook"))
    structure = workbook_structure(path)
    set_setting("latest_inspection", structure)
    set_setting("latest_workbook", str(path))
    return ActionResult(
        ok=True,
        message=f"Таблицю проаналізовано: {len(structure['sheets'])} арк.",
        data=structure,
    )
