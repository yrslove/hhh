from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from fastapi import UploadFile

from ...core.database import set_setting
from ...core.paths import IMPORT_DIR


def save_upload(file: UploadFile) -> dict:
    """Зберігає локальну робочу копію Excel та повертає технічний результат."""
    original_name = Path(file.filename or "workbook.xlsx").name
    suffix = Path(original_name).suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        raise ValueError("Підтримуються тільки .xlsx та .xlsm")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    destination = IMPORT_DIR / f"{stamp}_{original_name}"
    with destination.open("wb") as out:
        shutil.copyfileobj(file.file, out)

    set_setting("latest_workbook", str(destination.resolve()))
    set_setting("latest_workbook_original_name", original_name)
    set_setting("active_sheet", None)
    set_setting("latest_inspection", None)
    return {
        "filename": original_name,
        "path": str(destination.resolve()),
        "size": destination.stat().st_size,
        "suffix": suffix,
    }
