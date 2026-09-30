from __future__ import annotations

import os
import platform
import sqlite3
import sys
from pathlib import Path

from ...core.paths import CONFIG_DIR, DATA_DIR, DB_PATH, IMPORT_DIR, LOG_DIR, RUNTIME_DIR
from ...core.result import ActionResult


def run(payload: dict) -> ActionResult:
    checks = []
    for name, path in (
        ("config", CONFIG_DIR),
        ("data", DATA_DIR),
        ("imports", IMPORT_DIR),
        ("logs", LOG_DIR),
        ("runtime", RUNTIME_DIR),
    ):
        writable = path.exists() and os.access(path, os.W_OK)
        checks.append({"name": name, "path": str(path), "ok": bool(writable)})

    sqlite_ok = False
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("SELECT 1")
        conn.close()
        sqlite_ok = True
    except Exception:
        sqlite_ok = False

    checks.append({"name": "sqlite", "path": str(DB_PATH), "ok": sqlite_ok})
    ok = all(item["ok"] for item in checks)
    return ActionResult(
        ok=ok,
        message="Середовище готове." if ok else "Є проблеми з локальним середовищем.",
        data={
            "python": sys.version.split()[0],
            "platform": platform.platform(),
            "checks": checks,
        },
    )
