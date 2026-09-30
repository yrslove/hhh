"""Cooperative cancellation and progress reporting for every workflow.

Python cannot safely kill a worker thread in the middle of an Excel write or a
browser click. Workflows therefore check the shared cancellation flag at safe
boundaries; already committed Doctor Eleks/Excel changes are left intact.
"""
from __future__ import annotations

from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime, timezone
from threading import Event, RLock
from uuid import uuid4


class ActionCancelled(RuntimeError):
    """Raised when the operator requested cancellation of a running action."""


_lock = RLock()
_active: dict[str, dict] = {}
_current: ContextVar[str | None] = ContextVar("current_action_token", default=None)


def _public(item: dict) -> dict:
    return {
        "id": item["id"],
        "function_id": item["function_id"],
        "started_at": item["started_at"],
        "progress": item["progress"],
        "detail": item["detail"],
        "cancel_requested": item["event"].is_set(),
    }


@contextmanager
def action_scope(function_id: str):
    token = uuid4().hex
    item = {
        "id": token,
        "function_id": function_id,
        "started_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "progress": 2,
        "detail": "Підготовка запуску",
        "event": Event(),
    }
    with _lock:
        _active[token] = item
    context_token = _current.set(token)
    try:
        yield token
    finally:
        _current.reset(context_token)
        with _lock:
            _active.pop(token, None)


def set_progress(progress: int | float, detail: str = "") -> None:
    """Update the current action's approximate progress (0–100)."""
    token = _current.get()
    if not token:
        return
    with _lock:
        item = _active.get(token)
        if item:
            item["progress"] = max(0, min(100, round(float(progress))))
            if detail:
                item["detail"] = detail


def set_progress_count(done: int, total: int, detail: str = "") -> None:
    # Leave a little room for finalization; an action reports 100 only on exit.
    progress = 5 if total <= 0 else 5 + 90 * max(0, min(done, total)) / total
    set_progress(progress, detail)


def raise_if_cancelled() -> None:
    token = _current.get()
    if not token:
        return
    with _lock:
        item = _active.get(token)
        cancelled = bool(item and item["event"].is_set())
    if cancelled:
        raise ActionCancelled("Виконання зупинено оператором. Уже збережені зміни не скасовуються.")


def cancel_all() -> int:
    """Signal active actions; the API stop route also stops the browser driver."""
    with _lock:
        for item in _active.values():
            item["event"].set()
            item["detail"] = "Отримано запит на зупинку"
        return len(_active)


def active_actions() -> list[dict]:
    with _lock:
        return [_public(item) for item in _active.values()]
