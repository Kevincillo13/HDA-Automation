from __future__ import annotations

from datetime import datetime


_run_id: str | None = None
_run_name: str = "run"


def start_run(run_name: str) -> str:
    global _run_id, _run_name
    _run_name = run_name
    # Force short ID for SAP 128-char path limit
    _run_id = datetime.now().strftime('%Y%m%d_%H%M%S')
    return _run_id


def get_run_id() -> str:
    return _run_id or start_run(_run_name)


def get_run_name() -> str:
    return _run_name
