"""Append-only usage log when users download exports."""
from __future__ import annotations

import csv
from datetime import date, datetime
from pathlib import Path
from typing import Any

from modules.config import DATA_DIR

USAGE_LOG_CSV = DATA_DIR / "usage_log.csv"

USAGE_LOG_COLUMNS = [
    "TOOL_TYPE",
    "ACCOUNT",
    "PROJECT",
    "PHASE",
    "WEIGHT",
    "SYSTEM_ETA",
    "CRITICAL_FEEDBACK",
    "SYSTEM_NUMBER",
    "FUNCTIONAL",
    "RECORDED_AT",
]


def _fmt_date(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    if len(text) >= 10 and text[4] == "-" and text[7] == "-":
        return text[:10]
    return text


def append_usage_record(
    *,
    tool_type: str,
    account: str = "",
    project: str = "",
    phase: str = "",
    weight: Any = "",
    system_eta: Any = "",
    critical_feedback: Any = "",
    system_number: Any = "",
    functional: str = "",
) -> Path:
    """
    Append one usage row. tool_type is TEST PLAN / NRE / HC.
    Creates data/usage_log.csv with header on first write.
    """
    USAGE_LOG_CSV.parent.mkdir(parents=True, exist_ok=True)
    write_header = not USAGE_LOG_CSV.exists() or USAGE_LOG_CSV.stat().st_size == 0
    row = {
        "TOOL_TYPE": str(tool_type or "").strip(),
        "ACCOUNT": str(account or "").strip(),
        "PROJECT": str(project or "").strip(),
        "PHASE": str(phase or "").strip(),
        "WEIGHT": "" if weight is None or weight == "" else str(weight),
        "SYSTEM_ETA": _fmt_date(system_eta),
        "CRITICAL_FEEDBACK": _fmt_date(critical_feedback),
        "SYSTEM_NUMBER": "" if system_number is None or system_number == "" else str(system_number),
        "FUNCTIONAL": str(functional or "").strip(),
        "RECORDED_AT": datetime.now().isoformat(timespec="seconds"),
    }
    with USAGE_LOG_CSV.open("a", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=USAGE_LOG_COLUMNS)
        if write_header:
            writer.writeheader()
        writer.writerow(row)
    return USAGE_LOG_CSV
