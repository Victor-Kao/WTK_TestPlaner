"""Synced horizontal-scroll calendar: marks table + one table per system."""
from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import streamlit.components.v1 as components

from modules.config import CELL_AHEAD_OPT, CELL_POSTPONE_OPT

from modules.timeline import sorted_system_keys

_FRONTEND = Path(__file__).parent / "frontend"
_component = components.declare_component("synced_calendar", path=str(_FRONTEND))


def render_synced_calendar(
    timeline: dict[str, Any],
    options: list[str],
    *,
    key: str | None = None,
) -> dict | None:
    """
    Render split tables (Calendar marks + one table per System).

    Returns event dict when user changes a cell:
      - {kind: "system", system, date, value, nonce}
      - {kind: "mark_empty", date, value, nonce}  value is "" or "Empty"
      - {kind: "shift_system", system, delta, nonce}  whole-row ±1
      - {kind: "shift_from_date", system, date, delta, nonce}  from selected date ±1
      - {kind: "rename_system", system, name, nonce}  rename row label
    """
    dates: list[str] = list(timeline.get("dates", []))
    markers: dict = timeline.get("markers", {})
    calendar_blocked = list(
        timeline.get("calendar_blocked")
        or [
            d
            for d, tags in markers.items()
            if "Weekend" in tags or "Holiday" in tags
        ]
    )
    empty_marks = list(timeline.get("empty_marks", []))
    blocked = list(timeline.get("blocked", []))
    grid: dict = timeline.get("grid", {})

    weekdays = [date.fromisoformat(d).strftime("%a") for d in dates]
    marks = [" | ".join(markers.get(d, [])) for d in dates]

    systems = []
    for name in sorted_system_keys(timeline):
        cells = {}
        for d in dates:
            cell = grid[name].get(d)
            if not cell:
                cells[d] = ""
            elif cell.get("is_empty"):
                cells[d] = "Empty"
            else:
                tid = cell.get("test_id", "")
                abbrv = cell.get("abbrv") or tid
                cells[d] = f"{abbrv} ({tid})"
        systems.append({"name": name, "cells": cells})

    return _component(
        dates=dates,
        weekdays=weekdays,
        marks=marks,
        calendar_blocked=calendar_blocked,
        empty_marks=empty_marks,
        blocked=blocked,
        systems=systems,
        options=options,
        shift_ahead_opt=CELL_AHEAD_OPT,
        shift_postpone_opt=CELL_POSTPONE_OPT,
        key=key,
        default=None,
    )
