"""Synced horizontal-scroll calendar: marks table + one table per system."""
from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

import streamlit.components.v1 as components

from modules.config import (
    CELL_AHEAD_OPT,
    CELL_POSTPONE_OPT,
    EVENT_CRITICAL,
    EVENT_DETAIL_CF,
    EVENT_DETAIL_CRITICAL_OPTS,
    EVENT_DETAIL_ETA,
    EVENT_NONE,
    EVENT_OCCUPIED,
    EVENT_OPTS,
)
from modules.data_loader import profile_column_for_weight
from modules.timeline import ensure_event_state, sorted_system_keys

_FRONTEND = Path(__file__).parent / "frontend"
_component = components.declare_component("synced_calendar", path=str(_FRONTEND))


def render_synced_calendar(
    timeline: dict[str, Any],
    options: list[str],
    *,
    weight_kg: float | None = None,
    detail_by_id: dict[str, dict[str, str]] | None = None,
    key: str | None = None,
) -> dict | None:
    """
    Render split tables (Calendar marks + one table per System).

    Calendar marks: Weekday, Event (- / Occupied / Critical Event), Detail.
    Under each system schedule row, a read-only PROFILE note row is shown.

    Returns event dict when user changes a cell:
      - {kind: "system", system, date, value, nonce}
      - {kind: "mark_event", date, value, nonce}  Event row
      - {kind: "event_detail", date, value, nonce}  Detail row
      - {kind: "shift_system", system, delta, nonce}
      - {kind: "shift_from_date", system, date, delta, nonce}
      - {kind: "rename_system", system, name, nonce}
    """
    ensure_event_state(timeline)
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
    # Occupied days (legacy empty_marks) — system cells locked
    occupied_marks = list(timeline.get("empty_marks", []))
    blocked = list(timeline.get("blocked", []))
    grid: dict = timeline.get("grid", {})
    event_marks = dict(timeline.get("event_marks") or {})
    event_details = dict(timeline.get("event_details") or {})

    profile_col = None
    if weight_kg is not None and detail_by_id:
        profile_col = profile_column_for_weight(float(weight_kg))

    weekdays = [date.fromisoformat(d).strftime("%a") for d in dates]
    marks = [" | ".join(markers.get(d, [])) for d in dates]
    events = [event_marks.get(d, EVENT_NONE) for d in dates]
    details = [event_details.get(d, EVENT_NONE) for d in dates]

    systems = []
    for name in sorted_system_keys(timeline):
        cells = {}
        notes = {}
        for d in dates:
            cell = grid[name].get(d)
            if not cell:
                cells[d] = ""
                notes[d] = ""
            elif cell.get("is_empty"):
                cells[d] = "Empty"
                notes[d] = ""
            else:
                tid = cell.get("test_id", "")
                abbrv = cell.get("abbrv") or tid
                cells[d] = f"{abbrv} ({tid})"
                note = ""
                if profile_col and detail_by_id:
                    meta = detail_by_id.get(str(tid).strip())
                    if meta:
                        note = meta.get(profile_col, "") or ""
                notes[d] = note
        systems.append({"name": name, "cells": cells, "notes": notes})

    return _component(
        dates=dates,
        weekdays=weekdays,
        marks=marks,
        events=events,
        details=details,
        calendar_blocked=calendar_blocked,
        empty_marks=occupied_marks,
        blocked=blocked,
        systems=systems,
        options=options,
        event_opts=list(EVENT_OPTS),
        detail_critical_opts=list(EVENT_DETAIL_CRITICAL_OPTS),
        event_none=EVENT_NONE,
        event_occupied=EVENT_OCCUPIED,
        event_critical=EVENT_CRITICAL,
        detail_eta=EVENT_DETAIL_ETA,
        detail_cf=EVENT_DETAIL_CF,
        shift_ahead_opt=CELL_AHEAD_OPT,
        shift_postpone_opt=CELL_POSTPONE_OPT,
        key=key,
        default=None,
    )
