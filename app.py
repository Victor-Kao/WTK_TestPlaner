"""
WTK Test Planner — Streamlit schedule management tool.
Accounts: ROSA, NAOMI (per-account data under data/<ACCOUNT>/).

Main-page layout only (no sidebar) so a parent tool can own the sidebar.
"""
from __future__ import annotations

import html
from datetime import date, timedelta

import pandas as pd
import streamlit as st

from modules.config import (
    ACCOUNTS,
    CELL_AHEAD_OPT,
    CELL_POSTPONE_OPT,
    EMPTY_LABEL,
    EMPTY_TOKEN,
    FUNCTIONALITY_OPTS,
    HC_CALENDAR_MONTHS,
    HC_DISTRIBUTION_OPTS,
    HC_DIST_PER_MONTH,
    HC_DIST_PER_PHASE,
    HC_ENGINEER_TYPES,
    HC_MONTH_DEFAULTS,
    HC_MONTH_DISPLAY_OPTS,
    HC_SHOW_MONTH,
    HC_SHOW_SEQUENCE,
    NRE_AUX_COST_DEFAULTS_USD,
    NRE_AUX_COST_IDS,
    NRE_DEFAULT_PHASES,
    NRE_DNP_ID,
    NRE_EMPTY_CONTENT_DEFAULT_PHASES,
    NRE_OM_ID,
    NRE_PHASE_OPTION_DEFAULTS,
    NRE_SELECTOR_EXTRA_IDS,
    NRE_UFIT_ID,
    ORV3_MGX_WO_L11_COL,
    ORV3_MGX_WO_L11_LABEL,
    PROJECT_PHASES,
    SEQUENCE_TEMPLATE_MAX_SYSTEMS,
    STANDARDS,
    SYSTEM_NUMBERS,
    YES_NO,
)
from modules.data_loader import (
    PROFILE_COLUMNS,
    load_convert_table,
    load_location_info,
    load_test_plan_info,
    lookup_convert_id,
    load_case_sequence,
    test_detail_by_id,
    test_info_by_id,
)
from modules.headcount import (
    build_headcount_detail_table,
    build_headcount_table,
    export_headcount_xlsx,
)
from modules.nre import (
    aux_costs_have_content,
    available_nre_phases,
    build_nre_table,
    export_nre_xlsx,
    first_nre_phase,
    nre_grand_total,
    phase_has_nre_content,
    pivot_nre_for_template,
    sanitize_nre_phases,
)
from modules.timeline import (
    apply_sequence_template,
    clear_all_test_items,
    clear_cell_span,
    ensure_event_state,
    init_timeline_state,
    place_item,
    refresh_durations_from_catalog,
    rename_system_row,
    set_calendar_event,
    set_event_detail,
    shift_system_from_date,
    shift_system_schedule,
    style_timeline_display,
    timeline_to_export_df,
    export_timeline_xlsx,
)
from modules.synced_calendar import render_synced_calendar
from modules.usage_log import append_usage_record

st.set_page_config(
    page_title="WTK Test Planner",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Session defaults
# ---------------------------------------------------------------------------
def _init_state() -> None:
    defaults = {
        "timeline": None,
        "export_df": None,
        "nre_df": None,
        "hc_phases": None,
        "hc_nre_hours_by_phase": None,
        "hc_nre_func_by_phase": None,
        "hc_df": None,
        "hc_detail_df": None,
        "hc_table_phases": None,
        "hc_table_months": None,
        "last_convert_id": None,
        "editor_version": 0,
        "editor_error": None,
        "last_cal_nonce": None,
        "active_account": None,
        "plan_started": False,
        "started_account": None,
        "started_weight_kg": None,
        "started_standard": None,
        "started_project_name": None,
        "started_project_phase": None,
        "plan_log_context": None,
        "catalog_account": None,
        "work_test_plan": None,
        "work_location": None,
        "work_convert": None,
        "catalog_edit_rev": 0,
        "refresh_msg": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()

st.markdown(
    """
    <style>
    /* Parent shell owns the sidebar — hide this app's empty sidebar */
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="stSidebarCollapsedControl"] { display: none !important; }

    .stApp [data-stale="true"],
    .stApp [data-stale="true"] iframe,
    iframe[title*="synced_calendar"] {
      opacity: 1 !important;
      transition: none !important;
      filter: none !important;
    }
    div[data-testid="stCustomComponentV1"] {
      opacity: 1 !important;
      transition: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def _cached_catalog(account: str):
    return (
        load_test_plan_info(account),
        load_location_info(account),
        load_convert_table(account),
    )


def _normalize_test_plan(df):
    out = df.copy()
    required = [
        "Test_ID",
        "Testplan_Item",
        "Duration_Days",
        "Duration_for_NRE",
        "Abbrv_Name",
        "Lab_Rate",
    ]
    missing = [c for c in required if c not in out.columns]
    if missing:
        raise ValueError(f"Test plan table missing columns: {missing}")
    out["Test_ID"] = out["Test_ID"].astype(str).str.strip()
    out["Duration_Days"] = (
        pd.to_numeric(out["Duration_Days"], errors="coerce").fillna(0).astype(int)
    )
    out["Duration_for_NRE"] = pd.to_numeric(
        out["Duration_for_NRE"], errors="coerce"
    ).fillna(0)
    out["Lab_Rate"] = pd.to_numeric(out["Lab_Rate"], errors="coerce").fillna(0)
    for col in PROFILE_COLUMNS:
        if col not in out.columns:
            out[col] = ""
        else:
            out[col] = out[col].fillna("").astype(str).str.strip()
            out.loc[out[col].str.lower().isin(("nan", "none")), col] = ""
    return out


def _normalize_location(df):
    out = df.copy()
    if "Test_ID" not in out.columns or "Location" not in out.columns:
        raise ValueError("Location table needs Test_ID and Location columns")
    out["Test_ID"] = out["Test_ID"].astype(str).str.strip()
    return out


def _normalize_convert(df):
    out = df.copy()
    if (
        ORV3_MGX_WO_L11_COL not in out.columns
        and "Ufit_for_SoR" in out.columns
    ):
        out = out.rename(columns={"Ufit_for_SoR": ORV3_MGX_WO_L11_COL})
    required = [
        "Standard",
        "Functionality",
        "Gold_Rail_Selection",
        ORV3_MGX_WO_L11_COL,
        "System_Number",
        "Convert_ID",
    ]
    missing = [c for c in required if c not in out.columns]
    if missing:
        raise ValueError(f"Convert table missing columns: {missing}")
    out["System_Number"] = out["System_Number"].astype(int)
    out["Convert_ID"] = out["Convert_ID"].astype(int)
    return out


def _load_working_catalog_from_disk(account: str) -> None:
    """Load folder CSVs into session working copies (does not write back)."""
    plan, loc, conv = _cached_catalog(account)
    st.session_state.catalog_account = account
    st.session_state.work_test_plan = plan.copy()
    st.session_state.work_location = loc.copy()
    st.session_state.work_convert = conv.copy()
    st.session_state.catalog_edit_rev = int(st.session_state.get("catalog_edit_rev", 0)) + 1


def _ensure_working_catalog(account: str) -> None:
    if (
        st.session_state.catalog_account != account
        or st.session_state.work_test_plan is None
        or st.session_state.work_location is None
        or st.session_state.work_convert is None
    ):
        _load_working_catalog_from_disk(account)


def _excluded_test_ids_for_plan(
    standard: str,
    functionality: str,
    gold_rail: str,
) -> set[str]:
    """Test_IDs hidden from the timeline cell selector for the current plan options."""
    excluded: set[str] = set()
    std = str(standard or "")
    if std.startswith("EIA"):
        excluded.update({"SV0126", "SV0125", "SV0127", "SV0128"})
    if "OCP" in std:
        excluded.update({"SV0117", "SV0118", "SV0105", "SV0107"})
    if functionality == "Non-functional":
        excluded.update(
            {
                "SV0107",
                "SV0127",
                "SV0128",
                "SV0105",
                "REL0164_OM",
                "REL0164_DNP",
            }
        )
    if gold_rail == "No":
        excluded.add("ENG0013791_GRS")
    return excluded


def _set_nre_selected_ids(ids_key: str, ids: list[str]) -> None:
    """on_click helper: set NRE multiselect selection before widgets render."""
    st.session_state[ids_key] = list(ids)


def _select_options(
    info: dict,
    *,
    standard: str = "",
    functionality: str = "",
    gold_rail: str = "",
) -> list[str]:
    """Dropdown choices for calendar cells (Excel-like), filtered by plan options."""
    excluded = _excluded_test_ids_for_plan(standard, functionality, gold_rail)
    opts = ["", CELL_AHEAD_OPT, CELL_POSTPONE_OPT]
    for tid, meta in info.items():
        if str(tid).strip() in excluded:
            continue
        opts.append(f"{meta['Abbrv_Name']} ({tid})")
    return opts


def _parse_editor_value(value: str) -> tuple[str, str, bool] | None:
    """
    Parse a selectbox cell value.
    Returns (test_id, abbrv, is_empty), or None for blank/clear.
    """
    if value is None:
        return None
    text = str(value).strip()
    if text == "":
        return None
    if text == EMPTY_LABEL:
        return EMPTY_TOKEN, EMPTY_LABEL, True
    if text.endswith(")") and " (" in text:
        abbrv, tid_part = text.rsplit(" (", 1)
        return tid_part[:-1].strip(), abbrv.strip(), False
    return None


def _clear_plan_work() -> None:
    st.session_state.timeline = None
    st.session_state.export_df = None
    st.session_state.nre_df = None
    st.session_state.hc_phases = None
    st.session_state.hc_nre_hours_by_phase = None
    st.session_state.hc_nre_func_by_phase = None
    st.session_state.hc_df = None
    st.session_state.hc_detail_df = None
    st.session_state.hc_table_phases = None
    st.session_state.hc_table_months = None
    st.session_state.plan_log_context = None
    _clear_hc_default_seeds()
    st.session_state.editor_version = int(st.session_state.editor_version) + 1
    st.session_state.editor_error = None


def _hc_month_default(phase: str) -> int:
    return int(HC_MONTH_DEFAULTS.get(str(phase), 3))


def _hc_eng_slug(engineer: str) -> str:
    return (
        str(engineer)
        .lower()
        .replace("&", "and")
        .replace(" ", "_")
        .replace("/", "_")
    )


def _hc_task_count_key(phase: str, engineer: str) -> str:
    return f"hc_ntasks_{phase}_{_hc_eng_slug(engineer)}"


# Bump when Structure / S&V / CAE default task lists change (forces re-seed once).
_HC_DEFAULTS_SEED_VERSION = 15
_HC_NRE_DURATION_TASK = "S&V Test"
_HC_BASELINE_TASK = "Baseline"
_HC_NRE_HOURS_PER_PERSON = 160.0  # lab hours → 1.0 HeadCount per Person
_HC_NRE_DURATION_STRUCTURE_PEOPLE = 2
_HC_NRE_DURATION_SV_PEOPLE = 3
_HC_SV_BASELINE_PEOPLE = 3
# System weight > 91 kg (~200 lbs) → +1 Number of HeadCount on NRE Duration
# (Structure + S&V) and S&V Baseline.
_HC_HEAVY_WEIGHT_KG = 91.0
_HC_NRE_DURATION_STRUCTURE_PHASES = frozenset({"Concept", "BU", "BCT"})
_HC_NRE_DURATION_SV_PHASES = frozenset({"CT", "NT", "OT", "NOT"})
# S&V Baseline only when that phase has NRE test hours (CT → NT/OT/NOT)
_HC_SV_BASELINE_PHASES = frozenset({"CT", "NT", "OT", "NOT"})
_HC_CAE_OVERALL_RISK_PHASES = frozenset({"BU", "BCT"})
_HC_CAE_PRELIM_TASK = "Simulation - Preliminary Risk Assessment"
_HC_CAE_OVERALL_TASK = "Simulation - Overall Risk Assessment"
_HC_CAE_ISSUE_TASK = "Simulation - Issue Analysis"


def _hc_task_total(number_of_headcount: int, headcount_per_person: float) -> float:
    """Total task headcount = Number of HeadCount × HeadCount per Person."""
    return float(number_of_headcount or 0) * float(headcount_per_person or 0)


def _hc_weight_people_bonus(weight_kg: float | None = None) -> int:
    """+1 Number of HeadCount when system weight is over 91 kg (200 lbs)."""
    w = weight_kg
    if w is None:
        w = st.session_state.get("started_weight_kg")
    try:
        return 1 if float(w or 0) > _HC_HEAVY_WEIGHT_KG else 0
    except (TypeError, ValueError):
        return 0


def _nre_phase_hours(
    cfg: dict,
    *,
    info_map: dict[str, dict] | None = None,
) -> float:
    """
    Sum Duration_for_NRE (hours) for NRE items in one phase config:
    selected test_ids + quantity extras (UFIT / OM / DnP) × qty.
    """
    info = info_map or {}
    hours = 0.0
    for tid in cfg.get("test_ids") or []:
        meta = info.get(str(tid))
        if not meta:
            continue
        hours += float(meta.get("Duration_for_NRE") or 0)
    for tid, qty in (cfg.get("qty_by_id") or {}).items():
        q = int(qty or 0)
        if q <= 0:
            continue
        meta = info.get(str(tid))
        if not meta:
            continue
        hours += float(meta.get("Duration_for_NRE") or 0) * q
    return hours


def _nre_duration_task(
    hours: float, *, number_of_headcount: int = _HC_NRE_DURATION_STRUCTURE_PEOPLE
) -> dict | None:
    """
    NRE hours → HeadCount per Person = hours / 160.
    Structure default people = 2; S&V default people = 3.
    Example: 16 h → 0.1 per person × N people.
    """
    h = float(hours or 0)
    if h <= 0:
        return None
    return {
        "task": _HC_NRE_DURATION_TASK,
        "number_of_headcount": int(number_of_headcount),
        "headcount_per_person": round(h / _HC_NRE_HOURS_PER_PERSON, 6),
        "distribution": HC_DIST_PER_PHASE,
    }


def _phase_has_sv_test_task(phase: str, nre_hours: float) -> bool:
    """True when Structure or S&V would get an S&V Test task in this phase."""
    p = str(phase)
    if float(nre_hours or 0) <= 0:
        return False
    return (
        p in _HC_NRE_DURATION_STRUCTURE_PHASES
        or p in _HC_NRE_DURATION_SV_PHASES
    )


def _clear_hc_default_seeds() -> None:
    """Force headcount default tasks to re-seed (e.g. after NRE regenerate)."""
    for k in list(st.session_state.keys()):
        if str(k).startswith("_hc_defaults_ver_"):
            del st.session_state[k]


def _structure_default_tasks(
    phase: str,
    *,
    nre_hours: float = 0.0,
    people_bonus: int = 0,
) -> list[dict]:
    """
    Structure Engineer defaults (number_of_headcount=1 unless noted):
      All phases: Regular Meeting / Discussion — 0.1 per person Per Month
      Concept: Q&A, Document Study & NRE & Test Plan Analysis — 1 per person Per Phase
      BCT: Issue Analysis & Discussion — 3 per person Per Phase;
           S&V Sample Preparation — 1 per person Per Phase
      BU / CT (split of BCT): same tasks at BU:CT = 2:1
      NOT: Issue Analysis & Discussion — 1 per person Per Phase
      NT / OT (split of NOT): same task at NT:OT = 1:1 (0.5 each)
      Concept / BU / BCT: S&V Test — (Σ hours / 160) × (2 + heavy)
    """
    bonus = max(int(people_bonus or 0), 0)

    def _task(
        name: str,
        per_person: float,
        distribution: str,
        *,
        number: int = 1,
    ) -> dict:
        return {
            "task": name,
            "number_of_headcount": number,
            "headcount_per_person": per_person,
            "distribution": distribution,
        }

    tasks = [
        _task("Regular Meeting / Discussion", 0.1, HC_DIST_PER_MONTH),
    ]
    p = str(phase)
    if p == "Concept":
        tasks.append(
            _task(
                "Q&A, Document Study & NRE & Test Plan Analysis",
                1.0,
                HC_DIST_PER_PHASE,
            )
        )
    elif p == "BCT":
        tasks.extend(
            [
                _task("Issue Analysis & Discussion", 3.0, HC_DIST_PER_PHASE),
                _task("S&V Sample Preparation", 1.0, HC_DIST_PER_PHASE),
            ]
        )
    elif p == "BU":
        tasks.extend(
            [
                _task("Issue Analysis & Discussion", 2.0, HC_DIST_PER_PHASE),
                _task("S&V Sample Preparation", round(2.0 / 3.0, 6), HC_DIST_PER_PHASE),
            ]
        )
    elif p == "CT":
        tasks.extend(
            [
                _task("Issue Analysis & Discussion", 1.0, HC_DIST_PER_PHASE),
                _task("S&V Sample Preparation", round(1.0 / 3.0, 6), HC_DIST_PER_PHASE),
            ]
        )
    elif p == "NOT":
        tasks.append(_task("Issue Analysis & Discussion", 1.0, HC_DIST_PER_PHASE))
    elif p in ("NT", "OT"):
        tasks.append(_task("Issue Analysis & Discussion", 0.5, HC_DIST_PER_PHASE))

    if p in _HC_NRE_DURATION_STRUCTURE_PHASES:
        nre_task = _nre_duration_task(
            nre_hours,
            number_of_headcount=_HC_NRE_DURATION_STRUCTURE_PEOPLE + bonus,
        )
        if nre_task:
            tasks.append(nre_task)
    return tasks


def _cae_default_tasks(
    phase: str,
    *,
    nre_hours: float = 0.0,
) -> list[dict]:
    """
    CAE Engineer defaults:
      Concept only: Simulation - Preliminary Risk Assessment — 1 × 1.0 Per Phase
      BU / BCT: Simulation - Overall Risk Assessment — 1 × 3.0 Per Phase
      Any phase with Structure/S&V S&V Test: Simulation - Issue Analysis
        — 1 × 0.2 Per Phase
    """
    p = str(phase)
    tasks: list[dict] = []
    if p == "Concept":
        tasks.append(
            {
                "task": _HC_CAE_PRELIM_TASK,
                "number_of_headcount": 1,
                "headcount_per_person": 1.0,
                "distribution": HC_DIST_PER_PHASE,
            }
        )
    elif p in _HC_CAE_OVERALL_RISK_PHASES:
        tasks.append(
            {
                "task": _HC_CAE_OVERALL_TASK,
                "number_of_headcount": 1,
                "headcount_per_person": 3.0,
                "distribution": HC_DIST_PER_PHASE,
            }
        )
    if _phase_has_sv_test_task(p, nre_hours):
        tasks.append(
            {
                "task": _HC_CAE_ISSUE_TASK,
                "number_of_headcount": 1,
                "headcount_per_person": 0.2,
                "distribution": HC_DIST_PER_PHASE,
            }
        )
    return tasks


def _sv_baseline_task(functionality: str, *, people_bonus: int = 0) -> dict:
    """
    Baseline for S&V when the phase has NRE test content:
      Functional → (3+heavy) × 0.25 Per Phase
      Non-functional → (3+heavy) × 0.1 Per Phase
    """
    per = (
        0.25
        if str(functionality or "").strip() == "Functional"
        else 0.1
    )
    return {
        "task": _HC_BASELINE_TASK,
        "number_of_headcount": _HC_SV_BASELINE_PEOPLE + max(int(people_bonus or 0), 0),
        "headcount_per_person": per,
        "distribution": HC_DIST_PER_PHASE,
    }


def _sv_default_tasks(
    phase: str,
    *,
    nre_hours: float = 0.0,
    functionality: str = "Functional",
    people_bonus: int = 0,
) -> list[dict]:
    """
    S&V defaults:
      Regular Meeting / Discussion — 0.1 per person Per Month (except Concept)
      CT / NT / OT / NOT: Baseline only if that phase has NRE test hours —
        Functional (3+heavy)×0.25, Non-functional (3+heavy)×0.1 Per Phase
        (NT and OT each get their own Baseline when each has tests)
      CT / NT / OT / NOT: S&V Test — (Σ hours / 160) × (3+heavy)
    """
    bonus = max(int(people_bonus or 0), 0)
    tasks: list[dict] = []
    p = str(phase)
    if p != "Concept":
        tasks.append(
            {
                "task": "Regular Meeting / Discussion",
                "number_of_headcount": 1,
                "headcount_per_person": 0.1,
                "distribution": HC_DIST_PER_MONTH,
            }
        )
    if p in _HC_SV_BASELINE_PHASES and float(nre_hours or 0) > 0:
        tasks.append(_sv_baseline_task(functionality, people_bonus=bonus))
    if p in _HC_NRE_DURATION_SV_PHASES:
        nre_task = _nre_duration_task(
            nre_hours,
            number_of_headcount=_HC_NRE_DURATION_SV_PEOPLE + bonus,
        )
        if nre_task:
            tasks.append(nre_task)
    return tasks


def _seed_hc_engineer_defaults(phase: str, engineer: str) -> None:
    """Seed (or refresh once per version) default tasks for a phase/engineer."""
    n_key = _hc_task_count_key(phase, engineer)
    slug = _hc_eng_slug(engineer)
    ver_key = f"_hc_defaults_ver_{phase}_{slug}"
    people_bonus = _hc_weight_people_bonus()
    # Version tuple includes heavy-weight bonus so crossing 91 kg re-seeds.
    seed_stamp = (_HC_DEFAULTS_SEED_VERSION, int(people_bonus))
    if st.session_state.get(ver_key) == seed_stamp:
        return

    nre_hours = float(
        (st.session_state.get("hc_nre_hours_by_phase") or {}).get(phase, 0) or 0
    )
    functionality = str(
        (st.session_state.get("hc_nre_func_by_phase") or {}).get(
            phase, "Functional"
        )
        or "Functional"
    )
    if engineer == "Structure Engineer":
        defaults = _structure_default_tasks(
            phase, nre_hours=nre_hours, people_bonus=people_bonus
        )
    elif engineer == "CAE Engineer":
        defaults = _cae_default_tasks(phase, nre_hours=nre_hours)
    elif engineer == "S&V Engineer":
        defaults = _sv_default_tasks(
            phase,
            nre_hours=nre_hours,
            functionality=functionality,
            people_bonus=people_bonus,
        )
    else:
        defaults = []

    st.session_state[n_key] = len(defaults)
    for i, task in enumerate(defaults):
        st.session_state[f"hc_task_name_{phase}_{slug}_{i}"] = task["task"]
        st.session_state[f"hc_task_n_{phase}_{slug}_{i}"] = int(
            task.get("number_of_headcount", 1)
        )
        per = task.get("headcount_per_person", task.get("headcount", 0.0))
        st.session_state[f"hc_task_hc_{phase}_{slug}_{i}"] = float(per)
        st.session_state[f"hc_task_dist_{phase}_{slug}_{i}"] = task[
            "distribution"
        ]
        # 1-based within phase; end=999 clamps to current phase month count in UI
        st.session_state[f"hc_task_start_{phase}_{slug}_{i}"] = int(
            task.get("start_month", 1) or 1
        )
        default_end = (
            int(task["end_month"])
            if task.get("end_month") is not None
            else (
                999
                if task.get("distribution") == HC_DIST_PER_PHASE
                else 1
            )
        )
        st.session_state[f"hc_task_end_{phase}_{slug}_{i}"] = default_end
    st.session_state[ver_key] = seed_stamp


def _collect_hc_tasks_for_phase(
    phase: str, *, n_months: int = 0
) -> dict[str, list[dict]]:
    """Read task widgets for one phase → engineer → task list."""
    out: dict[str, list[dict]] = {}
    n_months = max(int(n_months or 0), 0)
    for eng in HC_ENGINEER_TYPES:
        n_key = _hc_task_count_key(phase, eng)
        n_tasks = int(st.session_state.get(n_key, 0) or 0)
        tasks: list[dict] = []
        slug = _hc_eng_slug(eng)
        for i in range(n_tasks):
            name = str(
                st.session_state.get(f"hc_task_name_{phase}_{slug}_{i}", "")
                or ""
            ).strip()
            n_hc = int(
                st.session_state.get(f"hc_task_n_{phase}_{slug}_{i}", 1) or 1
            )
            per_person = float(
                st.session_state.get(f"hc_task_hc_{phase}_{slug}_{i}", 0.0)
                or 0.0
            )
            total_hc = _hc_task_total(n_hc, per_person)
            dist = str(
                st.session_state.get(
                    f"hc_task_dist_{phase}_{slug}_{i}", HC_DIST_PER_MONTH
                )
                or HC_DIST_PER_MONTH
            )
            start_m: int | None = None
            end_m: int | None = None
            if dist == HC_DIST_PER_PHASE:
                start_m = int(
                    st.session_state.get(
                        f"hc_task_start_{phase}_{slug}_{i}", 1
                    )
                    or 1
                )
                end_m = int(
                    st.session_state.get(
                        f"hc_task_end_{phase}_{slug}_{i}",
                        n_months if n_months else 1,
                    )
                    or 1
                )
                if n_months > 0:
                    start_m = max(1, min(start_m, n_months))
                    end_m = max(start_m, min(end_m, n_months))
                else:
                    start_m = max(1, start_m)
                    end_m = max(start_m, end_m)
            if not name and total_hc == 0:
                continue
            tasks.append(
                {
                    "task": name or f"Task {i + 1}",
                    "number_of_headcount": n_hc,
                    "headcount_per_person": per_person,
                    "headcount": total_hc,
                    "distribution": dist,
                    "start_month": start_m,
                    "end_month": end_m,
                }
            )
        out[eng] = tasks
    return out


def _hc_task_field_keys(phase: str, slug: str, i: int) -> list[str]:
    return [
        f"hc_task_name_{phase}_{slug}_{i}",
        f"hc_task_n_{phase}_{slug}_{i}",
        f"hc_task_hc_{phase}_{slug}_{i}",
        f"hc_task_dist_{phase}_{slug}_{i}",
        f"hc_task_start_{phase}_{slug}_{i}",
        f"hc_task_end_{phase}_{slug}_{i}",
    ]


def _render_hc_detail_table(df: pd.DataFrame) -> None:
    """
    Show HC detail with task lines visible (st.dataframe collapses \\n until expand).
    """
    if df is None or df.empty:
        return
    cols = list(df.columns)
    ths = "".join(
        f"<th style='text-align:left;padding:8px 10px;border:1px solid "
        f"#d0d7de;background:#f6f8fa;white-space:nowrap;'>{html.escape(str(c))}</th>"
        for c in cols
    )
    body_rows: list[str] = []
    for _, rec in df.iterrows():
        tds: list[str] = []
        for c in cols:
            raw = rec.get(c, "")
            if raw is None or (isinstance(raw, float) and pd.isna(raw)):
                text = ""
            else:
                text = str(raw)
            cell = html.escape(text).replace("\n", "<br>")
            tds.append(
                "<td style='text-align:left;padding:8px 10px;border:1px solid "
                "#d0d7de;vertical-align:top;white-space:pre-line;'>"
                f"{cell}</td>"
            )
        body_rows.append("<tr>" + "".join(tds) + "</tr>")
    table_html = (
        "<div style='overflow-x:auto;width:100%;'>"
        "<table style='border-collapse:collapse;width:100%;font-size:0.92rem;'>"
        f"<thead><tr>{ths}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def _bump_hc_task_count(phase: str, engineer: str) -> None:
    key = _hc_task_count_key(phase, engineer)
    new_i = int(st.session_state.get(key, 0) or 0)
    st.session_state[key] = new_i + 1
    slug = _hc_eng_slug(engineer)
    st.session_state[f"hc_task_n_{phase}_{slug}_{new_i}"] = 1
    st.session_state[f"hc_task_hc_{phase}_{slug}_{new_i}"] = 0.0
    st.session_state[f"hc_task_dist_{phase}_{slug}_{new_i}"] = HC_DIST_PER_MONTH
    st.session_state[f"hc_task_start_{phase}_{slug}_{new_i}"] = 1
    st.session_state[f"hc_task_end_{phase}_{slug}_{new_i}"] = 1


def _shrink_hc_task_count(phase: str, engineer: str) -> None:
    key = _hc_task_count_key(phase, engineer)
    n = int(st.session_state.get(key, 0) or 0)
    if n <= 0:
        return
    _remove_hc_task_at(phase, engineer, n - 1)


def _remove_hc_task_at(phase: str, engineer: str, index: int) -> None:
    """Remove one task at index and shift later tasks down."""
    key = _hc_task_count_key(phase, engineer)
    n = int(st.session_state.get(key, 0) or 0)
    if index < 0 or index >= n:
        return
    slug = _hc_eng_slug(engineer)
    for j in range(index, n - 1):
        src_keys = _hc_task_field_keys(phase, slug, j + 1)
        dest_keys = _hc_task_field_keys(phase, slug, j)
        for src, dest in zip(src_keys, dest_keys):
            if src in st.session_state:
                st.session_state[dest] = st.session_state[src]
            else:
                st.session_state.pop(dest, None)
    for field in _hc_task_field_keys(phase, slug, n - 1):
        st.session_state.pop(field, None)
    st.session_state[key] = n - 1


def _record_download(tool_type: str, **overrides) -> None:
    """Write a usage-log row for TEST PLAN / NRE / HC downloads.

    NRE and HC log only Account / Project / Weight; TEST PLAN logs full context.
    """
    ctx = dict(st.session_state.get("plan_log_context") or {})
    ctx.update({k: v for k, v in overrides.items() if v is not None})
    tool = str(tool_type or "").strip()
    if tool in ("NRE", "HC"):
        append_usage_record(
            tool_type=tool,
            account=ctx.get("account", ""),
            project=ctx.get("project", ""),
            weight=ctx.get("weight", ""),
        )
        return
    append_usage_record(
        tool_type=tool,
        account=ctx.get("account", ""),
        project=ctx.get("project", ""),
        phase=ctx.get("phase", ""),
        weight=ctx.get("weight", ""),
        system_eta=ctx.get("system_eta", ""),
        critical_feedback=ctx.get("critical_feedback", ""),
        system_number=ctx.get("system_number", ""),
        functional=ctx.get("functional", ""),
    )


# ---------------------------------------------------------------------------
# Main page — setup
# ---------------------------------------------------------------------------
st.title("WTK Test Planner")
st.caption("Schedule · NRE · Headcount")

st.subheader("Setup")
s1, s2, s3 = st.columns(3)
with s1:
    account = st.selectbox("Account (Brand)", ACCOUNTS, index=0, key="setup_account")
with s2:
    weight_kg = st.number_input(
        "System weight (kg)", min_value=0.0, value=25.0, step=0.5, key="setup_weight"
    )
with s3:
    standard = st.selectbox("Standard", STANDARDS, index=0, key="setup_standard")

s4, _, _ = st.columns(3)
with s4:
    project_name = st.text_input(
        "Project name",
        value="",
        key="setup_project_name",
        placeholder="e.g. Project Apollo",
    )

# Switching account before start clears stale catalog-bound work
if not st.session_state.plan_started:
    if st.session_state.active_account != account:
        st.session_state.active_account = account
        _clear_plan_work()

start_col, reset_col, _ = st.columns([2, 1, 3])
with start_col:
    if st.button("Start Arranging Test Plan", type="primary", use_container_width=True):
        missing: list[str] = []
        if not account:
            missing.append("Account (Brand)")
        if weight_kg is None:
            missing.append("System weight (kg)")
        if not standard:
            missing.append("Standard")
        if not (project_name or "").strip():
            missing.append("Project name")
        if missing:
            st.warning(
                "Please fill in all setup fields before starting: **"
                + "**, **".join(missing)
                + "**."
            )
        else:
            prev = st.session_state.started_account
            st.session_state.plan_started = True
            st.session_state.started_account = account
            st.session_state.started_weight_kg = weight_kg
            st.session_state.started_standard = standard
            st.session_state.started_project_name = (project_name or "").strip()
            st.session_state.active_account = account
            if prev is not None and prev != account:
                _clear_plan_work()
                _load_working_catalog_from_disk(account)
            st.rerun()
with reset_col:
    if st.session_state.plan_started and st.button("Reset setup", use_container_width=True):
        st.session_state.plan_started = False
        st.session_state.started_account = None
        st.session_state.started_weight_kg = None
        st.session_state.started_standard = None
        st.session_state.started_project_name = None
        st.session_state.started_project_phase = None
        st.session_state.pop("tp_project_phase", None)
        _clear_plan_work()
        st.rerun()

# Working catalog (session only — never overwrites folder files until we choose to)
try:
    _ensure_working_catalog(account)
except Exception as exc:
    st.error(f"Failed to load data for {account}: {exc}")
    st.stop()

test_plan_df = st.session_state.work_test_plan
location_df = st.session_state.work_location
convert_df = st.session_state.work_convert
info_map = test_info_by_id(test_plan_df)

if not st.session_state.plan_started:
    st.info(
        "Fill in **all** setup fields (**Account**, **System weight**, **Standard**, "
        "**Project name**), then click **Start Arranging Test Plan**."
    )
else:
    account = st.session_state.started_account or account
    weight_kg = (
        st.session_state.started_weight_kg
        if st.session_state.started_weight_kg is not None
        else weight_kg
    )
    standard = st.session_state.started_standard or standard
    project_name = st.session_state.started_project_name or project_name or ""
    project_phase = (
        st.session_state.get("tp_project_phase")
        or st.session_state.get("started_project_phase")
        or PROJECT_PHASES[0]
    )

    proj_bit = f"**Project:** {project_name} &nbsp;|&nbsp; " if project_name else ""
    st.write(
        f"{proj_bit}**Phase:** {project_phase} &nbsp;|&nbsp; "
        f"**Account:** {account} &nbsp;|&nbsp; **Weight:** {weight_kg:g} kg "
        f"&nbsp;|&nbsp; **Standard:** {standard}"
    )

    def _export_stem() -> str:
        bits = [account, project_phase]
        if project_name:
            safe = "".join(
                c if c.isalnum() or c in "-_" else "_" for c in project_name
            ).strip("_")
            if safe:
                bits.insert(1, safe)
        return "_".join(bits)

    try:
        _ensure_working_catalog(account)
    except Exception as exc:
        st.error(f"Failed to load data for {account}: {exc}")
        st.stop()
    test_plan_df = st.session_state.work_test_plan
    location_df = st.session_state.work_location
    convert_df = st.session_state.work_convert
    info_map = test_info_by_id(test_plan_df)

    tab_plan, tab_nre_hc = st.tabs(
        ["Test Plan Timeline", "NRE / Headcount"]
    )


    if account != "ROSA":
        with tab_plan:
            st.subheader("Test Plan Timeline")
            st.info(
                "TBD — NAOMI uses a different test-plan workflow "
                "(not implemented yet)."
            )
        with tab_nre_hc:
            st.subheader("NRE / Headcount")
            st.info(
                "TBD — NAOMI NRE and headcount estimation are not "
                "implemented yet."
            )
    else:
        # ============================= TEST PLAN ==================================
        with tab_plan:
            st.subheader(f"1. {account} Test Plan Options")

            left, right = st.columns(2)
            with left:
                project_phase = st.selectbox(
                    "Phase",
                    PROJECT_PHASES,
                    index=0,
                    key="tp_project_phase",
                )
                st.session_state.started_project_phase = project_phase
                system_eta = st.date_input(
                    "System ETA", value=date.today() + timedelta(days=7)
                )
                critical_fb = st.date_input(
                    "Critical feedback", value=date.today() + timedelta(days=21)
                )
                n_systems = st.selectbox("Number of System", SYSTEM_NUMBERS, index=0)
                if int(n_systems) > SEQUENCE_TEMPLATE_MAX_SYSTEMS:
                    st.caption(
                        f"Convert table & sequence sheets only cover "
                        f"**1–{SEQUENCE_TEMPLATE_MAX_SYSTEMS}** systems. "
                        f"For **{n_systems}** systems, fill the timeline yourself."
                    )
                st.caption(
                    "Timeline starts at **System ETA** and ends at "
                    "**5 business days after Critical feedback** "
                    "(or later if the filled plan runs past that)."
                )
            with right:
                functionality = st.radio(
                    "Functional / Non-functional",
                    FUNCTIONALITY_OPTS,
                    horizontal=True,
                    key="tp_func",
                )
                is_ocp = "OCP" in str(standard)
                if is_ocp:
                    ufit_sor = st.radio(
                        ORV3_MGX_WO_L11_LABEL,
                        YES_NO,
                        index=1,
                        horizontal=True,
                        key="tp_ufit",
                    )
                else:
                    ufit_sor = "No"
                gold_rail = st.radio(
                    "Gold Rail Selection for this plan?",
                    YES_NO,
                    horizontal=True,
                    key="tp_gold",
                )

            n_sys = int(n_systems)
            template_supported = n_sys <= SEQUENCE_TEMPLATE_MAX_SYSTEMS
            convert_id = None
            if template_supported:
                convert_id = lookup_convert_id(
                    account,
                    functionality,
                    gold_rail,
                    ufit_sor,
                    n_sys,
                    standard=standard,
                    convert_df=convert_df,
                )
            st.session_state.last_convert_id = convert_id
            if convert_id:
                st.info(
                    f"Case Convert_ID = **{convert_id}** "
                    f"(Standard={standard}, Functionality={functionality}, "
                    f"Gold Rail={gold_rail}"
                    + (
                        f", {ORV3_MGX_WO_L11_LABEL}={ufit_sor}"
                        if is_ocp
                        else ""
                    )
                    + f", Systems={n_sys})"
                )
            elif not template_supported:
                st.info(
                    f"**{n_sys} systems** — no Convert_ID / Case sequence template "
                    f"(templates only for 1–{SEQUENCE_TEMPLATE_MAX_SYSTEMS}). "
                    "Generate a blank timeline and fill each system row yourself."
                )

            preload = False
            if template_supported:
                preload = st.checkbox(
                    "Pre-fill timeline from sequence sheet (rename Row labels freely; "
                    "one token per cell: 1 | T003 | T004 | …)",
                    value=False,
                )

            gen_col, _ = st.columns([1, 3])
            with gen_col:
                generate = st.button(
                    "Generate Test Plan", type="primary", use_container_width=True
                )

            if generate:
                if critical_fb < system_eta:
                    st.warning(
                        "Critical feedback is before System ETA — timeline still starts at ETA."
                    )
                tl = init_timeline_state(
                    system_eta,
                    critical_fb,
                    n_sys,
                )
                msgs: list[str] = []
                if preload and convert_id:
                    seq = load_case_sequence(account, convert_id)
                    if seq is not None:
                        tl, msgs = apply_sequence_template(
                            tl, seq, system_eta.isoformat(), info=info_map
                        )
                    else:
                        msgs.append(
                            f"Sequence sheet Case_{int(convert_id):02d} not found in "
                            f"data/{account}/test_item_sequence_all_cases.xlsx."
                        )
                elif preload and not convert_id:
                    msgs.append(
                        "No Convert_ID for these options — timeline left blank to fill manually."
                    )
                st.session_state.timeline = tl
                st.session_state.export_df = None
                st.session_state.editor_version = int(st.session_state.editor_version) + 1
                st.session_state.editor_error = None
                st.session_state.plan_log_context = {
                    "account": account,
                    "project": project_name,
                    "phase": project_phase,
                    "weight": weight_kg,
                    "system_eta": system_eta,
                    "critical_feedback": critical_fb,
                    "system_number": n_sys,
                    "functional": functionality,
                }
                if msgs:
                    for m in msgs:
                        st.warning(m)
                st.success(
                    f"Timeline generated: {tl['dates'][0]} → {tl['dates'][-1]} "
                    f"(Weekend/Holiday columns locked)."
                )
            timeline = st.session_state.timeline
            if timeline is not None:
                # Keep log context fresh while options widgets are on screen
                st.session_state.plan_log_context = {
                    "account": account,
                    "project": project_name,
                    "phase": project_phase,
                    "weight": weight_kg,
                    "system_eta": system_eta,
                    "critical_feedback": critical_fb,
                    "system_number": n_sys,
                    "functional": functionality,
                }
            if timeline is None:
                st.caption("Configure options above, then click **Generate Test Plan**.")
            else:
                st.subheader("2. Timeline")
                st.caption(
                    "Top table = Weekday / Event / Detail. "
                    "Event: **-** (none), **Occupied** (unfillable day), "
                    "**Critical Event** (extra System ETA / Critical Feedback). "
                    "Detail: type a comment, or for Critical Event pick "
                    "**System ETA** / **Critical Feedback** from the list, "
                    "or choose **Custom (type…)** to enter free text. "
                    "Occupied days stay white and cannot be filled. "
                    "Multi-day items keep their start day and skip Occupied/weekend "
                    "(e.g. 09/30–10/02 with Occupied on 10/01 → 09/30, 10/02, 10/05); "
                    "blank days before the start stay blank. "
                    "Weekend / Holiday columns are light red. "
                    "Edit the first **Row** cell to rename a system (e.g. DUT-A). "
                    "The **Profile** row under each system shows weight-band notes "
                    "(truncated; **double-click** a cell to read the full text). "
                    "Row buttons Ahead/Postpone shift the **whole** system line. "
                    "In a cell dropdown, **Ahead (−1) from here** / **Postpone (+1) from here** "
                    "shift only items from that date onward (blocked with a warning if Ahead would overlap)."
                )

                options = _select_options(
                    info_map,
                    standard=standard,
                    functionality=functionality,
                    gold_rail=gold_rail,
                )
                if st.session_state.editor_error:
                    st.error(st.session_state.editor_error)
                if st.session_state.get("refresh_msg"):
                    st.success(st.session_state.refresh_msg)
                    st.session_state.refresh_msg = None

                ref_col, clear_col, _ = st.columns([1, 1, 2])
                with ref_col:
                    if st.button(
                        "Refresh durations from loaded data",
                        use_container_width=True,
                        help=(
                            "Update each placed item’s length from the session catalog "
                            "(Data preview → Update). Keeps current sequence / starts; "
                            "does not reload the Case template."
                        ),
                    ):
                        fresh_info = test_info_by_id(st.session_state.work_test_plan)
                        new_tl, refresh_errs = refresh_durations_from_catalog(
                            st.session_state.timeline, fresh_info
                        )
                        st.session_state.timeline = new_tl
                        st.session_state.export_df = None
                        st.session_state.editor_version = (
                            int(st.session_state.editor_version) + 1
                        )
                        if refresh_errs:
                            st.session_state.editor_error = "; ".join(refresh_errs)
                            st.session_state.refresh_msg = None
                        else:
                            st.session_state.editor_error = None
                            st.session_state.refresh_msg = (
                                "Durations refreshed from loaded data. Sequence unchanged."
                            )
                        st.rerun()
                with clear_col:
                    if st.button(
                        "Clear all test items",
                        use_container_width=True,
                        help=(
                            "Remove every placed test item from all system rows. "
                            "Keeps dates, Occupied / Critical Event marks, and system names."
                        ),
                    ):
                        st.session_state.timeline = clear_all_test_items(
                            st.session_state.timeline
                        )
                        st.session_state.export_df = None
                        st.session_state.editor_error = None
                        st.session_state.editor_version = (
                            int(st.session_state.editor_version) + 1
                        )
                        st.session_state.refresh_msg = (
                            "All test items cleared. Occupied / Critical Event marks kept."
                        )
                        st.rerun()

                ensure_event_state(timeline)
                detail_map = test_detail_by_id(st.session_state.work_test_plan)
                event = render_synced_calendar(
                    timeline,
                    options,
                    weight_kg=float(weight_kg),
                    detail_by_id=detail_map,
                    key="synced_cal",
                )

                nonce = event.get("nonce") if isinstance(event, dict) else None
                if (
                    event
                    and isinstance(event, dict)
                    and nonce is not None
                    and nonce != st.session_state.last_cal_nonce
                ):
                    st.session_state.last_cal_nonce = nonce
                    kind = str(event.get("kind") or "system")
                    err_msg = None
                    new_tl = timeline

                    if kind == "rename_system":
                        sk = str(event.get("system", ""))
                        new_name = (
                            "" if event.get("name") is None else str(event.get("name"))
                        )
                        new_tl, err_msg = rename_system_row(timeline, sk, new_name)
                    elif kind == "shift_system":
                        sk = str(event.get("system", ""))
                        try:
                            delta = int(event.get("delta", 0))
                        except (TypeError, ValueError):
                            delta = 0
                        new_tl, shift_errs = shift_system_schedule(timeline, sk, delta)
                        if shift_errs and new_tl is timeline:
                            err_msg = "; ".join(shift_errs)
                        elif shift_errs:
                            st.session_state.editor_error = None
                            for msg in shift_errs:
                                if "could not" in msg.lower():
                                    err_msg = msg
                                    break
                    elif kind == "shift_from_date":
                        sk = str(event.get("system", ""))
                        d = str(event.get("date", ""))
                        try:
                            delta = int(event.get("delta", 0))
                        except (TypeError, ValueError):
                            delta = 0
                        new_tl, shift_errs = shift_system_from_date(timeline, sk, d, delta)
                        if shift_errs:
                            err_msg = "; ".join(shift_errs)
                    elif kind == "mark_event" or kind == "mark_empty":
                        d = str(event.get("date", ""))
                        after = "" if event.get("value") is None else str(event.get("value"))
                        new_tl, shift_errs = set_calendar_event(timeline, d, after)
                        if shift_errs:
                            err_msg = "; ".join(shift_errs)
                    elif kind == "event_detail":
                        d = str(event.get("date", ""))
                        after = "" if event.get("value") is None else str(event.get("value"))
                        new_tl, shift_errs = set_event_detail(timeline, d, after)
                        if shift_errs:
                            err_msg = "; ".join(shift_errs)
                    else:
                        sk = str(event.get("system", ""))
                        d = str(event.get("date", ""))
                        after = "" if event.get("value") is None else str(event.get("value"))
                        if after in (CELL_AHEAD_OPT, CELL_POSTPONE_OPT):
                            delta = -1 if after == CELL_AHEAD_OPT else 1
                            new_tl, shift_errs = shift_system_from_date(
                                timeline, sk, d, delta
                            )
                            if shift_errs:
                                err_msg = "; ".join(shift_errs)
                        elif not d:
                            err_msg = "Missing date for cell edit."
                        elif d in set(timeline.get("blocked", [])):
                            err_msg = f"{d} is Weekend/Holiday/Occupied and cannot be filled."
                        elif sk not in timeline.get("grid", {}):
                            err_msg = f"Unknown system row: {sk}"
                        else:
                            new_tl = clear_cell_span(new_tl, sk, d)
                            parsed = _parse_editor_value(after)
                            if parsed is not None:
                                tid, abbrv, is_empty = parsed
                                if is_empty:
                                    new_tl, err_msg = place_item(
                                        new_tl, sk, d, EMPTY_TOKEN, EMPTY_LABEL, 1, True
                                    )
                                elif tid not in info_map:
                                    err_msg = f"Unknown test item: {after}"
                                else:
                                    dur = int(info_map[tid]["Duration_Days"])
                                    abbrv = info_map[tid]["Abbrv_Name"]
                                    new_tl, err_msg = place_item(
                                        new_tl, sk, d, tid, abbrv, dur, False
                                    )

                    if err_msg:
                        st.session_state.editor_error = err_msg
                    else:
                        st.session_state.timeline = new_tl
                        st.session_state.editor_error = None
                    old_dates = timeline.get("dates", [])
                    new_dates = (new_tl or timeline).get("dates", [])
                    old_sys = list(timeline.get("grid", {}).keys())
                    new_sys = list((new_tl or timeline).get("grid", {}).keys())
                    if old_dates != new_dates or old_sys != new_sys:
                        st.session_state.editor_version = (
                            int(st.session_state.editor_version) + 1
                        )
                    st.rerun()

                st.divider()
                st.subheader("3. Update → Export Table")
                if st.button("Update", type="primary"):
                    detail_map = test_detail_by_id(st.session_state.work_test_plan)
                    st.session_state.export_df = timeline_to_export_df(
                        st.session_state.timeline,
                        weight_kg=float(weight_kg),
                        detail_by_id=detail_map,
                    )
                    st.success(
                        "Timeline converted to export table "
                        "(system schedule row + profile row; Excel merges the system name)."
                    )

                if st.session_state.export_df is not None:
                    export_styled = style_timeline_display(
                        st.session_state.export_df, st.session_state.timeline
                    )
                    st.dataframe(export_styled, use_container_width=True, hide_index=True)
                    xlsx_bytes = export_timeline_xlsx(
                        st.session_state.export_df, st.session_state.timeline
                    )
                    if st.download_button(
                        "Download timeline Excel (with colors)",
                        data=xlsx_bytes,
                        file_name=f"{_export_stem()}_timeline_{date.today().isoformat()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ):
                        _record_download("TEST PLAN")
                        st.toast("Usage recorded · TEST PLAN")

        # ============================= NRE / HEADCOUNT ============================
        with tab_nre_hc:
            st.subheader("NRE Estimation")
            st.caption(
                "Select one or more testing phases. Fixture / rack USD costs are "
                "one-time (under Testing Phase) and land on the earliest selected "
                "phase in the Excel template. Each phase has its own options + "
                "test items. "
                "BCT is exclusive with BU/CT; NOT is exclusive with NT/OT. "
                "Generated table matches `NRE_TEMPLATE_<account>.xlsx` "
                "(Test Item · Location · Rate · Concept/BU/CT/BCT/NT/OT/NOT · "
                "Sub Total). "
                "**Sub Total** in Excel = `SUM(Concept:NOT) × Lab Rate`. "
                "After **Generate NRE Table**, Headcount Estimation appears below."
            )

            _nre_default_phases = sanitize_nre_phases(list(NRE_DEFAULT_PHASES))
            if "nre_phases" not in st.session_state:
                st.session_state.nre_phases = list(_nre_default_phases)
            else:
                cleaned = sanitize_nre_phases(
                    list(st.session_state.nre_phases or [])
                )
                opts_now = available_nre_phases(cleaned)
                st.session_state.nre_phases = [
                    p for p in cleaned if p in opts_now
                ] or list(_nre_default_phases)
            nre_phase_options = available_nre_phases(
                list(st.session_state.nre_phases or [])
            )
            nre_phases = st.multiselect(
                "Testing Phase (multi-select)",
                nre_phase_options,
                key="nre_phases",
                help=(
                    "Default: Concept, BU, CT, NT, OT. "
                    "OT is listed with no tests / U-fit by default. "
                    "BU/CT hide BCT (and the reverse). "
                    "NT/OT hide NOT (and the reverse). "
                    "Generated columns always use order: "
                    "Concept → BU → CT → BCT → NT → OT → NOT."
                ),
            )
            # Fixed generate / UI section order (ignore multiselect pick order)
            nre_phases = sanitize_nre_phases(list(nre_phases))

            st.caption(
                "One-time fixture / rack charges (USD) — applied once to the "
                "earliest selected phase "
                f"({first_nre_phase(list(nre_phases)) or '—'})."
            )
            aux_costs: dict[str, float] = {}
            aux_cols = st.columns(len(NRE_AUX_COST_IDS))
            for col, tid in zip(aux_cols, NRE_AUX_COST_IDS):
                label = (
                    info_map[tid]["Abbrv_Name"] if tid in info_map else tid
                )
                with col:
                    aux_costs[tid] = float(
                        st.number_input(
                            f"{label} (USD)",
                            min_value=0.0,
                            step=100.0,
                            value=float(
                                NRE_AUX_COST_DEFAULTS_USD.get(tid, 0.0)
                            ),
                            key=f"nre_aux_cost_{tid}",
                        )
                    )

            all_ids = list(test_plan_df["Test_ID"].astype(str))
            phase_configs: list[dict] = []

            for phase in nre_phases:
                st.divider()
                st.markdown(f"### {phase}")
                func_def, gold_def = NRE_PHASE_OPTION_DEFAULTS.get(
                    phase, ("Functional", "No")
                )
                func_key = f"nre_func_{phase}"
                gold_key = f"nre_gold_{phase}"
                if func_key not in st.session_state:
                    st.session_state[func_key] = func_def
                if gold_key not in st.session_state:
                    st.session_state[gold_key] = gold_def
                is_ocp = "OCP" in str(standard)
                if is_ocp:
                    pc1, pc2, pc3 = st.columns(3)
                else:
                    pc1, pc3 = st.columns(2)
                    pc2 = None
                with pc1:
                    functionality = st.radio(
                        "Functional / Non-functional",
                        FUNCTIONALITY_OPTS,
                        horizontal=True,
                        key=func_key,
                    )
                if is_ocp and pc2 is not None:
                    with pc2:
                        ufit_sor = st.radio(
                            ORV3_MGX_WO_L11_LABEL,
                            YES_NO,
                            index=1,
                            horizontal=True,
                            key=f"nre_ufit_{phase}",
                        )
                else:
                    ufit_sor = "No"
                with pc3:
                    gold_rail = st.radio(
                        "Gold Rail Selection",
                        YES_NO,
                        horizontal=True,
                        key=gold_key,
                    )

                ids_key = f"nre_ids_{phase}"
                excluded = _excluded_test_ids_for_plan(
                    standard, functionality, gold_rail
                )
                # OM / DnP / U-fit / AUX costs are entered outside the list
                allowed_ids = [
                    tid
                    for tid in all_ids
                    if str(tid).strip() not in excluded
                    and str(tid).strip() not in NRE_SELECTOR_EXTRA_IDS
                ]
                empty_content_default = phase in NRE_EMPTY_CONTENT_DEFAULT_PHASES
                if ids_key not in st.session_state:
                    # Most phases: Select all for current filters.
                    # OT (and similar): start empty — still listed in NRE/HC.
                    st.session_state[ids_key] = (
                        [] if empty_content_default else list(allowed_ids)
                    )
                else:
                    # Drop selections that the current EIA/OCP · Func · Gold filters exclude
                    kept = [
                        i
                        for i in list(st.session_state[ids_key])
                        if i in allowed_ids
                    ]
                    if kept != list(st.session_state[ids_key]):
                        st.session_state[ids_key] = kept

                selected_ids = st.multiselect(
                    f"Test items included in NRE — {phase}",
                    options=allowed_ids,
                    format_func=lambda tid: (
                        f"{info_map[tid]['Abbrv_Name']} ({tid})"
                        if tid in info_map
                        else str(tid)
                    ),
                    key=ids_key,
                    help=(
                        "Same filters as Test Plan: Standard (EIA/OCP), "
                        "Functional / Non-functional, and Gold Rail. "
                        "OM / DnP and U-fit quantity are below; "
                        "fixture / rack costs are under Testing Phase. "
                        + (
                            "OT defaults to no test items (phase still listed)."
                            if empty_content_default
                            else ""
                        )
                    ),
                )
                sel_col, clr_col, _ = st.columns([1, 1, 4])
                with sel_col:
                    st.button(
                        "Select all",
                        key=f"nre_sel_all_{phase}",
                        use_container_width=True,
                        on_click=_set_nre_selected_ids,
                        args=(ids_key, list(allowed_ids)),
                    )
                with clr_col:
                    st.button(
                        "Remove all",
                        key=f"nre_clr_all_{phase}",
                        use_container_width=True,
                        on_click=_set_nre_selected_ids,
                        args=(ids_key, []),
                    )

                qty_by_id: dict[str, int] = {}
                # U-fit / Leading Edge: default 1 (Gold Rail independent); OT → 0
                ufit_key = f"nre_qty_ufit_{phase}"
                if ufit_key not in st.session_state:
                    st.session_state[ufit_key] = (
                        0 if empty_content_default else 1
                    )
                st.caption("Run of U-fit / Leading Edge")
                qty_by_id[NRE_UFIT_ID] = int(
                    st.number_input(
                        "Run of U-fit / Leading Edge",
                        min_value=0,
                        step=1,
                        key=ufit_key,
                        help=(
                            "Default 0 for OT; otherwise 1 "
                            "(with or without Gold Rail)."
                            if empty_content_default
                            else "Default 1 for every phase "
                            "(with or without Gold Rail)."
                        ),
                    )
                )

                if functionality == "Functional":
                    st.caption("OM / DnP quantity (Functional)")
                    om_col, dnp_col = st.columns(2)
                    with om_col:
                        qty_by_id[NRE_OM_ID] = int(
                            st.number_input(
                                "OM quantity",
                                min_value=0,
                                step=1,
                                value=0,
                                key=f"nre_qty_om_{phase}",
                            )
                        )
                    with dnp_col:
                        qty_by_id[NRE_DNP_ID] = int(
                            st.number_input(
                                "DnP quantity",
                                min_value=0,
                                step=1,
                                value=0,
                                key=f"nre_qty_dnp_{phase}",
                            )
                        )

                phase_configs.append(
                    {
                        "phase": phase,
                        "functionality": functionality,
                        "gold_rail": gold_rail,
                        "ufit_sor": ufit_sor,
                        "test_ids": list(selected_ids),
                        "qty_by_id": qty_by_id,
                    }
                )

            if st.button("Generate NRE Table", type="primary"):
                has_phase = any(phase_has_nre_content(cfg) for cfg in phase_configs)
                has_aux = aux_costs_have_content(aux_costs)
                if not nre_phases:
                    st.error("Select at least one phase.")
                elif not has_phase and not has_aux:
                    st.error(
                        "Select at least one test item, OM/DnP/U-fit quantity, "
                        "or fixture/rack cost."
                    )
                else:
                    missing = [
                        cfg["phase"]
                        for cfg in phase_configs
                        if not phase_has_nre_content(cfg)
                    ]
                    if missing and has_phase:
                        st.warning(
                            "No NRE test content for: " + ", ".join(missing)
                            + " — those columns stay in the NRE table and "
                            "headcount with zero test hours "
                            "(fixture / rack still use the earliest phase)."
                        )
                    st.session_state.nre_df = build_nre_table(
                        phase_configs,
                        account=account,
                        test_plan_df=test_plan_df,
                        location_df=location_df,
                        aux_costs=aux_costs,
                    )
                    # Lock headcount phases + NRE hours used for duration tasks
                    st.session_state.hc_phases = list(nre_phases)
                    st.session_state.hc_nre_hours_by_phase = {
                        str(cfg.get("phase", "")): _nre_phase_hours(
                            cfg, info_map=info_map
                        )
                        for cfg in phase_configs
                        if cfg.get("phase")
                    }
                    st.session_state.hc_nre_func_by_phase = {
                        str(cfg.get("phase", "")): str(
                            cfg.get("functionality", "Functional")
                        )
                        for cfg in phase_configs
                        if cfg.get("phase")
                    }
                    # Re-seed HC defaults so NRE duration / Baseline match this generate
                    _clear_hc_default_seeds()
                    # Clear previous headcount grid until user regenerates it
                    st.session_state.hc_df = None
                    st.session_state.hc_detail_df = None
                    st.session_state.hc_table_phases = None
                    st.session_state.hc_table_months = None
                    st.session_state.hc_table_tasks = None

            if st.session_state.nre_df is not None and not st.session_state.nre_df.empty:
                nre_df = st.session_state.nre_df
                selected_phases = sanitize_nre_phases(
                    list(st.session_state.hc_phases or nre_phases)
                )
                nre_view = pivot_nre_for_template(
                    nre_df, phases=selected_phases
                )
                st.dataframe(nre_view, use_container_width=True, hide_index=True)
                st.metric(
                    "Sub Total (grand total)",
                    f"{nre_grand_total(nre_view):,.2f}",
                    help="Sum of all row Sub Total values.",
                )

                meta = {
                    "account": account,
                    "project_name": project_name,
                    "project_phase": project_phase,
                    "weight_kg": weight_kg,
                    "standard": standard,
                    "phases": selected_phases,
                    "gold_rail": "per phase",
                    "gold_rail_phase": "per phase",
                    "ufit_sor": "per phase",
                    "phase_configs": phase_configs,
                }
                xlsx_bytes = export_nre_xlsx(nre_df, meta)
                if st.download_button(
                    "Download NRE (XLSX template)",
                    data=xlsx_bytes,
                    file_name=f"{_export_stem()}_NRE_{date.today().isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="nre_dl_xlsx",
                ):
                    _record_download("NRE")
                    st.toast("Usage recorded · NRE")

                # ---------------- Headcount (after NRE generate) ----------------
                st.divider()
                st.subheader("Headcount Estimation")
                hc_phases = sanitize_nre_phases(
                    list(st.session_state.hc_phases or selected_phases)
                )
                # Backfill NRE hours if missing (e.g. session from before this feature)
                if (
                    (
                        not st.session_state.get("hc_nre_hours_by_phase")
                        or not st.session_state.get("hc_nre_func_by_phase")
                    )
                    and phase_configs
                ):
                    st.session_state.hc_nre_hours_by_phase = {
                        str(cfg.get("phase", "")): _nre_phase_hours(
                            cfg, info_map=info_map
                        )
                        for cfg in phase_configs
                        if cfg.get("phase")
                    }
                    st.session_state.hc_nre_func_by_phase = {
                        str(cfg.get("phase", "")): str(
                            cfg.get("functionality", "Functional")
                        )
                        for cfg in phase_configs
                        if cfg.get("phase")
                    }
                    _clear_hc_default_seeds()
                st.caption(
                    "Phases follow the NRE testing-phase selection. "
                    "Each phase has its own months and engineer task lists. "
                    "**S&V Test** = (Σ NRE hours ÷ 160) × people "
                    "(Structure **2** on Concept / BU / BCT; S&V **3** on CT / "
                    "NT / OT / NOT). "
                    "**Baseline** (S&V, CT→NOT): only if that phase has NRE "
                    "tests — Functional **3×0.25**, Non-functional **3×0.1** "
                    "Per Phase (NT and OT each independently). "
                    "If system weight **> 91 kg (200 lbs)**, Number of "
                    "HeadCount on those S&V Test / Baseline tasks is **+1**. "
                    "**Per Month** / **Per Phase** distribute task totals. "
                    "Task total = **Number of HeadCount × HeadCount per Person**."
                )

                hc_months: dict[str, int] = {}
                tasks_by_phase: dict[str, dict[str, list[dict]]] = {}

                for phase in hc_phases:
                    st.divider()
                    st.markdown(f"### {phase}")
                    month_key = f"hc_months_{phase}"
                    if month_key not in st.session_state:
                        st.session_state[month_key] = _hc_month_default(phase)
                    hc_months[phase] = int(
                        st.number_input(
                            f"{phase} — months",
                            min_value=0,
                            step=1,
                            key=month_key,
                            help=(
                                "Default 3 months."
                                if phase in ("Concept", "BCT", "NOT")
                                else "Default 2 months."
                            ),
                        )
                    )

                    for eng in HC_ENGINEER_TYPES:
                        _seed_hc_engineer_defaults(phase, eng)
                        slug = _hc_eng_slug(eng)
                        n_key = _hc_task_count_key(phase, eng)
                        with st.expander(eng, expanded=False):
                            n_tasks = int(st.session_state.get(n_key, 0) or 0)
                            if n_tasks == 0:
                                st.caption(
                                    "No tasks yet — add a task to enter "
                                    "headcount and distribution."
                                )
                            phase_months = max(int(hc_months.get(phase, 0) or 0), 0)
                            month_opts = list(range(1, phase_months + 1)) if phase_months else [1]
                            for i in range(n_tasks):
                                hdr_l, hdr_r = st.columns([5, 1])
                                with hdr_l:
                                    st.markdown(f"**Task {i + 1}**")
                                with hdr_r:
                                    st.button(
                                        "Remove task",
                                        key=f"hc_rm_task_{phase}_{slug}_{i}",
                                        use_container_width=True,
                                        on_click=_remove_hc_task_at,
                                        args=(phase, eng, i),
                                    )
                                t1, t2, t3 = st.columns([3, 2, 2])
                                with t1:
                                    st.text_input(
                                        "Task",
                                        key=f"hc_task_name_{phase}_{slug}_{i}",
                                        placeholder="Task name",
                                    )
                                with t2:
                                    st.number_input(
                                        "Number of HeadCount",
                                        min_value=0,
                                        step=1,
                                        key=f"hc_task_n_{phase}_{slug}_{i}",
                                        help="Default 1. Multiplied by HeadCount per Person.",
                                    )
                                with t3:
                                    st.number_input(
                                        "HeadCount per Person",
                                        min_value=0.0,
                                        step=0.1,
                                        format="%.2f",
                                        key=f"hc_task_hc_{phase}_{slug}_{i}",
                                        help=(
                                            "Total for this task = Number of "
                                            "HeadCount × HeadCount per Person."
                                        ),
                                    )
                                dist_key = f"hc_task_dist_{phase}_{slug}_{i}"
                                start_key = f"hc_task_start_{phase}_{slug}_{i}"
                                end_key = f"hc_task_end_{phase}_{slug}_{i}"
                                # Clamp start/end to current phase month list
                                if start_key not in st.session_state:
                                    st.session_state[start_key] = month_opts[0]
                                if end_key not in st.session_state:
                                    st.session_state[end_key] = month_opts[-1]
                                if st.session_state[start_key] not in month_opts:
                                    st.session_state[start_key] = month_opts[0]
                                if st.session_state[end_key] not in month_opts:
                                    st.session_state[end_key] = month_opts[-1]
                                if (
                                    st.session_state[end_key]
                                    < st.session_state[start_key]
                                ):
                                    st.session_state[end_key] = st.session_state[
                                        start_key
                                    ]

                                dist_val = st.session_state.get(
                                    dist_key, HC_DIST_PER_MONTH
                                )
                                if dist_val == HC_DIST_PER_PHASE:
                                    d1, d2, d3 = st.columns([2, 2, 2])
                                    with d1:
                                        st.selectbox(
                                            "Distribute",
                                            HC_DISTRIBUTION_OPTS,
                                            key=dist_key,
                                            help=(
                                                "Per Month: add task total to "
                                                "each month. Per Phase: divide "
                                                "total by (end−start+1) and put "
                                                "only in that month span."
                                            ),
                                        )
                                    with d2:
                                        st.selectbox(
                                            "Start month",
                                            month_opts,
                                            key=start_key,
                                            help=(
                                                f"First month within {phase} "
                                                f"(1…{phase_months or 1})."
                                            ),
                                        )
                                    with d3:
                                        st.selectbox(
                                            "End month",
                                            month_opts,
                                            key=end_key,
                                            help=(
                                                "Last month within this phase "
                                                "(must be ≥ start). Same as start "
                                                "for a 1-month task."
                                            ),
                                        )
                                    sm = int(st.session_state.get(start_key, 1) or 1)
                                    em = int(st.session_state.get(end_key, sm) or sm)
                                    if em < sm:
                                        st.caption(
                                            f"End month ({em}) is before start "
                                            f"({sm}) — table will use {sm}…{sm}."
                                        )
                                    else:
                                        span = em - sm + 1
                                        st.caption(
                                            f"Per Phase: total ÷ **{span}** → "
                                            f"months **{sm}–{em}** of {phase} only."
                                        )
                                else:
                                    st.selectbox(
                                        "Distribute",
                                        HC_DISTRIBUTION_OPTS,
                                        key=dist_key,
                                        help=(
                                            "Per Month: add task total to "
                                            "each month. Per Phase: divide "
                                            "across selected start→end months."
                                        ),
                                    )
                            st.button(
                                "Add task",
                                key=f"hc_add_task_{phase}_{slug}",
                                on_click=_bump_hc_task_count,
                                args=(phase, eng),
                            )

                    tasks_by_phase[phase] = _collect_hc_tasks_for_phase(
                        phase, n_months=int(hc_months.get(phase, 0) or 0)
                    )

                st.session_state.hc_months = hc_months
                st.session_state.hc_tasks_by_phase = tasks_by_phase

                st.divider()
                st.markdown("##### Generate table")
                disp_c1, disp_c2 = st.columns(2)
                with disp_c1:
                    if "hc_month_display" not in st.session_state:
                        st.session_state.hc_month_display = HC_SHOW_SEQUENCE
                    hc_month_display = st.selectbox(
                        "Month display",
                        HC_MONTH_DISPLAY_OPTS,
                        key="hc_month_display",
                        help=(
                            "Show in sequence: Month row is 1, 2, 3, …. "
                            "Show in month: Month row is Jan–Dec names from "
                            "Project Start Month."
                        ),
                    )
                with disp_c2:
                    hc_project_start = "Jan"
                    if hc_month_display == HC_SHOW_MONTH:
                        if "hc_project_start_month" not in st.session_state:
                            st.session_state.hc_project_start_month = "Jan"
                        hc_project_start = st.selectbox(
                            "Project Start Month",
                            HC_CALENDAR_MONTHS,
                            key="hc_project_start_month",
                            help=(
                                "First calendar month on the headcount timeline. "
                                "E.g. Jun → Jun, Jul, Aug, …"
                            ),
                        )
                hc_generate = st.button(
                    "Generate Headcount Table",
                    type="primary",
                    key="hc_generate_btn",
                )

                if hc_generate:
                    if not hc_phases:
                        st.error("No testing phases available for headcount.")
                    elif sum(int(hc_months.get(p, 0) or 0) for p in hc_phases) <= 0:
                        st.error("Enter at least one month across the phases.")
                    else:
                        st.session_state.hc_df = build_headcount_table(
                            hc_phases,
                            hc_months,
                            tasks_by_phase=tasks_by_phase,
                            month_display=hc_month_display,
                            project_start_month=hc_project_start,
                        )
                        st.session_state.hc_detail_df = (
                            build_headcount_detail_table(
                                hc_phases,
                                tasks_by_phase=tasks_by_phase,
                            )
                        )
                        st.session_state.hc_table_phases = list(hc_phases)
                        st.session_state.hc_table_months = dict(hc_months)
                        st.session_state.hc_table_tasks = tasks_by_phase
                        st.session_state.hc_table_month_display = hc_month_display
                        st.session_state.hc_table_project_start = hc_project_start

                if (
                    st.session_state.hc_df is not None
                    and not st.session_state.hc_df.empty
                ):
                    disp = st.session_state.get(
                        "hc_table_month_display", HC_SHOW_SEQUENCE
                    )
                    start_lbl = st.session_state.get(
                        "hc_table_project_start", "Jan"
                    )
                    if disp == HC_SHOW_MONTH:
                        mode_note = (
                            f"Month row = calendar names from **{start_lbl}**."
                        )
                    else:
                        mode_note = "Month row = sequence **1, 2, 3, …**."
                    st.caption(
                        "Headcount grid (template layout): Phase / Month / "
                        "engineer rows + Comments. "
                        f"{mode_note} "
                        "**Per Phase** tasks only fill the selected start→end "
                        "months (total ÷ duration). "
                        "Regenerates only when you click **Generate Headcount Table**."
                    )
                    st.dataframe(
                        st.session_state.hc_df,
                        use_container_width=True,
                        hide_index=True,
                    )
                    detail_df = st.session_state.get("hc_detail_df")
                    if detail_df is None and st.session_state.get(
                        "hc_table_tasks"
                    ) is not None:
                        detail_df = build_headcount_detail_table(
                            list(
                                st.session_state.get("hc_table_phases")
                                or hc_phases
                            ),
                            tasks_by_phase=st.session_state.hc_table_tasks,
                        )
                        st.session_state.hc_detail_df = detail_df
                    if detail_df is not None and not detail_df.empty:
                        st.caption(
                            "Headcount computation detail — one row per phase; "
                            "tasks listed on separate lines in each engineer "
                            "cell: "
                            '**N. "Task": Ax  B HC / Manpower (per Month|Phase)** '
                            "(total HC = A × B)."
                        )
                        _render_hc_detail_table(detail_df)

                    hc_xlsx = export_headcount_xlsx(
                        st.session_state.hc_df,
                        detail_df
                        if detail_df is not None
                        else pd.DataFrame(),
                    )
                    if st.download_button(
                        "Download Headcount (Excel)",
                        data=hc_xlsx,
                        file_name=(
                            f"{_export_stem()}_Headcount_"
                            f"{date.today().isoformat()}.xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        ),
                        key="hc_dl_xlsx",
                        help=(
                            "One Excel file with sheet **Headcount** (grid) "
                            "and sheet **Detail** (task formulas). "
                            "Multi-sheet export requires Excel format."
                        ),
                    ):
                        _record_download("HC")
                        st.toast("Usage recorded · HC")
                else:
                    st.caption(
                        "Configure phase months and engineer tasks above, then "
                        "click **Generate Headcount Table**."
                    )

# ---------------------------------------------------------------------------
# Data preview / edit (end of page, collapsed by default)
# Edits apply only on Update — session working copy; does NOT overwrite folder files.
# Form prevents rerun-on-every-keystroke lag.
# ---------------------------------------------------------------------------
if account == "ROSA":
    st.divider()
    with st.expander("Data preview / edit (session only)", expanded=False):
        st.caption(
            f"Source folder: `data/{account}/` · catalog from "
            f"`test_plan_info_detail.csv` · working copy has "
            f"**{len(st.session_state.work_test_plan)}** test items. "
            "Edit below, then click **Update loaded data**. "
            "Folder CSVs are not overwritten."
        )
        reload_col, _ = st.columns([1, 3])
        with reload_col:
            if st.button("Reload from folder", use_container_width=True):
                try:
                    _load_working_catalog_from_disk(account)
                    st.success("Reloaded from folder into session (disk files unchanged).")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Reload failed: {exc}")

        with st.form("catalog_edit_form", clear_on_submit=False):
            rev = int(st.session_state.get("catalog_edit_rev", 0))
            prev_tab1, prev_tab2, prev_tab3 = st.tabs(
                ["Test Plan Info (detail)", "Location Info", "Convert Table"]
            )
            with prev_tab1:
                edited_plan = st.data_editor(
                    st.session_state.work_test_plan,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key=f"editor_test_plan_{rev}",
                )
            with prev_tab2:
                edited_loc = st.data_editor(
                    st.session_state.work_location,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key=f"editor_location_{rev}",
                )
            with prev_tab3:
                edited_conv = st.data_editor(
                    st.session_state.work_convert,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key=f"editor_convert_{rev}",
                )
            submitted = st.form_submit_button(
                "Update loaded data", type="primary", use_container_width=True
            )

        if submitted:
            try:
                st.session_state.work_test_plan = _normalize_test_plan(edited_plan)
                st.session_state.work_location = _normalize_location(edited_loc)
                st.session_state.work_convert = _normalize_convert(edited_conv)
                st.session_state.catalog_account = account
                st.session_state.catalog_edit_rev = rev + 1
                st.success(
                    "Session data updated. Folder files were not changed. "
                    "Generate / NRE will use this working copy."
                )
                st.rerun()
            except Exception as exc:
                st.error(f"Could not apply table edits: {exc}")
