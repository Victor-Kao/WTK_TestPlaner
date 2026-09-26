"""
WTK Test Planner — Streamlit schedule management tool.
Accounts: ROSA, NAOMI (per-account data under data/<ACCOUNT>/).

Main-page layout only (no sidebar) so a parent tool can own the sidebar.
"""
from __future__ import annotations

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
    NRE_AUX_COST_DEFAULTS_USD,
    NRE_AUX_COST_IDS,
    NRE_DNP_ID,
    NRE_OM_ID,
    NRE_SELECTOR_EXTRA_IDS,
    NRE_UFIT_ID,
    ORV3_MGX_WO_L11_COL,
    ORV3_MGX_WO_L11_LABEL,
    PHASES,
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
from modules.nre import (
    aux_costs_have_content,
    build_nre_table,
    export_nre_xlsx,
    first_nre_phase,
    phase_has_nre_content,
    pivot_nre_for_template,
)
from modules.timeline import (
    apply_sequence_template,
    clear_cell_span,
    init_timeline_state,
    place_item,
    refresh_durations_from_catalog,
    rename_system_row,
    set_calendar_empty_mark,
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
    st.session_state.plan_log_context = None
    st.session_state.editor_version = int(st.session_state.editor_version) + 1
    st.session_state.editor_error = None


def _record_download(tool_type: str, **overrides) -> None:
    """Write a usage-log row for TEST PLAN / NRE / HC downloads."""
    ctx = dict(st.session_state.get("plan_log_context") or {})
    ctx.update({k: v for k, v in overrides.items() if v is not None})
    append_usage_record(
        tool_type=tool_type,
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

    tab_plan, tab_nre, tab_hc = st.tabs(
        ["Test Plan Timeline", "NRE Estimation", "Headcount Estimation"]
    )


    if account != "ROSA":
        with tab_plan:
            st.subheader("Test Plan Timeline")
            st.info(
                "TBD — NAOMI uses a different test-plan workflow "
                "(not implemented yet)."
            )
        with tab_nre:
            st.subheader("NRE Estimation")
            st.info("TBD — NAOMI NRE estimation is not implemented yet.")
        with tab_hc:
            st.subheader("Headcount Estimation")
            st.info(
                "TBD — NAOMI headcount estimation is not implemented yet."
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
                    "Top table = Weekday / Event (choose Empty in Event). "
                    "Empty days stay white and cannot be filled. "
                    "Multi-day items keep their start day and skip Empty/weekend "
                    "(e.g. 09/30–10/02 with Empty on 10/01 → 09/30, 10/02, 10/05); "
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

                ref_col, _ = st.columns([1, 3])
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
                    elif kind == "mark_empty":
                        d = str(event.get("date", ""))
                        after = "" if event.get("value") is None else str(event.get("value"))
                        enabled = after == EMPTY_LABEL or after.lower() == "empty"
                        new_tl, shift_errs = set_calendar_empty_mark(timeline, d, enabled)
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
                            err_msg = f"{d} is Weekend/Holiday/Empty and cannot be filled."
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

        # ============================= NRE ========================================
        with tab_nre:
            st.subheader("NRE Estimation")
            st.caption(
                "Select one or more testing phases. Fixture / rack USD costs are "
                "one-time (under Testing Phase) and land on the earliest selected "
                "phase in the Excel template. Each phase has its own options + "
                "test items. "
                "Generated table matches `NRE_TEMPLATE_<account>.xlsx` "
                "(Test Item (ID) · Location · Rate · Concept/BCT/NOT · Sub Total). "
                "**Sub Total** in Excel = `SUM(Concept:NOT) × Lab Rate`."
            )

            nre_phases = st.multiselect(
                "Testing Phase (multi-select)",
                PHASES,
                default=[PHASES[0]],
                key="nre_phases",
            )

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
                        key=f"nre_func_{phase}",
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
                        key=f"nre_gold_{phase}",
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
                if ids_key not in st.session_state:
                    # Each newly added phase defaults to Select all for its own filters
                    st.session_state[ids_key] = list(allowed_ids)
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
                        "fixture / rack costs are under Testing Phase."
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
                # U-fit / Leading Edge: default 4, or 3 when Gold Rail = Yes
                ufit_default = 3 if gold_rail == "Yes" else 4
                ufit_key = f"nre_qty_ufit_{phase}"
                ufit_gold_track = f"_nre_ufit_gold_{phase}"
                if (
                    ufit_gold_track not in st.session_state
                    or st.session_state[ufit_gold_track] != gold_rail
                ):
                    st.session_state[ufit_key] = ufit_default
                    st.session_state[ufit_gold_track] = gold_rail
                st.caption("U-fit / Leading Edge quantity")
                qty_by_id[NRE_UFIT_ID] = int(
                    st.number_input(
                        "U-fit / Leading Edge quantity",
                        min_value=0,
                        step=1,
                        key=ufit_key,
                        help="Default 4; default 3 when Gold Rail is Yes.",
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
                            + " — those phases will be skipped "
                            "(fixture / rack still use the earliest phase)."
                        )
                    st.session_state.nre_df = build_nre_table(
                        phase_configs,
                        account=account,
                        test_plan_df=test_plan_df,
                        location_df=location_df,
                        aux_costs=aux_costs,
                    )

            if st.session_state.nre_df is not None and not st.session_state.nre_df.empty:
                nre_df = st.session_state.nre_df
                nre_view = pivot_nre_for_template(nre_df)
                st.dataframe(nre_view, use_container_width=True, hide_index=True)
                grand = float(nre_view["Sub Total"].fillna(0).sum())
                st.metric("Sub Total (grand total)", f"{grand:,.0f}")

                used_phases = list(dict.fromkeys(nre_df["Phase"].astype(str).tolist()))
                meta = {
                    "account": account,
                    "project_name": project_name,
                    "project_phase": project_phase,
                    "weight_kg": weight_kg,
                    "standard": standard,
                    "phases": used_phases,
                    "gold_rail": "per phase",
                    "gold_rail_phase": "per phase",
                    "ufit_sor": "per phase",
                    "phase_configs": phase_configs,
                }
                xlsx_bytes = export_nre_xlsx(nre_df, meta)
                nre_functional = ""
                if phase_configs:
                    nre_functional = ", ".join(
                        f"{c['phase']}:{c['functionality']}" for c in phase_configs
                    )
                if st.download_button(
                    "Download NRE (XLSX template)",
                    data=xlsx_bytes,
                    file_name=f"{_export_stem()}_NRE_{date.today().isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="nre_dl_xlsx",
                ):
                    _record_download("NRE", functional=nre_functional)
                    st.toast("Usage recorded · NRE")

        # ============================= HEADCOUNT ==================================
        with tab_hc:
            st.subheader("Headcount Estimation")
            st.info("TBD — placeholder for future headcount model.")
            st.write(
                "Planned inputs (not implemented yet): phase mix, system count, "
                "lab shifts, and overlapping SoR / Gold Rail workload."
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
