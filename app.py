"""
WTK Test Planner — Streamlit schedule management tool.
Accounts: ROSA, NAOMI (per-account data under data/<ACCOUNT>/).
"""
from __future__ import annotations

from datetime import date, timedelta

import streamlit as st

from modules.config import (
    ACCOUNTS,
    CELL_AHEAD_OPT,
    CELL_POSTPONE_OPT,
    EMPTY_LABEL,
    EMPTY_TOKEN,
    FUNCTIONALITY_OPTS,
    PHASES,
    SYSTEM_NUMBERS,
    YES_NO,
    account_data_dir,
)
from modules.data_loader import (
    load_convert_table,
    load_location_info,
    load_test_plan_info,
    lookup_convert_id,
    load_case_sequence,
    test_info_by_id,
)
from modules.nre import (
    build_nre_table,
    collect_test_ids_from_timeline,
    default_test_ids_for_filters,
    export_nre_xlsx,
)
from modules.timeline import (
    apply_sequence_template,
    clear_cell_span,
    init_timeline_state,
    place_item,
    set_calendar_empty_mark,
    shift_system_from_date,
    shift_system_schedule,
    style_timeline_display,
    timeline_to_export_df,
)
from modules.synced_calendar import render_synced_calendar

st.set_page_config(
    page_title="WTK Test Planner",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
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
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_state()

# Kill Streamlit's stale-element fade/shine on rerun
st.markdown(
    """
    <style>
    /* Remove the white/opacity flash when widgets rerun */
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


def _select_options(info: dict) -> list[str]:
    """Dropdown choices for calendar cells (Excel-like)."""
    opts = ["", CELL_AHEAD_OPT, CELL_POSTPONE_OPT]
    for tid, meta in info.items():
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


# ---------------------------------------------------------------------------
# Sidebar — account / weight
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("WTK Test Planner")
    st.caption("Schedule · NRE · Headcount")
    account = st.selectbox("Account (Brand)", ACCOUNTS, index=0)
    weight_kg = st.number_input("System weight (kg)", min_value=0.0, value=25.0, step=0.5)

    # Clear timeline/NRE when switching brands so data stays account-scoped
    if st.session_state.active_account != account:
        st.session_state.active_account = account
        st.session_state.timeline = None
        st.session_state.export_df = None
        st.session_state.nre_df = None
        st.session_state.editor_version = int(st.session_state.editor_version) + 1
        st.session_state.editor_error = None

    st.divider()
    data_path = account_data_dir(account)
    try:
        test_plan_df, location_df, convert_df = _cached_catalog(account)
        st.success(f"Loaded {len(test_plan_df)} test items from `{data_path.name}/`")
        with st.expander("Test Plan Info Table", expanded=False):
            st.dataframe(test_plan_df, width="stretch", hide_index=True)
        with st.expander("Location Info", expanded=False):
            st.dataframe(location_df, width="stretch", hide_index=True)
        with st.expander("Convert Table (24 cases)", expanded=False):
            st.dataframe(convert_df, width="stretch", hide_index=True)
        st.caption(f"Data folder: `data/{account}/`")
    except Exception as exc:
        st.error(f"Failed to load data for {account}: {exc}")
        st.stop()

info_map = test_info_by_id(test_plan_df)

st.title("Schedule Manager")
st.write(f"**Account:** {account} &nbsp;|&nbsp; **Weight:** {weight_kg:g} kg")

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_plan, tab_nre, tab_hc = st.tabs(
    ["Test Plan Timeline", "NRE Estimation", "Headcount Estimation"]
)

# ============================= TEST PLAN ==================================
with tab_plan:
    st.subheader(f"1. {account} Test Plan Options")

    c1, c2, c3 = st.columns(3)
    with c1:
        phase = st.selectbox("Current Testing Phase (pick one)", PHASES, key="tp_phase")
        functionality = st.radio(
            "Functional / Non-functional",
            FUNCTIONALITY_OPTS,
            horizontal=True,
            key="tp_func",
        )
        ufit_sor = st.radio(
            "U-fit for SoR / ORV3",
            YES_NO,
            horizontal=True,
            key="tp_ufit",
        )
    with c2:
        system_eta = st.date_input("System ETA", value=date.today() + timedelta(days=7))
        critical_fb = st.date_input(
            "Critical feedback", value=date.today() + timedelta(days=21)
        )
        n_systems = st.selectbox("Number of System", SYSTEM_NUMBERS, index=0)
    with c3:
        gold_rail = st.radio(
            "Gold Rail Selection for this plan?",
            YES_NO,
            horizontal=True,
            key="tp_gold",
        )
        st.caption(
            "Timeline starts at **System ETA** and ends at "
            "**5 business days after Critical feedback** "
            "(or later if the filled plan runs past that)."
        )

    convert_id = lookup_convert_id(
        account, functionality, gold_rail, ufit_sor, int(n_systems)
    )
    st.session_state.last_convert_id = convert_id
    if convert_id:
        st.info(
            f"Case Convert_ID = **{convert_id}** "
            f"(Functionality={functionality}, Gold Rail={gold_rail}, "
            f"U-fit SoR={ufit_sor}, Systems={n_systems})"
        )

    preload = st.checkbox(
        "Pre-fill timeline from sequence sheet (e.g. 1, T003, T004, 5, T010)",
        value=False,
    )

    gen_col, _ = st.columns([1, 3])
    with gen_col:
        generate = st.button("Generate Test Plan", type="primary", width="stretch")

    if generate:
        if critical_fb < system_eta:
            st.warning(
                "Critical feedback is before System ETA — timeline still starts at ETA."
            )
        tl = init_timeline_state(
            system_eta,
            critical_fb,
            int(n_systems),
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
        st.session_state.timeline = tl
        st.session_state.export_df = None
        st.session_state.editor_version = int(st.session_state.editor_version) + 1
        st.session_state.editor_error = None
        if msgs:
            for m in msgs:
                st.warning(m)
        st.success(
            f"Timeline generated: {tl['dates'][0]} → {tl['dates'][-1]} "
            f"(Weekend/Holiday columns locked)."
        )
    timeline = st.session_state.timeline
    if timeline is None:
        st.caption("Configure options above, then click **Generate Test Plan**.")
    else:
        st.subheader("2. Timeline")
        st.caption(
            "Top table = Weekday / Marked (choose Empty in Marked). "
            "Empty days stay white and cannot be filled. "
            "Multi-day items keep their start day and skip Empty/weekend "
            "(e.g. 09/30–10/02 with Empty on 10/01 → 09/30, 10/02, 10/05); "
            "blank days before the start stay blank. "
            "Weekend / Holiday columns are light red. "
            "Row buttons Ahead/Postpone shift the **whole** system line. "
            "In a cell dropdown, **Ahead (−1) from here** / **Postpone (+1) from here** "
            "shift only items from that date onward (blocked with a warning if Ahead would overlap)."
        )

        options = _select_options(info_map)
        if st.session_state.editor_error:
            st.error(st.session_state.editor_error)

        event = render_synced_calendar(
            timeline,
            options,
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

            if kind == "shift_system":
                sk = str(event.get("system", ""))
                try:
                    delta = int(event.get("delta", 0))
                except (TypeError, ValueError):
                    delta = 0
                new_tl, shift_errs = shift_system_schedule(timeline, sk, delta)
                # non-fatal notices (e.g. already at earliest day)
                if shift_errs and new_tl is timeline:
                    err_msg = "; ".join(shift_errs)
                elif shift_errs:
                    st.session_state.editor_error = None
                    # soft warnings only
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
                # Guard: treat shift options as no-op if they slip through as system edits
                if after in (CELL_AHEAD_OPT, CELL_POSTPONE_OPT):
                    delta = -1 if after == CELL_AHEAD_OPT else 1
                    new_tl, shift_errs = shift_system_from_date(timeline, sk, d, delta)
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
                st.session_state.editor_version = int(st.session_state.editor_version) + 1
            st.rerun()

        st.divider()
        st.subheader("3. Update → Export Table")
        if st.button("Update", type="primary"):
            st.session_state.export_df = timeline_to_export_df(st.session_state.timeline)
            st.success("Timeline converted to export table.")

        if st.session_state.export_df is not None:
            export_styled = style_timeline_display(
                st.session_state.export_df, st.session_state.timeline
            )
            st.dataframe(export_styled, width="stretch", hide_index=True)
            csv_bytes = st.session_state.export_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Download timeline CSV",
                data=csv_bytes,
                file_name=f"{account}_timeline_{phase}_{date.today().isoformat()}.csv",
                mime="text/csv",
            )

# ============================= NRE ========================================
with tab_nre:
    st.subheader("NRE Estimation")
    st.caption(
        "Select one or more phases. Lab fee / location come from external CSVs. "
        "Final fee applies phase · functionality · gold-rail · U-fit multipliers."
    )

    nre_phases = st.multiselect("Testing Phase (multi-select)", PHASES, default=[phase])
    phase_func: dict[str, str] = {}
    if nre_phases:
        cols = st.columns(len(nre_phases))
        for i, p in enumerate(nre_phases):
            with cols[i]:
                phase_func[p] = st.radio(
                    f"{p} type",
                    FUNCTIONALITY_OPTS,
                    horizontal=True,
                    key=f"nre_func_{p}",
                )

    nc1, nc2, nc3, nc4 = st.columns(4)
    with nc1:
        gold_rail_phase = st.selectbox(
            "Phase for Gold Rail Selection",
            PHASES,
            key="nre_gold_phase",
        )
    with nc2:
        nre_gold = st.radio("Gold Rail Selection", YES_NO, horizontal=True, key="nre_gold")
    with nc3:
        nre_ufit = st.radio("U-fit for SoR", YES_NO, horizontal=True, key="nre_ufit")
    with nc4:
        nre_qty = st.number_input("Qty per test item", min_value=1, value=1)

    # Source of test IDs
    from_timeline = collect_test_ids_from_timeline(st.session_state.timeline)
    default_ids = from_timeline or default_test_ids_for_filters(
        phase_func.get(nre_phases[0], "Functional") if nre_phases else "Functional",
        nre_gold,
        nre_ufit,
        account=account,
    )
    all_ids = list(test_plan_df["Test_ID"].astype(str))
    selected_ids = st.multiselect(
        "Test items included in NRE",
        options=all_ids,
        default=[i for i in default_ids if i in all_ids],
        format_func=lambda tid: f"{info_map[tid]['Abbrv_Name']} ({tid})",
    )

    if st.button("Generate NRE Table", type="primary"):
        if not nre_phases:
            st.error("Select at least one phase.")
        elif not selected_ids:
            st.error("Select at least one test item.")
        else:
            st.session_state.nre_df = build_nre_table(
                selected_ids,
                nre_phases,
                phase_func,
                nre_gold,
                nre_ufit,
                int(nre_qty),
                account=account,
            )

    if st.session_state.nre_df is not None and not st.session_state.nre_df.empty:
        nre_df = st.session_state.nre_df
        st.dataframe(nre_df, width="stretch", hide_index=True)
        grand = float(nre_df["Final_Fee"].sum())
        st.metric("Final Fee (grand total)", f"{grand:,.0f}")

        meta = {
            "account": account,
            "weight_kg": weight_kg,
            "phases": nre_phases,
            "gold_rail": nre_gold,
            "gold_rail_phase": gold_rail_phase,
            "ufit_sor": nre_ufit,
        }
        xlsx_bytes = export_nre_xlsx(nre_df, meta)
        st.download_button(
            "Download NRE (XLSX template)",
            data=xlsx_bytes,
            file_name=f"{account}_NRE_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Download NRE CSV",
            data=nre_df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{account}_NRE_{date.today().isoformat()}.csv",
            mime="text/csv",
        )

# ============================= HEADCOUNT ==================================
with tab_hc:
    st.subheader("Headcount Estimation")
    st.info("TBD — placeholder for future headcount model.")
    st.write(
        "Planned inputs (not implemented yet): phase mix, system count, "
        "lab shifts, and overlapping SoR / Gold Rail workload."
    )
