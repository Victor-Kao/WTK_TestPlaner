from __future__ import annotations

from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from modules.config import (
    NRE_AUX_COST_IDS,
    NRE_DNP_ID,
    NRE_MULTIPLIERS,
    NRE_OM_ID,
    NRE_TEMPLATE_XLSX,
    NRE_UFIT_ID,
    ORV3_MGX_WO_L11_COL,
    PHASES,
    nre_template_path,
)
from modules.data_loader import load_location_info, load_test_plan_info, test_info_by_id

# ROSA NRE template columns (Sheet1)
_NRE_PHASE_COLS = {"Concept": "Concept", "BCT": "BCT", "NOT": "NOT"}
_NRE_TEMPLATE_HEADERS = [
    "Test Item",
    "Lab Location",
    "Lab Rate",
    "Concept",
    "BCT",
    "NOT",
    "Sub Total",
    "Comments",
]
_NRE_SUBTOTAL_FORMULA = "=SUM(D{row}:F{row})*C{row}"


def compute_multiplier(
    phase: str,
    functionality: str,
    gold_rail: str,
    ufit_sor: str,
) -> float:
    m = 1.0
    m *= NRE_MULTIPLIERS["phase"].get(phase, 1.0)
    m *= NRE_MULTIPLIERS["functionality"].get(functionality, 1.0)
    m *= NRE_MULTIPLIERS["gold_rail"].get(gold_rail, 1.0)
    m *= NRE_MULTIPLIERS["ufit_sor"].get(ufit_sor, 1.0)
    return m


def lab_fee_for_item(meta: dict[str, Any]) -> float:
    """NRE lab fee = Duration_for_NRE (hours) × Lab_Rate (per hour)."""
    hours = float(meta.get("Duration_for_NRE", 0) or 0)
    rate = float(meta.get("Lab_Rate", 0) or 0)
    return round(hours * rate, 2)


def phase_has_nre_content(cfg: dict[str, Any]) -> bool:
    """True if the phase config will produce at least one non-aux NRE row."""
    if cfg.get("test_ids"):
        return True
    qty_by_id = cfg.get("qty_by_id") or {}
    qty_ids = [NRE_UFIT_ID]
    if str(cfg.get("functionality", "")) == "Functional":
        qty_ids.extend([NRE_OM_ID, NRE_DNP_ID])
    return any(int(qty_by_id.get(tid, 0) or 0) > 0 for tid in qty_ids)


def aux_costs_have_content(aux_costs: dict[str, Any] | None) -> bool:
    """True if any one-time fixture / rack USD cost is > 0."""
    costs = aux_costs or {}
    return any(float(costs.get(tid, 0) or 0) > 0 for tid in NRE_AUX_COST_IDS)


def first_nre_phase(phases: list[str]) -> str | None:
    """Earliest phase among selection in Concept → BCT → NOT order."""
    selected = {str(p) for p in phases}
    for p in PHASES:
        if p in selected:
            return p
    return str(phases[0]) if phases else None



def _nre_row(
    *,
    meta: dict[str, Any],
    loc: dict[str, str],
    phase: str,
    functionality: str,
    gold_rail: str,
    ufit_sor: str,
    hours: float,
    rate: float,
    lab_fee: float,
    qty: int,
    total_fee: float,
) -> dict[str, Any]:
    return {
        "Test_ID": meta["Test_ID"],
        "Test_Item": meta["Testplan_Item"],
        "Location": loc.get(meta["Test_ID"], ""),
        "Duration_for_NRE": hours,
        "Lab_Rate": rate,
        "Lab_Fee": lab_fee,
        "Qty": qty,
        "Total_Fee": total_fee,
        "Phase": phase,
        "Functionality": functionality,
        "Gold_Rail_Selection": gold_rail,
        ORV3_MGX_WO_L11_COL: ufit_sor,
        "Final_Fee": total_fee,
    }


def build_nre_table(
    phase_configs: list[dict[str, Any]],
    *,
    account: str,
    test_plan_df: pd.DataFrame | None = None,
    location_df: pd.DataFrame | None = None,
    aux_costs: dict[str, Any] | None = None,
) -> pd.DataFrame:
    """
    Build NRE rows from per-phase configs plus one-time fixture / rack USD costs.

    Each config:
      phase, functionality, gold_rail, ufit_sor, test_ids, qty_by_id

    aux_costs (Wooden Fixture / Dummy / Rack) are charged once and attributed to the
    earliest selected phase (Concept → BCT → NOT).

    Lab_Fee = Duration_for_NRE × Lab_Rate (hours × hourly rate), unless aux USD override.
    Final_Fee = Lab_Fee × Qty (aux rows use typed USD as Lab_Fee / Final_Fee).
    """
    plan = test_plan_df if test_plan_df is not None else load_test_plan_info(account)
    loc_df = location_df if location_df is not None else load_location_info(account)
    info = test_info_by_id(plan)
    loc = loc_df.set_index("Test_ID")["Location"].to_dict()
    rows: list[dict] = []
    cfg_by_phase = {
        str(cfg.get("phase", "")): cfg for cfg in phase_configs if cfg.get("phase")
    }

    for cfg in phase_configs:
        phase = str(cfg.get("phase", ""))
        functionality = str(cfg.get("functionality", "Functional"))
        gold_rail = str(cfg.get("gold_rail", "No"))
        ufit_sor = str(cfg.get("ufit_sor", "No"))
        test_ids = list(cfg.get("test_ids") or [])
        qty_by_id = dict(cfg.get("qty_by_id") or {})

        for tid in test_ids:
            meta = info.get(str(tid))
            if not meta:
                continue
            hours = float(meta.get("Duration_for_NRE") or 0)
            rate = float(meta.get("Lab_Rate") or 0)
            lab_fee = lab_fee_for_item(meta)
            rows.append(
                _nre_row(
                    meta=meta,
                    loc=loc,
                    phase=phase,
                    functionality=functionality,
                    gold_rail=gold_rail,
                    ufit_sor=ufit_sor,
                    hours=hours,
                    rate=rate,
                    lab_fee=lab_fee,
                    qty=1,
                    total_fee=lab_fee,
                )
            )

        # Quantity extras: U-fit always; OM / DnP when Functional (skip if qty ≤ 0)
        qty_tids = [NRE_UFIT_ID]
        if functionality == "Functional":
            qty_tids.extend([NRE_OM_ID, NRE_DNP_ID])
        for tid in qty_tids:
            qty = int(qty_by_id.get(tid, 0) or 0)
            if qty <= 0:
                continue
            meta = info.get(tid)
            if not meta:
                continue
            hours = float(meta.get("Duration_for_NRE") or 0)
            rate = float(meta.get("Lab_Rate") or 0)
            unit = lab_fee_for_item(meta)
            rows.append(
                _nre_row(
                    meta=meta,
                    loc=loc,
                    phase=phase,
                    functionality=functionality,
                    gold_rail=gold_rail,
                    ufit_sor=ufit_sor,
                    hours=hours,
                    rate=rate,
                    lab_fee=unit,
                    qty=qty,
                    total_fee=round(unit * qty, 2),
                )
            )

    # One-time fixture / rack USD → earliest selected phase only
    costs = dict(aux_costs or {})
    aux_phase = first_nre_phase(list(cfg_by_phase.keys()))
    aux_cfg = cfg_by_phase.get(aux_phase or "", {}) if aux_phase else {}
    if aux_phase and aux_costs_have_content(costs):
        functionality = str(aux_cfg.get("functionality", "Functional"))
        gold_rail = str(aux_cfg.get("gold_rail", "No"))
        ufit_sor = str(aux_cfg.get("ufit_sor", "No"))
        for tid in NRE_AUX_COST_IDS:
            cost = float(costs.get(tid, 0) or 0)
            if cost <= 0:
                continue
            meta = info.get(tid)
            if not meta:
                continue
            rows.append(
                _nre_row(
                    meta=meta,
                    loc=loc,
                    phase=aux_phase,
                    functionality=functionality,
                    gold_rail=gold_rail,
                    ufit_sor=ufit_sor,
                    hours=float(meta.get("Duration_for_NRE") or 0),
                    rate=0.0,
                    lab_fee=round(cost, 2),
                    qty=1,
                    total_fee=round(cost, 2),
                )
            )
    return pd.DataFrame(rows)


def collect_test_ids_from_timeline(timeline: dict | None) -> list[str]:
    if not timeline:
        return []
    seen: list[str] = []
    for row in timeline.get("grid", {}).values():
        for cell in row.values():
            if not cell or not cell.get("is_start") or cell.get("is_empty"):
                continue
            tid = cell.get("test_id")
            if tid and tid not in seen:
                seen.append(tid)
    return seen


def _is_aux_cost_row(tid: str) -> bool:
    return str(tid).strip() in NRE_AUX_COST_IDS


def _nre_test_item_label(item: str, tid: str) -> str:
    """
    Template Test Item column:
      SV0… → "Testplan_Item (Test_ID)"
      otherwise (AUX / ENG / REL / …) → Testplan_Item only
    """
    tid_s = str(tid or "").strip()
    item_s = str(item or "").strip()
    if tid_s.upper().startswith("SV0"):
        return f"{item_s} ({tid_s})" if item_s else tid_s
    return item_s or tid_s


def pivot_nre_for_template(nre_df: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse detailed NRE rows into the account template shape:
    Test Item | Lab Location | Lab Rate | Concept | BCT | NOT | Sub Total | Comments

    Phase columns hold hours (Duration_for_NRE × Qty). Aux USD rows put the cost in
    Lab Rate and 1 in the phase column so Sub Total (=SUM(phases)*rate) equals the cost.
    SV0* Test_IDs show as "name (ID)"; other IDs show name only.
    """
    if nre_df is None or nre_df.empty:
        return pd.DataFrame(columns=_NRE_TEMPLATE_HEADERS)

    # Preserve first-seen order of Test_ID
    order: list[str] = []
    buckets: dict[str, dict[str, Any]] = {}

    for rec in nre_df.to_dict(orient="records"):
        tid = str(rec.get("Test_ID", "")).strip()
        if not tid:
            continue
        phase = str(rec.get("Phase", "")).strip()
        if tid not in buckets:
            order.append(tid)
            item = str(rec.get("Test_Item", "") or "").strip()
            buckets[tid] = {
                "Test Item": _nre_test_item_label(item, tid),
                "Lab Location": str(rec.get("Location", "") or ""),
                "Lab Rate": 0.0,
                "Concept": 0.0,
                "BCT": 0.0,
                "NOT": 0.0,
                "Comments": "",
                "_aux": _is_aux_cost_row(tid),
            }
        b = buckets[tid]
        if _is_aux_cost_row(tid):
            cost = float(rec.get("Final_Fee") or rec.get("Lab_Fee") or 0)
            b["Lab Rate"] = cost
            if phase in _NRE_PHASE_COLS:
                b[phase] = float(b.get(phase, 0) or 0) + 1.0
        else:
            rate = float(rec.get("Lab_Rate") or 0)
            if rate:
                b["Lab Rate"] = rate
            hours = float(rec.get("Duration_for_NRE") or 0) * float(rec.get("Qty") or 1)
            if phase in _NRE_PHASE_COLS:
                b[phase] = float(b.get(phase, 0) or 0) + hours
            if not b["Lab Location"] and rec.get("Location"):
                b["Lab Location"] = str(rec.get("Location") or "")

    rows: list[dict[str, Any]] = []
    for tid in order:
        b = buckets[tid]
        concept = float(b["Concept"] or 0)
        bct = float(b["BCT"] or 0)
        not_ = float(b["NOT"] or 0)
        rate = float(b["Lab Rate"] or 0)
        sub = round((concept + bct + not_) * rate, 2)
        rows.append(
            {
                "Test Item": b["Test Item"],
                "Lab Location": b["Lab Location"],
                "Lab Rate": rate,
                "Concept": concept if concept else None,
                "BCT": bct if bct else None,
                "NOT": not_ if not_ else None,
                "Sub Total": sub,
                "Comments": b["Comments"] or None,
            }
        )
    return pd.DataFrame(rows, columns=_NRE_TEMPLATE_HEADERS)


def _excel_num(value: Any) -> float | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    num = float(value)
    return num if num else None


def _copy_nre_template_row(ws: Worksheet, src_row: int, dest_row: int) -> None:
    """Copy cell styles/number formats from src_row and install Sub Total formula."""
    for col in range(1, 9):
        src = ws.cell(src_row, col)
        dst = ws.cell(dest_row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)
            dst.alignment = copy(src.alignment)
    ws.cell(dest_row, 7).value = _NRE_SUBTOTAL_FORMULA.format(row=dest_row)


def export_nre_xlsx(
    nre_df: pd.DataFrame,
    meta: dict,
    template_path: Path | None = None,
) -> bytes:
    """
    Fill the account NRE template (colors + Sub Total formulas) and return bytes.

    Layout: Test Item | Lab Location | Lab Rate | Concept | BCT | NOT | Sub Total | Comments
    (SV0* IDs appended as "name (ID)"; AUX / ENG / REL show name only.)
    """
    account = str(meta.get("account") or "")
    template = template_path or (
        nre_template_path(account) if account else NRE_TEMPLATE_XLSX
    )
    pivoted = pivot_nre_for_template(nre_df)
    buf = BytesIO()

    if template.exists():
        wb = load_workbook(template)
        ws = wb.active
        data_start = 2
        style_src = data_start
        # Pre-styled blank rows in the ROSA template
        prefilled_last = max(ws.max_row, data_start)

        records = pivoted.to_dict(orient="records")
        for i, rec in enumerate(records):
            r = data_start + i
            if r > prefilled_last:
                _copy_nre_template_row(ws, style_src, r)
            else:
                # Ensure Sub Total formula is present (template already has it)
                if not ws.cell(r, 7).value:
                    ws.cell(r, 7).value = _NRE_SUBTOTAL_FORMULA.format(row=r)
            ws.cell(r, 1).value = rec.get("Test Item")
            ws.cell(r, 2).value = rec.get("Lab Location") or None
            rate = rec.get("Lab Rate")
            ws.cell(r, 3).value = (
                None if rate is None or (isinstance(rate, float) and pd.isna(rate)) else float(rate)
            )
            ws.cell(r, 4).value = _excel_num(rec.get("Concept"))
            ws.cell(r, 5).value = _excel_num(rec.get("BCT"))
            ws.cell(r, 6).value = _excel_num(rec.get("NOT"))
            # column G = formula (do not overwrite with computed value)
            comments = rec.get("Comments")
            ws.cell(r, 8).value = (
                None
                if comments is None or (isinstance(comments, float) and pd.isna(comments))
                else comments
            )

        # Clear unused pre-styled rows (keep fills + Sub Total formula)
        for r in range(data_start + len(records), prefilled_last + 1):
            for c in (1, 2, 3, 4, 5, 6, 8):
                ws.cell(r, c).value = None
            ws.cell(r, 7).value = _NRE_SUBTOTAL_FORMULA.format(row=r)

        wb.save(buf)
    else:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            pivoted.to_excel(writer, sheet_name="NRE", index=False)

    return buf.getvalue()


def default_test_ids_for_filters(
    functionality: str,
    gold_rail: str,
    ufit_sor: str,
    *,
    account: str,
    test_plan_df: pd.DataFrame | None = None,
) -> list[str]:
    """
    Heuristic pool filter for NRE when timeline is empty.
    Uses abbrv / name cues from sample catalog; user can override selection.
    """
    df = test_plan_df if test_plan_df is not None else load_test_plan_info(account)
    ids = set(df["Test_ID"].astype(str))

    # Always include baseline items
    keep = set(ids)

    if gold_rail == "No":
        keep -= {tid for tid in ids if tid in ("T007", "T008")}
    if ufit_sor == "No":
        keep -= {tid for tid in ids if tid in ("T005", "T006")}
    if functionality == "Functional":
        keep -= {"T010"}
    else:
        keep -= {"T009"}

    # Preserve catalog order
    ordered = [str(t) for t in df["Test_ID"] if str(t) in keep]
    return ordered
