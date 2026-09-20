from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook

from modules.config import NRE_MULTIPLIERS, NRE_TEMPLATE_XLSX
from modules.data_loader import load_location_info, load_test_plan_info, test_info_by_id


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


def build_nre_table(
    phase_configs: list[dict[str, Any]],
    *,
    account: str,
    test_plan_df: pd.DataFrame | None = None,
    location_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """
    Build NRE rows from per-phase configs.

    Each config:
      phase, functionality, gold_rail, ufit_sor, test_ids

    Lab_Fee = Duration_for_NRE × Lab_Rate (hours × hourly rate).
    Final_Fee = Lab_Fee (same base charge; phase / options are recorded as columns only).
    """
    plan = test_plan_df if test_plan_df is not None else load_test_plan_info(account)
    loc_df = location_df if location_df is not None else load_location_info(account)
    info = test_info_by_id(plan)
    loc = loc_df.set_index("Test_ID")["Location"].to_dict()
    rows: list[dict] = []

    for cfg in phase_configs:
        phase = str(cfg.get("phase", ""))
        functionality = str(cfg.get("functionality", "Functional"))
        gold_rail = str(cfg.get("gold_rail", "No"))
        ufit_sor = str(cfg.get("ufit_sor", "No"))
        test_ids = list(cfg.get("test_ids") or [])
        for tid in test_ids:
            meta = info.get(str(tid))
            if not meta:
                continue
            hours = float(meta["Duration_for_NRE"])
            rate = float(meta["Lab_Rate"])
            lab_fee = lab_fee_for_item(meta)
            rows.append(
                {
                    "Test_ID": meta["Test_ID"],
                    "Test_Item": meta["Testplan_Item"],
                    "Location": loc.get(meta["Test_ID"], ""),
                    "Duration_for_NRE": hours,
                    "Lab_Rate": rate,
                    "Lab_Fee": lab_fee,
                    "Qty": 1,
                    "Total_Fee": lab_fee,
                    "Phase": phase,
                    "Functionality": functionality,
                    "Gold_Rail_Selection": gold_rail,
                    "Ufit_for_SoR": ufit_sor,
                    "Final_Fee": lab_fee,
                }
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


def export_nre_xlsx(
    nre_df: pd.DataFrame,
    meta: dict,
    template_path: Path | None = None,
) -> bytes:
    """Fill NRE template (or create workbook) and return bytes."""
    template = template_path or NRE_TEMPLATE_XLSX
    buf = BytesIO()

    if template.exists():
        wb = load_workbook(template)
        ws = wb["NRE_Estimate"] if "NRE_Estimate" in wb.sheetnames else wb.active
        # clear old data rows
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for record in nre_df.to_dict(orient="records"):
            ws.append(
                [
                    record.get("Test_ID"),
                    record.get("Test_Item"),
                    record.get("Location"),
                    record.get("Duration_for_NRE"),
                    record.get("Lab_Rate"),
                    record.get("Lab_Fee"),
                    record.get("Qty"),
                    record.get("Total_Fee"),
                    record.get("Phase"),
                    record.get("Functionality"),
                    record.get("Gold_Rail_Selection"),
                    record.get("Ufit_for_SoR"),
                    record.get("Final_Fee"),
                ]
            )
        if "Summary" in wb.sheetnames:
            summary = wb["Summary"]
            phases = meta.get("phases", [])
            mapping = {
                "Account": meta.get("account", ""),
                "System_Weight_kg": meta.get("weight_kg", ""),
                "Selected_Phases": ", ".join(phases),
                "Gold_Rail_Selection": meta.get("gold_rail", "per phase"),
                "Phase_for_Gold_Rail_Selection": meta.get(
                    "gold_rail_phase", "per phase"
                ),
                "Ufit_for_SoR": meta.get("ufit_sor", "per phase"),
                "Grand_Total_Fee": float(nre_df["Final_Fee"].sum())
                if not nre_df.empty
                else 0,
            }
            for row in summary.iter_rows(min_row=1, max_row=summary.max_row, max_col=2):
                key = row[0].value
                if key in mapping:
                    row[1].value = mapping[key]
        wb.save(buf)
    else:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            nre_df.to_excel(writer, sheet_name="NRE_Estimate", index=False)
            pd.DataFrame([meta]).to_excel(writer, sheet_name="Summary", index=False)

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
