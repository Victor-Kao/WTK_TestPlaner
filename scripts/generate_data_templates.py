"""Generate template CSV / XLSX data files for WTK Test Planner.

Writes per-account folders under data/<ACCOUNT>/:
  - test_plan_info.csv
  - location_info.csv
  - convert_table.csv   (32 case combos; Standard × … × systems 1–2)
  - test_item_sequence_all_cases.xlsx  (48 sheets: Case_01 … Case_48)
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from modules.config import ORV3_MGX_WO_L11_COL

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
ACCOUNTS = ["ROSA", "NAOMI"]
STANDARDS = ['EIA - 19"', 'OCP - 21" / MGX rack']

# ---------------------------------------------------------------------------
# Sample Test Plan Info Table (user can replace / extend)
# Columns: Test_ID, Testplan_Item, Duration_Days, Duration_for_NRE, Abbrv_Name, Lab_Rate
# Duration_Days = timeline days; Duration_for_NRE = billable hours; Lab_Rate = NT$/hr
# Lab fee for NRE = Duration_for_NRE * Lab_Rate
# ---------------------------------------------------------------------------
TEST_PLAN_ROWS = [
    # Test_ID, item, days, nre_hours, abbrv, lab_rate_per_hr
    ("T001", "System Boot Functional Check", 1, 8, "Boot-FC", 1875.00),
    ("T002", "Power Rail Bring-up", 2, 16, "PWR-BU", 1562.50),
    ("T003", "Thermal Soak Test", 3, 24, "THM-SK", 1666.67),
    ("T004", "Signal Integrity Sweep", 2, 16, "SI-SW", 2187.50),
    ("T005", "ORV3 Connector Validation", 2, 16, "ORV3-CV", 1875.00),
    ("T006", "System on Rail Fit Check", 1, 8, "SoR-Fit", 2500.00),
    ("T007", "Gold Rail Alignment", 2, 16, "GR-ALN", 1750.00),
    ("T008", "Gold Rail Continuity", 1, 8, "GR-CON", 2250.00),
    ("T009", "Functional Stress Run", 3, 24, "FNC-ST", 1875.00),
    ("T010", "Non-Functional Endurance", 4, 32, "NF-END", 1562.50),
    ("T011", "Concept Smoke Test", 1, 8, "CPT-SM", 1500.00),
    ("T012", "BCT Full Sequence", 5, 40, "BCT-FS", 2000.00),
    ("T013", "NOT Regression Pack", 3, 24, "NOT-RG", 2291.67),
    ("T014", "EMI / EMC Scan", 2, 16, "EMI-SC", 2625.00),
    ("T015", "Acoustic Noise Profile", 1, 8, "ACQ-NP", 2000.00),
    ("T016", "Idle Power Baseline", 1, 8, "IDL-PW", 1750.00),
    ("T017", "Peak Load Characterization", 2, 16, "PK-LD", 2000.00),
    ("T018", "Failover / Recovery", 2, 16, "FO-RC", 2250.00),
    ("T019", "Firmware Flash Verify", 1, 8, "FW-FV", 1250.00),
    ("T020", "Mechanical Fit Gauge", 1, 8, "MECH-FG", 2750.00),
]

# Location lookup for NRE (separate external CSV)
LOCATION_ROWS = [
    ("T001", "Lab-A / Hsinchu"),
    ("T002", "Lab-A / Hsinchu"),
    ("T003", "Lab-B / Taipei"),
    ("T004", "Lab-B / Taipei"),
    ("T005", "Lab-C / Taoyuan"),
    ("T006", "Lab-C / Taoyuan"),
    ("T007", "Lab-A / Hsinchu"),
    ("T008", "Lab-A / Hsinchu"),
    ("T009", "Lab-B / Taipei"),
    ("T010", "Lab-B / Taipei"),
    ("T011", "Lab-A / Hsinchu"),
    ("T012", "Lab-C / Taoyuan"),
    ("T013", "Lab-C / Taoyuan"),
    ("T014", "Lab-EMC / Linkou"),
    ("T015", "Lab-ACO / Linkou"),
    ("T016", "Lab-A / Hsinchu"),
    ("T017", "Lab-A / Hsinchu"),
    ("T018", "Lab-B / Taipei"),
    ("T019", "Lab-A / Hsinchu"),
    ("T020", "Lab-MECH / Taoyuan"),
]


def all_case_combinations() -> list[dict]:
    """
    2 * 2 * 2 * 2 * 3 = 48 combinations (Convert_ID 1..48).
    Dimensions: Standard × Functionality × Gold_Rail × Only_ORv3_MGX_wo_L11 × System_Number(1–3).
    Systems 4–10: no template — user fills the timeline manually.
    """
    rows: list[dict] = []
    convert_id = 1
    for standard in STANDARDS:
        for functionality in ("Functional", "Non-functional"):
            for gold_rail in ("Yes", "No"):
                for ufit_sor in ("Yes", "No"):
                    for system_number in (1, 2, 3):
                        rows.append(
                            {
                                "Standard": standard,
                                "Functionality": functionality,
                                "Gold_Rail_Selection": gold_rail,
                                ORV3_MGX_WO_L11_COL: ufit_sor,
                                "System_Number": system_number,
                                "Convert_ID": convert_id,
                                "Sheet_Name": f"Case_{convert_id:02d}",
                            }
                        )
                        convert_id += 1
    return rows


def _style_header(ws) -> None:
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def write_account_data(account: str) -> dict[str, Path]:
    """Write all catalog / sequence files into data/<account>/."""
    base = DATA / account
    base.mkdir(parents=True, exist_ok=True)
    written: dict[str, Path] = {}

    plan = pd.DataFrame(
        TEST_PLAN_ROWS,
        columns=[
            "Test_ID",
            "Testplan_Item",
            "Duration_Days",
            "Duration_for_NRE",
            "Abbrv_Name",
            "Lab_Rate",
        ],
    )
    path = base / "test_plan_info.csv"
    plan.to_csv(path, index=False)
    written["test_plan_info"] = path

    loc = pd.DataFrame(LOCATION_ROWS, columns=["Test_ID", "Location"])
    path = base / "location_info.csv"
    loc.to_csv(path, index=False)
    written["location_info"] = path

    cases = all_case_combinations()
    convert_df = pd.DataFrame(cases)
    path = base / "convert_table.csv"
    convert_df.to_csv(path, index=False)
    written["convert_table"] = path

    # Remove obsolete duplicate index if present
    obsolete_index = base / "test_item_sequence_all_cases_index.csv"
    if obsolete_index.exists():
        obsolete_index.unlink()

    # Sequence workbook: one sheet per Convert_ID
    # Format: Row | 1 | 2 | 3 | …  (one token per cell; add more numbered columns in Excel as needed)
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    # Starter columns only — not a hard limit; loader reads every numbered column present
    starter_steps = 50
    step_headers = [str(i) for i in range(1, starter_steps + 1)]
    header_note = (
        "Column A = system label (rename freely, e.g. DUT-A / Chamber-1). "
        "One token per cell under columns 1, 2, 3, … (add more numbered columns if needed). "
        "Integer = blank fillable days after ETA; Test_ID = run that item "
        "(duration from test_plan_info.csv). "
        "Example: 1 | T003 | T004 | T006 | 5 | T010 → blank 1 day, then T003, T004, T006, "
        "blank 5 days, then T010."
    )

    for case in cases:
        sheet_name = case["Sheet_Name"]
        n_sys = int(case["System_Number"])
        ws = wb.create_sheet(title=sheet_name)

        meta = (
            f"Convert_ID={case['Convert_ID']} | "
            f"Standard={case['Standard']} | "
            f"Functionality={case['Functionality']} | "
            f"Gold_Rail={case['Gold_Rail_Selection']} | "
            f"Only_ORv3_MGX_wo_L11={case[ORV3_MGX_WO_L11_COL]} | "
            f"Systems={n_sys}"
        )
        ws.append([meta] + [""] * starter_steps)
        ws.append([header_note] + [""] * starter_steps)
        ws.append(["Row"] + step_headers)
        fill = PatternFill("solid", fgColor="1F4E79")
        font = Font(color="FFFFFF", bold=True)
        for cell in ws[3]:
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

        for s in range(1, n_sys + 1):
            ws.append([f"System {s}"] + [""] * starter_steps)

        ws.column_dimensions["A"].width = 14
        for idx in range(2, min(starter_steps + 2, 22)):
            ws.column_dimensions[ws.cell(row=3, column=idx).column_letter].width = 10
        ws.row_dimensions[2].height = 48
        ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")

    xlsx_path = base / "test_item_sequence_all_cases.xlsx"
    wb.save(xlsx_path)
    written["sequence_xlsx"] = xlsx_path

    return written


def write_nre_template() -> Path:
    """Blank NRE export template matching expected columns."""
    wb = Workbook()
    ws = wb.active
    ws.title = "NRE_Estimate"
    headers = [
        "Test_ID",
        "Test_Item",
        "Location",
        "Duration_for_NRE",
        "Lab_Rate",
        "Lab_Fee",
        "Qty",
        "Total_Fee",
        "Phase",
        "Functionality",
        "Gold_Rail_Selection",
        "Only_ORv3_MGX_Mini_Rack_wo_L11_Rack",
        "Final_Fee",
    ]
    ws.append(headers)
    _style_header(ws)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18

    summary = wb.create_sheet("Summary")
    summary.append(["Account", ""])
    summary.append(["System_Weight_kg", ""])
    summary.append(["Selected_Phases", ""])
    summary.append(["Gold_Rail_Selection", ""])
    summary.append(["Only_ORv3_MGX_Mini_Rack_wo_L11_Rack", ""])
    summary.append(["Grand_Total_Fee", ""])

    path = ROOT / "templates" / "nre_template.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def main() -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    for account in ACCOUNTS:
        print(f"=== {account} ===")
        for name, path in write_account_data(account).items():
            print(f"  {name}: {path}")
    print("Writing NRE template ...", write_nre_template())
    print("Done.")


if __name__ == "__main__":
    main()
