from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
TEMPLATES_DIR = ROOT / "templates"

ACCOUNTS = ["ROSA", "NAOMI"]
PHASES = ["Concept", "BCT", "NOT"]  # NRE testing phases
PROJECT_PHASES = ["C0", "C1", "C2", "C3", "C4", "C5", "C6"]  # setup project phase
FUNCTIONALITY_OPTS = ["Functional", "Non-functional"]
YES_NO = ["Yes", "No"]
SYSTEM_NUMBERS = list(range(1, 11))  # 1…10 systems on the timeline
# Convert table + Case sequence sheets only define templates through this many systems
SEQUENCE_TEMPLATE_MAX_SYSTEMS = 3
STANDARDS = ['EIA - 19"', 'OCP - 21" / MGX rack']

# Convert-table / sequence flag (Yes/No): Only ORv3 / MGX Mini Rack w.o L11 Rack
ORV3_MGX_WO_L11_COL = "Only_ORv3_MGX_Mini_Rack_wo_L11_Rack"
ORV3_MGX_WO_L11_LABEL = "Only ORv3 / MGX Mini Rack w.o L11 Rack"

EMPTY_TOKEN = "EMPTY"
EMPTY_LABEL = "Empty"

# Cell-dropdown actions: shift schedule from the selected date onward
CELL_AHEAD_OPT = "◀ Ahead (−1) from here"
CELL_POSTPONE_OPT = "Postpone (+1) from here ▶"

NRE_TEMPLATE_XLSX = TEMPLATES_DIR / "nre_template.xlsx"

# NRE extras shown under the test-item selector (not in the multiselect list)
NRE_OM_ID = "REL0164_OM"
NRE_DNP_ID = "REL0164_DNP"
NRE_UFIT_ID = "ENG0013791_UFIT"  # U-fit / Leading Edge
NRE_AUX_COST_IDS = ("AUX_1", "AUX_2", "AUX_3")  # Wooden Fixture, DummyWeight, Rack Building
NRE_AUX_COST_DEFAULTS_USD = {
    "AUX_1": 3000.0,  # Wooden Fixture
    "AUX_2": 0.0,  # DummyWeight
    "AUX_3": 0.0,  # Rack Building / Modification
}
# Free / non-billable catalog items — never offered in the NRE multiselect
NRE_FREE_TEST_IDS = frozenset({"ENG0013791_CHECK"})  # Pre-Process
NRE_SELECTOR_EXTRA_IDS = frozenset(
    {
        NRE_OM_ID,
        NRE_DNP_ID,
        NRE_UFIT_ID,
        *NRE_AUX_COST_IDS,
        *NRE_FREE_TEST_IDS,
    }
)

# Taiwan holidays: CDN is source of truth; local JSON is a refreshable cache
HOLIDAYS_DIR = DATA_DIR / "holidays"
HOLIDAY_CACHE_MAX_AGE_DAYS = 30  # re-fetch from CDN when cache older than this
TW_HOLIDAY_CDN = "https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/{year}.json"
TW_HOLIDAY_CDN_FALLBACK = (
    "https://raw.githubusercontent.com/ruyut/TaiwanCalendar/master/data/{year}.json"
)

# Fee multipliers used for NRE final fee (editable business rules)
NRE_MULTIPLIERS = {
    "phase": {"Concept": 1.0, "BCT": 1.15, "NOT": 1.25},
    "functionality": {"Functional": 1.0, "Non-functional": 0.9},
    "gold_rail": {"Yes": 1.1, "No": 1.0},
    "ufit_sor": {"Yes": 1.05, "No": 1.0},  # Only_ORv3_MGX_Mini_Rack_wo_L11_Rack flag
}


def account_data_dir(account: str) -> Path:
    """Per-account data folder: data/ROSA, data/NAOMI, …"""
    return DATA_DIR / account


def account_paths(account: str) -> dict[str, Path]:
    base = account_data_dir(account)
    return {
        "dir": base,
        "test_plan_info": base / "test_plan_info.csv",
        "test_plan_info_detail": base / "test_plan_info_detail.csv",
        "location_info": base / "location_info.csv",
        "convert_table": base / "convert_table.csv",
        "sequence_xlsx": base / "test_item_sequence_all_cases.xlsx",
        "nre_template": base / f"NRE_TEMPLATE_{account}.xlsx",
    }


def nre_template_path(account: str) -> Path:
    """Account NRE workbook template (falls back to templates/nre_template.xlsx)."""
    path = account_paths(account)["nre_template"]
    return path if path.exists() else NRE_TEMPLATE_XLSX
