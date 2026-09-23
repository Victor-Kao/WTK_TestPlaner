from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
TEMPLATES_DIR = ROOT / "templates"

ACCOUNTS = ["ROSA", "NAOMI"]
PHASES = ["Concept", "BCT", "NOT"]
FUNCTIONALITY_OPTS = ["Functional", "Non-functional"]
YES_NO = ["Yes", "No"]
SYSTEM_NUMBERS = list(range(1, 11))  # 1…10 systems on the timeline
# Convert table + Case sequence sheets only define templates through this many systems
SEQUENCE_TEMPLATE_MAX_SYSTEMS = 2
STANDARDS = ['EIA - 19"', 'OCP - 21" / MGX rack']

EMPTY_TOKEN = "EMPTY"
EMPTY_LABEL = "Empty"

# Cell-dropdown actions: shift schedule from the selected date onward
CELL_AHEAD_OPT = "◀ Ahead (−1) from here"
CELL_POSTPONE_OPT = "Postpone (+1) from here ▶"

NRE_TEMPLATE_XLSX = TEMPLATES_DIR / "nre_template.xlsx"

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
    "ufit_sor": {"Yes": 1.05, "No": 1.0},
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
    }
