# WTK Test Planner

Streamlit schedule tool for brand test plans (ROSA, NAOMI), timeline editing, NRE estimate, and case templates.

**ROSA** — full Test Plan Timeline + NRE flows are implemented.  
**NAOMI** — different workflow; Test Plan / NRE / Headcount are TBD placeholders for now.

## Requirements

- Python 3.10+ recommended
- Packages listed in `requirements.txt` (Streamlit, pandas, openpyxl, requests, …)

---

## Install & run (with virtual environment — recommended)

A venv keeps project packages separate from your system Python.

### macOS / Linux

```bash
cd WTK_TestPlaner

# 1) Create venv (once)
python3 -m venv .venv

# 2) Activate venv (every new terminal)
source .venv/bin/activate

# 3) Install libraries
pip install -r requirements.txt

# 4) (Optional) regenerate data templates
python scripts/generate_data_templates.py

# 5) Start Streamlit
streamlit run app.py
```

### Windows (Command Prompt)

```bat
cd WTK_TestPlaner

python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
streamlit run app.py
```

### Windows (PowerShell)

```powershell
cd WTK_TestPlaner

python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

After activate, your prompt usually shows `(.venv)`. To leave the venv:

```bash
deactivate
```

---

## Install & run (without virtual environment)

Installs packages into the current Python (user or system). Use if you prefer not to use a venv.

```bash
cd WTK_TestPlaner

# Install libraries for the current user (safer than system-wide)
python3 -m pip install --user -r requirements.txt

# Or system / active Python:
# python3 -m pip install -r requirements.txt

# Start Streamlit
python3 -m streamlit run app.py
```

On Windows, replace `python3` with `python` if that is your command.

---

## Open the app

After `streamlit run app.py` (or `python3 -m streamlit run app.py`):

1. Terminal prints a **Local URL**, usually: http://localhost:8501  
2. Open that URL in your browser.  
3. Stop the server with `Ctrl+C` in the terminal.

### Useful Streamlit options

```bash
# Fixed port
streamlit run app.py --server.port 8501

# Don’t open a browser automatically
streamlit run app.py --server.headless true
```

---

## Shared calendar data (`data/holidays/`)

**CDN is the source of truth** ([ruyut/TaiwanCalendar](https://github.com/ruyut/TaiwanCalendar)). The app:

1. Requests the year(s) needed for the timeline (from System ETA / feedback span — not a fixed year list in code).
2. Caches each year as `data/holidays/{year}.json`.
3. Re-fetches when a cache file is older than 30 days (`HOLIDAY_CACHE_MAX_AGE_DAYS`).
4. If the network/proxy fails, uses the last good cache so planning still works offline.

You do not need to hand-maintain new years: when upstream publishes e.g. `2028.json`, the next timeline that spans 2028 will download and cache it automatically.

## Data layout (`data/<ACCOUNT>/`)

Each account has its own folder (`data/ROSA/`, `data/NAOMI/`):

| File | Purpose |
|------|---------|
| `test_plan_info.csv` | Test_ID, Testplan_Item, Duration_Days (timeline days), Duration_for_NRE (billable hours), Abbrv_Name, Lab_Rate (per hour) |
| `location_info.csv` | Test_ID → Location (for NRE) |
| `convert_table.csv` | 32 case combos → Convert_ID + Sheet_Name (Standard × … × systems 1–2) |
| `test_item_sequence_all_cases.xlsx` | 32 sheets (`Case_01` … `Case_32`); one token per cell |

### Convert dimensions (2×2×2×2×2 = 32)

- Standard: **EIA - 19"** / **OCP - 21"**
- Functionality: Functional / Non-functional
- Gold Rail Selection: Yes / No
- U-fit for L10.5 SoR / ORv3 mini Rack: Yes / No
- System Number: **1 / 2** (templates only)

The timeline UI allows **1–10** systems. From **3** systems up, fill the timeline manually.

### Sequence sheet format

Each system row uses columns **1, 2, 3, …** (one token per cell). Add more numbered columns in Excel anytime — there is no step limit; the app reads all of them.

The first column (**Row**) can be renamed from `System 1` / `System 2` to any label (e.g. `DUT-A`). Those names appear on the timeline after pre-fill.

| Row | 1 | 2 | 3 | 4 | 5 | 6 |
|-----|---|---|---|---|---|---|
| DUT-A | 1 | T003 | T004 | T006 | 5 | T010 |

- **Integer** → that many **blank** fillable days (after System ETA)
- **Test_ID** → run that item next (length = `Duration_Days` from `test_plan_info.csv`)

Example above → blank 1 day, then T003, T004, T006, blank 5 days, then T010.

## Flow

1. On the main page, set **Account**, **System weight**, and **Standard**, then click **Start Arranging Test Plan** (no app sidebar — reserved for a parent tool).
2. Set System ETA, critical feedback, # systems (left); functional type, U-fit, gold rail (right).
3. **Generate Test Plan** → timeline with weekdays, weekends, Taiwan holidays, ETA, feedback.
4. Select a cell → assign a test item (or use Event row for Empty). Same-row no overlap; multi-day duration auto-spans.
5. **Update** → export table (Date / Event / System rows) → download CSV.
6. **NRE Estimation** tab → pick phases → each phase section sets functionality / gold rail / U-fit / test items → generate & download.
7. **Data preview** at the bottom of the page is collapsed by default.
8. Headcount tab is TBD.
