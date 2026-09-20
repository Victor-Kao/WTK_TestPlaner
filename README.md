# WTK Test Planner

Streamlit schedule tool for brand test plans (ROSA, NAOMI), timeline editing, NRE estimate, and case templates.

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
| `test_plan_info.csv` | Test_ID, Testplan_Item, Duration_Days, Abbrv_Name, Lab_Fee |
| `location_info.csv` | Test_ID → Location (for NRE) |
| `convert_table.csv` | 24 case combos → Convert_ID + Sheet_Name (single index) |
| `test_item_sequence_all_cases.xlsx` | 24 sheets (`Case_01` … `Case_24`); one **Sequence** cell per system |

### Convert dimensions (2×2×2×3 = 24)

- Functionality: Functional / Non-functional
- Gold Rail Selection: Yes / No
- U-fit for SoR: Yes / No
- System Number: 1 / 2 / 3

### Sequence sheet format

Each system row has one `Sequence` cell (comma-separated tokens):

- **Integer** → that many **blank** fillable days (after System ETA)
- **Test_ID** → run that item next (length = `Duration_Days` from `test_plan_info.csv`)

Example: `1, T003, T004, T006, 5, T010` → blank 1 day, then T003, T004, T006, blank 5 days, then T010.

## Flow

1. Select account (**ROSA** or **NAOMI**) and system weight.
2. Set phase, functional type, U-fit, System ETA, critical feedback, # systems, gold rail.
3. **Generate Test Plan** → timeline with weekdays, weekends, Taiwan holidays, ETA, feedback.
4. Select a cell → assign a test item (or use Marked row for Empty). Same-row no overlap; multi-day duration auto-spans.
5. **Update** → export table (Date / Marked / System rows) → download CSV.
6. **NRE Estimation** tab → multi-phase fees → download XLSX from template.
7. Headcount tab is TBD.
