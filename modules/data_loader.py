from __future__ import annotations

import json
import time
from functools import lru_cache
from typing import Any

import pandas as pd
import requests
import streamlit as st

from modules.config import (
    HOLIDAY_CACHE_MAX_AGE_DAYS,
    HOLIDAYS_DIR,
    ORV3_MGX_WO_L11_COL,
    TW_HOLIDAY_CDN,
    TW_HOLIDAY_CDN_FALLBACK,
    account_paths,
)


@st.cache_data(show_spinner=False)
def load_test_plan_info(account: str) -> pd.DataFrame:
    path = account_paths(account)["test_plan_info"]
    df = pd.read_csv(path)
    required = [
        "Test_ID",
        "Testplan_Item",
        "Duration_Days",
        "Duration_for_NRE",
        "Abbrv_Name",
        "Lab_Rate",
    ]
    # Migrate older CSVs that stored a flat Lab_Fee instead of rate × hours
    if "Lab_Rate" not in df.columns and "Lab_Fee" in df.columns:
        days = df["Duration_Days"].astype(int).clip(lower=1)
        if "Duration_for_NRE" not in df.columns:
            df["Duration_for_NRE"] = days * 8
        hours = pd.to_numeric(df["Duration_for_NRE"], errors="coerce").fillna(1).clip(lower=1)
        old_fee = pd.to_numeric(df["Lab_Fee"], errors="coerce").fillna(0)
        df["Lab_Rate"] = (old_fee / hours).round(2)
        df = df.drop(columns=["Lab_Fee"])
    if "Duration_for_NRE" not in df.columns:
        df["Duration_for_NRE"] = df["Duration_Days"].astype(int).clip(lower=1) * 8
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} missing columns: {missing}")
    df["Test_ID"] = df["Test_ID"].astype(str).str.strip()
    df["Duration_Days"] = df["Duration_Days"].astype(int)
    df["Duration_for_NRE"] = pd.to_numeric(df["Duration_for_NRE"], errors="coerce").fillna(0)
    df["Lab_Rate"] = pd.to_numeric(df["Lab_Rate"], errors="coerce").fillna(0)
    return df


PROFILE_COLUMNS = (
    "PROFILE_S40LBS",
    "PROFILE_S75LBS",
    "PROFILE_S200LBS",
    "PROFILE_L200LBS",
)


def profile_column_for_weight(weight_kg: float) -> str:
    """
    Pick detail profile column from system weight (kg):
      < 18      → PROFILE_S40LBS
      [18, 34)  → PROFILE_S75LBS
      [34, 90)  → PROFILE_S200LBS
      ≥ 90      → PROFILE_L200LBS
    """
    w = float(weight_kg)
    if w < 18:
        return "PROFILE_S40LBS"
    if w < 34:
        return "PROFILE_S75LBS"
    if w < 90:
        return "PROFILE_S200LBS"
    return "PROFILE_L200LBS"


@st.cache_data(show_spinner=False)
def load_test_plan_info_detail(account: str) -> pd.DataFrame:
    """Load weight-profile detail text per Test_ID (optional file)."""
    path = account_paths(account)["test_plan_info_detail"]
    if not path.exists():
        return pd.DataFrame(columns=["Test_ID", *PROFILE_COLUMNS])
    df = pd.read_csv(path)
    # Drop trailing unnamed empty columns from Excel exports
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    if "Test_ID" not in df.columns:
        raise ValueError(f"{path.name} missing Test_ID column")
    df["Test_ID"] = df["Test_ID"].astype(str).str.strip()
    for col in PROFILE_COLUMNS:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].fillna("").astype(str).str.strip()
            df.loc[df[col].str.lower().isin(("nan", "none")), col] = ""
    return df


def test_detail_by_id(
    df: pd.DataFrame | None = None,
    *,
    account: str | None = None,
) -> dict[str, dict[str, str]]:
    if df is None:
        if not account:
            raise ValueError("Provide df or account for test_detail_by_id")
        df = load_test_plan_info_detail(account)
    out: dict[str, dict[str, str]] = {}
    for _, row in df.iterrows():
        tid = str(row["Test_ID"]).strip()
        if not tid:
            continue
        out[tid] = {col: str(row.get(col, "") or "").strip() for col in PROFILE_COLUMNS}
    return out


@st.cache_data(show_spinner=False)
def load_location_info(account: str) -> pd.DataFrame:
    path = account_paths(account)["location_info"]
    df = pd.read_csv(path)
    df["Test_ID"] = df["Test_ID"].astype(str).str.strip()
    return df


@st.cache_data(show_spinner=False)
def load_convert_table(account: str) -> pd.DataFrame:
    path = account_paths(account)["convert_table"]
    df = pd.read_csv(path)
    if (
        ORV3_MGX_WO_L11_COL not in df.columns
        and "Ufit_for_SoR" in df.columns
    ):
        df = df.rename(columns={"Ufit_for_SoR": ORV3_MGX_WO_L11_COL})
    return df


def lookup_convert_id(
    account: str,
    functionality: str,
    gold_rail: str,
    ufit_sor: str,
    system_number: int,
    *,
    standard: str | None = None,
    convert_df: pd.DataFrame | None = None,
) -> int | None:
    df = convert_df if convert_df is not None else load_convert_table(account)
    mask = (
        (df["Functionality"] == functionality)
        & (df["Gold_Rail_Selection"] == gold_rail)
        & (df[ORV3_MGX_WO_L11_COL] == ufit_sor)
        & (df["System_Number"].astype(int) == int(system_number))
    )
    if standard is not None and "Standard" in df.columns:
        mask = mask & (df["Standard"] == standard)
    hits = df.loc[mask]
    if hits.empty:
        return None
    return int(hits.iloc[0]["Convert_ID"])


def load_case_sequence(account: str, convert_id: int) -> pd.DataFrame | None:
    """
    Load sequence template for a Convert_ID from the account workbook
    (one sheet per case: Case_01 … Case_48; Standard × … × systems 1–3).

    Sheet layout (header on Excel row 3) — one token per column:
      Row | 1 | 2 | 3 | 4 | 5 | …
      DUT-A | 1 | T003 | T004 | T006 | 5 | T010
    Column A is the system label (any name; used on the timeline).
    Tokens: integer = blank fillable days; Test_ID = run that item (Duration_Days).
    Add as many numbered columns as needed — there is no fixed step limit.
    """
    xlsx = account_paths(account)["sequence_xlsx"]
    if not xlsx.exists():
        return None
    sheet = f"Case_{int(convert_id):02d}"
    try:
        # Row 1 meta, row 2 note, row 3 headers
        return pd.read_excel(xlsx, sheet_name=sheet, header=2)
    except Exception:
        return None


def sequence_step_columns(columns) -> list:
    """Ordered step columns (1, 2, 3… or Step_1…). Skips Row / Sequence / Day_*."""
    numbered: list[tuple[int, Any]] = []
    other: list[Any] = []
    for c in columns:
        s = str(c).strip()
        low = s.lower()
        if low in ("row", "sequence") or s.startswith("Unnamed"):
            continue
        if s.startswith("Day_"):
            continue
        if s.isdigit():
            numbered.append((int(s), c))
        elif low.startswith("step_") and s[5:].isdigit():
            numbered.append((int(s[5:]), c))
        elif low.startswith("step") and s[4:].isdigit():
            numbered.append((int(s[4:]), c))
        else:
            other.append(c)
    if numbered:
        return [c for _, c in sorted(numbered, key=lambda x: x[0])]
    return other


def tokens_from_sequence_row(seq_row: Any, columns) -> list[str]:
    """
    Collect sequence tokens from a system row.
    Prefer one-token-per-cell step columns; fall back to legacy comma Sequence cell.
    Empty cells are skipped (not treated as blank days).
    """
    step_cols = sequence_step_columns(columns)
    if step_cols:
        tokens: list[str] = []
        for c in step_cols:
            raw = seq_row.get(c) if hasattr(seq_row, "get") else seq_row[c]
            if raw is None or (isinstance(raw, float) and pd.isna(raw)):
                continue
            text = str(raw).strip()
            if not text:
                continue
            tokens.append(text)
        return tokens
    cols = list(columns)
    if "Sequence" in cols:
        return parse_sequence_tokens(seq_row.get("Sequence"))
    return []


def parse_sequence_tokens(raw: Any) -> list[str]:
    """Legacy: split a single Sequence cell into tokens (commas / whitespace)."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return []
    text = str(raw).strip()
    if not text:
        return []
    for sep in (";", "|", "\n", "\t"):
        text = text.replace(sep, ",")
    if "," not in text:
        return [p for p in text.split() if p]
    return [p.strip() for p in text.split(",") if p.strip()]


def _holiday_cache_path(year: int):
    return HOLIDAYS_DIR / f"{int(year)}.json"


def _cache_is_fresh(year: int) -> bool:
    path = _holiday_cache_path(year)
    if not path.exists():
        return False
    age_sec = time.time() - path.stat().st_mtime
    return age_sec <= HOLIDAY_CACHE_MAX_AGE_DAYS * 86400


def _load_local_year_calendar(year: int) -> tuple[dict[str, Any], ...] | None:
    path = _holiday_cache_path(year)
    if not path.exists():
        return None
    try:
        with path.open(encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, list):
        return None
    return tuple(data)


def _save_local_year_calendar(year: int, data: tuple[dict[str, Any], ...]) -> None:
    try:
        HOLIDAYS_DIR.mkdir(parents=True, exist_ok=True)
        path = _holiday_cache_path(year)
        path.write_text(
            json.dumps(list(data), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass


def _download_year_calendar(year: int) -> tuple[dict[str, Any], ...]:
    """
    Fetch from CDN.
    Try without env proxies first (avoids broken corporate HTTP_PROXY),
    then with system proxies as a second attempt.
    """
    last_err: Exception | None = None
    for trust_env in (False, True):
        session = requests.Session()
        session.trust_env = trust_env
        for template in (TW_HOLIDAY_CDN, TW_HOLIDAY_CDN_FALLBACK):
            url = template.format(year=year)
            try:
                resp = session.get(url, timeout=15)
                resp.raise_for_status()
                data = resp.json()
                if isinstance(data, list) and data:
                    return tuple(data)
            except Exception as exc:
                last_err = exc
                continue
    raise RuntimeError(f"CDN fetch failed for {year}: {last_err}")


@lru_cache(maxsize=16)
def _fetch_year_calendar(year: int) -> tuple[dict[str, Any], ...]:
    """
    Source of truth = online TaiwanCalendar JSON (auto-updates each year).

    Flow:
      1. If local cache is fresh (< HOLIDAY_CACHE_MAX_AGE_DAYS) → use it
      2. Else try CDN; on success write/overwrite data/holidays/{year}.json
      3. If CDN fails → use stale local cache if present
      4. Otherwise raise (caller shows a soft warning)
    """
    year = int(year)
    if _cache_is_fresh(year):
        local = _load_local_year_calendar(year)
        if local is not None:
            return local

    try:
        data = _download_year_calendar(year)
        _save_local_year_calendar(year, data)
        return data
    except Exception:
        local = _load_local_year_calendar(year)
        if local is not None:
            return local
        raise


@st.cache_data(show_spinner="Loading Taiwan holidays…", ttl=86400)
def load_taiwan_holidays(years: tuple[int, ...]) -> dict[str, str]:
    """
    Return {YYYY-MM-DD: description} for official Taiwan holidays / days off.

    Online calendar is preferred and cached under data/holidays/{year}.json so
    new years appear automatically when the upstream publishes them — no hard-coded
    year list in code.

    If a year is not on the CDN yet (gov usually publishes ~June for the next year),
    weekends still count as non-working; named holidays for that year are skipped.
    """
    from datetime import date as _date

    holidays: dict[str, str] = {}
    # Upstream typically has current year + next; further-out years often 404.
    warn_through = _date.today().year + 1
    for year in years:
        y = int(year)
        try:
            records = _fetch_year_calendar(y)
        except Exception:
            if y <= warn_through:
                st.warning(
                    f"Taiwan holidays for {y} unavailable (CDN/proxy). "
                    "Weekends are still treated as non-working; "
                    f"optional offline cache: data/holidays/{y}.json"
                )
            # else: unpublished future year — silent; weekends only
            continue
        for rec in records:
            if not rec.get("isHoliday"):
                continue
            desc = (rec.get("description") or "").strip()
            if not desc:
                continue
            raw = str(rec.get("date", ""))
            if len(raw) == 8 and raw.isdigit():
                iso = f"{raw[0:4]}-{raw[4:6]}-{raw[6:8]}"
            else:
                iso = raw
            holidays[iso] = desc
    return holidays


def test_info_by_id(
    df: pd.DataFrame | None = None,
    *,
    account: str | None = None,
) -> dict[str, dict]:
    if df is None:
        if not account:
            raise ValueError("Provide df or account for test_info_by_id")
        df = load_test_plan_info(account)
    out: dict[str, dict] = {}
    for _, row in df.iterrows():
        out[str(row["Test_ID"])] = {
            "Test_ID": str(row["Test_ID"]),
            "Testplan_Item": row["Testplan_Item"],
            "Duration_Days": int(row["Duration_Days"]),
            "Duration_for_NRE": float(row["Duration_for_NRE"]),
            "Abbrv_Name": row["Abbrv_Name"],
            "Lab_Rate": float(row["Lab_Rate"]),
        }
    return out
