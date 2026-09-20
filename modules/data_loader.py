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
    TW_HOLIDAY_CDN,
    TW_HOLIDAY_CDN_FALLBACK,
    account_paths,
)


@st.cache_data(show_spinner=False)
def load_test_plan_info(account: str) -> pd.DataFrame:
    path = account_paths(account)["test_plan_info"]
    df = pd.read_csv(path)
    required = ["Test_ID", "Testplan_Item", "Duration_Days", "Abbrv_Name", "Lab_Fee"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} missing columns: {missing}")
    df["Test_ID"] = df["Test_ID"].astype(str).str.strip()
    df["Duration_Days"] = df["Duration_Days"].astype(int)
    df["Lab_Fee"] = pd.to_numeric(df["Lab_Fee"], errors="coerce").fillna(0)
    return df


@st.cache_data(show_spinner=False)
def load_location_info(account: str) -> pd.DataFrame:
    path = account_paths(account)["location_info"]
    df = pd.read_csv(path)
    df["Test_ID"] = df["Test_ID"].astype(str).str.strip()
    return df


@st.cache_data(show_spinner=False)
def load_convert_table(account: str) -> pd.DataFrame:
    path = account_paths(account)["convert_table"]
    return pd.read_csv(path)


def lookup_convert_id(
    account: str,
    functionality: str,
    gold_rail: str,
    ufit_sor: str,
    system_number: int,
) -> int | None:
    df = load_convert_table(account)
    mask = (
        (df["Functionality"] == functionality)
        & (df["Gold_Rail_Selection"] == gold_rail)
        & (df["Ufit_for_SoR"] == ufit_sor)
        & (df["System_Number"] == int(system_number))
    )
    hits = df.loc[mask]
    if hits.empty:
        return None
    return int(hits.iloc[0]["Convert_ID"])


def load_case_sequence(account: str, convert_id: int) -> pd.DataFrame | None:
    """
    Load sequence template for a Convert_ID from the account workbook
    (one sheet per case: Case_01 … Case_24).

    Sheet layout (header on Excel row 3):
      Row | Sequence
      System 1 | 1, T003, T004, T006, 5, T010
    Tokens: integer = blank fillable days; Test_ID = run that item (uses Duration_Days).
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


def parse_sequence_tokens(raw: Any) -> list[str]:
    """Split a Sequence cell into tokens (commas / whitespace)."""
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
    """
    holidays: dict[str, str] = {}
    for year in years:
        try:
            records = _fetch_year_calendar(int(year))
        except Exception:
            st.warning(
                f"Taiwan holidays for {year} unavailable (CDN/proxy). "
                "Weekends are still treated as non-working; "
                f"optional offline cache: data/holidays/{year}.json"
            )
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
            "Abbrv_Name": row["Abbrv_Name"],
            "Lab_Fee": float(row["Lab_Fee"]),
        }
    return out
