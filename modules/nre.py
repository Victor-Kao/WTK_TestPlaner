from __future__ import annotations

from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.worksheet import Worksheet

from modules.config import (
    NRE_AUX_COST_IDS,
    NRE_DNP_ID,
    NRE_MULTIPLIERS,
    NRE_OM_ID,
    NRE_UFIT_ID,
    ORV3_MGX_WO_L11_COL,
    PHASES,
    nre_template_path,
)
from modules.data_loader import load_location_info, load_test_plan_info, test_info_by_id

# ROSA NRE template columns (Sheet1)
_NRE_PHASE_ORDER = ["Concept", "BU", "CT", "BCT", "NT", "OT", "NOT"]
_NRE_PHASE_COLS = {p: p for p in _NRE_PHASE_ORDER}
_NRE_PHASE_EXCLUSIVE_GROUPS = (
    ("BCT", frozenset({"BU", "CT"})),
    ("NOT", frozenset({"NT", "OT"})),
)


def _nre_phase_columns(phases: list[str] | None) -> list[str]:
    """Selected phase columns in canonical order (Concept → … → NOT)."""
    if not phases:
        return list(_NRE_PHASE_ORDER)
    wanted = {str(p) for p in phases}
    ordered = [p for p in _NRE_PHASE_ORDER if p in wanted]
    return ordered or list(_NRE_PHASE_ORDER)


def _nre_headers(phase_cols: list[str]) -> list[str]:
    return [
        "Test Item",
        "Lab Location",
        "Lab Rate",
        *phase_cols,
        "Sub Total",
        "Comments",
    ]


def _nre_subtotal_formula(n_phases: int) -> str:
    """Excel formula template: SUM(phase cols) × Lab Rate."""
    from openpyxl.utils import get_column_letter

    if n_phases <= 0:
        return "=0*C{row}"
    first = get_column_letter(4)
    last = get_column_letter(3 + n_phases)
    return f"=SUM({first}{{row}}:{last}{{row}})*C{{row}}"


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
    """Earliest phase among selection in Concept → BU/CT/BCT → NT/OT/NOT order."""
    selected = {str(p) for p in phases}
    for p in PHASES:
        if p in selected:
            return p
    return str(phases[0]) if phases else None


def sanitize_nre_phases(phases: list[str] | None) -> list[str]:
    """
    Drop conflicting phases (BCT vs BU/CT, NOT vs NT/OT), then return the
    selection in fixed generate order:
    Concept → BU → CT → BCT → NT → OT → NOT
    (independent of the order the user picked them in the multiselect).
    """
    if not phases:
        return []
    selected = {str(p) for p in phases if str(p) in _NRE_PHASE_COLS}
    for combined, parts in _NRE_PHASE_EXCLUSIVE_GROUPS:
        if combined in selected and selected & parts:
            selected -= parts
    return [p for p in _NRE_PHASE_ORDER if p in selected]


def available_nre_phases(selected: list[str] | None) -> list[str]:
    """
    Phase options for the multiselect given the current selection:
      - BU or CT selected → hide BCT
      - BCT selected → hide BU and CT
      - NT or OT selected → hide NOT
      - NOT selected → hide NT and OT
    """
    sel = set(sanitize_nre_phases(list(selected or [])))
    out: list[str] = []
    for p in PHASES:
        skip = False
        for combined, parts in _NRE_PHASE_EXCLUSIVE_GROUPS:
            if p == combined and sel & parts:
                skip = True
                break
            if p in parts and combined in sel:
                skip = True
                break
        if not skip:
            out.append(p)
    return out


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
    earliest selected phase (Concept → BU/CT/BCT → NT/OT/NOT).

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


def pivot_nre_for_template(
    nre_df: pd.DataFrame,
    *,
    phases: list[str] | None = None,
) -> pd.DataFrame:
    """
    Collapse detailed NRE rows into the account template shape:
    Test Item | Lab Location | Lab Rate | <selected phases> | Sub Total | Comments

    Only selected phase columns are included (canonical order). Hours come from
    Duration_for_NRE × Qty. Aux USD rows put the cost in Lab Rate and 1 in the
    phase column so Sub Total (=SUM(phases)*rate) equals the cost.

    U-envelop measurement (U-fit) Comments list runs per phase, one per line, e.g.
    "1x for Concept\\n1x for BCT".
    """
    phase_cols = _nre_phase_columns(phases)
    headers = _nre_headers(phase_cols)
    if nre_df is None or nre_df.empty:
        return pd.DataFrame(columns=headers)

    # Preserve first-seen order of Test_ID
    order: list[str] = []
    buckets: dict[str, dict[str, Any]] = {}
    phase_set = set(phase_cols)

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
                **{p: 0.0 for p in phase_cols},
                "Comments": "",
                "_aux": _is_aux_cost_row(tid),
                "_ufit_qty_by_phase": {},
            }
        b = buckets[tid]
        if _is_aux_cost_row(tid):
            cost = float(rec.get("Final_Fee") or rec.get("Lab_Fee") or 0)
            b["Lab Rate"] = cost
            if phase in phase_set:
                b[phase] = float(b.get(phase, 0) or 0) + 1.0
        else:
            rate = float(rec.get("Lab_Rate") or 0)
            if rate:
                b["Lab Rate"] = rate
            qty = float(rec.get("Qty") or 1)
            hours = float(rec.get("Duration_for_NRE") or 0) * qty
            if phase in phase_set:
                b[phase] = float(b.get(phase, 0) or 0) + hours
            if tid == NRE_UFIT_ID and phase in phase_set and qty > 0:
                qty_map = b["_ufit_qty_by_phase"]
                qty_map[phase] = float(qty_map.get(phase, 0) or 0) + qty
            if not b["Lab Location"] and rec.get("Location"):
                b["Lab Location"] = str(rec.get("Location") or "")

    rows: list[dict[str, Any]] = []
    for tid in order:
        b = buckets[tid]
        phase_hours = {p: float(b.get(p, 0) or 0) for p in phase_cols}
        rate = float(b["Lab Rate"] or 0)
        sub = round(sum(phase_hours.values()) * rate, 2)
        comment = b["Comments"] or ""
        if tid == NRE_UFIT_ID:
            qty_map = b.get("_ufit_qty_by_phase") or {}
            parts = []
            for p in phase_cols:
                q = float(qty_map.get(p, 0) or 0)
                if q <= 0:
                    continue
                q_disp = int(q) if abs(q - int(q)) < 1e-9 else q
                parts.append(f"{q_disp}x for {p}")
            if parts:
                comment = "\n".join(parts)
        rows.append(
            {
                "Test Item": b["Test Item"],
                "Lab Location": b["Lab Location"],
                "Lab Rate": rate,
                **{p: (v if v else None) for p, v in phase_hours.items()},
                "Sub Total": sub,
                "Comments": comment or None,
            }
        )
    return pd.DataFrame(rows, columns=headers)


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


def nre_grand_total(pivoted: pd.DataFrame) -> float:
    """Sum of all row Sub Total values (canonical grand total)."""
    if pivoted is None or pivoted.empty or "Sub Total" not in pivoted.columns:
        return 0.0
    return round(
        float(pd.to_numeric(pivoted["Sub Total"], errors="coerce").fillna(0).sum()),
        2,
    )


def _copy_nre_template_row(
    ws: Worksheet,
    src_row: int,
    dest_row: int,
    *,
    ncols: int,
    subtotal_col: int,
    subtotal_formula: str,
) -> None:
    """Copy cell styles/number formats from src_row and install Sub Total formula."""
    for col in range(1, ncols + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dest_row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)
            dst.alignment = copy(src.alignment)
    ws.cell(dest_row, subtotal_col).value = subtotal_formula.format(row=dest_row)


def _nre_fit_sheet_layout(
    ws: Worksheet,
    *,
    data_start: int,
    last_data: int,
    ncols: int,
    comments_col: int,
) -> None:
    """
    Widen columns, wrap text, and grow row heights so multi-line Comments /
    long Test Item text are fully visible (not clipped by default cell size).
    """
    from openpyxl.utils import get_column_letter

    # Reasonable widths: Test Item wide, Comments wide enough for "1x for Concept"
    min_widths = {
        1: 42,  # Test Item
        2: 18,  # Lab Location
        3: 12,  # Lab Rate
        comments_col: 24,  # Comments
    }
    for col in range(1, ncols + 1):
        letter = get_column_letter(col)
        current = float(ws.column_dimensions[letter].width or 0)
        if col == 1:
            target = 42.0
        elif col == 2:
            target = 18.0
        elif col == 3:
            target = 12.0
        elif col == comments_col:
            target = 24.0
        elif col == comments_col - 1:  # Sub Total
            target = 14.0
        else:
            target = 10.0
        ws.column_dimensions[letter].width = max(current, target)

    for r in range(1, max(last_data, 1) + 1):
        for c in (1, 2, comments_col):
            if c > ncols:
                continue
            cell = ws.cell(r, c)
            cell.alignment = Alignment(
                wrap_text=True,
                vertical="center" if r == 1 else "top",
            )

        if r < data_start or last_data < data_start:
            continue

        lines = 1
        comment = ws.cell(r, comments_col).value
        if comment is not None and str(comment).strip():
            comment_s = str(comment).replace("\r\n", "\n").replace("\r", "\n")
            comment_width = float(
                ws.column_dimensions[get_column_letter(comments_col)].width or 24
            )
            line_count = 0
            for part in comment_s.split("\n"):
                # chars that fit roughly per wrapped line in this column
                chars_per_line = max(int(comment_width), 8)
                line_count += max(1, (len(part) + chars_per_line - 1) // chars_per_line)
            lines = max(lines, line_count)

        item = ws.cell(r, 1).value
        if item is not None and str(item).strip():
            item_width = float(ws.column_dimensions["A"].width or 42)
            chars_per_line = max(int(item_width), 8)
            item_lines = max(
                1, (len(str(item)) + chars_per_line - 1) // chars_per_line
            )
            lines = max(lines, item_lines)

        # ~15 pt per line + padding so wrapped text is not clipped
        ws.row_dimensions[r].height = max(18.0, 15.0 * lines + 6.0)


def export_nre_xlsx(
    nre_df: pd.DataFrame,
    meta: dict,
    template_path: Path | None = None,
) -> bytes:
    """
    Fill the account NRE template (colors + Sub Total formulas) and return bytes.

    Only phase columns listed in meta["phases"] are kept (e.g. BU / CT / NOT).
    Layout: Test Item | Lab Location | Lab Rate | <phases> | Sub Total | Comments
    """
    account = str(meta.get("account") or "")
    phase_cols = _nre_phase_columns(list(meta.get("phases") or []))
    if template_path is not None:
        template = template_path
    elif account:
        template = nre_template_path(account)
    else:
        template = None
    pivoted = pivot_nre_for_template(nre_df, phases=phase_cols)
    buf = BytesIO()

    n_phases = len(phase_cols)
    subtotal_col = 3 + n_phases + 1
    comments_col = subtotal_col + 1
    ncols = comments_col
    subtotal_formula = _nre_subtotal_formula(n_phases)
    phase_col_by_name = {name: 4 + i for i, name in enumerate(phase_cols)}

    if template is not None and template.exists():
        wb = load_workbook(template)
        ws = wb.active
        # Drop unselected phase columns from the full template (right → left)
        full_phase_cols = {
            name: 4 + i for i, name in enumerate(_NRE_PHASE_ORDER)
        }
        for col_idx in sorted(
            (
                full_phase_cols[p]
                for p in _NRE_PHASE_ORDER
                if p not in phase_col_by_name
            ),
            reverse=True,
        ):
            ws.delete_cols(col_idx)

        data_start = 2
        style_src = data_start
        prefilled_last = max(ws.max_row, data_start)
        records = pivoted.to_dict(orient="records")

        for i, rec in enumerate(records):
            r = data_start + i
            if r > prefilled_last:
                _copy_nre_template_row(
                    ws,
                    style_src,
                    r,
                    ncols=ncols,
                    subtotal_col=subtotal_col,
                    subtotal_formula=subtotal_formula,
                )
            else:
                ws.cell(r, subtotal_col).value = subtotal_formula.format(row=r)
            ws.cell(r, 1).value = rec.get("Test Item")
            ws.cell(r, 2).value = rec.get("Lab Location") or None
            rate = rec.get("Lab Rate")
            ws.cell(r, 3).value = (
                None
                if rate is None or (isinstance(rate, float) and pd.isna(rate))
                else float(rate)
            )
            for phase_name, col_idx in phase_col_by_name.items():
                ws.cell(r, col_idx).value = _excel_num(rec.get(phase_name))
            comments = rec.get("Comments")
            comment_cell = ws.cell(r, comments_col)
            if (
                comments is None
                or (isinstance(comments, float) and pd.isna(comments))
            ):
                comment_cell.value = None
            else:
                comment_cell.value = str(comments)
                comment_cell.alignment = Alignment(
                    wrap_text=True, vertical="top"
                )
            # Wrap long Test Item / Location too
            for c in (1, 2):
                cell = ws.cell(r, c)
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        last_data = data_start + len(records) - 1
        # Clear unused pre-styled rows (keep fills; restore Sub Total formulas)
        for r in range(max(last_data + 1, data_start), prefilled_last + 1):
            for c in (1, 2, 3, *phase_col_by_name.values(), comments_col):
                ws.cell(r, c).value = None
            ws.cell(r, subtotal_col).value = subtotal_formula.format(row=r)

        if records:
            _nre_fit_sheet_layout(
                ws,
                data_start=data_start,
                last_data=last_data,
                ncols=ncols,
                comments_col=comments_col,
            )

        wb.save(buf)
    else:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            pivoted.to_excel(writer, sheet_name="NRE", index=False)
            ws = writer.sheets["NRE"]
            if not pivoted.empty:
                _nre_fit_sheet_layout(
                    ws,
                    data_start=2,
                    last_data=1 + len(pivoted),
                    ncols=len(pivoted.columns),
                    comments_col=len(pivoted.columns),
                )

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
