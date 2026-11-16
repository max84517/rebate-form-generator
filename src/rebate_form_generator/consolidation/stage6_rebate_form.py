"""Stage 6 — Generate rebate form input.xlsx from rebate raw.xlsx.

Quarter rules
-------------
Q1 : Nov (prev year), Dec (prev year), Jan  (FY year)
Q2 : Feb, Mar, Apr  (FY year)
Q3 : May, Jun, Jul  (FY year)
Q4 : Aug, Sep, Oct  (FY year)

Example: FY26 Q1 → Nov 2025, Dec 2025, Jan 2026
"""
from __future__ import annotations

import re
from collections import defaultdict
from datetime import date
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook, Workbook

from . import issues
from .headers import (
    SOURCE_FILE_COLUMN,
    header_map,
    is_error_value,
    is_header_row,
    normalize_header,
    safe_filename_part,
)
from .stage1_ingest import load_provenance_dir, provenance_by_supplier

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

REBATE_COL_RE = re.compile(
    r"^Rebate\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})$",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Quarter helpers
# ---------------------------------------------------------------------------

def quarter_months(fy: int, q: int) -> list[tuple[int, int]]:
    """Return [(month, year), ...] for the 3 months of FY<fy> Q<q>.

    *fy* is the 2-digit year (e.g. 26 for FY26).
    """
    full_year = 2000 + fy
    if q == 1:
        return [(11, full_year - 1), (12, full_year - 1), (1, full_year)]
    elif q == 2:
        return [(2, full_year), (3, full_year), (4, full_year)]
    elif q == 3:
        return [(5, full_year), (6, full_year), (7, full_year)]
    else:  # q == 4
        return [(8, full_year), (9, full_year), (10, full_year)]


def current_fy_quarter() -> tuple[int, int]:
    """Return (fy_2digit, quarter) based on today's date."""
    today = date.today()
    m, y = today.month, today.year
    if m in (11, 12):
        return (y + 1) % 100, 1
    elif m == 1:
        return y % 100, 1
    elif m in (2, 3, 4):
        return y % 100, 2
    elif m in (5, 6, 7):
        return y % 100, 3
    else:  # m in (8, 9, 10)
        return y % 100, 4


def _col_header(month: int, year: int) -> str:
    return f"Rebate {MONTH_ABBR[month - 1]} {year}"


def _coerce_price(value) -> tuple[float | None, str]:
    """Return ``(price, problem)``.

    *problem* is an empty string when the cell is usable.  A blank cell is
    usable and means "no rebate", which is written as $0.00 downstream.
    """
    if value is None:
        return None, ""
    if isinstance(value, bool):
        return None, "boolean value"
    if isinstance(value, (int, float)):
        return round(float(value), 2), ""
    text = str(value).strip()
    if not text:
        return None, ""
    if is_error_value(value):
        return None, text
    try:
        return round(float(text.replace(",", "").replace("$", "")), 2), ""
    except ValueError:
        return None, f"not a number: {text[:30]!r}"


# ---------------------------------------------------------------------------
# Main function
# ---------------------------------------------------------------------------

def generate_rebate_form(
    rebate_raw_path: Path,
    fy: int,
    quarter: int,
    selected_columns: list[str],
    output_dir: Path,
    log: Callable[[str, str], None],
) -> list[Path]:
    """Read rebate raw.xlsx and write per-supplier contract input files.

    For each data row the function emits 1–3 output rows depending on whether
    the rebate price changes across the three months of the quarter.
    Rows are split by GTK Suppliers and saved as
    ``contract input - <Supplier>.xlsx``.

    Returns a list of paths to the written files (empty list on failure).
    """
    if not rebate_raw_path.exists():
        log(f"rebate raw.xlsx not found: {rebate_raw_path}", "ERROR")
        return []

    wb_in = load_workbook(rebate_raw_path, data_only=True)
    ws_in = wb_in.active

    months = quarter_months(fy, quarter)          # [(month, year), ...]
    target_headers = [_col_header(m, y) for m, y in months]
    start_dates = [date(y, m, 1) for m, y in months]

    log(f"  Quarter months: {target_headers}", "INFO")

    # Where each supplier's numbers originally came from, so warnings can
    # name the workbook the user has to fix.
    source_data_dir = rebate_raw_path.parent.parent / "source data"
    origin_map = provenance_by_supplier(source_data_dir)
    file_paths = {k: v.get("path", "") for k, v in load_provenance_dir(source_data_dir).items()}

    def origin_of(supplier: str, source_file: str = "") -> tuple[str, str]:
        """Prefer the exact workbook a row came from; fall back to the supplier."""
        if source_file:
            return source_file, file_paths.get(source_file, "")
        entries = origin_map.get(supplier) or []
        if not entries:
            return "", ""
        return ", ".join(e["file"] for e in entries), entries[0].get("path", "")

    # ── Locate columns in header row ─────────────────────────────────────
    header_iter = ws_in.iter_rows(min_row=1, max_row=1)
    header_cells = next(header_iter, None)
    if header_cells is None:
        log("rebate raw.xlsx has no header row", "ERROR")
        wb_in.close()
        return []
    all_headers: list = [cell.value for cell in header_cells]
    header_lower_map: dict[str, int] = header_map(tuple(all_headers))
    known_header_names = set(header_lower_map)

    target_col_indices: list[int | None] = []
    for h in target_headers:
        idx = header_lower_map.get(normalize_header(h))
        if idx is None:
            log(f"  Column '{h}' not found in rebate raw.xlsx — will treat as None", "WARNING")
            issues.add(
                issues.MISSING_COLUMN,
                f"No '{h}' column exists anywhere in the consolidated data, so that "
                f"month is blank for every supplier.",
                file="rebate raw.xlsx", sheet=str(ws_in.title),
            )
        target_col_indices.append(idx)

    # ── Build feature columns: selected optional + GTK Suppliers ─────────
    gtk_col_idx: int | None = header_lower_map.get(normalize_header("GTK Suppliers"))

    feature_col_indices: list[int] = []
    feature_headers: list[str] = []
    for col_name in selected_columns:
        idx = header_lower_map.get(normalize_header(col_name))
        if idx is not None:
            feature_col_indices.append(idx)
            feature_headers.append(all_headers[idx])
        else:
            log(f"  Optional column '{col_name}' not found — skipping", "WARNING")

    if gtk_col_idx is not None:
        feature_col_indices.append(gtk_col_idx)
        feature_headers.append(all_headers[gtk_col_idx])
    else:
        log("  'GTK Suppliers' column not found in data", "WARNING")

    # Index of the supplier value inside each output row tuple
    supplier_tuple_idx: int | None = (
        len(feature_col_indices) - 1 if gtk_col_idx is not None else None
    )

    out_headers = feature_headers + ["Per-Unit Rebate Amount $USD", "Rebate Period Start Date"]
    price_col = len(feature_headers) + 1  # 1-based

    # ── Collect all rows ──────────────────────────────────────────────────
    def _supplier_of(row: tuple) -> str:
        if gtk_col_idx is None or gtk_col_idx >= len(row):
            return "Unknown"
        return str(row[gtk_col_idx] or "Unknown").strip()

    src_file_idx = header_lower_map.get(normalize_header(SOURCE_FILE_COLUMN))

    all_rows: list[tuple] = []
    stray_headers = 0
    # (supplier, source workbook) → ["rebate raw row 5223: Rebate Dec 2026 (#VALUE!)", …]
    bad_rows: dict[tuple[str, str], list[str]] = {}
    has_data: dict[str, bool] = {}               # supplier → any usable price this quarter

    for rno, row in enumerate(ws_in.iter_rows(min_row=2, values_only=True), start=2):
        # A header row that slipped into the data (reported in detail at stage 2)
        if is_header_row(row, known_header_names):
            stray_headers += 1
            continue

        supplier = _supplier_of(row)
        has_data.setdefault(supplier, False)

        feature_vals = [row[i] if i < len(row) else None for i in feature_col_indices]

        prices = [
            (row[i] if i < len(row) else None) if i is not None else None
            for i in target_col_indices
        ]

        # Reject the whole row if any of this quarter's cells is unusable;
        # a wrong number in a contract is worse than a missing row.
        broken: list[str] = []
        clean: list[float | None] = []
        for price, header in zip(prices, target_headers):
            value, problem = _coerce_price(price)
            if problem:
                broken.append(f"{header} ({problem})")
            clean.append(value)
        if broken:
            src_file = ""
            if src_file_idx is not None and src_file_idx < len(row):
                src_file = str(row[src_file_idx] or "").strip()
            bad_rows.setdefault((supplier, src_file), []).append(
                f"rebate raw row {rno}: {', '.join(broken)}")
            continue

        if any(v is not None for v in clean):
            has_data[supplier] = True

        # Build price segments: start a new segment whenever the price changes
        segments: list[tuple] = []
        for price, seg_date in zip(clean, start_dates):
            if segments and segments[-1][0] == price:
                continue  # same price — extend current segment
            segments.append((price, seg_date))

        for price, seg_date in segments:
            rounded_price = price if price is not None else 0.0
            all_rows.append(tuple(feature_vals + [rounded_price, seg_date]))

    if stray_headers:
        log(f"  Skipped {stray_headers} repeated header row(s) in rebate raw.xlsx", "WARNING")

    for (supplier, src_file), entries in sorted(bad_rows.items()):
        fname, fpath = origin_of(supplier, src_file)
        log(f"  [{supplier}] dropped {len(entries)} row(s) with unusable rebate values "
            f"({entries[0]}{'…' if len(entries) > 1 else ''})", "WARNING")
        issues.add(
            issues.ERROR_VALUE,
            f"{len(entries)} row(s) were dropped from this quarter because a rebate cell "
            f"could not be read — {'; '.join(entries[:4])}"
            + (" …" if len(entries) > 4 else "")
            + ". Fix the formula in the supplier workbook and re-run.",
            file=fname or "rebate raw.xlsx", path=fpath,
            supplier=supplier, count=len(entries),
        )

    # ── Deduplicate while preserving order ───────────────────────────────
    seen: set[tuple] = set()
    unique_rows: list[tuple] = []
    for row_tuple in all_rows:
        if row_tuple not in seen:
            seen.add(row_tuple)
            unique_rows.append(row_tuple)

    dropped = len(all_rows) - len(unique_rows)
    if dropped:
        log(f"  Dropped {dropped} duplicate row(s)", "INFO")

    # ── Group by supplier ─────────────────────────────────────────────────
    supplier_rows: dict[str, list[tuple]] = defaultdict(list)
    for row_tuple in unique_rows:
        if supplier_tuple_idx is not None:
            supplier = str(row_tuple[supplier_tuple_idx] or "Unknown").strip()
        else:
            supplier = "Unknown"
        supplier_rows[supplier].append(row_tuple)

    # ── Write one file per supplier ───────────────────────────────────────
    output_dir.mkdir(parents=True, exist_ok=True)
    out_paths: list[Path] = []
    for supplier, rows in sorted(supplier_rows.items()):
        fname_src, fpath_src = origin_of(supplier)

        # A leftover header row, or a blank GTK cell, is not a supplier.
        if normalize_header(supplier) in ("gtk suppliers", "", "unknown"):
            log(f"  Ignoring non-supplier group {supplier!r} ({len(rows)} row(s))", "WARNING")
            continue

        # No usable price for any of this quarter's months → do not invent a
        # contract for them; say which workbook has to be fixed.
        if not has_data.get(supplier, False):
            log(f"  [{supplier}] no rebate value for {', '.join(target_headers)} "
                f"— no contract input written", "ERROR")
            issues.add(
                issues.SUPPLIER_SKIPPED,
                f"No usable rebate value for {target_headers[0]} / {target_headers[1]} / "
                f"{target_headers[2]}, so no contract input file was produced. Check that "
                f"this supplier's workbook has month columns for this FY.",
                file=fname_src, path=fpath_src, supplier=supplier,
            )
            continue

        safe = safe_filename_part(supplier)
        if safe != supplier:
            log(f"  [{supplier}] supplier name contains characters not allowed in a "
                f"filename — saving as '{safe}'", "WARNING")

        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.title = "Input"
        ws_out.append(out_headers)
        for i, row_tuple in enumerate(rows, start=2):
            ws_out.append(list(row_tuple))
            ws_out.cell(row=i, column=price_col).number_format = '"$"#,##0.00'
        out_path = output_dir / f"contract input - {safe}.xlsx"
        wb_out.save(out_path)
        log(f"  Saved {len(rows)} rows → {out_path.name}", "INFO")
        out_paths.append(out_path)

    return out_paths
