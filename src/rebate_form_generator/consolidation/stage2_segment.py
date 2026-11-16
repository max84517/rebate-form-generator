"""Stage 2 — Consolidate all supplier files for a single segment.

Reads the per-supplier workbooks produced by Stage 1, stacks them vertically
per FY sheet and writes
``<output_base>/tmp/segment/Consolidated_<segment>_<YYYYMMDD>.xlsx``.

Columns are matched **by name**, never by position: supplier workbooks are
hand-maintained and their column counts and order differ (50 … 68 columns in
practice).  Matching by position silently files one supplier's numbers under
another supplier's column headings.

Every anomaly is reported against the originating Excel file so the user knows
which workbook to go and fix.
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from openpyxl.workbook import Workbook

from . import issues
from .headers import (
    REBATE_HEADER_RE,
    SOURCE_FILE_COLUMN,
    expected_rebate_headers,
    find_header_row,
    header_map,
    is_error_value,
    is_header_row,
    uses_nonstandard_month,
)
from .stage1_ingest import load_provenance

FY_SHEET_RE = re.compile(r"^FY\d{2}$", re.IGNORECASE)


def _is_all_none(row: tuple | list) -> bool:
    return all(v is None for v in row)


def consolidate_segment(
    segment_name: str,
    raw_dir: Path,
    output_base: Path,
    log: Callable[[str, str], None],
    file_glob: str = "*.xlsx",
    sheet_suffix: str | None = None,
    only_fy: str | None = None,
) -> Path | None:
    """Consolidate all supplier xlsx files for *segment_name* into one workbook.

    *raw_dir* is the folder containing the per-supplier .xlsx files.
    *file_glob* filters which files to include (default ``"*.xlsx"``).
    *sheet_suffix* selects sheets by suffix, e.g. ``"bNB"`` matches ``"FY26 bNB"``;
    when ``None`` the plain ``FY##`` sheet names are used (DT / Peripheral).
    Returns the path of the saved file, or ``None`` on failure.
    """
    if not raw_dir.exists():
        log(f"  [{segment_name}] raw dir not found: {raw_dir}", "ERROR")
        return None

    xlsx_files = sorted(f for f in raw_dir.glob(file_glob) if not f.name.startswith("~$"))
    if not xlsx_files:
        log(f"  [{segment_name}] no .xlsx files in {raw_dir}", "WARNING")
        return None

    provenance = load_provenance(raw_dir)

    def origin(path: Path) -> tuple[str, str, str]:
        """(display name, original full path, supplier) for *path*."""
        meta = provenance.get(path.name, {})
        return path.name, meta.get("path", str(path)), meta.get("supplier", "")

    # -----------------------------------------------------------------------
    # Discover which FY sheets exist across all supplier files
    # -----------------------------------------------------------------------
    if sheet_suffix:
        _suffix_re = re.compile(
            r"^FY(\d{2})\s+" + re.escape(sheet_suffix) + r"$", re.IGNORECASE
        )
    all_fy: set[str] = set()
    for xlsx_path in xlsx_files:
        wb = load_workbook(xlsx_path, data_only=True, read_only=True)
        for ws in wb.worksheets:
            t = ws.title.strip()
            if sheet_suffix:
                m = _suffix_re.match(t)
                if m:
                    all_fy.add(f"FY{m.group(1).upper()}")
            elif FY_SHEET_RE.match(t.upper()):
                all_fy.add(t.upper())
        wb.close()

    if only_fy:
        # Never touch a year the user did not pick, so no warning can be
        # raised about one either.
        all_fy = {f for f in all_fy if f.upper() == only_fy.upper()}

    if not all_fy:
        log(f"  [{segment_name}] no {only_fy or 'FY'} sheets found", "WARNING")
        return None

    out_wb: Workbook = Workbook()
    out_wb.remove(out_wb.active)  # type: ignore[arg-type]

    # -----------------------------------------------------------------------
    # Per-FY consolidation
    # -----------------------------------------------------------------------
    for fy in sorted(all_fy):
        out_ws = out_wb.create_sheet(fy)
        target = (f"{fy} {sheet_suffix}".upper() if sheet_suffix else fy.upper())
        sheet_label = f"{fy} {sheet_suffix}" if sheet_suffix else fy

        # ── Pass A: read every file's header, build the column union ──────
        per_file: list[dict] = []
        column_order: list[str] = []          # normalised names, in output order
        display_name: dict[str, str] = {}     # normalised → header text to write

        for xlsx_path in xlsx_files:
            wb = load_workbook(xlsx_path, data_only=True, read_only=True)
            ws = next((s for s in wb.worksheets if s.title.strip().upper() == target), None)
            if ws is None:
                wb.close()
                continue

            rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
            wb.close()

            fname, fpath, supplier = origin(xlsx_path)
            hidx = find_header_row(rows)
            if hidx is None:
                log(f"  [{segment_name}] {sheet_label}: no header row in {fname} — skipped",
                    "ERROR")
                issues.add(
                    issues.HEADER_MISSING,
                    f"Sheet '{sheet_label}' has no recognisable header row, so none of "
                    f"its rows could be used.",
                    file=fname, path=fpath, sheet=sheet_label, supplier=supplier,
                )
                continue

            # "Rebate June 2027" and "Rebate Jun 2027" are the same column;
            # header_map folds them together. Worth a log line, not a warning.
            odd = [str(v).replace("\n", " ").strip()
                   for v in rows[hidx] if uses_nonstandard_month(v)]
            if odd:
                log(f"  [{segment_name}] {sheet_label}: {fname} writes "
                    f"{len(odd)} month name(s) in full ({', '.join(odd[:3])}"
                    f"{' …' if len(odd) > 3 else ''}) — read as the short form",
                    "INFO")

            hmap = header_map(rows[hidx])
            for name in hmap:
                if name not in display_name:
                    raw = rows[hidx][hmap[name]]
                    display_name[name] = (
                        str(raw).replace("\n", " ").replace("\r", " ").strip()
                    )
                    column_order.append(name)

            per_file.append({
                "path": xlsx_path, "rows": rows, "hidx": hidx, "hmap": hmap,
                "file": fname, "fpath": fpath, "supplier": supplier,
            })

        if not per_file:
            continue

        # ── Month sanity check: does each file carry this FY's months? ────
        expected = expected_rebate_headers(fy)
        expected_set = set(expected)
        for info in per_file:
            own_rebate = {n for n in info["hmap"] if REBATE_HEADER_RE.match(n)}
            if not own_rebate:
                continue
            matched = own_rebate & expected_set
            if not matched:
                got = sorted(own_rebate)[:3]
                log(f"  [{segment_name}] {sheet_label}: {info['file']} has NO {fy} month "
                    f"columns (it has e.g. {', '.join(got)})", "ERROR")
                issues.add(
                    issues.MONTH_MISMATCH,
                    f"Sheet '{sheet_label}' should hold {expected[0].title()} … "
                    f"{expected[-1].title()}, but its rebate columns are "
                    f"{', '.join(g.title() for g in got)} … — the sheet looks like it "
                    f"was copied from a previous FY without updating the month headers.",
                    file=info["file"], path=info["fpath"], sheet=sheet_label,
                    supplier=info["supplier"],
                )
            else:
                missing = [m for m in expected if m not in own_rebate]
                if missing:
                    log(f"  [{segment_name}] {sheet_label}: {info['file']} is missing "
                        f"{len(missing)} month column(s): "
                        f"{', '.join(m.title() for m in missing[:4])}", "WARNING")
                    issues.add(
                        issues.MISSING_COLUMN,
                        f"Sheet '{sheet_label}' is missing {len(missing)} rebate month "
                        f"column(s): {', '.join(m.title() for m in missing)}",
                        file=info["file"], path=info["fpath"], sheet=sheet_label,
                        supplier=info["supplier"],
                    )

        # ── Pass B: write header, then every file's data mapped by name ───
        # SOURCE_FILE_COLUMN travels with every row so that later stages can
        # name the exact workbook a bad value came from.
        out_ws.append([display_name[n] for n in column_order] + [SOURCE_FILE_COLUMN])
        known_names = set(column_order)
        rebate_names = [n for n in column_order if REBATE_HEADER_RE.match(n)]
        row_count = 0

        for info in per_file:
            rows, hidx, hmap = info["rows"], info["hidx"], info["hmap"]
            fname, fpath, supplier = info["file"], info["fpath"], info["supplier"]

            stray_headers = 0
            error_cells: list[str] = []

            for offset, row in enumerate(rows[hidx + 1:], start=hidx + 2):
                if _is_all_none(row):
                    continue
                if is_header_row(row, known_names):
                    stray_headers += 1
                    continue

                mapped = [
                    row[hmap[n]] if n in hmap and hmap[n] < len(row) else None
                    for n in column_order
                ]
                if _is_all_none(mapped):
                    continue

                for n in rebate_names:
                    idx = hmap.get(n)
                    if idx is None or idx >= len(row):
                        continue
                    if is_error_value(row[idx]):
                        error_cells.append(f"row {offset} / {n.title()}")

                out_ws.append(mapped + [fname])
                row_count += 1

            if stray_headers:
                # Recognised and skipped, so no data is lost: a log line, not
                # something to stop the user with.
                log(f"  [{segment_name}] {sheet_label}: skipped {stray_headers} repeated "
                    f"header row(s) inside {fname}", "INFO")
            if error_cells:
                log(f"  [{segment_name}] {sheet_label}: {len(error_cells)} Excel error "
                    f"cell(s) in {fname} ({', '.join(error_cells[:3])}…)", "WARNING")
                issues.add(
                    issues.ERROR_VALUE,
                    f"{len(error_cells)} rebate cell(s) in sheet '{sheet_label}' contain an "
                    f"Excel error (#VALUE!, #REF! …): {', '.join(error_cells[:6])}"
                    + (" …" if len(error_cells) > 6 else "")
                    + ". Those rows are dropped from the contract input.",
                    file=fname, path=fpath, sheet=sheet_label, supplier=supplier,
                    count=len(error_cells),
                )

        log(f"  [{segment_name}] {fy}: {row_count} data rows from {len(per_file)} supplier(s), "
            f"{len(column_order)} column(s)", "INFO")

    seg_dir = output_base / "tmp" / "segment"
    seg_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    out_path = seg_dir / f"Consolidated_{segment_name}_{date_str}.xlsx"
    out_wb.save(out_path)
    log(f"  [{segment_name}] saved → {out_path.name}", "INFO")
    return out_path
