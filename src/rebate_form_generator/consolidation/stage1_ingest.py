"""Stage 1 — Ingest raw supplier Excel files.

Reads the latest .xlsx from each supplier sub-folder, applies the
GTK Suppliers fix, and writes split/normalised workbooks to
the caller-supplied *raw_dir*.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from openpyxl.workbook import Workbook

from . import issues
from .headers import HEADER_SCAN_ROWS, find_header_row, normalize_header

PROVENANCE_FILE = "_provenance.json"

# ---------------------------------------------------------------------------
# Regexes
# ---------------------------------------------------------------------------
NB_SHEET_RE = re.compile(r"^FY(\d{2})\s*([cb])\s*NB$", re.IGNORECASE)
FY_SHEET_RE = re.compile(r"^FY\d{2}$", re.IGNORECASE)


def _make_supplier_folder_re(segment: str) -> re.Pattern:
    """Return a regex matching 'Master price table_<segment>_<Supplier>'."""
    return re.compile(
        rf"^Master\s+price\s+table_{re.escape(segment)}_.+$", re.IGNORECASE
    )


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_latest_xlsx(folder: Path) -> Path | None:
    """Return the most recently modified .xlsx that is not an Office temp file."""
    files = [f for f in folder.glob("*.xlsx") if not f.name.startswith("~$")]
    return max(files, key=lambda f: f.stat().st_mtime) if files else None


def _open_source(path: Path, supplier: str, segment: str,
                 log: Callable[[str, str], None]):
    """Open a supplier workbook, reporting instead of crashing on failure.

    A workbook left open in Excel cannot be read. One supplier's open file
    must not take the whole run down with it, so the rest still processes and
    the user is told exactly which file to close.
    """
    try:
        return load_workbook(path, data_only=True)
    except PermissionError:
        log(f"  [{segment}] {supplier}: '{path.name}' is open in another program "
            f"(most likely Excel) — skipped", "ERROR")
        issues.add(
            issues.FILE_LOCKED,
            f"This workbook is open in another program (most likely Excel), so "
            f"{supplier}'s {segment} data was left out. Close it and run again.",
            file=path.name, path=str(path), supplier=supplier,
        )
    except Exception as exc:
        log(f"  [{segment}] {supplier}: cannot read '{path.name}': {exc}", "ERROR")
        issues.add(
            issues.FILE_LOCKED,
            f"This workbook could not be read ({exc}), so {supplier}'s "
            f"{segment} data was left out.",
            file=path.name, path=str(path), supplier=supplier,
        )
    return None


def _parse_supplier_name(folder_name: str) -> str:
    """Extract supplier from folder name, normalised to title-case.

    ``"Master price table_NB_Chicony"``  →  ``"Chicony"``
    ``"Master price table_DT_CHICONY"``  →  ``"Chicony"``
    ``"Master price table_NB_LiteOn"``   →  ``"Liteon"``
    """
    # Split on first underscore: "Master price table" + "NB_Chicony"
    parts = folder_name.split("_", 1)
    if len(parts) < 2:
        return folder_name.title()
    rest = parts[1]          # "NB_Chicony"
    # Split on first underscore: "NB" + "Chicony"
    parts2 = rest.split("_", 1)
    raw = parts2[1] if len(parts2) > 1 else rest
    return raw.title()


def _find_segment_root(parent: Path, keyword: str) -> Path | None:
    """Find a sub-folder whose name contains *keyword* (case-insensitive)."""
    keyword_lower = keyword.lower()
    for sub in parent.iterdir():
        if sub.is_dir() and keyword_lower in sub.name.lower():
            return sub
    return None


def _record_provenance(raw_dir: Path, ingested_name: str, source: Path, supplier: str) -> None:
    """Remember which original file an ingested workbook came from.

    Later stages only see the copy under ``source data/``; this map lets their
    warnings name the real file the user has to go and fix.
    """
    store = raw_dir.parent / PROVENANCE_FILE
    data: dict = {}
    if store.exists():
        try:
            data = json.loads(store.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            data = {}
    data[ingested_name] = {"path": str(source), "supplier": supplier}
    try:
        store.parent.mkdir(parents=True, exist_ok=True)
        store.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def load_provenance_dir(source_data_dir: Path) -> dict[str, dict]:
    """Read the map written by :func:`_record_provenance`."""
    store = source_data_dir / PROVENANCE_FILE
    if not store.exists():
        return {}
    try:
        return json.loads(store.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return {}


def load_provenance(raw_dir: Path) -> dict[str, dict]:
    """As :func:`load_provenance_dir`, for a per-segment raw directory."""
    return load_provenance_dir(raw_dir.parent)


def provenance_by_supplier(source_data_dir: Path) -> dict[str, list[dict]]:
    """``{supplier: [{"file": name, "path": original path}, ...]}``."""
    out: dict[str, list[dict]] = {}
    for name, meta in load_provenance_dir(source_data_dir).items():
        supplier = meta.get("supplier", "")
        if supplier:
            out.setdefault(supplier, []).append({"file": name, "path": meta.get("path", "")})
    return out


def _fix_gtk_suppliers(
    wb: Workbook,
    supplier_name: str,
    *,
    file_name: str,
    file_path: str,
    log: Callable[[str, str], None],
    fy: str | None = None,
) -> None:
    """Overwrite the 'GTK Suppliers' column with *supplier_name* for data rows.

    The header is **not** assumed to be on row 2: suppliers sometimes insert a
    row above it, which previously made this fix silently do nothing.
    """
    for ws in wb.worksheets:
        title = ws.title.strip()
        # Only FY data sheets are worth reporting on; notes/lookup tabs are not.
        nb_match = NB_SHEET_RE.match(title)
        is_data_sheet = bool(nb_match or FY_SHEET_RE.match(title))
        if is_data_sheet and fy:
            sheet_fy = f"FY{nb_match.group(1)}" if nb_match else title.upper()
            if sheet_fy.upper() != fy.upper():
                # A different year than the one being built: not our business,
                # so neither fix it nor complain about it.
                continue

        rows = [tuple(r) for r in ws.iter_rows(min_row=1, max_row=HEADER_SCAN_ROWS,
                                               values_only=True)]
        hidx = find_header_row(rows)
        if hidx is None:
            if is_data_sheet:
                log(f"    [{file_name}] sheet '{ws.title}': header row not found "
                    f"— GTK Suppliers not normalised", "WARNING")
                issues.add(
                    issues.HEADER_MISSING,
                    f"No row starting with 'Segment' was found in the first "
                    f"{HEADER_SCAN_ROWS} rows, so the header could not be located "
                    f"and the GTK Suppliers column was not normalised.",
                    file=file_name, path=file_path, sheet=ws.title, supplier=supplier_name,
                )
            continue

        header_excel_row = hidx + 1          # openpyxl rows are 1-based
        if header_excel_row != 2 and is_data_sheet:
            # The header is located per sheet, so a row inserted above it
            # costs nothing: the data is read correctly either way. Worth a
            # log line, not a warning the user has to dismiss.
            log(f"    [{file_name}] sheet '{ws.title}': header found on row "
                f"{header_excel_row} (not the usual row 2) - read from there",
                "INFO")

        for col_idx, value in enumerate(rows[hidx], start=1):
            if normalize_header(value) == "gtk suppliers":
                for row_idx in range(header_excel_row + 1, ws.max_row + 1):
                    ws.cell(row=row_idx, column=col_idx).value = supplier_name
                break


def _copy_ws_to_wb(src_ws, target_wb: Workbook, sheet_name: str) -> None:
    """Copy all rows from *src_ws* into a new sheet in *target_wb*."""
    new_ws = target_wb.create_sheet(sheet_name)
    for row in src_ws.iter_rows(values_only=True):
        new_ws.append(list(row))


# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------

def scan_nb(
    nb_kb_path: Path,
    log: Callable[[str, str], None],
) -> tuple[dict[str, dict[str, list[str]]], set[str]]:
    """Which FY years each NB supplier has, from sheet names alone.

    Reads no cell data, so it is fast enough to run before asking which FY to
    build. A year counts only when the supplier has both bNB and cNB, which
    is the same rule :func:`ingest_nb` applies.
    """
    seg_root = _find_segment_root(nb_kb_path, "Master price table_NB")
    if seg_root is None:
        log(f"Cannot find 'Master price table_NB' sub-folder in {nb_kb_path}", "ERROR")
        return {"bNB": {}, "cNB": {}}, set()

    folder_re = _make_supplier_folder_re("NB")
    cov_b: dict[str, list[str]] = {}
    cov_c: dict[str, list[str]] = {}
    all_fy: set[str] = set()

    for folder in sorted(seg_root.iterdir()):
        if not folder.is_dir() or not folder_re.match(folder.name):
            continue
        latest = _get_latest_xlsx(folder)
        if latest is None:
            continue
        supplier = _parse_supplier_name(folder.name)
        try:
            wb = load_workbook(latest, data_only=True, read_only=True)
            titles = list(wb.sheetnames)
            wb.close()
        except Exception as exc:
            log(f"  NB: cannot read '{latest.name}': {exc}", "WARNING")
            continue

        fy_types: dict[str, set[str]] = {}
        for title in titles:
            m = NB_SHEET_RE.match(title.strip())
            if m:
                fy_types.setdefault(f"FY{m.group(1)}", set()).add(m.group(2).lower())
        both = sorted(fy for fy, kinds in fy_types.items() if {"b", "c"} <= kinds)
        if both:
            cov_b[supplier] = both
            cov_c[supplier] = both
            all_fy.update(both)

    return {"bNB": cov_b, "cNB": cov_c}, all_fy


def scan_segment(
    segment_path: Path,
    segment_name: str,
    log: Callable[[str, str], None],
) -> tuple[dict[str, list[str]], set[str]]:
    """As :func:`scan_nb`, for DT and Peripheral."""
    seg_root = _find_segment_root(segment_path, f"Master price table_{segment_name}")
    if seg_root is None:
        log(f"Cannot find 'Master price table_{segment_name}' sub-folder "
            f"in {segment_path}", "ERROR")
        return {}, set()

    folder_re = _make_supplier_folder_re(segment_name)
    coverage: dict[str, list[str]] = {}
    all_fy: set[str] = set()

    for folder in sorted(seg_root.iterdir()):
        if not folder.is_dir() or not folder_re.match(folder.name):
            continue
        latest = _get_latest_xlsx(folder)
        if latest is None:
            continue
        supplier = _parse_supplier_name(folder.name)
        try:
            wb = load_workbook(latest, data_only=True, read_only=True)
            titles = list(wb.sheetnames)
            wb.close()
        except Exception as exc:
            log(f"  {segment_name}: cannot read '{latest.name}': {exc}", "WARNING")
            continue

        found = sorted({t.strip().upper() for t in titles
                        if FY_SHEET_RE.match(t.strip())})
        if found:
            coverage[supplier] = found
            all_fy.update(found)

    return coverage, all_fy


def ingest_nb(
    nb_kb_path: Path,
    raw_nb_dir: Path,
    log: Callable[[str, str], None],
    fy: str | None = None,
) -> dict[str, list[str]]:
    """Stage 1 for NB KB — split each supplier into bNB and cNB workbooks.

    *raw_nb_dir* is ``source data/NB``; outputs go to
    ``raw_nb_dir/bNB`` and ``raw_nb_dir/cNB``.

    Returns a dict ``{"bNB": [supplier, ...], "cNB": [supplier, ...]}``.
    """
    seg_root = _find_segment_root(nb_kb_path, "Master price table_NB")
    if seg_root is None:
        log(f"Cannot find 'Master price table_NB' sub-folder in {nb_kb_path}", "ERROR")
        return {"bNB": [], "cNB": []}

    # -----------------------------------------------------------------------
    # Collect all supplier workbooks (apply GTK fix in-memory)
    # -----------------------------------------------------------------------
    nb_folder_re = _make_supplier_folder_re("NB")
    supplier_data: dict[str, dict] = {}
    for supplier_folder in sorted(seg_root.iterdir()):
        if not supplier_folder.is_dir():
            continue
        if not nb_folder_re.match(supplier_folder.name):
            log(
                f"  NB: skipping '{supplier_folder.name}' "
                f"(must match 'Master price table_NB_<Supplier>')",
                "WARNING",
            )
            continue
        latest = _get_latest_xlsx(supplier_folder)
        if latest is None:
            log(f"  NB: no .xlsx found in '{supplier_folder.name}', skipping", "WARNING")
            continue

        supplier_name = _parse_supplier_name(supplier_folder.name)
        log(f"  NB supplier: {supplier_name}  ←  {latest.name}", "INFO")

        wb = _open_source(latest, supplier_name, "NB", log)
        if wb is None:
            continue
        _fix_gtk_suppliers(wb, supplier_name, file_name=latest.name,
                           file_path=str(latest), log=log, fy=fy)
        supplier_data[supplier_name] = {"path": latest, "wb": wb}

    if not supplier_data:
        log("  NB: no supplier data found", "WARNING")
        return {"bNB": [], "cNB": []}

    # -----------------------------------------------------------------------
    # Pass 1 — find FY years that have BOTH cNB and bNB per supplier
    # (each supplier is evaluated independently; no cross-supplier intersection)
    # -----------------------------------------------------------------------
    valid_fy_per_supplier: dict[str, set[str]] = {}
    for supplier_name, data in supplier_data.items():
        fy_types: dict[str, set[str]] = {}
        for ws in data["wb"].worksheets:
            m = NB_SHEET_RE.match(ws.title.strip())
            if m:
                fy = f"FY{m.group(1)}"
                nb_type = m.group(2).lower()
                fy_types.setdefault(fy, set()).add(nb_type)
        # Only include FY years where THIS supplier has both 'b' and 'c'
        both = {y for y, types in fy_types.items() if "c" in types and "b" in types}
        if fy:
            both = {y for y in both if y.upper() == fy.upper()}
        valid_fy_per_supplier[supplier_name] = both
        log(f"  NB {supplier_name} valid FY: {sorted(both)}", "INFO")

    # -----------------------------------------------------------------------
    # Pass 2 — combine bNB + cNB into one workbook per supplier
    # -----------------------------------------------------------------------
    raw_nb_dir.mkdir(parents=True, exist_ok=True)

    coverage_bnb: dict[str, list[str]] = {}
    coverage_cnb: dict[str, list[str]] = {}
    for supplier_name, data in supplier_data.items():
        supplier_valid_fy = valid_fy_per_supplier.get(supplier_name, set())
        wb_out: Workbook = Workbook()
        wb_out.remove(wb_out.active)  # type: ignore[arg-type]

        fy_bnb: list[str] = []
        fy_cnb: list[str] = []

        for ws in data["wb"].worksheets:
            m = NB_SHEET_RE.match(ws.title.strip())
            if not m:
                continue
            fy = f"FY{m.group(1)}"
            if fy not in supplier_valid_fy:
                continue
            nb_type = m.group(2).lower()
            sheet_name = f"{fy} bNB" if nb_type == "b" else f"{fy} cNB"
            _copy_ws_to_wb(ws, wb_out, sheet_name)
            if nb_type == "b":
                fy_bnb.append(fy)
            else:
                fy_cnb.append(fy)

        if wb_out.worksheets:
            fname = f"{data['path'].stem}.xlsx"
            wb_out.save(raw_nb_dir / fname)
            _record_provenance(raw_nb_dir, fname, data["path"], supplier_name)
            log(
                f"    → NB/{fname}  "
                f"({len(fy_bnb)} bNB + {len(fy_cnb)} cNB sheets)",
                "INFO",
            )
            if fy_bnb:
                coverage_bnb[supplier_name] = fy_bnb
            if fy_cnb:
                coverage_cnb[supplier_name] = fy_cnb

    return {"bNB": coverage_bnb, "cNB": coverage_cnb}


def ingest_segment(
    segment_path: Path,
    segment_name: str,
    raw_dir: Path,
    log: Callable[[str, str], None],
    fy: str | None = None,
) -> list[str]:
    """Stage 1 for DT or Peripheral.

    *raw_dir* is the destination folder (e.g. ``source data/DT``).
    *segment_name* should be ``"DT"`` or ``"Peripheral"``.
    Returns ``{supplier_name: [FY sheets]}`` for processed suppliers.
    """
    seg_root = _find_segment_root(
        segment_path, f"Master price table_{segment_name}"
    )
    if seg_root is None:
        log(
            f"Cannot find 'Master price table_{segment_name}' sub-folder "
            f"in {segment_path}",
            "ERROR",
        )
        return {}

    raw_dir.mkdir(parents=True, exist_ok=True)

    supplier_folder_re = _make_supplier_folder_re(segment_name)
    coverage: dict[str, list[str]] = {}
    for supplier_folder in sorted(seg_root.iterdir()):
        if not supplier_folder.is_dir():
            continue
        if not supplier_folder_re.match(supplier_folder.name):
            log(
                f"  {segment_name}: skipping '{supplier_folder.name}' "
                f"(must match 'Master price table_{segment_name}_<Supplier>')",
                "WARNING",
            )
            continue
        latest = _get_latest_xlsx(supplier_folder)
        if latest is None:
            log(f"  {segment_name}: no .xlsx in '{supplier_folder.name}', skipping", "WARNING")
            continue

        supplier_name = _parse_supplier_name(supplier_folder.name)
        log(f"  {segment_name} supplier: {supplier_name}  ←  {latest.name}", "INFO")

        wb = _open_source(latest, supplier_name, segment_name, log)
        if wb is None:
            continue
        _fix_gtk_suppliers(wb, supplier_name, file_name=latest.name,
                           file_path=str(latest), log=log, fy=fy)

        # Build output workbook with only FY sheets, normalised names
        wb_out: Workbook = Workbook()
        wb_out.remove(wb_out.active)  # type: ignore[arg-type]

        for ws in wb.worksheets:
            stripped = ws.title.strip()
            if not FY_SHEET_RE.match(stripped):
                continue
            if fy and stripped.upper() != fy.upper():
                continue          # only the FY that was asked for
            _copy_ws_to_wb(ws, wb_out, stripped.upper())

        if wb_out.worksheets:
            wb_out.save(raw_dir / latest.name)
            _record_provenance(raw_dir, latest.name, latest, supplier_name)
            log(f"    → {segment_name}/{latest.name}  ({len(wb_out.worksheets)} sheets)", "INFO")
            coverage[supplier_name] = [ws.title for ws in wb_out.worksheets]

    return coverage
