"""Public pipeline API.

Usage::

    from rebate_form_generator.consolidation.pipeline import (
        get_available_fy_sheets,
        run_full_pipeline,
    )

    source_paths = {
        "nb_kb":      r"C:\\Data\\NB KB",
        "dt_kb":      r"C:\\Data\\DT KB",
        "peripheral": r"C:\\Data\\Peripheral",
    }
    output_path = Path("data/output")

    def log(msg, level):
        print(f"[{level}] {msg}")

    # Run Stage 1-4, get available FY sheets
    fy_list = get_available_fy_sheets(source_paths, output_path, log)

    # Run Stage 5 (reuses cached Stage 1-4 results)
    out = run_full_pipeline(source_paths, "FY25", output_path, log)
"""
from __future__ import annotations

import os
import shutil
import stat
import tempfile
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook

from . import issues
from .stage1_ingest import ingest_nb, ingest_segment, scan_nb, scan_segment
from .stage2_segment import consolidate_segment
from .stage3_all import consolidate_all
from .stage4_rebate import build_rebate_only
from .stage5_template import write_pricing_template
from .stage6_rebate_form import generate_rebate_form
from .stage7_report import generate_report

# Module-level cache: stores results from the last successful Stage 1-4 run
_cache: dict = {
    "rebate_path": None,
    "output_path": None,
    "pricing_data_dir": None,
    "coverage": None,
}


def scan_fy_coverage(
    source_paths: dict,
    log: Callable[[str, str], None],
) -> tuple[list[str], dict]:
    """List the FY sheets each supplier has, without ingesting anything.

    Only sheet *names* are read, so this is quick enough to run before the FY
    dialog. Nothing is copied or written: the heavy work happens afterwards,
    against the one FY that was chosen, so a run never processes — or
    complains about — a year you did not ask for.
    """
    issues.reset()

    nb_cov, nb_fy = scan_nb(Path(source_paths["nb_kb"]), log)
    dt_cov, dt_fy = scan_segment(Path(source_paths["dt_kb"]), "DT", log)
    peri_cov, peri_fy = scan_segment(Path(source_paths["peripheral"]), "Peripheral", log)

    coverage: dict[str, dict[str, list[str]]] = {
        "bNB":        nb_cov.get("bNB", {}),
        "cNB":        nb_cov.get("cNB", {}),
        "DT":         dt_cov,
        "Peripheral": peri_cov,
    }

    def _fy_key(name: str) -> int:
        digits = "".join(c for c in name if c.isdigit())
        return int(digits) if digits else 0

    fy_sheets = sorted(nb_fy | dt_fy | peri_fy, key=_fy_key)
    log(f"Available FY: {fy_sheets}", "INFO")
    return fy_sheets, coverage


def build_rebate_raw(
    source_paths: dict,
    fy_sheet: str,
    output_path: Path,
    log: Callable[[str, str], None],
) -> list[str]:
    """Run Stages 1–4 for a single FY and return the sheet names produced."""
    issues.reset()
    fy = fy_sheet.strip().upper()

    # Directories derived from output_path
    #   data/source data  → ingested supplier workbooks
    #   data/pricing data → final output (rebate raw.xlsx)
    #   system temp       → stage 2-4 intermediates (never written to data/)
    source_data_dir = output_path.parent / "source data"
    pricing_data_dir = output_path.parent / "rebate raw"

    def _rmtree(path: Path) -> None:
        def _on_err(func, p, exc_info):
            try:
                os.chmod(p, stat.S_IWRITE)
                func(p)
            except Exception:
                pass
        shutil.rmtree(path, onexc=_on_err)

    # Clear source data sub-dirs (NB, DT, Peripheral) before fresh ingest
    for seg_dir in (
        source_data_dir / "NB",
        source_data_dir / "DT",
        source_data_dir / "Peripheral",
    ):
        if seg_dir.exists():
            log(f"  Clearing {seg_dir.name} source data …", "INFO")
            _rmtree(seg_dir)

    # Remove legacy output/tmp if it still exists from an older run
    legacy_tmp = output_path / "tmp"
    if legacy_tmp.exists():
        _rmtree(legacy_tmp)

    # Stage 2-4 intermediates go to the system temp directory
    tmp_dir = Path(tempfile.gettempdir()) / "rebate_form_generator"
    if tmp_dir.exists():
        _rmtree(tmp_dir)
    tmp_dir.mkdir(parents=True, exist_ok=True)

    nb_kb = Path(source_paths["nb_kb"])
    dt_kb = Path(source_paths["dt_kb"])
    peripheral = Path(source_paths["peripheral"])

    # ------------------------------------------------------------------
    # Stage 1 — Ingest
    # ------------------------------------------------------------------
    log(f"=== Stage 1: Ingest ({fy} only) ===", "INFO")
    log("Processing NB KB …", "INFO")
    nb_cov = ingest_nb(nb_kb, source_data_dir / "NB", log, fy=fy)

    log("Processing DT KB …", "INFO")
    dt_cov = ingest_segment(dt_kb, "DT", source_data_dir / "DT", log, fy=fy)

    log("Processing Peripheral …", "INFO")
    peri_cov = ingest_segment(peripheral, "Peripheral", source_data_dir / "Peripheral",
                              log, fy=fy)

    # Build combined coverage: segment → {supplier: [FY sheets]}
    coverage: dict[str, dict[str, list[str]]] = {
        "bNB":        nb_cov.get("bNB", {}),
        "cNB":        nb_cov.get("cNB", {}),
        "DT":         dt_cov,
        "Peripheral": peri_cov,
    }

    # ------------------------------------------------------------------
    # Stage 2 — Consolidate by segment
    # ------------------------------------------------------------------
    log("=== Stage 2: Consolidate by Segment ===", "INFO")
    nb_raw_dir = source_data_dir / "NB"
    segment_files: list[Path] = []
    for seg, raw_dir, suffix in [
        ("bNB",        nb_raw_dir,                     "bNB"),
        ("cNB",        nb_raw_dir,                     "cNB"),
        ("DT",         source_data_dir / "DT",         None),
        ("Peripheral", source_data_dir / "Peripheral", None),
    ]:
        log(f"Consolidating {seg} …", "INFO")
        path = consolidate_segment(seg, raw_dir, tmp_dir, log,
                                   sheet_suffix=suffix, only_fy=fy)
        if path is not None:
            segment_files.append(path)

    # ------------------------------------------------------------------
    # Stage 3 — Consolidate all segments
    # ------------------------------------------------------------------
    log("=== Stage 3: Consolidate All Segments ===", "INFO")
    all_path = consolidate_all(segment_files, tmp_dir, log)

    # ------------------------------------------------------------------
    # Stage 4 — Rebate-only (remove cost columns)
    # ------------------------------------------------------------------
    log("=== Stage 4: Rebate Only ===", "INFO")
    rebate_path = build_rebate_only(all_path, tmp_dir, log)

    # Cache results
    _cache["rebate_path"] = rebate_path
    _cache["output_path"] = output_path
    _cache["pricing_data_dir"] = pricing_data_dir
    _cache["coverage"] = coverage
    _cache["fy"] = fy

    if rebate_path is None:
        return []

    wb = load_workbook(rebate_path, data_only=True, read_only=True)
    fy_sheets = [ws.title for ws in wb.worksheets]
    wb.close()
    return fy_sheets


def run_full_pipeline(
    source_paths: dict,
    fy_sheet: str,
    output_path: Path,
    log: Callable[[str, str], None],
) -> Path | None:
    """Run Stages 1–5 for *fy_sheet* only.

    Returns the path to ``rebate raw.xlsx``, or ``None``.
    """
    fy_sheets = build_rebate_raw(source_paths, fy_sheet, output_path, log)
    rebate_path: Path | None = _cache.get("rebate_path")
    if rebate_path is None:
        log("Nothing to write: no data survived stages 1-4.", "ERROR")
        return None
    if fy_sheets and fy_sheet.strip().upper() not in {s.strip().upper() for s in fy_sheets}:
        log(f"'{fy_sheet}' produced no data. Sheets built: {fy_sheets}", "ERROR")
        return None

    pricing_data_dir: Path = _cache.get("pricing_data_dir") or (output_path.parent / "pricing data")

    log("=== Stage 5: Write Pricing Template ===", "INFO")
    return write_pricing_template(rebate_path, fy_sheet, pricing_data_dir, log)


def check_latest_files(
    source_paths: dict,
    log: Callable[[str, str], None],
) -> dict[str, dict[str, str]]:
    """Preview the latest .xlsx file chosen per supplier (no file copying)."""
    from .stage1_ingest import _find_segment_root, _get_latest_xlsx, _parse_supplier_name
    import re

    NB_ROOT_KW = "Master price table_NB"
    DT_ROOT_KW = "Master price table_DT"
    PERI_ROOT_KW = "Master price table_Peripheral"

    result: dict[str, dict[str, str]] = {}
    spec = [
        ("NB KB", Path(source_paths["nb_kb"]), NB_ROOT_KW),
        ("DT KB", Path(source_paths["dt_kb"]), DT_ROOT_KW),
        ("Peripheral", Path(source_paths["peripheral"]), PERI_ROOT_KW),
    ]
    for label, root, kw in spec:
        seg_root = _find_segment_root(root, kw)
        if seg_root is None:
            log(f"[{label}] segment root not found", "WARNING")
            continue
        result[label] = {}
        for supplier_folder in sorted(seg_root.iterdir()):
            if not supplier_folder.is_dir():
                continue
            latest = _get_latest_xlsx(supplier_folder)
            supplier = _parse_supplier_name(supplier_folder.name)
            result[label][supplier] = str(latest) if latest else "(no xlsx)"
            log(f"  [{label}] {supplier}: {latest.name if latest else 'missing'}", "INFO")
    return result


def check_missing_fy_sheets(
    source_paths: dict,
    fy: str,
    log: Callable[[str, str], None],
) -> list[str]:
    """Return supplier names that are missing the specified *fy* sheet."""
    from .stage1_ingest import _find_segment_root, _get_latest_xlsx, _parse_supplier_name
    import re

    FY_RE = re.compile(r"^FY\d{2}$", re.IGNORECASE)
    NB_SHEET_RE = re.compile(r"^FY(\d{2})\s*([cb])\s*NB$", re.IGNORECASE)
    missing: list[str] = []
    target = fy.strip().upper()

    spec = [
        (Path(source_paths["nb_kb"]), "Master price table_NB", True),
        (Path(source_paths["dt_kb"]), "Master price table_DT", False),
        (Path(source_paths["peripheral"]), "Master price table_Peripheral", False),
    ]
    for root, kw, is_nb in spec:
        seg_root = _find_segment_root(root, kw)
        if seg_root is None:
            continue
        for supplier_folder in sorted(seg_root.iterdir()):
            if not supplier_folder.is_dir():
                continue
            latest = _get_latest_xlsx(supplier_folder)
            if latest is None:
                continue
            supplier = _parse_supplier_name(supplier_folder.name)
            wb = load_workbook(latest, data_only=True, read_only=True)
            if is_nb:
                found = any(
                    f"FY{m.group(1)}" == target
                    for s in wb.worksheets
                    if (m := NB_SHEET_RE.match(s.title.strip()))
                )
            else:
                found = any(
                    s.title.strip().upper() == target for s in wb.worksheets
                )
            wb.close()
            if not found:
                missing.append(supplier)
                log(f"  MISSING {target}: {supplier}", "WARNING")
    return missing


def run_rebate_form_pipeline(
    output_path: Path,
    fy: int,
    quarter: int,
    selected_columns: list[str],
    log: Callable[[str, str], None],
) -> list[Path]:
    """Stage 6: generate per-supplier contract input files from ``rebate raw.xlsx``.

    *fy* is the 2-digit FY year (e.g. 26 for FY26).
    *quarter* is 1–4.
    *selected_columns* are the optional feature columns to keep.
    Returns a list of paths to the written files (empty list on failure).
    """
    issues.reset()
    rebate_raw_path = output_path.parent / "rebate raw" / "rebate raw.xlsx"
    output_dir = output_path.parent / "rebate form input"
    log(f"=== Stage 6: Generate Form Data (FY{fy:02d} Q{quarter}) ===", "INFO")
    return generate_rebate_form(rebate_raw_path, fy, quarter, selected_columns, output_dir, log)


def run_report_pipeline(
    output_path: Path,
    suppliers: list[str],
    form_numbers: dict[str, str],
    log: Callable[[str, str], None],
    fy: int | None = None,
    quarter: int | None = None,
    icertis_codes: dict[str, str] | None = None,
    destination: Path | None = None,
) -> tuple[list[Path], dict[str, list[str]]]:
    """Stage 7: generate per-supplier Word contracts.

    Returns ``(out_paths, blanks)`` where *blanks* maps supplier name to a
    list of column names that contain at least one blank cell.
    """
    issues.reset()
    base = output_path.parent
    supplier_info_dir = base / "supplier info"
    template_dir = base / "template"
    rebate_form_input_dir = base / "rebate form input"
    log(f"=== Stage 7: Generate Report ===", "INFO")
    return generate_report(
        supplier_info_dir, template_dir, rebate_form_input_dir,
        destination or output_path, suppliers, form_numbers, log,
        fy, quarter, icertis_codes,
    )
