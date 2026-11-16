"""Stage 3 — Merge all segment Consolidated files into one workbook.

Reads the four ``Consolidated_<segment>_<YYYYMMDD>.xlsx`` files produced by
Stage 2 and stacks them vertically per FY sheet.

Like Stage 2, columns are matched **by name**: bNB, cNB, DT and Peripheral do
not all carry the same column set, so a positional stack would misfile values.

Output: ``<output_base>/tmp/all/consolidate_all_<YYYYMMDD>.xlsx``
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from openpyxl.workbook import Workbook

from .headers import header_map, is_header_row


def _is_all_none(row: tuple | list) -> bool:
    return all(v is None for v in row)


def consolidate_all(
    segment_paths: list[Path],
    output_base: Path,
    log: Callable[[str, str], None],
) -> Path | None:
    """Merge all segment consolidated files into a single workbook.

    Returns the path of the saved file, or ``None`` on failure.
    """
    if not segment_paths:
        log("  [All] no segment files to merge", "ERROR")
        return None

    # -----------------------------------------------------------------------
    # Discover all FY sheets across all segment files
    # -----------------------------------------------------------------------
    all_fy: set[str] = set()
    for p in segment_paths:
        wb = load_workbook(p, data_only=True, read_only=True)
        for ws in wb.worksheets:
            all_fy.add(ws.title.strip().upper())
        wb.close()

    out_wb: Workbook = Workbook()
    out_wb.remove(out_wb.active)  # type: ignore[arg-type]

    # -----------------------------------------------------------------------
    # Per-FY merge
    # -----------------------------------------------------------------------
    for fy in sorted(all_fy):
        out_ws = out_wb.create_sheet(fy)

        # ── Pass A: collect each segment's rows and build the column union ──
        per_seg: list[dict] = []
        column_order: list[str] = []
        display_name: dict[str, str] = {}

        for seg_path in segment_paths:
            wb = load_workbook(seg_path, data_only=True, read_only=True)
            ws = next((s for s in wb.worksheets if s.title.strip().upper() == fy), None)
            if ws is None:
                wb.close()
                continue
            rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
            if not rows:
                continue

            hmap = header_map(rows[0])
            if not hmap:
                log(f"  [All] {fy}: {seg_path.name} has no header row — skipped", "WARNING")
                continue
            for name in hmap:
                if name not in display_name:
                    display_name[name] = str(rows[0][hmap[name]]).strip()
                    column_order.append(name)
            per_seg.append({"name": seg_path.name, "rows": rows, "hmap": hmap})

        if not per_seg:
            continue

        # ── Pass B: write header then map every segment's rows by name ─────
        out_ws.append([display_name[n] for n in column_order])
        known = set(column_order)
        merged = 0

        for seg in per_seg:
            hmap = seg["hmap"]
            for row in seg["rows"][1:]:
                if _is_all_none(row) or is_header_row(row, known):
                    continue
                mapped = [
                    row[hmap[n]] if n in hmap and hmap[n] < len(row) else None
                    for n in column_order
                ]
                if not _is_all_none(mapped):
                    out_ws.append(mapped)
            merged += 1

        log(f"  [All] {fy}: merged {merged} segment(s), {len(column_order)} column(s)", "INFO")

    all_dir = output_base / "tmp" / "all"
    all_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    out_path = all_dir / f"consolidate_all_{date_str}.xlsx"
    out_wb.save(out_path)
    log(f"  [All] saved → {out_path.name}", "INFO")
    return out_path
