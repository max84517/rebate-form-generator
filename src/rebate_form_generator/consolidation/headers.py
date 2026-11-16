"""Shared header handling.

Supplier workbooks are hand-maintained, so the header row is not always on
Excel row 2 and the column order is not the same between suppliers.  Every
stage therefore has to find the header row for *each* file and map columns by
name rather than by position.
"""
from __future__ import annotations

import re

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

REBATE_HEADER_RE = re.compile(
    r"^rebate\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+(\d{4})$"
)

# Cells Excel writes when a formula fails
ERROR_VALUES = {"#VALUE!", "#REF!", "#N/A", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "#SPILL!"}

# The first column of every data sheet's header row. This is what locks the
# header position down; everything else is a fallback.
HEADER_FIRST_COLUMN = "segment"

# How far down to look for it before giving up.
HEADER_SCAN_ROWS = 15

# Columns that identify a header row
_ANCHORS = (HEADER_FIRST_COLUMN, "gtk suppliers", "hp/odm part#", "platforms/project")

# Synthetic column added by stage 2 so later stages can name the workbook a
# row came from.  It never reaches the supplier-facing output.
SOURCE_FILE_COLUMN = "Source File"

# Windows-illegal filename characters
_ILLEGAL_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_NAMES = {
    "con", "prn", "aux", "nul",
    *(f"com{i}" for i in range(1, 10)),
    *(f"lpt{i}" for i in range(1, 10)),
}


# Month spellings seen in supplier workbooks, folded to the 3-letter form.
# Without this, "Rebate June 2027" and "Rebate Jun 2027" become two separate
# columns and one supplier's month silently goes missing.
_MONTH_ALIASES: dict[str, str] = {}
for _i, _abbr in enumerate(MONTH_ABBR):
    _full = ["january", "february", "march", "april", "may", "june", "july",
             "august", "september", "october", "november", "december"][_i]
    _MONTH_ALIASES[_full] = _abbr.lower()
    _MONTH_ALIASES[_abbr.lower()] = _abbr.lower()
_MONTH_ALIASES["sept"] = "sep"

_REBATE_ANY_MONTH = re.compile(
    r"^(rebate|hp cost|odm cost)\s+([a-z]+)\.?\s+(\d{4})$"
)


def normalize_header(value: object) -> str:
    """Fold a header cell to a comparable key.

    ``"ODM\\n(Regional Site)"`` → ``"odm (regional site)"``
    ``"Rebate June 2027"``      → ``"rebate jun 2027"``
    """
    if value is None:
        return ""
    text = re.sub(r"\s+", " ", str(value).replace("\n", " ").replace("\r", " ")).strip().lower()
    m = _REBATE_ANY_MONTH.match(text)
    if m:
        month = _MONTH_ALIASES.get(m.group(2))
        if month:
            return f"{m.group(1)} {month} {m.group(3)}"
    return text


def uses_nonstandard_month(value: object) -> bool:
    """True when a month column is spelled out ('June') rather than 'Jun'."""
    if value is None:
        return False
    text = re.sub(r"\s+", " ", str(value).replace("\n", " ").replace("\r", " ")).strip().lower()
    m = _REBATE_ANY_MONTH.match(text)
    return bool(m and m.group(2) in _MONTH_ALIASES and len(m.group(2)) > 3)


def find_header_row(rows: list[tuple], max_scan: int = HEADER_SCAN_ROWS) -> int | None:
    """Return the 0-based index of the header row in *rows*, or ``None``.

    Every data sheet starts its header with ``Segment``, so that is the lock:
    the first row whose opening cell reads ``Segment`` is the header, however
    many title or spacer rows somebody has inserted above it. Two looser
    passes follow in case a column was inserted to the left of it or the word
    was renamed.

    Never assume a fixed row: these workbooks are hand-maintained and a row
    added at the top silently turns the real header into a data row.
    """
    scanned = rows[:max_scan]

    # 1. The header begins with the key column — the normal case.
    for i, row in enumerate(scanned):
        if row and normalize_header(row[0]) == HEADER_FIRST_COLUMN:
            return i

    # 2. The key column is there but not first (a column was inserted before
    #    it); require a second known column so a stray word cannot match.
    others = [a for a in _ANCHORS if a != HEADER_FIRST_COLUMN]
    for i, row in enumerate(scanned):
        if not row:
            continue
        names = {normalize_header(v) for v in row}
        if HEADER_FIRST_COLUMN in names and any(a in names for a in others):
            return i

    # 3. No Segment column at all: fall back to whichever row carries the
    #    most known column names.
    best: tuple[int, int] | None = None   # (hits, index)
    for i, row in enumerate(scanned):
        if not row:
            continue
        names = {normalize_header(v) for v in row}
        hits = sum(1 for a in _ANCHORS if a in names)
        if hits >= 2 and (best is None or hits > best[0]):
            best = (hits, i)
    return best[1] if best else None


def header_map(row: tuple) -> dict[str, int]:
    """``{normalised header name: column index}``, first occurrence wins."""
    out: dict[str, int] = {}
    for i, v in enumerate(row):
        name = normalize_header(v)
        if name and name not in out:
            out[name] = i
    return out


def is_header_row(row: tuple, known_names: set[str], min_hits: int = 3) -> bool:
    """True when *row* is a stray copy of the header sitting in the data."""
    if not row:
        return False
    hits = sum(1 for v in row if normalize_header(v) in known_names and normalize_header(v))
    return hits >= min_hits


def fy_to_year(fy_label: str) -> int | None:
    """``"FY27"`` / ``"FY27 bNB"`` → ``2027``."""
    m = re.match(r"^fy\s*(\d{2})", str(fy_label).strip(), re.IGNORECASE)
    return 2000 + int(m.group(1)) if m else None


def expected_rebate_headers(fy_label: str) -> list[str]:
    """The 12 ``Rebate <Mon> <Year>`` headers an ``FY##`` sheet should carry.

    FY27 runs Nov 2026 → Oct 2027.
    """
    year = fy_to_year(fy_label)
    if year is None:
        return []
    months = [(11, year - 1), (12, year - 1)] + [(m, year) for m in range(1, 11)]
    return [f"rebate {MONTH_ABBR[m - 1].lower()} {y}" for m, y in months]


def is_error_value(value: object) -> bool:
    return isinstance(value, str) and value.strip().upper() in ERROR_VALUES


def safe_filename_part(name: str, fallback: str = "Unknown") -> str:
    """Make *name* safe to drop into a Windows filename."""
    cleaned = _ILLEGAL_FILENAME.sub("_", str(name or "")).strip().strip(".")
    cleaned = re.sub(r"\s+", " ", cleaned)
    if not cleaned or cleaned.lower() in _RESERVED_NAMES:
        return fallback
    return cleaned[:120]
