"""Structured source-file problems collected during a pipeline run.

The log callback is good for a running narrative, but the UI also needs a
machine-readable list so it can pop up "these Excel files need fixing, here
is which one and why".  Stages append to the module-level collector; the
pipeline resets it at the start of a run and hands the result to the caller.
"""
from __future__ import annotations

import threading
from dataclasses import dataclass, asdict, field

# ── Issue kinds ────────────────────────────────────────────────────────
HEADER_SHIFTED = "header_shifted"     # someone inserted a row above the header
HEADER_MISSING = "header_missing"     # no header row found at all
MONTH_MISMATCH = "month_mismatch"     # sheet says FY27 but columns are FY26 months
MISSING_COLUMN = "missing_column"     # expected rebate month column absent
ERROR_VALUE = "error_value"           # #VALUE! / #REF! … in a rebate cell
BAD_VALUE = "bad_value"               # non-numeric text in a rebate cell
REPEATED_HEADER = "repeated_header"   # a header row showed up inside the data
MONTH_SPELLING = "month_spelling"     # "Rebate June 2027" instead of "Rebate Jun 2027"
FILE_LOCKED = "file_locked"           # workbook open in Excel, or unreadable
SUPPLIER_SKIPPED = "supplier_skipped" # supplier dropped from the output

KIND_LABEL = {
    HEADER_SHIFTED:   "Header row moved",
    HEADER_MISSING:   "Header row not found",
    MONTH_MISMATCH:   "Wrong months for this FY",
    MISSING_COLUMN:   "Missing month column",
    ERROR_VALUE:      "Excel error value",
    BAD_VALUE:        "Non-numeric rebate value",
    REPEATED_HEADER:  "Header row inside the data",
    MONTH_SPELLING:   "Month spelled out",
    FILE_LOCKED:      "File could not be read",
    SUPPLIER_SKIPPED: "Supplier skipped",
}

# Severity drives ordering and colour in the UI
SEVERITY = {
    MONTH_MISMATCH:   "error",
    SUPPLIER_SKIPPED: "error",
    HEADER_MISSING:   "error",
    MISSING_COLUMN:   "warning",
    HEADER_SHIFTED:   "warning",
    ERROR_VALUE:      "warning",
    BAD_VALUE:        "warning",
    REPEATED_HEADER:  "warning",
    MONTH_SPELLING:   "warning",
    FILE_LOCKED:      "error",
}


@dataclass
class Issue:
    kind: str
    detail: str
    file: str = ""        # display name of the offending workbook
    path: str = ""        # full path to the original source file, when known
    sheet: str = ""
    supplier: str = ""
    count: int = 1

    @property
    def severity(self) -> str:
        return SEVERITY.get(self.kind, "warning")

    @property
    def label(self) -> str:
        return KIND_LABEL.get(self.kind, self.kind)

    def to_dict(self) -> dict:
        d = asdict(self)
        d["severity"] = self.severity
        d["label"] = self.label
        return d


@dataclass
class _Collector:
    items: list[Issue] = field(default_factory=list)
    _lock: threading.Lock = field(default_factory=threading.Lock)

    def reset(self) -> None:
        with self._lock:
            self.items = []

    def add(self, issue: Issue) -> None:
        """Append *issue*, merging with an identical earlier one."""
        with self._lock:
            for existing in self.items:
                if (existing.kind == issue.kind and existing.file == issue.file
                        and existing.sheet == issue.sheet
                        and existing.detail == issue.detail):
                    existing.count += issue.count
                    return
            self.items.append(issue)

    def all(self) -> list[Issue]:
        with self._lock:
            return list(self.items)

    def as_dicts(self) -> list[dict]:
        order = {"error": 0, "warning": 1}
        return [
            i.to_dict()
            for i in sorted(self.all(), key=lambda x: (order.get(x.severity, 9), x.file, x.kind))
        ]


collector = _Collector()


def reset() -> None:
    collector.reset()


def add(kind: str, detail: str, *, file: str = "", path: str = "",
        sheet: str = "", supplier: str = "", count: int = 1) -> None:
    collector.add(Issue(kind=kind, detail=detail, file=file, path=path,
                        sheet=sheet, supplier=supplier, count=count))


def as_dicts() -> list[dict]:
    return collector.as_dicts()
