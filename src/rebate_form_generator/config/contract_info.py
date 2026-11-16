"""Supplier contract details, owned by the app rather than by an Excel file.

Historically these values lived in ``supplier info/Contract Source info.xlsx``
and had to be edited in Excel.  They now live in ``data/contract_info.json``
and are edited in the Settings dialog; the spreadsheet is read exactly once,
to seed the store.
"""
from __future__ import annotations

import json
from pathlib import Path

# Columns a new install starts with. Each column other than the key is a
# placeholder keyword in the Word template, so adding a column here — or in
# the Settings dialog — is all that a new keyword needs.
DEFAULT_FIELDS: list[str] = [
    "Supplier Name",
    "ICMRebateAgreementCode",
    "CW#",
    "ICMPartyName1",
    "ICMExternalSignatory",
    "ICMExternalSignatoryTitle",
]

KEY_FIELD = DEFAULT_FIELDS[0]

# Placeholders Stage 7 computes or takes from the Generate Report dialog.
# They must not become columns, or the typed value would be overwritten.
RESERVED_FIELDS: list[str] = [
    "ICMAgreementCode",      # iCertis Code, entered per supplier
    "ICMAgreementNumber1",   # Form #, entered per supplier
    "ICMEffectiveDate",      # derived from the selected FY + quarter
]

# HP-side values that are the same on every contract, so they are held once
# instead of being repeated down a column.
#
# A value is *required* only when the template still shows the keyword itself
# rather than real text — see web.docx_preview.unfilled_keywords. HP's legal
# entity is deliberately not here: the template spells it out already.
DEFAULT_SHARED_FIELDS: list[str] = [
    "ICMInternalSignatory",   # HP signer name
    "ICMHPSignatoryTitle",    # HP signer title
]

SHARED_LABELS: dict[str, str] = {
    "ICMInternalSignatory": "Internal signer",
    "ICMHPSignatoryTitle": "Internal signer title",
    "ICMHPLegalEntity": "HP legal entity",
}

# Kept for callers that only want the stock layout
FIELDS = DEFAULT_FIELDS


class ContractInfoStore:
    """Read/write ``contract_info.json``, seeding it from the old workbook."""

    def __init__(self, base_dir: Path) -> None:
        self.base_dir = base_dir
        self.path = base_dir / "contract_info.json"
        self.legacy_dir = base_dir / "supplier info"

    # ── load ─────────────────────────────────────────────────────────
    def load_all(self) -> dict:
        """Return ``{"fields": [...], "rows": [...], "shared": {...}}``."""
        blank = {"fields": list(DEFAULT_FIELDS), "rows": [],
                 "shared": {f: "" for f in DEFAULT_SHARED_FIELDS}}
        if not self.path.exists():
            rows = self._import_legacy()
            if rows:
                return self.save(DEFAULT_FIELDS, rows, blank["shared"])
            return blank
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            return blank

        if isinstance(data, list):                      # very old shape
            data = {"fields": DEFAULT_FIELDS, "rows": data}
        fields = self._clean_fields(data.get("fields") or DEFAULT_FIELDS)
        rows = [self._clean(r, fields) for r in (data.get("rows") or [])
                if isinstance(r, dict)]
        return {"fields": fields, "rows": rows,
                "shared": self._clean_shared(data.get("shared"))}

    def load(self) -> list[dict]:
        return self.load_all()["rows"]

    def fields(self) -> list[str]:
        return self.load_all()["fields"]

    def placeholder_fields(self) -> list[str]:
        """Columns that map onto a Word placeholder (everything but the key)."""
        return [f for f in self.fields() if f != KEY_FIELD]

    def save(self, fields: list[str], rows: list[dict],
             shared: dict | None = None) -> dict:
        clean_fields = self._clean_fields(fields)
        cleaned = [self._clean(r, clean_fields) for r in rows if isinstance(r, dict)]
        cleaned = [r for r in cleaned if r[KEY_FIELD]]
        clean_shared = self._clean_shared(
            shared if shared is not None else self.load_all()["shared"])
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps({"fields": clean_fields, "shared": clean_shared,
                        "rows": cleaned}, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        return {"fields": clean_fields, "rows": cleaned, "shared": clean_shared}

    # ── lookup used by Stage 7 ───────────────────────────────────────
    def by_supplier(self) -> dict[str, dict]:
        return {r[KEY_FIELD]: r for r in self.load() if r[KEY_FIELD]}

    def shared(self) -> dict[str, str]:
        """HP-side values applied to every contract. Blanks mean 'leave as is'."""
        return {k: v for k, v in self.load_all()["shared"].items() if v}

    # ── helpers ──────────────────────────────────────────────────────
    @staticmethod
    def _clean_fields(fields) -> list[str]:
        """Normalise a column list: key first, no blanks, no duplicates."""
        out: list[str] = [KEY_FIELD]
        seen = {KEY_FIELD.lower()}
        for f in fields or []:
            name = str(f or "").strip()
            if not name or name.lower() in seen:
                continue
            seen.add(name.lower())
            out.append(name)
        return out

    @staticmethod
    def _clean_shared(shared) -> dict[str, str]:
        """Always offer the default shared fields; carry over anything else
        only while it still holds a value, so retired keys fade away."""
        out = {f: "" for f in DEFAULT_SHARED_FIELDS}
        if isinstance(shared, dict):
            for k, v in shared.items():
                name = str(k or "").strip()
                if not name or name in RESERVED_FIELDS:
                    continue
                value = "" if v is None else str(v).strip()
                if name in out or value:
                    out[name] = value
        return out

    @staticmethod
    def _clean(row: dict, fields: list[str]) -> dict:
        out = {}
        for f in fields:
            value = row.get(f, "")
            out[f] = "" if value is None else str(value).strip()
        return out

    def _import_legacy(self) -> list[dict]:
        """One-time import from the newest ``.xlsx`` in ``supplier info/``."""
        if not self.legacy_dir.exists():
            return []
        candidates = [f for f in self.legacy_dir.glob("*.xlsx")
                      if not f.name.startswith("~$")]
        if not candidates:
            return []
        newest = max(candidates, key=lambda f: f.stat().st_mtime)

        try:
            from openpyxl import load_workbook
            wb = load_workbook(newest, data_only=True)
        except Exception:
            return []

        ws = wb.active
        header_cells = next(ws.iter_rows(min_row=1, max_row=1), None)
        if header_cells is None:
            wb.close()
            return []
        headers = [str(c.value).strip() if c.value is not None else "" for c in header_cells]
        lookup = {h.lower(): i for i, h in enumerate(headers) if h}

        rows: list[dict] = []
        for raw in ws.iter_rows(min_row=2, values_only=True):
            row = {}
            for field in DEFAULT_FIELDS:
                idx = lookup.get(field.lower())
                value = raw[idx] if idx is not None and idx < len(raw) else None
                row[field] = "" if value is None else str(value).strip()
            if row[KEY_FIELD]:
                rows.append(row)
        wb.close()
        return rows
