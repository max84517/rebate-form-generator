"""Preview support for Word templates.

Two independent things:

* :func:`inspect_docx` — instant. Reads the document with python-docx and
  reports which placeholders are present, plus a short outline. Shown while
  the PDF is still converting, and the whole answer when conversion is not
  possible.
* :func:`to_pdf` — slow (Word takes ~15s per document). Drives the installed
  Word through PowerShell, so no extra Python package is needed. Results are
  cached and only regenerated when the source file changes.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

# Placeholders Stage 7 substitutes. Keep in step with stage7_report.
EXPECTED_PLACEHOLDERS = [
    "ICMRebateAgreementCode",
    "ICMAgreementCode",
    "ICMAgreementNumber1",
    "CW#",
    "ICMPartyName1",
    "ICMExternalSignatory",
    "ICMExternalSignatoryTitle",
    "ICMEffectiveDate",
]

CONVERT_TIMEOUT_SECONDS = 180


def _document_text(path: Path) -> tuple[str, list[dict], int]:
    """Return (all text, outline, table count) for a .docx."""
    from docx import Document

    doc = Document(path)
    chunks: list[str] = []
    outline: list[dict] = []

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        if text:
            chunks.append(text)
            if len(outline) < 40:
                style = (para.style.name if para.style is not None else "") or ""
                outline.append({
                    "kind": "heading" if "head" in style.lower() or "title" in style.lower()
                            else "paragraph",
                    "text": text[:300],
                })

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = (cell.text or "").strip()
                if text:
                    chunks.append(text)
        if len(outline) < 45:
            header = [c.text.strip()[:40] for c in table.rows[0].cells] if table.rows else []
            outline.append({
                "kind": "table",
                "text": f"{len(table.rows)} x {len(table.columns)}",
                "header": header,
            })

    # Content controls (w:sdt) hold text that doc.paragraphs/doc.tables do not
    # expose, and this template keeps nearly everything inside them.
    table_count = len(doc.tables)
    try:
        from docx.oxml.ns import qn
        chunks.append(doc.element.body.xml)
        table_count = sum(1 for _ in doc.element.body.iter(qn("w:tbl")))
        if not outline:
            for node in doc.element.body.iter(qn("w:t")):
                text = (node.text or "").strip()
                if text and len(outline) < 40:
                    outline.append({"kind": "paragraph", "text": text[:300]})
    except Exception:
        pass

    return "\n".join(chunks), outline, table_count


class TemplateUnreadable(Exception):
    """The template could not be opened — usually because Word has it open."""


def visible_text(path: Path) -> str:
    """The document's readable text, raising a clear error when it is locked."""
    try:
        from docx import Document
        doc = Document(path)
    except PermissionError as exc:
        raise TemplateUnreadable(
            f"'{path.name}' is open in another program (most likely Word). "
            f"Close it and try again."
        ) from exc
    except Exception as exc:
        raise TemplateUnreadable(f"'{path.name}' could not be read: {exc}") from exc

    import re
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", doc.element.body.xml))


def unfilled_keywords(path: Path, keywords: list[str]) -> list[str]:
    """Which of *keywords* the template still shows as the keyword itself.

    A control whose content is real text (HP's legal entity, say) is already
    written and needs no setting; one that still reads ``ICMInternalSignatory``
    is an empty blank that somebody has to fill.
    """
    text = visible_text(path)
    return [k for k in keywords if k and k in text]


def inspect_docx(path: Path, placeholders: list[str] | None = None) -> dict:
    """Placeholder check plus a short outline. Fast.

    *placeholders* defaults to the stock set; the server passes the live list,
    which follows whatever columns and shared values are configured.
    """
    wanted = list(placeholders or EXPECTED_PLACEHOLDERS)
    try:
        text, outline, table_count = _document_text(path)
    except Exception as exc:          # a corrupt or non-Word file
        return {
            "ok": False,
            "error": f"Could not read this document: {exc}",
            "found": [], "missing": wanted,
            "outline": [], "tables": 0,
        }

    found = [p for p in wanted if p in text]
    missing = [p for p in wanted if p not in found]
    return {
        "ok": True,
        "error": None,
        "found": found,
        "missing": missing,
        "outline": outline,
        "tables": table_count,
    }


# ---------------------------------------------------------------------------
# PDF conversion through the installed Word
# ---------------------------------------------------------------------------

_PS_CONVERT = r"""
$ErrorActionPreference = 'Stop'
$in  = $args[0]
$out = $args[1]
$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($in, [ref]$false, [ref]$true)
    $doc.SaveAs([ref]$out, [ref]17)
    Write-Output 'OK'
} catch {
    Write-Output ('ERR:' + $_.Exception.Message)
} finally {
    if ($doc  -ne $null) { try { $doc.Close([ref]$false) } catch {} }
    if ($word -ne $null) { try { $word.Quit() } catch {} }
}
"""


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer"):
            return True
    except OSError:
        return False


def to_pdf(source: Path, target: Path) -> tuple[bool, str]:
    """Convert *source* to *target*. Returns ``(ok, message)``.

    Cached: an up-to-date PDF is reused rather than re-converting, because a
    conversion costs roughly fifteen seconds.
    """
    if target.exists() and target.stat().st_mtime >= source.stat().st_mtime:
        return True, "cached"
    if not word_available():
        return False, "Microsoft Word is not available on this machine."

    target.parent.mkdir(parents=True, exist_ok=True)
    script = target.parent / "_convert.ps1"
    try:
        script.write_text(_PS_CONVERT, encoding="utf-8")
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", str(script), str(source.resolve()), str(target.resolve())],
            capture_output=True, text=True, timeout=CONVERT_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return False, "Word did not finish converting in time."
    except OSError as exc:
        return False, f"Could not start the converter: {exc}"
    finally:
        try:
            script.unlink()
        except OSError:
            pass

    output = (proc.stdout or "").strip()
    if output.startswith("ERR:"):
        return False, output[4:].strip() or "Word refused to convert this document."
    if not target.exists():
        return False, (proc.stderr or "Word produced no output.").strip()[:300]
    return True, "converted"
