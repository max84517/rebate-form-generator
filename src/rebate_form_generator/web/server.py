"""Local HTTP server backing the dark-mode web UI.

Standard library only (``http.server``) so the tool keeps zero extra
dependencies.  Business logic lives untouched in ``consolidation.pipeline``;
this module only translates HTTP requests into pipeline calls and streams the
log back to the browser over Server-Sent Events.
"""
from __future__ import annotations

import json
import mimetypes
import os
import queue
import shutil
import socket
import string
import tempfile
import threading
import time
import uuid
import zipfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, urlparse, parse_qs, unquote

from rebate_form_generator.config.contract_info import KEY_FIELD as CONTRACT_KEY_FIELD
from rebate_form_generator.config.contract_info import (
    RESERVED_FIELDS,
    SHARED_LABELS,
    ContractInfoStore,
)
from rebate_form_generator.config.settings import Settings
from rebate_form_generator.config.templates import TemplateStore
from rebate_form_generator.consolidation import issues as issue_log
from rebate_form_generator.consolidation.pipeline import (
    run_full_pipeline,
    run_rebate_form_pipeline,
    run_report_pipeline,
    scan_fy_coverage,
)
from rebate_form_generator.consolidation.stage6_rebate_form import current_fy_quarter

from .docx_preview import (
    EXPECTED_PLACEHOLDERS,
    TemplateUnreadable,
    inspect_docx,
    to_pdf,
    unfilled_keywords,
    word_available,
)
from .jobs import Hub, sse_format

STATIC_DIR = Path(__file__).resolve().parent / "static"

# Returned by /api/ping so a second launch can tell our server apart from
# whatever else might be sitting on the port.
APP_SIGNATURE = "rebate-form-generator"

# Changes on every start, and on every edit to a static file, so the browser
# can never pair a new page with an old stylesheet or script.
BUILD_ID = str(int(time.time()))

# Stop the server once the browser has been gone this long. Without it a
# closed window leaves the process running, every later launch lands on a
# different port, and an old tab can still be showing a stale page.
# EventSource reconnects within ~3s, so this is well clear of a blip.
IDLE_SHUTDOWN_SECONDS = 12.0

# How often the SSE stream writes a comment. The write is what detects a
# closed window, so it also bounds how fast the shutdown above kicks in.
SSE_KEEPALIVE_SECONDS = 3.0

# Optional columns offered in the "Generate Form Data" dialog
OPTIONAL_COLUMNS: list[str] = [
    "Segment", "Category", "SPM (Project Owner)", "HP/ODM Part#",
    "Series", "Platforms/Project", "Product", "Size",
    "Product Type", "Color", "ODM (Regional Site)", "IncoTerm",
]
DEFAULT_CHECKED: list[str] = [
    "Segment", "Color", "HP/ODM Part#", "Platforms/Project",
    "Product", "Size", "ODM (Regional Site)",
]

CONTRACT_INPUT_PREFIX = "contract input - "

# Each source folder must directly contain a sub-folder whose name holds this
# keyword (see consolidation.stage1_ingest._find_segment_root).
SEGMENT_KEYWORD = {
    "nb_kb": "Master price table_NB",
    "dt_kb": "Master price table_DT",
    "peripheral": "Master price table_Peripheral",
}
SEGMENT_TITLE = {"nb_kb": "NB KB", "dt_kb": "DT KB", "peripheral": "Peripheral"}


#: How many generated contract bundles to keep before the oldest is dropped.
KEEP_BUNDLES = 3


def prune_bundles(cache_dir: Path, keep: int = KEEP_BUNDLES) -> int:
    """Keep only the newest *keep* contract zips.

    Each Generate Report leaves a zip behind so the browser can fetch it. They
    used to accumulate forever in the system temp folder; now they live in the
    cache and the old ones are dropped.
    """
    removed = 0
    try:
        zips = sorted(cache_dir.glob("rfg_report_*.zip"),
                      key=lambda f: f.stat().st_mtime, reverse=True)
    except OSError:
        return 0
    for old in zips[keep:]:
        try:
            old.unlink()
            removed += 1
        except OSError:
            pass
    # Scratch folders only exist while a report is being built.
    for stale in cache_dir.glob("rfg_report_*"):
        if stale.is_dir():
            shutil.rmtree(stale, ignore_errors=True)
    return removed


def sweep_legacy_temp() -> int:
    """Delete report zips earlier versions left in the system temp folder."""
    removed = 0
    try:
        for old in Path(tempfile.gettempdir()).glob("rfg_report_*"):
            try:
                if old.is_dir():
                    shutil.rmtree(old, ignore_errors=True)
                else:
                    old.unlink()
                removed += 1
            except OSError:
                pass
    except OSError:
        pass
    return removed


def file_tracker(source_paths: dict) -> dict:
    """The newest workbook in every supplier folder, and whether it is current.

    Lets the user see at a glance which suppliers have not sent an update
    this month, without opening three folder trees.
    """
    from rebate_form_generator.consolidation.stage1_ingest import (
        _find_segment_root,
        _get_latest_xlsx,
        _make_supplier_folder_re,
        _parse_supplier_name,
    )

    today = datetime.now()
    segments: list[dict] = []
    stale = 0
    total = 0

    for key, seg in (("nb_kb", "NB"), ("dt_kb", "DT"), ("peripheral", "Peripheral")):
        entry: dict = {"key": key, "label": SEGMENT_TITLE[key], "segment": seg,
                       "error": None, "files": []}
        root = Path((source_paths.get(key) or "").strip() or ".")
        if not (source_paths.get(key) or "").strip():
            entry["error"] = "This source folder is not set."
            segments.append(entry)
            continue
        if not root.is_dir():
            entry["error"] = "Folder not found."
            segments.append(entry)
            continue

        seg_root = _find_segment_root(root, f"Master price table_{seg}")
        if seg_root is None:
            entry["error"] = f"No “Master price table_{seg}…” sub-folder here."
            segments.append(entry)
            continue

        folder_re = _make_supplier_folder_re(seg)
        try:
            children = sorted(seg_root.iterdir())
        except OSError as exc:
            entry["error"] = f"Cannot read this folder: {exc}"
            segments.append(entry)
            continue

        for folder in children:
            if not folder.is_dir() or not folder_re.match(folder.name):
                continue
            supplier = _parse_supplier_name(folder.name)
            latest = _get_latest_xlsx(folder)
            total += 1
            if latest is None:
                stale += 1
                entry["files"].append({
                    "supplier": supplier, "folder": folder.name, "file": None,
                    "path": str(folder), "modified": None, "size": 0,
                    "thisMonth": False, "note": "No .xlsx in this folder",
                })
                continue
            st = latest.stat()
            modified = datetime.fromtimestamp(st.st_mtime)
            current = (modified.year, modified.month) == (today.year, today.month)
            if not current:
                stale += 1
            entry["files"].append({
                "supplier": supplier, "folder": folder.name, "file": latest.name,
                "path": str(latest),
                "modified": modified.isoformat(timespec="seconds"),
                "size": st.st_size, "thisMonth": current, "note": None,
            })
        segments.append(entry)

    return {
        "segments": segments,
        "month": today.strftime("%B %Y"),
        "checkedAt": today.strftime("%H:%M:%S"),
        "total": total,
        "stale": stale,
    }


def check_source_path(key: str, raw: str) -> dict:
    """Tell the user whether *raw* is the folder the pipeline expects."""
    keyword = SEGMENT_KEYWORD[key]
    raw = (raw or "").strip()
    if not raw:
        return {"state": "empty",
                "detail": f"Not set — pick the folder containing “{keyword}_…”"}

    folder = Path(raw)
    if not folder.is_dir():
        return {"state": "bad", "detail": "Folder not found"}

    try:
        children = [c for c in folder.iterdir() if c.is_dir()]
    except OSError:
        return {"state": "bad", "detail": "Cannot read this folder"}

    low = keyword.lower()
    match = next((c for c in children if low in c.name.lower()), None)
    if match is None:
        return {"state": "bad",
                "detail": f"No “{keyword}_…” sub-folder here — go one level up or down"}

    # Supplier folders follow "Master price table_<Segment>_<Supplier>";
    # anything else in there (e.g. "Consolidated Master price table_NB")
    # is not a supplier and must not be counted.
    prefix = keyword.lower()
    try:
        names = sorted(
            c.name[len(keyword):].lstrip("_")
            for c in match.iterdir()
            if c.is_dir() and c.name.lower().startswith(prefix) and len(c.name) > len(keyword)
        )
    except OSError:
        names = []
    if not names:
        return {"state": "bad",
                "detail": f"{match.name} has no “{keyword}_<Supplier>” folders"}

    return {"state": "ok",
            "detail": f"{match.name} → {len(names)} supplier(s): "
                      f"{', '.join(names)}"}


# ---------------------------------------------------------------------------
# Application state
# ---------------------------------------------------------------------------

class App:
    """Holds settings, the event hub and the cached source paths."""

    def __init__(self) -> None:
        self.settings = Settings()
        self.hub = Hub()
        self.source_paths: dict[str, str] = {
            "nb_kb": self.settings.nb_kb,
            "dt_kb": self.settings.dt_kb,
            "peripheral": self.settings.peripheral,
        }
        self.shutdown_event = threading.Event()
        # Generated contract bundles, kept in temp until the browser fetches
        # them: {token: (zip path, download filename)}
        self.bundles: dict[str, tuple[Path, str]] = {}

    @property
    def contract_info(self) -> ContractInfoStore:
        return ContractInfoStore(self.base_dir)

    @property
    def templates(self) -> TemplateStore:
        return TemplateStore(self.base_dir / "template")

    def missing_shared_values(self) -> tuple[list[dict], str]:
        """Shared values the template needs but nobody has set.

        Returns ``(missing, blocking_error)``. *missing* entries carry the
        keyword and its friendly label so the dialog can name them.
        """
        template = self.templates.active_path()
        if template is None:
            return [], ("No Word template is available. Open Settings → "
                        "Contract Template and upload one.")

        data = self.contract_info.load_all()
        shared = data["shared"]
        try:
            unfilled = unfilled_keywords(template, list(shared))
        except TemplateUnreadable as exc:
            return [], str(exc)

        missing = [
            {"key": key, "label": SHARED_LABELS.get(key, key)}
            for key in unfilled
            if not (shared.get(key) or "").strip()
        ]
        return missing, ""

    def expected_placeholders(self) -> list[str]:
        """Every keyword Stage 7 will substitute, given the current settings."""
        data = self.contract_info.load_all()
        names: list[str] = []
        for name in (list(RESERVED_FIELDS)
                     + [f for f in data["fields"] if f != CONTRACT_KEY_FIELD]
                     + list(data["shared"])):
            if name and name not in names:
                names.append(name)
        return names

    # ── paths ────────────────────────────────────────────────────────
    @property
    def data_dir(self) -> Path:
        """User-maintained inputs, inside the project."""
        return self.settings.data_dir

    @property
    def work_dir(self) -> Path:
        """Derived data, outside the project so OneDrive does not sync it."""
        return self.settings.work_dir

    # Kept so existing call sites reading inputs stay correct.
    @property
    def base_dir(self) -> Path:
        return self.data_dir

    # ── config ───────────────────────────────────────────────────────
    def config_payload(self) -> dict:
        s = self.settings
        cur_fy, cur_q = current_fy_quarter()
        return {
            "nb_kb": s.nb_kb,
            "dt_kb": s.dt_kb,
            "peripheral": s.peripheral,
            "data_dir": str(self.data_dir),
            "work_dir": str(self.work_dir),
            "last_fy": s.last_fy,
            "last_quarter": s.last_quarter,
            "form_numbers": s.form_numbers,
            "icertis_codes": s.icertis_codes,
            "segment_keyword": SEGMENT_KEYWORD,
            "segment_title": SEGMENT_TITLE,
            "optional_columns": OPTIONAL_COLUMNS,
            "default_checked": DEFAULT_CHECKED,
            "current_fy": cur_fy,
            "current_quarter": cur_q,
            "rebate_raw_exists": (self.work_dir / "rebate raw" / "rebate raw.xlsx").exists(),
            "form_input_exists": (self.work_dir / "rebate form input").exists(),
        }

    def save_paths(self, nb_kb: str, dt_kb: str, peripheral: str) -> None:
        self.settings.nb_kb = nb_kb.strip()
        self.settings.dt_kb = dt_kb.strip()
        self.settings.peripheral = peripheral.strip()
        self.settings.save()
        self.source_paths = {
            "nb_kb": self.settings.nb_kb,
            "dt_kb": self.settings.dt_kb,
            "peripheral": self.settings.peripheral,
        }

    def remember_fy(self, fy: str) -> None:
        self.settings.last_fy = fy
        self.settings.save()

    def remember_quarter(self, fy_int: int, quarter: int) -> None:
        self.settings.last_fy = f"FY{fy_int:02d}"
        self.settings.last_quarter = quarter
        self.settings.save()

    def fy_int(self) -> int | None:
        raw = (self.settings.last_fy or "").upper().lstrip("FY")
        return int(raw) if raw.isdigit() else None

    # ── suppliers ────────────────────────────────────────────────────
    def detect_suppliers(self) -> list[str]:
        folder = self.work_dir / "rebate form input"
        if not folder.exists():
            return []
        return [
            f.stem[len(CONTRACT_INPUT_PREFIX):].strip()
            for f in sorted(folder.glob(f"{CONTRACT_INPUT_PREFIX}*.xlsx"))
            if not f.name.startswith("~$")
        ]


# ---------------------------------------------------------------------------
# Folder browsing
# ---------------------------------------------------------------------------

def list_drives() -> list[str]:
    return [f"{d}:/" for d in string.ascii_uppercase if Path(f"{d}:/").exists()]


def browse(path: str) -> dict:
    """Return the sub-folders of *path*; an empty path lists the drives."""
    drives = list_drives()
    if not path:
        return {"path": "", "parent": None, "dirs": [], "drives": drives,
                "crumbs": [], "error": None}

    p = Path(path)
    if not p.exists() or not p.is_dir():
        return {"path": path, "parent": None, "dirs": [], "drives": drives,
                "crumbs": [], "error": "Folder not found"}

    p = p.resolve()
    dirs: list[str] = []
    error = None
    try:
        for child in sorted(p.iterdir(), key=lambda c: c.name.lower()):
            try:
                if child.is_dir() and not child.name.startswith("$"):
                    dirs.append(child.name)
            except OSError:
                continue
    except PermissionError:
        error = "Permission denied"
    except OSError as exc:
        error = str(exc)

    parent = str(p.parent).replace("\\", "/") if p.parent != p else None
    crumbs: list[dict] = []
    acc = p
    chain: list[Path] = []
    while True:
        chain.append(acc)
        if acc.parent == acc:
            break
        acc = acc.parent
    for node in reversed(chain):
        crumbs.append({"name": node.name or str(node).replace("\\", "/"),
                       "path": str(node).replace("\\", "/")})

    return {
        "path": str(p).replace("\\", "/"),
        "parent": parent,
        "dirs": dirs,
        "drives": drives,
        "crumbs": crumbs,
        "error": error,
    }


# ---------------------------------------------------------------------------
# HTTP handler
# ---------------------------------------------------------------------------

class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "RebateFormGenerator"
    app: App  # injected by make_server

    # ── helpers ──────────────────────────────────────────────────────
    def log_message(self, fmt: str, *args) -> None:  # silence request spam
        pass

    def _send(self, status: int, body: bytes, ctype: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, data, status: int = 200) -> None:
        body = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
        self._send(status, body, "application/json; charset=utf-8")

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if not length:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return {}

    def _busy(self) -> None:
        self._json({"ok": False, "error": "A job is already running."}, 409)

    # ── GET ──────────────────────────────────────────────────────────
    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        route = parsed.path
        # keep_blank_values: a cleared field arrives as "key=" and must be
        # treated as empty, not as "absent" (which falls back to the config).
        qs = parse_qs(parsed.query, keep_blank_values=True)

        if route in ("/", "/index.html"):
            return self._static("index.html")
        if route.startswith("/static/"):
            return self._static(route[len("/static/"):])

        if route == "/api/ping":
            # Lets a second launch recognise an instance that is already up
            # instead of drifting to another port.
            return self._json({"app": APP_SIGNATURE, "pid": os.getpid()})
        if route == "/api/config":
            return self._json(self.app.config_payload())
        if route == "/api/state":
            return self._json({"running": self.app.hub.running})
        if route == "/api/suppliers":
            return self._json({"suppliers": self.app.detect_suppliers()})
        if route == "/api/browse":
            raw = unquote(qs.get("path", [""])[0])
            return self._json(browse(raw))
        if route == "/api/contract-info":
            data = self.app.contract_info.load_all()
            data["keyField"] = CONTRACT_KEY_FIELD
            data["reserved"] = RESERVED_FIELDS
            data["sharedLabels"] = SHARED_LABELS
            return self._json(data)
        if route == "/api/templates":
            data = self.app.templates.list()
            data["word_available"] = word_available()
            data["placeholders"] = self.app.expected_placeholders()
            return self._json(data)
        if route == "/api/templates/inspect":
            entry = self.app.templates.get(qs.get("id", [""])[0])
            if entry is None:
                return self._json({"ok": False, "error": "Template not found"}, 404)
            return self._json(inspect_docx(self.app.templates.path_of(entry),
                                           self.app.expected_placeholders()))
        if route == "/api/templates/preview":
            return self._template_pdf(qs.get("id", [""])[0])
        if route == "/api/report/download":
            return self._bundle(qs.get("token", [""])[0])
        if route == "/api/report/preflight":
            missing, blocking = self.app.missing_shared_values()
            return self._json({"ok": not (missing or blocking),
                               "needsSettings": missing, "error": blocking or None})
        if route == "/api/file-tracker":
            paths = {k: (qs.get(k, [None])[0] if qs.get(k) is not None
                         else getattr(self.app.settings, k, ""))
                     for k in SEGMENT_KEYWORD}
            return self._json(file_tracker(paths))
        if route == "/api/check-paths":
            result = {}
            for key in SEGMENT_KEYWORD:
                supplied = qs.get(key)
                raw = supplied[0] if supplied else getattr(self.app.settings, key, "")
                result[key] = check_source_path(key, raw)
            return self._json(result)
        if route == "/api/events":
            return self._events()

        self._send(404, b"Not found", "text/plain; charset=utf-8")

    # ── POST ─────────────────────────────────────────────────────────
    def do_POST(self) -> None:  # noqa: N802
        parsed_post = urlparse(self.path)
        route = parsed_post.path
        app = self.app
        hub = app.hub

        # Template upload sends raw bytes, not JSON.
        if route == "/api/templates/upload":
            return self._upload_template(parse_qs(parsed_post.query, keep_blank_values=True))

        data = self._read_json()

        if route == "/api/contract-info":
            rows = data.get("rows")
            if not isinstance(rows, list):
                return self._json({"ok": False, "error": "rows must be a list"}, 400)
            fields = data.get("fields")
            if fields is not None and not isinstance(fields, list):
                return self._json({"ok": False, "error": "fields must be a list"}, 400)
            clash = [f for f in (fields or []) if f in RESERVED_FIELDS]
            if clash:
                return self._json({"ok": False, "error":
                    f"{', '.join(clash)} is filled in automatically and cannot be "
                    f"a column."}, 400)
            saved = app.contract_info.save(
                fields or app.contract_info.fields(), rows, data.get("shared"))
            return self._json({"ok": True, **saved})

        if route == "/api/templates/activate":
            ok = app.templates.activate(str(data.get("id", "")))
            return self._json({"ok": ok} if ok
                              else {"ok": False, "error": "Template not found"},
                              200 if ok else 404)

        if route == "/api/templates/rename":
            name = str(data.get("name", "")).strip()
            if not name:
                return self._json({"ok": False, "error": "Name cannot be empty"}, 400)
            ok = app.templates.rename(str(data.get("id", "")), name)
            return self._json({"ok": ok} if ok
                              else {"ok": False, "error": "Template not found"},
                              200 if ok else 404)

        if route == "/api/templates/delete":
            ok = app.templates.delete(str(data.get("id", "")))
            return self._json({"ok": ok} if ok
                              else {"ok": False, "error": "Could not delete that template"},
                              200 if ok else 400)

        if route == "/api/config":
            app.save_paths(data.get("nb_kb", ""), data.get("dt_kb", ""),
                           data.get("peripheral", ""))
            return self._json({"ok": True})

        if route == "/api/log/clear":
            hub.clear_history()
            return self._json({"ok": True})

        if route == "/api/open-folder":
            which = str(data.get("which", "work"))
            target = {"work": app.work_dir, "data": app.data_dir,
                      "rebate-raw": app.work_dir / "rebate raw",
                      "form-input": app.work_dir / "rebate form input"}.get(which)
            if target is None:
                return self._json({"ok": False, "error": "Unknown folder"}, 400)
            target.mkdir(parents=True, exist_ok=True)
            try:
                os.startfile(str(target))      # noqa: S606 - local, user-initiated
            except AttributeError:
                return self._json({"ok": False,
                                   "error": "Only supported on Windows."}, 400)
            except OSError as exc:
                return self._json({"ok": False, "error": str(exc)}, 400)
            return self._json({"ok": True, "path": str(target)})

        if route == "/api/client-report":
            # The page describes itself right after it boots. Written to disk
            # so a launch can be diagnosed afterwards without a debugger.
            data["port"] = self.server.server_address[1]
            data["received_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
            try:
                target = app.base_dir / "last-client.json"
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_text(json.dumps(data, indent=2, ensure_ascii=False),
                                  encoding="utf-8")
            except OSError:
                pass
            modals = data.get("visibleModals") or []
            if modals:
                log_line = f"Page reported visible dialogs on load: {', '.join(modals)}"
                hub.log(log_line, "WARNING")
                print(f"  [client] {log_line}")
            else:
                print(f"  [client] page loaded cleanly "
                      f"(config={data.get('configLoaded')}, paths={data.get('pathsFilled')})")
            return self._json({"ok": True})

        if route == "/api/shutdown":
            self._json({"ok": True})
            app.shutdown_event.set()
            return

        # ── Stage 1-4 ────────────────────────────────────────────────
        if route == "/api/consolidate":
            missing = [label for label, key in
                       (("NB KB", "nb_kb"), ("DT KB", "dt_kb"), ("Peripheral", "peripheral"))
                       if not (data.get(key) or "").strip()]
            if missing:
                return self._json(
                    {"ok": False, "error": f"Please fill in: {', '.join(missing)}"}, 400)
            app.save_paths(data["nb_kb"], data["dt_kb"], data["peripheral"])
            run_all = bool(data.get("run_all"))
            source_paths = dict(app.source_paths)
            work_dir = app.work_dir

            def job(log):
                hub.clear_history()
                log("=== Scanning source folders for available FY ===", "INFO")
                fy_sheets, coverage = scan_fy_coverage(source_paths, log)
                if fy_sheets:
                    log("=== Scan complete \u2014 choose the FY to build ===", "INFO")
                else:
                    log("No FY sheets found. Check the source folder paths.", "WARNING")
                return {"fy_sheets": fy_sheets, "coverage": coverage, "run_all": run_all,
                        "issues": issue_log.as_dicts()}

            if not hub.start("consolidate", job):
                return self._busy()
            return self._json({"ok": True})

        # ── Stage 5 ──────────────────────────────────────────────────
        if route == "/api/stage5":
            fy_sheet = (data.get("fy_sheet") or "").strip()
            if not fy_sheet:
                return self._json({"ok": False, "error": "fy_sheet is required"}, 400)
            app.remember_fy(fy_sheet)
            run_all = bool(data.get("run_all"))
            source_paths = dict(app.source_paths)
            work_dir = app.work_dir

            def job(log):
                result = run_full_pipeline(source_paths, fy_sheet, work_dir, log)
                if result:
                    log(f"=== {fy_sheet} rebate raw written \u2192 {result} ===", "INFO")
                else:
                    log(f"Failed to build {fy_sheet}. See log for details.", "ERROR")
                return {"path": str(result) if result else None, "run_all": run_all,
                        "issues": issue_log.as_dicts()}

            if not hub.start("stage5", job):
                return self._busy()
            return self._json({"ok": True})

        # ── Stage 6 ──────────────────────────────────────────────────
        if route == "/api/form-data":
            rebate_raw = app.base_dir / "rebate raw" / "rebate raw.xlsx"
            if not rebate_raw.exists():
                return self._json({"ok": False, "error":
                    "rebate raw.xlsx not found. Run 'Consolidate Rebate Data' "
                    "and select a FY first."}, 400)
            try:
                fy = int(data["fy"])
                quarter = int(data["quarter"])
            except (KeyError, TypeError, ValueError):
                return self._json({"ok": False, "error": "fy and quarter are required"}, 400)
            columns = [c for c in data.get("columns", []) if c in OPTIONAL_COLUMNS]
            app.remember_quarter(fy, quarter)
            run_all = bool(data.get("run_all"))
            work_dir = app.work_dir

            def job(log):
                result = run_rebate_form_pipeline(work_dir, fy, quarter, columns, log)
                if result:
                    log(f"=== Form Data saved: {len(result)} file(s) "
                        f"\u2192 {result[0].parent} ===", "INFO")
                else:
                    log("Failed to generate form data. See log for details.", "ERROR")
                return {"files": [str(p) for p in result], "run_all": run_all,
                        "issues": issue_log.as_dicts()}

            if not hub.start("form-data", job):
                return self._busy()
            return self._json({"ok": True})

        # ── Stage 7 ──────────────────────────────────────────────────
        if route == "/api/report":
            # Refuse to produce contracts that would go out with
            # "ICMInternalSignatory" printed where a signer's name belongs.
            missing, blocking = app.missing_shared_values()
            if blocking:
                return self._json({"ok": False, "error": blocking}, 400)
            if missing:
                names = ", ".join(m["label"] for m in missing)
                return self._json({
                    "ok": False,
                    "needsSettings": missing,
                    "error": f"Set {names} in Settings → Contract Info before "
                             f"generating contracts.",
                }, 400)

            suppliers = [s for s in data.get("suppliers", []) if s]
            if not suppliers:
                return self._json({"ok": False, "error": "Select at least one supplier."}, 400)
            form_numbers = {k: str(v).strip() for k, v in (data.get("form_numbers") or {}).items()
                            if str(v).strip()}
            missing = [s for s in suppliers if not form_numbers.get(s)]
            if missing:
                return self._json({"ok": False,
                                   "error": f"Form # required for: {', '.join(missing)}"}, 400)
            icertis = {k: str(v).strip() for k, v in (data.get("icertis_codes") or {}).items()
                       if str(v).strip()}
            app.settings.form_numbers.update(form_numbers)
            app.settings.icertis_codes.update(icertis)
            app.settings.save()

            fy = app.fy_int()
            quarter = app.settings.last_quarter or None
            work_dir = app.work_dir
            data_dir = app.data_dir
            cache_dir = app.settings.cache_dir

            def job(log):
                # Contracts are built in a scratch folder and handed to the
                # browser as one zip; nothing is kept in the project folder.
                work = Path(tempfile.mkdtemp(prefix="rfg_report_", dir=cache_dir))
                out_paths, blanks = run_report_pipeline(
                    work_dir, data_dir, suppliers, form_numbers, log, fy, quarter, icertis,
                    destination=work)
                if not out_paths:
                    log("Failed to generate report. See log for details.", "ERROR")
                    shutil.rmtree(work, ignore_errors=True)
                    return {"files": [], "blanks": blanks,
                            "issues": issue_log.as_dicts()}

                stamp = datetime.now().strftime("%Y%m%d")
                label = f"FY{fy:02d}Q{quarter}" if fy and quarter else stamp
                zip_name = f"Rebate Agreement Update Forms_{label}_{stamp}.zip"
                zip_path = work.parent / f"{work.name}.zip"
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for path in out_paths:
                        zf.write(path, arcname=path.name)
                shutil.rmtree(work, ignore_errors=True)

                token = uuid.uuid4().hex
                app.bundles[token] = (zip_path, zip_name)
                prune_bundles(cache_dir)
                log(f"=== {len(out_paths)} contract(s) packaged as {zip_name} "
                    f"\u2014 downloading \u2192 your Downloads folder ===", "INFO")
                return {
                    "files": [p.name for p in out_paths],
                    "download": f"/api/report/download?token={token}",
                    "zipName": zip_name,
                    "blanks": blanks,
                    "issues": issue_log.as_dicts(),
                }

            if not hub.start("report", job):
                return self._busy()
            return self._json({"ok": True})

        self._send(404, b"Not found", "text/plain; charset=utf-8")

    # ── binary endpoints ─────────────────────────────────────────────
    def _send_file(self, path: Path, ctype: str, filename: str = "",
                   attachment: bool = False) -> None:
        try:
            body = path.read_bytes()
        except OSError:
            return self._send(404, b"Not found", "text/plain; charset=utf-8")
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if filename:
            disposition = "attachment" if attachment else "inline"
            self.send_header(
                "Content-Disposition",
                f"{disposition}; filename*=UTF-8''{quote(filename)}",
            )
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _template_pdf(self, template_id: str) -> None:
        store = self.app.templates
        entry = store.get(template_id)
        if entry is None:
            return self._json({"ok": False, "error": "Template not found"}, 404)
        source = store.path_of(entry)
        target = store.preview_path(entry)
        ok, message = to_pdf(source, target)
        if not ok:
            # 422 so the page can fall back to the outline instead of showing
            # a broken PDF frame.
            return self._json({"ok": False, "error": message}, 422)
        self._send_file(target, "application/pdf", f"{entry['name']}.pdf")

    def _bundle(self, token: str) -> None:
        found = self.app.bundles.get(token)
        if found is None:
            return self._json(
                {"ok": False, "error": "That download has expired. Generate again."}, 404)
        path, filename = found
        if not path.exists():
            return self._json({"ok": False, "error": "The file is no longer there."}, 404)
        self._send_file(path, "application/zip", filename, attachment=True)

    def _upload_template(self, qs: dict) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return self._json({"ok": False, "error": "No file received"}, 400)
        if length > 40 * 1024 * 1024:
            return self._json({"ok": False, "error": "That file is larger than 40 MB"}, 400)
        content = self.rfile.read(length)
        filename = unquote(qs.get("filename", ["template.docx"])[0])
        name = unquote(qs.get("name", [""])[0])
        if not filename.lower().endswith(".docx"):
            return self._json({"ok": False, "error": "Only .docx files are supported"}, 400)
        try:
            entry = self.app.templates.add(filename, content, name)
        except ValueError as exc:
            return self._json({"ok": False, "error": str(exc)}, 400)
        except OSError as exc:
            return self._json({"ok": False, "error": f"Could not save the file: {exc}"}, 500)
        return self._json({"ok": True, "template": entry})

    # ── static files ─────────────────────────────────────────────────
    def _static(self, rel: str) -> None:
        rel = rel.split("?", 1)[0]
        target = (STATIC_DIR / rel).resolve()
        try:
            target.relative_to(STATIC_DIR)
        except ValueError:
            return self._send(403, b"Forbidden", "text/plain; charset=utf-8")
        if not target.is_file():
            return self._send(404, b"Not found", "text/plain; charset=utf-8")
        if target.name == "index.html":
            # Stamp the asset URLs with this run's id so a restarted server
            # can never be paired with a stylesheet or script the browser
            # kept from an earlier run.
            body = target.read_text(encoding="utf-8").replace("__BUILD__", BUILD_ID)
            return self._send(200, body.encode("utf-8"), "text/html; charset=utf-8")
        ctype, _ = mimetypes.guess_type(target.name)
        if ctype and (ctype.startswith("text/") or ctype.endswith("javascript")
                      or ctype == "image/svg+xml"):
            ctype = f"{ctype}; charset=utf-8"
        self._send(200, target.read_bytes(), ctype or "application/octet-stream")

    # ── SSE ──────────────────────────────────────────────────────────
    def _events(self) -> None:
        hub = self.app.hub
        q = hub.subscribe()
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "close")
        self.end_headers()
        self.close_connection = True
        try:
            while not self.app.shutdown_event.is_set():
                try:
                    event = q.get(timeout=SSE_KEEPALIVE_SECONDS)
                    self.wfile.write(sse_format(event))
                except queue.Empty:
                    self.wfile.write(b": keep-alive\n\n")  # detects a closed tab
                self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, OSError, ValueError):
            pass
        finally:
            hub.unsubscribe(q)


# ---------------------------------------------------------------------------
# Server bootstrap
# ---------------------------------------------------------------------------

class _Server(ThreadingHTTPServer):
    daemon_threads = True
    # Windows SO_REUSEADDR lets a second process bind a port that is already
    # in use, after which the OS splits connections between the two servers.
    # Refusing reuse makes a clash fail loudly instead.
    allow_reuse_address = False


def _port_free(port: int) -> bool:
    """True when nothing is listening on *port*.

    Probing by connecting rather than by binding: a test bind with
    SO_REUSEADDR succeeds on Windows even when the port is taken, which made
    every launch think 8765 was free and quietly stack servers on top of
    each other.
    """
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex(("127.0.0.1", port)) != 0


def pick_port(preferred: int, attempts: int = 20) -> int:
    for offset in range(attempts):
        if _port_free(preferred + offset):
            return preferred + offset
    raise RuntimeError(f"No free port in {preferred}-{preferred + attempts - 1}")


def make_server(port: int) -> tuple[ThreadingHTTPServer, App]:
    app = App()
    # Tidy up after earlier versions, then keep the cache bounded.
    sweep_legacy_temp()
    prune_bundles(app.settings.cache_dir)
    handler = type("BoundHandler", (Handler,), {"app": app})
    httpd = _Server(("127.0.0.1", port), handler)
    return httpd, app


def ping(port: int, timeout: float = 0.6) -> bool:
    """True when *port* is already serving this application."""
    import urllib.error
    import urllib.request
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/api/ping", timeout=timeout
        ) as resp:
            return json.loads(resp.read().decode("utf-8")).get("app") == APP_SIGNATURE
    except (urllib.error.URLError, OSError, ValueError, TimeoutError):
        return False


def find_running_instance(preferred: int, attempts: int = 20) -> int | None:
    """Return the port of an instance that is already running, if any."""
    for offset in range(attempts):
        candidate = preferred + offset
        if _port_free(candidate):
            continue                     # nothing is listening there
        if ping(candidate):
            return candidate
    return None


def start_idle_watchdog(app: App) -> threading.Thread:
    """Shut the server down once the browser has been closed.

    This replaces relying on the Chrome process we spawned: when Chrome hands
    the URL to an already-running instance, that process exits immediately and
    tells us nothing about the window.  Live SSE connections do.
    """
    def watch() -> None:
        gone_since: float | None = None
        while not app.shutdown_event.is_set():
            time.sleep(1.0)
            if app.hub.listener_count > 0:
                gone_since = None
                continue
            if not app.hub.ever_attached or app.hub.running is not None:
                # never opened yet, or a job is still running — keep going
                gone_since = None
                continue
            if gone_since is None:
                gone_since = time.monotonic()
            elif time.monotonic() - gone_since >= IDLE_SHUTDOWN_SECONDS:
                app.shutdown_event.set()
                return

    thread = threading.Thread(target=watch, daemon=True, name="idle-watchdog")
    thread.start()
    return thread


def serve(port: int) -> tuple[ThreadingHTTPServer, App, threading.Thread]:
    """Start the server on a background thread and return it."""
    httpd, app = make_server(port)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True, name="http")
    thread.start()
    # give the socket a moment so the browser never races the first request
    time.sleep(0.15)
    return httpd, app, thread
