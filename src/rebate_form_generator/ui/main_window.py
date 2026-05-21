"""Main application window (CustomTkinter dark-mode UI).

Appearance mode and colour theme are set at module level to prevent
the double-window bug on some Windows configurations.
"""
from __future__ import annotations

import sys
import threading
import traceback
from pathlib import Path
from tkinter import BooleanVar, filedialog

import customtkinter as ctk

from rebate_form_generator.config.settings import Settings
from rebate_form_generator.consolidation.pipeline import (
    get_available_fy_sheets,
    run_full_pipeline,
    run_rebate_form_pipeline,
    run_report_pipeline,
)
from rebate_form_generator.consolidation.stage6_rebate_form import current_fy_quarter

# Must be set at module level, before any CTk widget is instantiated
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Optional columns available in the Generate Form Data dialog
_OPTIONAL_COLUMNS: list[str] = [
    "Segment", "Category", "SPM (Project Owner)", "HP/ODM Part#",
    "Series", "Platforms/Project", "Product", "Size",
    "Product Type", "Color", "ODM (Regional Site)", "IncoTerm",
]
_DEFAULT_CHECKED: frozenset[str] = frozenset([
    "Segment", "Category", "HP/ODM Part#", "Platforms/Project",
    "Product", "Size", "ODM (Regional Site)",
])


class FySelectionDialog(ctk.CTkToplevel):
    """Pop-up window for selecting a FY sheet and writing the pricing template."""

    def __init__(
        self,
        parent: ctk.CTk,
        fy_sheets: list[str],
        source_paths: dict,
        output_path: Path,
        log_callback,
        coverage: dict | None = None,
        fy_confirmed_callback=None,
        **kwargs,
    ) -> None:
        super().__init__(parent)
        self.title("Generate")
        self.resizable(False, False)
        self.grab_set()  # modal

        self._source_paths = source_paths
        self._output_path = output_path
        self._log_callback = log_callback
        self._coverage = coverage or {}
        self._fy_confirmed_callback = fy_confirmed_callback
        self._on_pipeline_done = kwargs.get("on_pipeline_done")

        # ── Layout ──────────────────────────────────────────────────────
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self,
            text="Select FY Sheet",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, padx=24, pady=(24, 4))

        self._fy_menu = ctk.CTkOptionMenu(
            self, values=fy_sheets, width=200, command=self._update_info
        )
        self._fy_menu.grid(row=1, column=0, padx=24, pady=(4, 8))
        def _fy_key(s: str) -> int:
            digits = "".join(c for c in s if c.isdigit())
            return int(digits) if digits else 0

        default_fy = max(fy_sheets, key=_fy_key) if fy_sheets else ""
        if default_fy:
            self._fy_menu.set(default_fy)

        self._info_box = ctk.CTkTextbox(
            self,
            height=80,
            state="disabled",
            font=ctk.CTkFont(family="Consolas", size=11),
        )
        self._info_box.grid(row=2, column=0, padx=24, pady=(0, 12), sticky="ew")

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=3, column=0, padx=24, pady=(0, 24))
        ctk.CTkButton(
            btn_frame, text="Generate", width=110, height=34, command=self._on_write,
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            btn_frame, text="Close", width=80, height=34,
            fg_color="transparent", border_width=1, command=self.destroy,
        ).pack(side="left")

        if default_fy:
            self._update_info(default_fy)

        # ── Centre over parent ───────────────────────────────────────────
        W, H = 320, 300
        self.update_idletasks()
        px = parent.winfo_x()
        py = parent.winfo_y()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        x = px + (pw - W) // 2
        y = py + (ph - H) // 2
        self.geometry(f"{W}x{H}+{x}+{y}")

        # Ensure the dialog appears on top of the main window
        self.attributes("-topmost", True)
        self.lift()
        self.after(100, lambda: self.attributes("-topmost", False))

    def _update_info(self, fy: str) -> None:
        """Refresh the info box when the FY selection changes."""
        fy_up = fy.strip().upper()
        missing: list[str] = []
        for seg, suppliers in self._coverage.items():
            bad = [
                sup for sup, fys in suppliers.items()
                if fy_up not in [f.strip().upper() for f in fys]
            ]
            if bad:
                missing.append(f"{seg}: {', '.join(bad)}")

        if not self._coverage:
            text = "(No coverage data)"
        elif missing:
            text = f"Missing {fy}:\n" + "\n".join(f"  {m}" for m in missing)
        else:
            text = f"✓  All segments have {fy}"

        self._info_box.configure(state="normal")
        self._info_box.delete("1.0", "end")
        self._info_box.insert("1.0", text)
        self._info_box.configure(state="disabled")

    def _on_write(self) -> None:
        fy_sheet = self._fy_menu.get()
        source_paths = self._source_paths
        output_path = self._output_path
        log = self._log_callback
        on_done = self._on_pipeline_done
        parent = self.master
        if self._fy_confirmed_callback:
            self._fy_confirmed_callback(fy_sheet)
        self.destroy()  # close dialog immediately

        def worker() -> None:
            try:
                result = run_full_pipeline(source_paths, fy_sheet, output_path, log)
                if result:
                    log(f"=== Pricing Template written \u2192 {result} ===", "INFO")
                else:
                    log("Failed to write Pricing Template. See log for details.", "ERROR")
            except Exception as exc:
                log(f"Error: {exc}", "ERROR")
                log(traceback.format_exc(), "ERROR")
            finally:
                if on_done:
                    parent.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()


class QuarterSelectionDialog(ctk.CTkToplevel):
    """Pop-up for selecting FY + quarter and generating the rebate form input.xlsx."""

    def __init__(
        self,
        parent: ctk.CTk,
        output_path: Path,
        log_callback,
        last_fy: str | None = None,
        on_pipeline_done=None,
    ) -> None:
        super().__init__(parent)
        self.title("Generate Form Data")
        self.resizable(False, False)
        self.grab_set()

        self._output_path = output_path
        self._log_callback = log_callback
        self._last_fy = last_fy
        self._on_pipeline_done = on_pipeline_done

        self.grid_columnconfigure(0, weight=1)

        # ── Quarter selector ──────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text="Select Quarter",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, padx=24, pady=(20, 4))

        options = self._make_options()
        default = self._default_option()

        self._quarter_menu = ctk.CTkOptionMenu(self, values=options, width=200)
        self._quarter_menu.grid(row=1, column=0, padx=24, pady=(4, 12))
        self._quarter_menu.set(default)

        # ── Column checkboxes ─────────────────────────────────────────────
        chk_frame = ctk.CTkFrame(self)
        chk_frame.grid(row=2, column=0, padx=16, pady=(0, 12), sticky="ew")
        chk_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            chk_frame, text="Include Columns", font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=2, padx=10, pady=(8, 4), sticky="w")

        self._col_vars: dict[str, BooleanVar] = {}
        for idx, col in enumerate(_OPTIONAL_COLUMNS):
            var = BooleanVar(value=col in _DEFAULT_CHECKED)
            self._col_vars[col] = var
            r = (idx // 2) + 1
            c = idx % 2
            ctk.CTkCheckBox(
                chk_frame, text=col, variable=var,
                checkbox_width=18, checkbox_height=18, font=ctk.CTkFont(size=12),
            ).grid(row=r, column=c, padx=(10, 4), pady=2, sticky="w")

        # ── Buttons ───────────────────────────────────────────────────────
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=3, column=0, padx=24, pady=(4, 20))
        ctk.CTkButton(
            btn_frame, text="Generate", width=110, height=34, command=self._on_generate,
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            btn_frame, text="Close", width=80, height=34,
            fg_color="transparent", border_width=1, command=self.destroy,
        ).pack(side="left")

        W, H = 400, 440
        self.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        self.geometry(f"{W}x{H}+{px + (pw - W) // 2}+{py + (ph - H) // 2}")
        self.attributes("-topmost", True)
        self.lift()
        self.after(100, lambda: self.attributes("-topmost", False))

    def _make_options(self) -> list[str]:
        if self._last_fy:
            fy_label = self._last_fy.upper()  # e.g. "FY26"
            return [f"{fy_label} Q{q}" for q in range(1, 5)]
        fy, _ = current_fy_quarter()
        options = []
        for fy_off in range(-2, 2):
            fy_val = (fy + fy_off) % 100
            for q in range(1, 5):
                options.append(f"FY{fy_val:02d} Q{q}")
        return options

    def _default_option(self) -> str:
        cur_fy, cur_q = current_fy_quarter()
        if self._last_fy:
            fy_num = int(self._last_fy.upper().lstrip("FY"))
            if fy_num == cur_fy % 100:
                return f"{self._last_fy.upper()} Q{cur_q}"
            return f"{self._last_fy.upper()} Q1"
        return f"FY{cur_fy % 100:02d} Q{cur_q}"

    def _on_generate(self) -> None:
        label = self._quarter_menu.get()  # e.g. "FY26 Q3"
        selected_columns = [col for col, var in self._col_vars.items() if var.get()]
        output_path = self._output_path
        log = self._log_callback
        on_done = self._on_pipeline_done
        parent = self.master
        self.destroy()

        parts = label.split()
        fy = int(parts[0][2:])
        q = int(parts[1][1:])

        def worker() -> None:
            try:
                result = run_rebate_form_pipeline(output_path, fy, q, selected_columns, log)
                if result:
                    log(
                        f"=== Form Data saved: {len(result)} file(s) "
                        f"\u2192 {result[0].parent} ===",
                        "INFO",
                    )
                else:
                    log("Failed to generate form data. See log for details.", "ERROR")
            except Exception as exc:
                log(f"Error: {exc}", "ERROR")
                log(traceback.format_exc(), "ERROR")
            finally:
                if on_done:
                    parent.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()


class GenerateReportDialog(ctk.CTkToplevel):
    """Pop-up for selecting suppliers and Form# to generate Word contracts."""

    def __init__(
        self,
        parent: ctk.CTk,
        output_path: Path,
        log_callback,
    ) -> None:
        super().__init__(parent)
        self.title("Generate Report")
        self.resizable(False, False)
        self.grab_set()

        self._output_path = output_path
        self._log_callback = log_callback

        rebate_form_input_dir = output_path.parent / "rebate form input"
        self._suppliers = self._detect_suppliers(rebate_form_input_dir)

        self.grid_columnconfigure(0, weight=1)

        # ── Title ─────────────────────────────────────────────────────
        ctk.CTkLabel(
            self, text="Select Suppliers",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, padx=24, pady=(20, 4))

        # ── Supplier checkboxes (scrollable) ──────────────────────────
        scroll = ctk.CTkScrollableFrame(self, height=min(180, max(60, len(self._suppliers) * 32)))
        scroll.grid(row=1, column=0, padx=16, pady=(4, 8), sticky="ew")
        scroll.grid_columnconfigure(0, weight=1)

        self._supplier_vars: dict[str, BooleanVar] = {}
        if self._suppliers:
            for i, supplier in enumerate(self._suppliers):
                var = BooleanVar(value=True)
                self._supplier_vars[supplier] = var
                ctk.CTkCheckBox(
                    scroll, text=supplier, variable=var,
                    checkbox_width=18, checkbox_height=18,
                    font=ctk.CTkFont(size=12),
                ).grid(row=i, column=0, padx=10, pady=2, sticky="w")
        else:
            ctk.CTkLabel(
                scroll, text="No contract input files found.",
                text_color="gray",
            ).grid(row=0, column=0, padx=10, pady=8)

        # ── Form# input ───────────────────────────────────────────────
        form_frame = ctk.CTkFrame(self, fg_color="transparent")
        form_frame.grid(row=2, column=0, padx=16, pady=(4, 4), sticky="ew")
        form_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            form_frame, text="Form #", width=70, anchor="e",
            font=ctk.CTkFont(weight="bold"),
        ).grid(row=0, column=0, padx=(8, 6), pady=6)
        self._form_entry = ctk.CTkEntry(
            form_frame, placeholder_text="Enter form number…", width=200,
        )
        self._form_entry.grid(row=0, column=1, padx=(0, 8), pady=6, sticky="ew")

        self._error_label = ctk.CTkLabel(
            self, text="", text_color="#e05555", font=ctk.CTkFont(size=11),
        )
        self._error_label.grid(row=3, column=0, padx=24, pady=(0, 4))

        # ── Buttons ───────────────────────────────────────────────────
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=4, column=0, padx=24, pady=(4, 20))
        ctk.CTkButton(
            btn_frame, text="Generate", width=110, height=34,
            command=self._on_generate,
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            btn_frame, text="Close", width=80, height=34,
            fg_color="transparent", border_width=1, command=self.destroy,
        ).pack(side="left")

        W, H = 420, 420
        self.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        self.geometry(f"{W}x{H}+{px + (pw - W) // 2}+{py + (ph - H) // 2}")
        self.attributes("-topmost", True)
        self.lift()
        self.after(100, lambda: self.attributes("-topmost", False))

    # ------------------------------------------------------------------

    @staticmethod
    def _detect_suppliers(folder: Path) -> list[str]:
        if not folder.exists():
            return []
        prefix = "contract input - "
        return [
            f.stem[len(prefix):].strip()
            for f in sorted(folder.glob("contract input - *.xlsx"))
            if not f.name.startswith("~$")
        ]

    def _on_generate(self) -> None:
        form_num = self._form_entry.get().strip()
        selected = [s for s, var in self._supplier_vars.items() if var.get()]

        if not form_num:
            self._error_label.configure(text="Form # is required.")
            return
        if not selected:
            self._error_label.configure(text="Select at least one supplier.")
            return

        output_path = self._output_path
        log = self._log_callback
        self.destroy()

        def worker() -> None:
            try:
                result = run_report_pipeline(output_path, selected, form_num, log)
                if result:
                    log(
                        f"=== Report saved: {len(result)} file(s) "
                        f"\u2192 {result[0].parent} ===",
                        "INFO",
                    )
                else:
                    log("Failed to generate report. See log for details.", "ERROR")
            except Exception as exc:
                log(f"Error: {exc}", "ERROR")
                log(traceback.format_exc(), "ERROR")

        threading.Thread(target=worker, daemon=True).start()


class MainWindow(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Rebate Form Generator")
        self.geometry("820x580")
        self.minsize(640, 480)

        self._settings = Settings()
        self._source_paths: dict = {}
        self._output_path: Path | None = None
        self._is_running = False
        self._last_fy: str | None = None

        self._build_ui()
        self._load_config()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)  # log frame expands

        # Title
        ctk.CTkLabel(
            self,
            text="Rebate Form Generator",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(14, 6), sticky="w")

        # ── Source Folders ────────────────────────────────────────────────────────
        src_frame = ctk.CTkFrame(self)
        src_frame.grid(row=1, column=0, padx=16, pady=4, sticky="ew")
        src_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            src_frame, text="Source Folders", font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=3, padx=14, pady=(10, 2), sticky="w")

        self._nb_kb_var = ctk.StringVar()
        self._dt_kb_var = ctk.StringVar()
        self._peripheral_var = ctk.StringVar()

        self._add_path_row(src_frame, "NB KB:", self._nb_kb_var, row=1)
        self._add_path_row(src_frame, "DT KB:", self._dt_kb_var, row=2)
        self._add_path_row(src_frame, "Peripheral:", self._peripheral_var, row=3)

        # ── Output ──────────────────────────────────────────────────────
        out_frame = ctk.CTkFrame(self)
        out_frame.grid(row=2, column=0, padx=16, pady=4, sticky="ew")
        out_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            out_frame, text="Output", font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=3, padx=14, pady=(10, 2), sticky="w")

        self._output_var = ctk.StringVar()
        self._add_path_row(out_frame, "Output Path:", self._output_var, row=1)

        # ── Action buttons ──────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.grid(row=3, column=0, padx=16, pady=(8, 4))

        self._build_btn = ctk.CTkButton(
            btn_row,
            text="Consolidate Rebate Data",
            height=34,
            command=self._on_build,
        )
        self._build_btn.pack(side="left", padx=(0, 8))

        self._rebate_form_btn = ctk.CTkButton(
            btn_row,
            text="Generate Form Data",
            height=34,
            command=self._on_rebate_form,
        )
        self._rebate_form_btn.pack(side="left", padx=(0, 8))

        self._report_btn = ctk.CTkButton(
            btn_row,
            text="Generate Report",
            height=34,
            command=self._on_generate_report,
        )
        self._report_btn.pack(side="left", padx=(0, 8))

        self._run_all_btn = ctk.CTkButton(
            btn_row,
            text="Run All",
            height=34,
            fg_color="#2d7d2d",
            hover_color="#1e5c1e",
            command=self._on_run_all,
        )
        self._run_all_btn.pack(side="left")

        # ── Log ─────────────────────────────────────────────────────────
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=4, column=0, padx=16, pady=(0, 16), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_frame, text="Log", font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, padx=14, pady=(8, 2), sticky="w")

        self._log_box = ctk.CTkTextbox(
            log_frame,
            state="disabled",
            wrap="word",
            font=ctk.CTkFont(family="Consolas", size=11),
        )
        self._log_box.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")

    def _add_path_row(
        self, parent: ctk.CTkFrame, label: str, var: ctk.StringVar, row: int
    ) -> None:
        ctk.CTkLabel(parent, text=label, width=110, anchor="e").grid(
            row=row, column=0, padx=(14, 6), pady=5
        )
        ctk.CTkEntry(parent, textvariable=var, placeholder_text="Select folder\u2026").grid(
            row=row, column=1, padx=4, pady=5, sticky="ew"
        )
        ctk.CTkButton(
            parent,
            text="Browse",
            width=84,
            command=lambda v=var: self._browse_folder(v),
        ).grid(row=row, column=2, padx=(4, 14), pady=5)

    # ------------------------------------------------------------------
    # Config helpers
    # ------------------------------------------------------------------

    def _load_config(self) -> None:
        self._nb_kb_var.set(self._settings.nb_kb)
        self._dt_kb_var.set(self._settings.dt_kb)
        self._peripheral_var.set(self._settings.peripheral)
        self._output_var.set(self._settings.output_path)
        self._last_fy = self._settings.last_fy or None

    def _save_config(self) -> None:
        self._settings.nb_kb = self._nb_kb_var.get()
        self._settings.dt_kb = self._dt_kb_var.get()
        self._settings.peripheral = self._peripheral_var.get()
        self._settings.output_path = self._output_var.get()
        self._settings.last_fy = self._last_fy or ""
        self._settings.save()

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _browse_folder(self, var: ctk.StringVar) -> None:
        current = var.get()
        initial = current if current and Path(current).exists() else str(Path.home())
        folder = filedialog.askdirectory(initialdir=initial)
        if folder:
            var.set(folder)

    def _on_close(self) -> None:
        self._save_config()
        self.destroy()
        sys.exit(0)

    def _on_build(self) -> None:
        if self._is_running:
            return

        # Validate required fields
        missing = [
            label
            for label, var in (
                ("NB KB", self._nb_kb_var),
                ("DT KB", self._dt_kb_var),
                ("Peripheral", self._peripheral_var),
                ("Output Path", self._output_var),
            )
            if not var.get().strip()
        ]
        if missing:
            self._append_log(f"[ERROR] Please fill in: {', '.join(missing)}")
            return

        self._save_config()
        self._clear_log()
        self._set_running(True)

        self._source_paths = {
            "nb_kb": self._nb_kb_var.get(),
            "dt_kb": self._dt_kb_var.get(),
            "peripheral": self._peripheral_var.get(),
        }
        self._output_path = Path(self._output_var.get())

        def worker() -> None:
            try:
                fy_sheets, coverage = get_available_fy_sheets(
                    self._source_paths, self._output_path, self._log
                )
                self.after(0, lambda s=fy_sheets, c=coverage: self._on_build_done(s, c))
            except Exception as exc:  # pragma: no cover
                self._log(f"Unexpected error: {exc}", "ERROR")
                self._log(traceback.format_exc(), "ERROR")
                self.after(0, lambda: self._set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def _on_build_done(self, fy_sheets: list[str], coverage: dict) -> None:
        self._set_running(False)
        if fy_sheets:
            self._log("=== Build completed — opening FY selection window… ===", "INFO")
            self._open_fy_dialog(fy_sheets, coverage)
        else:
            self._log("Build completed but no FY sheets found. Check source paths.", "WARNING")

    def _on_fy_confirmed(self, fy: str) -> None:
        self._last_fy = fy
        self._settings.last_fy = fy
        self._settings.save()

    def _open_fy_dialog(self, fy_sheets: list[str], coverage: dict) -> None:
        dialog = FySelectionDialog(
            parent=self,
            fy_sheets=fy_sheets,
            source_paths=self._source_paths,
            output_path=self._output_path,
            log_callback=self._log,
            coverage=coverage,
            fy_confirmed_callback=self._on_fy_confirmed,
        )
        dialog.focus()

    def _on_rebate_form(self) -> None:
        output_str = self._output_var.get().strip()
        if not output_str:
            self._append_log("[ERROR] Please set the Output Path first.")
            return
        output_path = Path(output_str)
        rebate_raw = output_path.parent / "rebate raw" / "rebate raw.xlsx"
        if not rebate_raw.exists():
            self._append_log(
                "[ERROR] rebate raw.xlsx not found. "
                "Run 'Consolidate Rebate Data' and select a FY first."
            )
            return
        self._output_path = output_path
        self._open_quarter_dialog()

    def _open_quarter_dialog(self) -> None:
        dialog = QuarterSelectionDialog(
            parent=self,
            output_path=self._output_path,
            log_callback=self._log,
            last_fy=self._last_fy,
        )
        dialog.focus()

    def _on_run_all(self) -> None:
        if self._is_running:
            return
        missing = [
            label
            for label, var in (
                ("NB KB", self._nb_kb_var),
                ("DT KB", self._dt_kb_var),
                ("Peripheral", self._peripheral_var),
                ("Output Path", self._output_var),
            )
            if not var.get().strip()
        ]
        if missing:
            self._append_log(f"[ERROR] Please fill in: {', '.join(missing)}")
            return

        self._save_config()
        self._clear_log()
        self._set_running(True)

        self._source_paths = {
            "nb_kb": self._nb_kb_var.get(),
            "dt_kb": self._dt_kb_var.get(),
            "peripheral": self._peripheral_var.get(),
        }
        self._output_path = Path(self._output_var.get())

        def open_report() -> None:
            GenerateReportDialog(
                parent=self,
                output_path=self._output_path,
                log_callback=self._log,
            ).focus()

        def open_form() -> None:
            QuarterSelectionDialog(
                parent=self,
                output_path=self._output_path,
                log_callback=self._log,
                last_fy=self._last_fy,
                on_pipeline_done=open_report,
            ).focus()

        def worker() -> None:
            try:
                fy_sheets, coverage = get_available_fy_sheets(
                    self._source_paths, self._output_path, self._log
                )
                self.after(
                    0,
                    lambda s=fy_sheets, c=coverage: self._on_run_all_build_done(s, c, open_form),
                )
            except Exception as exc:
                self._log(f"Unexpected error: {exc}", "ERROR")
                self._log(traceback.format_exc(), "ERROR")
                self.after(0, lambda: self._set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def _on_run_all_build_done(
        self, fy_sheets: list[str], coverage: dict, on_fy_done
    ) -> None:
        self._set_running(False)
        if fy_sheets:
            self._log("=== Build completed — opening FY selection window… ===", "INFO")
            FySelectionDialog(
                parent=self,
                fy_sheets=fy_sheets,
                source_paths=self._source_paths,
                output_path=self._output_path,
                log_callback=self._log,
                coverage=coverage,
                fy_confirmed_callback=self._on_fy_confirmed,
                on_pipeline_done=on_fy_done,
            ).focus()
        else:
            self._log("Build completed but no FY sheets found. Check source paths.", "WARNING")

    def _on_generate_report(self) -> None:
        output_str = self._output_var.get().strip()
        if not output_str:
            self._append_log("[ERROR] Please set the Output Path first.")
            return
        output_path = Path(output_str)
        rebate_form_input_dir = output_path.parent / "rebate form input"
        if not rebate_form_input_dir.exists():
            self._append_log(
                "[ERROR] rebate form input folder not found. "
                "Run 'Generate Form Data' first."
            )
            return
        self._output_path = output_path
        GenerateReportDialog(
            parent=self,
            output_path=output_path,
            log_callback=self._log,
        ).focus()

    # ------------------------------------------------------------------
    # UI state helpers
    # ------------------------------------------------------------------

    def _set_running(self, running: bool) -> None:
        self._is_running = running
        state = "disabled" if running else "normal"
        self._build_btn.configure(state=state)
        self._rebate_form_btn.configure(state=state)
        self._report_btn.configure(state=state)
        self._run_all_btn.configure(state=state)

    def _append_log(self, msg: str) -> None:
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _clear_log(self) -> None:
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _log(self, msg: str, level: str = "INFO") -> None:
        """Thread-safe log helper — safe to call from background threads."""
        self.after(0, lambda m=f"[{level}] {msg}": self._append_log(m))

    # ------------------------------------------------------------------
    # Entry point
    # ------------------------------------------------------------------

    def run(self) -> None:
        self.mainloop()
