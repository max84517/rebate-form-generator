"""Main application window (CustomTkinter dark-mode UI).

Appearance mode and colour theme are set at module level to prevent
the double-window bug on some Windows configurations.
"""
from __future__ import annotations

import sys
import threading
import traceback
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from rebate_form_generator.config.settings import Settings
from rebate_form_generator.consolidation.pipeline import (
    get_available_fy_sheets,
    run_full_pipeline,
)

# Must be set at module level, before any CTk widget is instantiated
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


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
    ) -> None:
        super().__init__(parent)
        self.title("Generate")
        self.resizable(False, False)
        self.grab_set()  # modal

        self._source_paths = source_paths
        self._output_path = output_path
        self._log_callback = log_callback
        self._coverage = coverage or {}

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
        if fy_sheets:
            self._fy_menu.set(fy_sheets[0])

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

        if fy_sheets:
            self._update_info(fy_sheets[0])

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

        # ── Build button ────────────────────────────────────────────────
        self._build_btn = ctk.CTkButton(
            self,
            text="Consolidate Rebate Data",
            height=34,
            command=self._on_build,
        )
        self._build_btn.grid(row=3, column=0, padx=16, pady=(8, 4))

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

    def _save_config(self) -> None:
        self._settings.nb_kb = self._nb_kb_var.get()
        self._settings.dt_kb = self._dt_kb_var.get()
        self._settings.peripheral = self._peripheral_var.get()
        self._settings.output_path = self._output_var.get()
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

    def _open_fy_dialog(self, fy_sheets: list[str], coverage: dict) -> None:
        dialog = FySelectionDialog(
            parent=self,
            fy_sheets=fy_sheets,
            source_paths=self._source_paths,
            output_path=self._output_path,
            log_callback=self._log,
            coverage=coverage,
        )
        dialog.focus()

    # ------------------------------------------------------------------
    # UI state helpers
    # ------------------------------------------------------------------

    def _set_running(self, running: bool) -> None:
        self._is_running = running
        self._build_btn.configure(state="disabled" if running else "normal")

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
