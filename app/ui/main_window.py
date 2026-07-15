"""Main application window: file import, platform assignment, and export flow."""

from __future__ import annotations

import logging
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from config.naming import unique_component
from config.paths import OUTPUT_DIR
from config.settings import load_platform_config, load_settings
from config.themes import get_theme
from engine.data_pipeline import apply_platform_config, filter_data_by_campaigns
from engine.errors import IngestionError
from engine.excel_writer import write_to_excel
from engine.kpi import compute_kpis
from parsers.csv_parser import CSVParser
from parsers.excel_parser import ExcelParser
from parsers.html_parser import HTMLParser
from ui.client_wizard import ClientWizard
from ui.review_view import ReviewMixin
from ui.settings_window import SettingsWindow
from ui.utils import (bordered_card, brand_header, configure_ttk_styles,
                      enable_mousewheel, fit_window, modern_button,
                      run_in_background)

logger = logging.getLogger(__name__)


class IngestionEngine(ReviewMixin):
    """Main window; owns app state (files, platforms, parsed data, theme)."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Spectrum Reach — Reporting Ingestion Engine")
        fit_window(self.root, 940, 640)
        self.root.minsize(760, 520)
        self.selected_files = []
        self.parsed_data = []
        self._busy = False

        settings = load_settings()
        self.platforms = settings.get("platforms", {})
        self.theme_name = settings.get("theme", "light")
        self.t = get_theme(self.theme_name)
        self._build_ui()
        enable_mousewheel(self.root)

    # ── Main layout ──────────────────────────────────────────────────────

    def _build_ui(self):
        t = self.t
        self.root.configure(bg=t["bg"])
        for widget in self.root.winfo_children():
            widget.destroy()
        try:
            self.root.state("normal")
        except tk.TclError:
            pass
        fit_window(self.root, 940, 640)
        configure_ttk_styles(self.root, t)

        # Returning from review creates a clean import session. Theme changes
        # explicitly preserve and restore the current files in _on_settings_saved.
        self.selected_files = []

        brand_header(
            self.root, t,
            "Reporting Ingestion Engine",
            "Turn platform exports into client-ready Excel and PowerPoint reports",
            step=1,
            action=("Settings", self._open_settings),
        )

        main = tk.Frame(self.root, bg=t["bg"])
        main.pack(fill="both", expand=True, padx=28, pady=20)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        intro = tk.Frame(main, bg=t["bg"])
        intro.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        tk.Label(intro, text="Build this month’s report", font=("Segoe UI", 18, "bold"),
                 bg=t["bg"], fg=t["fg"]).pack(anchor="w")
        tk.Label(intro, text="Add one or more exports, assign their platforms, then continue.",
                 font=("Segoe UI", 10), bg=t["bg"], fg=t["muted"]).pack(anchor="w", pady=(3, 0))

        file_card = bordered_card(main, t)
        file_card.grid(row=1, column=0, sticky="nsew")
        file_card.columnconfigure(0, weight=1)
        file_card.rowconfigure(2, weight=1)

        card_head = tk.Frame(file_card, bg=t["card"])
        card_head.grid(row=0, column=0, sticky="ew", padx=18, pady=(15, 8))
        tk.Label(card_head, text="1", font=("Segoe UI", 10, "bold"),
                 bg=t["pill_info_bg"], fg=t["pill_info_fg"], padx=9, pady=4
                 ).pack(side="left", padx=(0, 10))
        head_text = tk.Frame(card_head, bg=t["card"])
        head_text.pack(side="left")
        tk.Label(head_text, text="Source files", font=("Segoe UI", 12, "bold"),
                 bg=t["card"], fg=t["card_fg"]).pack(anchor="w")
        tk.Label(head_text, text="CSV, XLSX, XLSM, or HTML exports",
                 font=("Segoe UI", 8), bg=t["card"], fg=t["muted"]).pack(anchor="w")
        self.file_count_label = tk.Label(
            card_head, text="0 files", font=("Segoe UI", 8, "bold"),
            bg=t["secondary"], fg=t["secondary_fg"], padx=9, pady=4)
        self.file_count_label.pack(side="right")

        controls = tk.Frame(file_card, bg=t["card"])
        controls.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 10))
        self.add_files_button = modern_button(
            controls, "+  Add files", self._add_files, t, kind="primary")
        self.add_files_button.pack(side="left")
        self.clear_files_button = modern_button(
            controls, "Clear", self._clear_files, t, kind="secondary", font_size=9)
        self.clear_files_button.pack(side="left", padx=(8, 0))

        default_wrap = tk.Frame(controls, bg=t["card"])
        default_wrap.pack(side="right")
        tk.Label(default_wrap, text="Default platform", font=("Segoe UI", 8),
                 bg=t["card"], fg=t["muted"]).pack(side="left", padx=(0, 7))
        self.default_platform_var = tk.StringVar(
            value=list(self.platforms.keys())[0] if self.platforms else "")
        self.default_platform_menu = ttk.Combobox(
            default_wrap, textvariable=self.default_platform_var,
            values=list(self.platforms.keys()), state="readonly", width=21)
        self.default_platform_menu.pack(side="left")

        list_shell = tk.Frame(file_card, bg=t["surface_alt"],
                              highlightbackground=t["border"], highlightthickness=1)
        list_shell.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 16))
        list_shell.columnconfigure(0, weight=1)
        list_shell.rowconfigure(0, weight=1)

        self.file_canvas = tk.Canvas(list_shell, bg=t["surface_alt"],
                                     highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(list_shell, orient="vertical",
                                  command=self.file_canvas.yview)
        self.file_canvas.configure(yscrollcommand=scrollbar.set)
        self.file_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.file_frame = tk.Frame(self.file_canvas, bg=t["surface_alt"])
        self._file_window_id = self.file_canvas.create_window(
            (0, 0), window=self.file_frame, anchor="nw")
        self.file_frame.bind(
            "<Configure>",
            lambda _e: self.file_canvas.configure(scrollregion=self.file_canvas.bbox("all")))
        self.file_canvas.bind(
            "<Configure>",
            lambda e: self.file_canvas.itemconfigure(self._file_window_id, width=e.width))
        self.file_widgets = []

        self.empty_state = tk.Frame(self.file_frame, bg=t["surface_alt"])
        self.empty_state.pack(fill="both", expand=True, padx=20, pady=38)
        tk.Label(self.empty_state, text="↥", font=("Segoe UI", 24),
                 bg=t["surface_alt"], fg=t["accent"]).pack()
        tk.Label(self.empty_state, text="No report files added yet",
                 font=("Segoe UI", 11, "bold"), bg=t["surface_alt"],
                 fg=t["fg"]).pack(pady=(4, 2))
        tk.Label(self.empty_state, text="Use Add files to select this month’s exports.",
                 font=("Segoe UI", 9), bg=t["surface_alt"],
                 fg=t["muted"]).pack()

        lower = tk.Frame(main, bg=t["bg"])
        lower.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        lower.columnconfigure(0, weight=1)
        lower.columnconfigure(1, weight=0)

        date_card = bordered_card(lower, t)
        date_card.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        tk.Label(date_card, text="Reporting period", font=("Segoe UI", 10, "bold"),
                 bg=t["card"], fg=t["fg"]).grid(row=0, column=0, columnspan=5,
                                                sticky="w", padx=15, pady=(12, 7))
        tk.Label(date_card, text="Start", font=("Segoe UI", 8), bg=t["card"],
                 fg=t["muted"]).grid(row=1, column=0, sticky="w", padx=(15, 5), pady=(0, 12))
        self.main_start_var = tk.StringVar(value="")
        self._date_entry(date_card, self.main_start_var).grid(
            row=1, column=1, sticky="ew", pady=(0, 12))
        modern_button(date_card, "Calendar", lambda: self._pick_main_date(self.main_start_var),
                      t, kind="secondary", font_size=8, padx=8, pady=5).grid(
                          row=1, column=2, padx=(5, 15), pady=(0, 12))
        tk.Label(date_card, text="End", font=("Segoe UI", 8), bg=t["card"],
                 fg=t["muted"]).grid(row=1, column=3, sticky="w", padx=(0, 5), pady=(0, 12))
        self.main_end_var = tk.StringVar(value="")
        self._date_entry(date_card, self.main_end_var).grid(
            row=1, column=4, sticky="ew", pady=(0, 12))
        modern_button(date_card, "Calendar", lambda: self._pick_main_date(self.main_end_var),
                      t, kind="secondary", font_size=8, padx=8, pady=5).grid(
                          row=1, column=5, padx=(5, 15), pady=(0, 12))
        date_card.columnconfigure(1, weight=1)
        date_card.columnconfigure(4, weight=1)

        action_card = bordered_card(lower, t)
        action_card.grid(row=0, column=1, sticky="ns")
        self.run_button = modern_button(
            action_card, "Continue  →", self._quick_run, t,
            kind="success", font_size=12, padx=28, pady=12)
        self.run_button.pack(padx=18, pady=(16, 7))
        tk.Label(action_card, text="Assign clients next", font=("Segoe UI", 8),
                 bg=t["card"], fg=t["muted"]).pack(pady=(0, 13))

        status = tk.Frame(main, bg=t["bg"])
        status.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        status.columnconfigure(0, weight=1)
        self.status_label = tk.Label(
            status, text="Ready", font=("Segoe UI", 9), bg=t["bg"], fg=t["muted"])
        self.status_label.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(
            status, mode="indeterminate", length=170,
            style="Modern.Horizontal.TProgressbar")
        self.progress.grid(row=0, column=1, sticky="e")
        self.progress.grid_remove()

    def _date_entry(self, parent, variable):
        return tk.Entry(parent, textvariable=variable, width=13, font=("Segoe UI", 10),
                        bg=self.t["input_bg"], fg=self.t["input_fg"],
                        insertbackground=self.t["insert"], relief="flat", borderwidth=0,
                        highlightbackground=self.t["border"],
                        highlightcolor=self.t["focus"], highlightthickness=1)

    def _set_busy(self, busy, message=None):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for button in (getattr(self, "run_button", None),
                       getattr(self, "add_files_button", None),
                       getattr(self, "clear_files_button", None)):
            if button is not None:
                button.configure(state=state)
        if message:
            self.status_label.configure(text=message,
                                        fg=self.t["accent"] if busy else self.t["muted"])
        if busy:
            self.progress.grid()
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.grid_remove()

    # ── File selection ───────────────────────────────────────────────────

    def _refresh_file_state(self):
        count = len(self.selected_files)
        self.file_count_label.configure(text=f"{count} file" if count == 1 else f"{count} files")
        if count:
            if self.empty_state.winfo_manager():
                self.empty_state.pack_forget()
        elif not self.empty_state.winfo_manager():
            self.empty_state.pack(fill="both", expand=True, padx=20, pady=38)
        self.file_canvas.configure(scrollregion=self.file_canvas.bbox("all"))

    def _add_file_row(self, filepath):
        t = self.t
        row = tk.Frame(self.file_frame, bg=t["card"],
                       highlightbackground=t["border"], highlightthickness=1)
        row.pack(fill="x", padx=8, pady=4)
        row.columnconfigure(0, weight=1)

        name_wrap = tk.Frame(row, bg=t["card"])
        name_wrap.grid(row=0, column=0, sticky="ew", padx=12, pady=9)
        tk.Label(name_wrap, text=os.path.basename(filepath), font=("Segoe UI", 9, "bold"),
                 bg=t["card"], fg=t["card_fg"], anchor="w").pack(anchor="w")
        tk.Label(name_wrap, text=os.path.dirname(filepath), font=("Segoe UI", 7),
                 bg=t["card"], fg=t["muted"], anchor="w").pack(anchor="w")

        pvar = tk.StringVar(value=self.default_platform_var.get())
        platform_menu = ttk.Combobox(
            row, textvariable=pvar, values=list(self.platforms.keys()),
            state="readonly", width=20)
        platform_menu.grid(row=0, column=1, padx=(8, 8))

        def remove(target=row):
            try:
                index = next(i for i, item in enumerate(self.file_widgets)
                             if item["frame"] is target)
            except StopIteration:
                return
            self.selected_files.pop(index)
            self.file_widgets.pop(index)
            target.destroy()
            self._refresh_file_state()

        modern_button(row, "Remove", remove, t, kind="secondary", font_size=8,
                      padx=8, pady=5).grid(row=0, column=2, padx=(0, 10))
        self.file_widgets.append({"frame": row, "platform_var": pvar,
                                  "platform_menu": platform_menu})
        self._refresh_file_state()

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Select Report Files",
            filetypes=[("All Supported", "*.html *.htm *.csv *.xlsx *.xlsm"),
                       ("CSV", "*.csv"), ("Excel", "*.xlsx *.xlsm"),
                       ("HTML", "*.html *.htm")])
        for filepath in files:
            if filepath not in self.selected_files:
                self.selected_files.append(filepath)
                self._add_file_row(filepath)
        if files:
            self.status_label.configure(text="Files added. Choose a platform for each export.")

    def _clear_files(self):
        if self._busy:
            return
        self.selected_files.clear()
        for item in self.file_widgets:
            item["frame"].destroy()
        self.file_widgets.clear()
        self.parsed_data = []
        self.status_label.configure(text="Ready", fg=self.t["muted"])
        self._refresh_file_state()

    # ── Settings and dates ───────────────────────────────────────────────

    def _open_settings(self):
        if not self._busy:
            SettingsWindow(self.root, self.theme_name, self._on_settings_saved)

    def _pick_main_date(self, target_var):
        try:
            from tkcalendar import Calendar
        except ImportError:
            messagebox.showinfo("Install", "Run: python -m pip install tkcalendar")
            return
        t = self.t
        win = tk.Toplevel(self.root)
        win.title("Select Date")
        fit_window(win, 315, 305)
        win.configure(bg=t["bg"])
        win.transient(self.root)
        win.grab_set()
        cal = Calendar(win, selectmode="day", date_pattern="yyyy-mm-dd")
        cal.pack(pady=12, padx=12, fill="both", expand=True)

        def confirm(_event=None):
            target_var.set(cal.get_date())
            win.destroy()

        cal.bind("<<CalendarSelected>>", confirm)
        win.bind("<Escape>", lambda _e: win.destroy())

    def _validate_date_range(self, start, end):
        parsed = {}
        for label, value in (("Start", start), ("End", end)):
            if not value:
                continue
            try:
                parsed[label] = datetime.strptime(value, "%Y-%m-%d").date()
            except ValueError:
                messagebox.showwarning(
                    "Invalid Date",
                    f"{label} date must use YYYY-MM-DD (for example, 2026-07-01).")
                return False
        if parsed.get("Start") and parsed.get("End") and parsed["Start"] > parsed["End"]:
            messagebox.showwarning("Invalid Date Range",
                                   "The start date must be on or before the end date.")
            return False
        return True

    def _on_settings_saved(self, new_theme):
        settings = load_settings()
        self.platforms = settings.get("platforms", {})
        if new_theme != self.theme_name:
            old_files = [(path, self.file_widgets[i]["platform_var"].get())
                         for i, path in enumerate(self.selected_files)]
            self.theme_name = new_theme
            self.t = get_theme(new_theme)
            self._build_ui()
            for filepath, platform in old_files:
                self.selected_files.append(filepath)
                self._add_file_row(filepath)
                self.file_widgets[-1]["platform_var"].set(platform)
        else:
            values = list(self.platforms.keys())
            self.default_platform_menu["values"] = values
            if self.default_platform_var.get() not in values:
                self.default_platform_var.set(values[0] if values else "")
            for item in self.file_widgets:
                item["platform_menu"]["values"] = values
                if item["platform_var"].get() not in values:
                    item["platform_var"].set(self.default_platform_var.get())

    # ── Parse and export workflow ────────────────────────────────────────

    def _quick_run(self):
        if self._busy:
            return
        if not self.selected_files:
            messagebox.showwarning("No Files", "Add report files first.")
            return

        pairs = []
        for index, filepath in enumerate(self.selected_files):
            platform = self.file_widgets[index]["platform_var"].get()
            if not platform:
                messagebox.showwarning(
                    "No Platform", f"Assign a platform to: {os.path.basename(filepath)}")
                return
            config = load_platform_config(platform)
            if not config:
                messagebox.showwarning(
                    "Not Configured",
                    f"'{platform}' has no config. Go to Settings → Add Platform.")
                return
            pairs.append((filepath, platform, config))

        start = self.main_start_var.get().strip()
        end = self.main_end_var.get().strip()
        if not self._validate_date_range(start, end):
            return

        self._set_busy(True, f"Reading {len(pairs)} file(s)…")

        def work():
            parsed = []
            failures = []
            for filepath, platform_name, platform_config in pairs:
                ext = os.path.splitext(filepath)[1].lower()
                try:
                    if ext == ".csv":
                        data = CSVParser(filepath).parse(build_outputs=False)
                    elif ext in (".xlsx", ".xlsm"):
                        data = ExcelParser(filepath).parse(build_outputs=False)
                    elif ext in (".html", ".htm"):
                        data = HTMLParser(filepath, platform_name).parse()
                    else:
                        failures.append((os.path.basename(filepath), "Unsupported file type"))
                        continue
                    apply_platform_config(data, platform_config)
                    data["source_platform"] = platform_name
                    parsed.append(data)
                    logger.info("Parsed %s (%s): %d campaign metrics, %d level rows",
                                os.path.basename(filepath), platform_name,
                                len(data.get("campaign_metrics", {})),
                                len(data.get("level_data", [])))
                except IngestionError as exc:
                    logger.warning("Parse failed for %s: %s", filepath, exc)
                    failures.append((os.path.basename(filepath), exc.user_message))
                except Exception as exc:
                    logger.exception("Parse failed for %s", filepath)
                    failures.append((os.path.basename(filepath), str(exc)))
            return parsed, failures

        def success(result):
            parsed, failures = result
            if failures:
                detail = "\n".join(f"• {name}: {reason}" for name, reason in failures)
                if not parsed:
                    self._set_busy(False, "Import failed")
                    messagebox.showerror("Import Failed", f"No files could be read:\n\n{detail}")
                    return
                if not messagebox.askyesno(
                        "Some Files Failed",
                        f"These files could not be read:\n\n{detail}\n\n"
                        "Continue with the remaining files?"):
                    self._set_busy(False, "Import cancelled")
                    return
            self.parsed_data = parsed
            self._set_busy(False, "Files parsed successfully")
            settings = load_settings()
            ClientWizard(
                self.root, parsed, self.platforms, self.t, self._on_wizard_complete,
                start_date=start, end_date=end,
                output_folder=settings.get("output_folder", OUTPUT_DIR))

        def error(exc):
            logger.error("Unexpected background parse failure", exc_info=(type(exc), exc, exc.__traceback__))
            self._set_busy(False, "Import failed")
            messagebox.showerror("Import Failed", f"The files could not be processed:\n{exc}")

        run_in_background(self.root, work, success, error)

    def _on_wizard_complete(self, client_assignments, start_date, end_date, output_folder):
        """Export per-client files without freezing the interface."""
        self._set_busy(True, f"Exporting {len(client_assignments)} client report(s)…")
        used_names = set()
        safe_clients = {
            client_name: unique_component(client_name, used_names, fallback="Client")
            for client_name in client_assignments
        }

        def work(progress):
            from engine.excel_vba import ExcelSession
            results = []
            errors = []
            vba_errors = []
            total = len(client_assignments)
            # One Excel process for the whole batch — launching Excel per
            # client was the dominant fixed cost of the export stage.
            with ExcelSession() as excel:
                for i, (client_name, campaigns) in enumerate(
                        client_assignments.items(), 1):
                    label = (f"{client_name}" if total == 1
                             else f"({i}/{total}) {client_name}")
                    safe_client = safe_clients[client_name]
                    client_folder = os.path.join(output_folder, safe_client)
                    client_data = filter_data_by_campaigns(self.parsed_data, campaigns)
                    for data in client_data:
                        data["client_name"] = client_name
                        data["campaign_type"] = ""
                        data["start_date"] = start_date
                        data["end_date"] = end_date

                    output_path = os.path.join(
                        client_folder, f"{safe_client}_unified_report.xlsx")
                    try:
                        os.makedirs(client_folder, exist_ok=True)
                        progress(f"{label}: writing workbook + search…")
                        final_path, vba_error = write_to_excel(
                            client_data, output_path, "", excel_app=excel.app)
                        if vba_error:
                            vba_errors.append(vba_error)
                        rows = sum(len(data.get("campaign_metrics", {})) +
                                   len(data.get("level_data", [])) for data in client_data)
                        progress(f"{label}: computing KPIs…")
                        results.append({
                            "client_name": client_name,
                            "campaigns": campaigns,
                            "client_data": client_data,
                            "output_path": final_path,
                            "rows": rows,
                            "folder": client_folder,
                            # Precomputed here, off the Tk thread — building the
                            # KPI DataFrame for a large client froze the UI when
                            # the review screen opened.
                            "kpi": compute_kpis(client_data),
                        })
                    except Exception as exc:
                        logger.exception("Excel export failed for %s", client_name)
                        errors.append((client_name, output_path, str(exc)))
            return results, errors, vba_errors

        def success(payload):
            results, errors, vba_errors = payload
            self._export_results = results
            self._set_busy(False, "Review exported data")
            if vba_errors and not getattr(self, "_vba_warning_shown", False):
                self._vba_warning_shown = True
                messagebox.showwarning(
                    "Interactive search not enabled",
                    "The workbooks were saved, but the live Excel search could not be added:\n\n"
                    + vba_errors[0])
            if errors:
                detail = "\n\n".join(
                    f"{client}: {reason}\nTarget: {path}" for client, path, reason in errors)
                messagebox.showerror("Export Error", detail)
            self._review_index = 0
            self._show_next_review()

        def error(exc):
            logger.error("Unexpected background export failure", exc_info=(type(exc), exc, exc.__traceback__))
            self._set_busy(False, "Export failed")
            messagebox.showerror("Export Error", f"The export could not be completed:\n{exc}")

        def show_progress(message):
            # Status text only — _set_busy would restart the progress bar
            # animation on every stage change
            self.status_label.configure(text=message, fg=self.t["accent"])

        run_in_background(self.root, work, success, error,
                          on_progress=show_progress)

    def _show_next_review(self):
        if self._review_index >= len(self._export_results):
            self._show_final_summary()
            return
        result = self._export_results[self._review_index]
        self._show_review_in_main(result, self._review_index + 1, len(self._export_results))

    def run(self):
        self.root.mainloop()
