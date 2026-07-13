"""Platform setup window: import a sample file, assign column roles."""

import logging
import os
import tkinter as tk
from ui.utils import fit_window, modern_button
from tkinter import ttk, filedialog, messagebox

from config.settings import DEFAULT_UNCHECKED, save_platform_config
from engine.data_pipeline import scan_file_structure
from parsers.dictionary import classify_columns

logger = logging.getLogger(__name__)


class PlatformSetupWindow:
    """Modal window that builds a platform's column-role config."""

    def __init__(self, parent, theme, on_save, existing_name=None, existing_config=None):
        self.window = tk.Toplevel(parent)
        self.window.title("Platform Setup" if not existing_name else f"Edit: {existing_name}")
        fit_window(self.window, 750, 700)
        t = self.theme = theme
        self.window.configure(bg=t["bg"])
        self.window.transient(parent)
        self.window.grab_set()
        self.window.lift()
        self.window.focus_force()
        self.on_save = on_save; self.sheet_configs = []

        top = tk.Frame(self.window, bg=t["bg"]); top.pack(fill="x", padx=18, pady=(16, 5))
        tk.Label(top, text="Platform name:", font=("Segoe UI", 11, "bold"),
                 bg=t["bg"], fg=t["fg"]).pack(side="left")
        self.name_var = tk.StringVar(value=existing_name or "")
        name_entry = tk.Entry(top, textvariable=self.name_var, width=25, font=("Segoe UI", 11),
                 bg=t["input_bg"], fg=t["input_fg"], insertbackground=t["insert"],
                 relief="solid", borderwidth=1)
        name_entry.pack(side="left", padx=10, fill="x", expand=True)

        bf = tk.Frame(self.window, bg=t["bg"]); bf.pack(fill="x", padx=18, pady=5)
        modern_button(bf, "Import Sample File", self._import_file, t,
                      kind="primary", padx=15, pady=6).pack(side="left")
        self.file_label = tk.Label(bf, text="No file selected", font=("Segoe UI", 9),
                                    bg=t["bg"], fg=t["muted"], anchor="w")
        self.file_label.pack(side="left", padx=10, fill="x", expand=True)

        # Keep actions in a dedicated bottom bar.  The previous layout packed
        # a full-height left canvas before the bottom bar, which made Tk place
        # Save Platform over the column headers on Windows.
        bottom = tk.Frame(self.window, bg=t["bg"])
        bottom.pack(side="bottom", fill="x", padx=18, pady=(8, 14))
        modern_button(bottom, "Save Platform", self._save, t, kind="success",
                      font_size=11, padx=20, pady=8).pack(side="right")
        modern_button(bottom, "Cancel", self.window.destroy, t, kind="secondary",
                      padx=15, pady=8).pack(side="right", padx=(0, 10))

        body = tk.Frame(self.window, bg=t["bg"])
        body.pack(fill="both", expand=True, padx=18, pady=(5, 0))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(body, bg=t["bg"], highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas, bg=t["bg"])
        self.scroll_frame.bind("<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._canvas_window = self.canvas.create_window(
            (0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfigure(self._canvas_window, width=e.width))
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns", padx=(8, 0))

        if existing_config:
            self._load_existing(existing_config)

    def _import_file(self):
        filepath = filedialog.askopenfilename(title="Select Sample File",
            filetypes=[("All Supported", "*.html *.htm *.csv *.xlsx *.xlsm"),
                       ("CSV", "*.csv"), ("Excel", "*.xlsx *.xlsm"),
                       ("HTML", "*.html *.htm")])
        if not filepath: return
        self.file_label.config(text=os.path.basename(filepath))
        try:
            sheets = scan_file_structure(filepath)
        except Exception as e:
            logger.exception("Sample file scan failed: %s", filepath)
            messagebox.showerror("Scan Error",
                f"Could not read {os.path.basename(filepath)}:\n{e}",
                parent=self.window)
            return
        if not sheets:
            messagebox.showwarning("No Data",
                "No data tables were found in that file.", parent=self.window)
            return
        for w in self.scroll_frame.winfo_children(): w.destroy()
        self.sheet_configs = []
        for sheet in sheets:
            self._add_sheet_config(sheet)

    def _add_sheet_config(self, sheet):
        t = self.theme; classification = sheet["classification"]
        sample = sheet.get("sample_row", {})
        config_data = {"sheet_name": sheet["name"], "columns": {}}

        hdr = tk.Frame(self.scroll_frame, bg=t["accent"])
        hdr.pack(fill="x", padx=5, pady=(12, 0))
        tk.Label(hdr, text=f"📊 {sheet['name']}  ({sheet['row_count']} rows)",
                 font=("Segoe UI", 11, "bold"), bg=t["accent"], fg="white"
                 ).pack(fill="x", padx=10, pady=5)

        # One shared grid per sheet: header and every column row live in the
        # SAME grid so the Role/Sample columns line up exactly. The previous
        # one-frame-per-row layout gave each row its own grid — long sample
        # text in one row shifted that row's columns, so the dropdowns
        # appeared staggered/floating relative to their neighbors.
        col_frame = tk.Frame(self.scroll_frame, bg=t["card"])
        col_frame.pack(fill="x", padx=5)
        col_frame.columnconfigure(1, weight=2, minsize=180)
        col_frame.columnconfigure(2, weight=1, minsize=150)
        col_frame.columnconfigure(3, weight=2, minsize=180)

        hdr_kw = {"font": ("Segoe UI", 9, "bold"), "bg": t["highlight"],
                  "fg": t["fg"]}
        tk.Label(col_frame, text="Extract", **hdr_kw).grid(
            row=0, column=0, sticky="nsew", ipadx=9, ipady=5)
        tk.Label(col_frame, text="Column", anchor="w", **hdr_kw).grid(
            row=0, column=1, sticky="nsew", ipadx=5, ipady=5)
        tk.Label(col_frame, text="Role", anchor="w", **hdr_kw).grid(
            row=0, column=2, sticky="nsew", ipadx=5, ipady=5)
        tk.Label(col_frame, text="Sample", anchor="w", **hdr_kw).grid(
            row=0, column=3, sticky="nsew", ipadx=5, ipady=5)

        role_options = ["metric", "campaign_id", "level: date", "level: device", "level: daypart",
                        "level: hour", "level: dow", "level: zip", "level: dma",
                        "level: audience", "level: creative", "level: keyword",
                        "level: network", "skip"]

        for grid_row, col_name in enumerate(sheet["headers"], start=1):
            auto_role = "metric"; auto_selected = True
            ctx_cols = classification["context"]
            lvl_cols = classification["levels"]
            skip_cols = classification["skip"]
            met_cols = classification["metrics"]

            if col_name in ctx_cols.values():
                auto_role = "campaign_id"; auto_selected = True
            elif any(col_name == lc for lc, _ in lvl_cols):
                ldef = next(ld for lc, ld in lvl_cols if lc == col_name)
                auto_role = f"level: {ldef['prefix']}"; auto_selected = True
            elif col_name in skip_cols:
                auto_role = "skip"; auto_selected = False
            else:
                universal = next((u for c, u in met_cols if c == col_name), col_name)
                auto_selected = universal not in DEFAULT_UNCHECKED

            sel_var = tk.BooleanVar(value=auto_selected)
            role_var = tk.StringVar(value=auto_role)
            tk.Checkbutton(col_frame, variable=sel_var, bg=t["card"],
                           selectcolor=t["input_bg"]).grid(
                               row=grid_row, column=0, padx=9, pady=2)
            tk.Label(col_frame, text=col_name, font=("Segoe UI", 10), bg=t["card"],
                     fg=t["card_fg"], anchor="w").grid(
                         row=grid_row, column=1, sticky="ew", padx=5, pady=1)
            ttk.Combobox(col_frame, textvariable=role_var, values=role_options,
                         state="readonly").grid(
                             row=grid_row, column=2, sticky="ew", padx=5, pady=1)
            sv = str(sample.get(col_name, ""))[:40]
            tk.Label(col_frame, text=sv, font=("Segoe UI", 9), bg=t["card"],
                     fg=t["muted"], anchor="w").grid(
                         row=grid_row, column=3, sticky="ew", padx=5, pady=1)
            config_data["columns"][col_name] = {"selected": sel_var, "role": role_var}

        btn_row = tk.Frame(col_frame, bg=t["card"])
        btn_row.grid(row=len(sheet["headers"]) + 1, column=0, columnspan=4,
                     sticky="ew", pady=5)
        def _all(cd=config_data):
            for v in cd["columns"].values(): v["selected"].set(True)
        def _none(cd=config_data):
            for v in cd["columns"].values(): v["selected"].set(False)
        tk.Button(btn_row, text="All", font=("Segoe UI", 8), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=6, command=_all
                  ).pack(side="left", padx=(10, 3))
        tk.Button(btn_row, text="None", font=("Segoe UI", 8), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=6, command=_none
                  ).pack(side="left", padx=3)
        self.sheet_configs.append(config_data)

    def _load_existing(self, config):
        for w in self.scroll_frame.winfo_children(): w.destroy()
        self.sheet_configs = []
        for sheet_cfg in config.get("sheets", []):
            fake = {"name": sheet_cfg["sheet_name"],
                    "headers": [c["column"] for c in sheet_cfg["columns"]],
                    "row_count": 0,
                    "classification": classify_columns([c["column"] for c in sheet_cfg["columns"]]),
                    "sample_row": {}}
            self._add_sheet_config(fake)
            sc = self.sheet_configs[-1]
            saved = {c["column"]: c for c in sheet_cfg["columns"]}
            for cn, widgets in sc["columns"].items():
                if cn in saved:
                    widgets["selected"].set(saved[cn].get("selected", True))
                    widgets["role"].set(saved[cn].get("role", "metric"))

    def _save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Name Required", "Enter a platform name.",
                                   parent=self.window)
            return
        if not self.sheet_configs:
            messagebox.showwarning("No Config", "Import a sample file first.",
                                   parent=self.window)
            return
        config = {"sheets": []}
        for sc in self.sheet_configs:
            sheet_out = {"sheet_name": sc["sheet_name"], "columns": []}
            for cn, widgets in sc["columns"].items():
                sheet_out["columns"].append({"column": cn,
                    "selected": widgets["selected"].get(), "role": widgets["role"].get()})
            config["sheets"].append(sheet_out)
        try:
            save_platform_config(name, config)
        except OSError as e:
            logger.exception("Could not save platform config for %s", name)
            messagebox.showerror("Save Failed",
                f"Could not save the platform config:\n{e}", parent=self.window)
            return
        logger.info("Platform config saved: %s (%d sheets)", name, len(config["sheets"]))
        self.on_save(name, config); self.window.destroy()
