"""Settings dialog: Platforms, Templates, Appearance, Export, Debug tabs."""

import logging
import os
import tkinter as tk
from ui.utils import fit_window, run_in_background
from tkinter import ttk, filedialog, messagebox

from config.naming import safe_component
from config.paths import LOGS_DIR, OUTPUT_DIR, TEMPLATES_DIR
from config.settings import (delete_platform_config, load_platform_config,
                             load_settings, save_settings)
from config.themes import get_theme
from ui.platform_setup import PlatformSetupWindow

logger = logging.getLogger(__name__)


class SettingsWindow:
    """Modal settings dialog; calls ``on_save(theme_name)`` when saved."""

    def __init__(self, parent, theme_name, on_save):
        self.window = tk.Toplevel(parent)
        self.window.title("Settings"); fit_window(self.window, 560, 520)
        t = self.theme = get_theme(theme_name)
        self.window.configure(bg=t["bg"])
        self.window.transient(parent)
        self.window.grab_set()
        self.window.lift()
        self.window.focus_force()
        self.on_save = on_save; self.theme_name = theme_name
        settings = load_settings()
        self.platforms = settings.get("platforms", {})
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)

        notebook = ttk.Notebook(self.window)
        notebook.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 0))

        # === Platforms tab ===
        pf = tk.Frame(notebook, bg=t["bg"]); notebook.add(pf, text="Platforms")
        lf = tk.Frame(pf, bg=t["bg"]); lf.pack(fill="both", expand=True, padx=15, pady=(10, 0))
        sb = tk.Scrollbar(lf); sb.pack(side="right", fill="y")
        self.p_listbox = tk.Listbox(lf, font=("Segoe UI", 11), selectbackground=t["accent"],
            selectforeground="white", relief="solid", borderwidth=1,
            bg=t["list_bg"], fg=t["fg"], yscrollcommand=sb.set)
        self.p_listbox.pack(fill="both", expand=True); sb.config(command=self.p_listbox.yview)
        self._refresh_platforms()
        bf = tk.Frame(pf, bg=t["bg"]); bf.pack(fill="x", padx=15, pady=10)
        for text, cmd, color in [("Add Platform", self._add, t["accent"]),
                                  ("Edit", self._edit, "#555555"),
                                  ("Remove", self._remove, t["danger"])]:
            tk.Button(bf, text=text, font=("Segoe UI", 10), bg=color, fg="white",
                      relief="flat", padx=10, pady=5, command=cmd).pack(side="left", padx=(0, 6))

        # === Templates tab ===
        tf_tab = tk.Frame(notebook, bg=t["bg"]); notebook.add(tf_tab, text="Templates")
        self._build_templates_tab(tf_tab, t)

        # === Appearance tab ===
        af = tk.Frame(notebook, bg=t["bg"]); notebook.add(af, text="Theme")
        tk.Label(af, text="Theme", font=("Segoe UI", 13, "bold"), bg=t["bg"], fg=t["fg"]).pack(pady=(20, 10))
        self.theme_var = tk.StringVar(value=theme_name)
        for label, val in [("Light Mode", "light"), ("Dark Mode", "dark")]:
            tk.Radiobutton(af, text=label, variable=self.theme_var, value=val,
                           font=("Segoe UI", 11), bg=t["bg"], fg=t["fg"],
                           selectcolor=t["input_bg"]).pack(anchor="w", padx=40, pady=3)

        # === Export tab ===
        ef = tk.Frame(notebook, bg=t["bg"]); notebook.add(ef, text="Export")
        tk.Label(ef, text="Export Options", font=("Segoe UI", 13, "bold"),
                 bg=t["bg"], fg=t["fg"]).pack(pady=(20, 10))
        tk.Label(ef, text="Default Output Folder:", font=("Segoe UI", 10),
                 bg=t["bg"], fg=t["fg"]).pack(anchor="w", padx=40, pady=(10, 3))
        out_frame = tk.Frame(ef, bg=t["bg"])
        out_frame.pack(fill="x", padx=40)
        self.output_folder_var = tk.StringVar(value=settings.get("output_folder", OUTPUT_DIR))
        tk.Entry(out_frame, textvariable=self.output_folder_var, font=("Segoe UI", 10),
                 bg=t["input_bg"], fg=t["input_fg"], relief="solid", borderwidth=1
                 ).pack(side="left", fill="x", expand=True)
        tk.Button(out_frame, text="Browse", font=("Segoe UI", 9), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=8,
                  command=lambda: self.output_folder_var.set(
                      filedialog.askdirectory(parent=self.window) or self.output_folder_var.get())
                  ).pack(side="right", padx=(5, 0))
        tk.Label(ef, text="Exports one workbook per client: the interactive "
                          "Search dashboard + Unified Data.",
                 font=("Segoe UI", 9), bg=t["bg"], fg=t["muted"]
                 ).pack(anchor="w", padx=40, pady=(15, 2))

        # === Debug tab ===
        self._build_debug_tab(notebook, t)

        tk.Button(self.window, text="Save & Close", font=("Segoe UI", 11, "bold"),
                  bg=t["accent"], fg="white", relief="flat", padx=20, pady=8,
                  command=self._save_and_close).grid(row=1, column=0, pady=12)

    # -- Templates tab --

    def _build_templates_tab(self, parent, t):
        """Template management: list, preview, rename, delete."""
        tk.Label(parent, text="Saved PowerPoint templates and their mappings",
                 font=("Segoe UI", 10), bg=t["bg"], fg=t["muted"]).pack(pady=(10, 5))
        content = tk.Frame(parent, bg=t["bg"])
        content.pack(fill="both", expand=True, padx=15)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=0)
        content.rowconfigure(0, weight=1)

        lb_frame = tk.Frame(content, bg=t["bg"])
        lb_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.tmpl_listbox = tk.Listbox(lb_frame, font=("Segoe UI", 10),
            selectbackground=t["accent"], selectforeground="white",
            relief="solid", borderwidth=1, bg=t["list_bg"], fg=t["fg"])
        self.tmpl_listbox.pack(fill="both", expand=True)
        self._refresh_templates()

        prev_frame = tk.LabelFrame(content, text="  Slide 1  ",
                                    font=("Segoe UI", 9, "bold"),
                                    bg=t["card"], fg=t["card_fg"],
                                    width=220, height=170)
        prev_frame.grid(row=0, column=1, sticky="n")
        prev_frame.pack_propagate(False)
        self._tmpl_preview_label = tk.Label(prev_frame, bg="#333333",
            text="Select a\ntemplate", fg="#888888", font=("Segoe UI", 8))
        self._tmpl_preview_label.pack(fill="both", expand=True)
        self._tmpl_thumb_image = None  # prevent GC
        self._tmpl_preview_token = 0
        self.tmpl_listbox.bind("<<ListboxSelect>>", self._on_template_select)
        if self._templates:
            self.tmpl_listbox.selection_set(0)
            self.tmpl_listbox.activate(0)
            self.window.after_idle(self._on_template_select)

        btn = tk.Frame(parent, bg=t["bg"]); btn.pack(fill="x", padx=15, pady=8)
        tk.Button(btn, text="Edit Mapping", font=("Segoe UI", 9), bg=t["accent"], fg="white",
                  relief="flat", padx=10, pady=4,
                  command=self._edit_template).pack(side="left", padx=(0, 5))
        tk.Button(btn, text="Rename", font=("Segoe UI", 9), bg="#555555", fg="white",
                  relief="flat", padx=10, pady=4,
                  command=self._rename_template).pack(side="left", padx=(0, 5))
        tk.Button(btn, text="Delete", font=("Segoe UI", 9), bg=t["danger"], fg="white",
                  relief="flat", padx=10, pady=4,
                  command=self._delete_template).pack(side="left")
        tk.Button(btn, text="Export…", font=("Segoe UI", 9), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=10, pady=4,
                  command=self._export_template).pack(side="left", padx=(14, 5))
        tk.Button(btn, text="Import…", font=("Segoe UI", 9), bg=t["success"],
                  fg="white", relief="flat", padx=10, pady=4,
                  command=self._import_template).pack(side="left")

    def _refresh_templates(self):
        self.tmpl_listbox.delete(0, tk.END)
        from engine.pptx_mapper import list_available_templates
        self._templates = list_available_templates()
        for tmpl in self._templates:
            status = "mapped" if tmpl["has_mapping"] else "unmapped"
            self.tmpl_listbox.insert(tk.END, f"{tmpl['filename']}  ({status})")
        if not self._templates:
            self.tmpl_listbox.insert(tk.END, "No templates saved yet")
            if hasattr(self, "_tmpl_preview_label"):
                self._tmpl_thumb_image = None
                self._tmpl_preview_label.config(
                    image="", text="No templates\nsaved yet", fg="#888888")
        elif hasattr(self, "_tmpl_preview_label"):
            self.tmpl_listbox.selection_set(0)
            self.tmpl_listbox.activate(0)
            self.window.after_idle(self._on_template_select)

    def _on_template_select(self, event=None):
        sel = self.tmpl_listbox.curselection()
        if not sel or sel[0] >= len(self._templates):
            return
        tmpl = self._templates[sel[0]]
        self._tmpl_preview_token += 1
        token = self._tmpl_preview_token
        self._tmpl_thumb_image = None
        self._tmpl_preview_label.config(
            image="", text="Loading preview…", fg="#B8B8B8",
            font=("Segoe UI", 8), justify="center")

        def work():
            from engine.pptx_thumbs import get_template_preview
            return get_template_preview(tmpl["path"])

        def success(preview):
            if token != self._tmpl_preview_token or not self.window.winfo_exists():
                return
            if preview.get("kind") == "image":
                from PIL import Image, ImageTk
                with Image.open(preview["value"]) as source:
                    img = source.copy()
                img.thumbnail((210, 158), Image.LANCZOS)
                self._tmpl_thumb_image = ImageTk.PhotoImage(img)
                self._tmpl_preview_label.config(
                    image=self._tmpl_thumb_image, text="")
            else:
                self._tmpl_thumb_image = None
                self._tmpl_preview_label.config(
                    text=preview.get("value") or tmpl["filename"], image="",
                    fg="#B8B8B8", font=("Segoe UI", 8), justify="left",
                    wraplength=200)

        def error(exc):
            logger.debug(
                "Thumbnail load failed for %s: %s", tmpl["filename"], exc,
                exc_info=(type(exc), exc, exc.__traceback__))
            if token == self._tmpl_preview_token and self.window.winfo_exists():
                self._tmpl_preview_label.config(
                    text=tmpl["filename"], image="", fg="#B8B8B8")

        run_in_background(self.window, work, success, error)

    def _export_template(self):
        """Package the selected mapped template (pptx + mapping + images)
        into a single shareable .zip for teammates."""
        from tkinter import filedialog, messagebox
        sel = self.tmpl_listbox.curselection()
        if not sel or sel[0] >= len(self._templates):
            messagebox.showinfo("Export Template",
                                "Select a template to export first.",
                                parent=self.window)
            return
        tmpl = self._templates[sel[0]]
        if not tmpl.get("has_mapping"):
            messagebox.showwarning("Export Template",
                                   "That template has no saved mapping yet — "
                                   "map it before exporting.",
                                   parent=self.window)
            return
        stem = os.path.splitext(tmpl["filename"])[0]
        dest = filedialog.asksaveasfilename(
            parent=self.window, title="Export template bundle",
            defaultextension=".zip",
            initialfile=f"{stem}_template.zip",
            filetypes=[("Template bundle (zip)", "*.zip")])
        if not dest:
            return
        try:
            from engine.template_bundle import export_template_bundle
            n = export_template_bundle(tmpl["filename"], dest)
            messagebox.showinfo(
                "Export Template",
                f"Exported '{tmpl['filename']}' with its mapping"
                + (f" and {n} image(s)" if n else "")
                + f".\n\nSend the file to a teammate — they import it "
                  f"from this same screen.",
                parent=self.window)
        except Exception as e:
            messagebox.showerror("Export Template", str(e), parent=self.window)

    def _import_template(self):
        """Install a teammate's template bundle: one click, ready to fill."""
        from tkinter import filedialog, messagebox
        path = filedialog.askopenfilename(
            parent=self.window, title="Import template bundle",
            filetypes=[("Template bundle (zip)", "*.zip")])
        if not path:
            return
        from engine.template_bundle import import_template_bundle
        try:
            try:
                name, n = import_template_bundle(path)
            except FileExistsError as fe:
                if not messagebox.askyesno(
                        "Import Template",
                        f"'{fe}' already exists here.\nReplace it and its "
                        f"mapping?", parent=self.window):
                    return
                name, n = import_template_bundle(path, overwrite=True)
            self._refresh_templates()
            messagebox.showinfo(
                "Import Template",
                f"Imported '{name}'"
                + (f" with {n} image(s)" if n else "")
                + " — mapped and ready to fill.",
                parent=self.window)
        except Exception as e:
            messagebox.showerror("Import Template", str(e), parent=self.window)

    def _edit_template(self):
        """Open the template mapper on the selected template with its saved
        mapping preloaded, so assignments can be fixed or added."""
        sel = self.tmpl_listbox.curselection()
        if not sel or sel[0] >= len(self._templates):
            return
        tmpl = self._templates[sel[0]]
        from mapper.mapper_window import PPTXWizard
        # No client data in the settings context — the mapper opens with
        # Quick Fill + query tools only; the saved mapping loads normally.
        dummy_export = {"client_name": "", "client_data": [], "folder": ""}
        PPTXWizard(self.window, dummy_export, self.theme,
                   template_path=tmpl["path"])

    def _rename_template(self):
        sel = self.tmpl_listbox.curselection()
        if not sel or sel[0] >= len(self._templates):
            return
        tmpl = self._templates[sel[0]]
        old_name = tmpl["filename"]
        t = self.theme
        win = tk.Toplevel(self.window)
        win.title("Rename Template"); fit_window(win, 350, 130)
        win.configure(bg=t["bg"]); win.transient(self.window)
        win.grab_set(); win.lift(); win.focus_force()
        tk.Label(win, text=f"Rename: {old_name}", font=("Segoe UI", 10),
                 bg=t["bg"], fg=t["fg"]).pack(pady=(15, 5))
        name_var = tk.StringVar(value=os.path.splitext(old_name)[0])
        entry = tk.Entry(win, textvariable=name_var, font=("Segoe UI", 11),
                         bg=t["input_bg"], fg=t["input_fg"], relief="solid", borderwidth=1)
        entry.pack(fill="x", padx=20, pady=5)
        entry.select_range(0, tk.END); entry.focus_set()

        def do_rename():
            new_base = name_var.get().strip()
            if not new_base:
                return
            raw_name = new_base if new_base.lower().endswith(".pptx") else new_base + ".pptx"
            new_name = safe_component(raw_name, fallback="template.pptx")
            if not new_name.lower().endswith(".pptx"):
                new_name += ".pptx"
            if new_name == old_name:
                win.destroy(); return
            old_path = tmpl["path"]
            new_path = os.path.join(TEMPLATES_DIR, new_name)
            if os.path.exists(new_path):
                messagebox.showerror(
                    "Rename Template",
                    f"A template named '{new_name}' already exists.",
                    parent=win)
                return
            try:
                os.rename(old_path, new_path)
                from engine.pptx_mapper import _mapping_path
                old_map = _mapping_path(old_name)
                new_map = _mapping_path(new_name)
                if os.path.exists(old_map):
                    os.rename(old_map, new_map)
                # Move the cached thumbnail with the template
                from engine.pptx_thumbs import _thumb_path
                old_thumb = _thumb_path(old_name)
                if os.path.exists(old_thumb):
                    os.rename(old_thumb, _thumb_path(new_name))
                win.destroy()
                self._refresh_templates()
            except OSError as e:
                messagebox.showerror("Error", f"Rename failed:\n{e}", parent=win)

        tk.Button(win, text="Rename", font=("Segoe UI", 10, "bold"),
                  bg=t["accent"], fg="white", relief="flat", padx=15, pady=5,
                  command=do_rename).pack(pady=5)

    def _delete_template(self):
        sel = self.tmpl_listbox.curselection()
        if not sel or sel[0] >= len(self._templates):
            return
        tmpl = self._templates[sel[0]]
        if not messagebox.askyesno("Confirm",
                f"Delete template '{tmpl['filename']}'?\n"
                f"This will also remove its mapping.",
                parent=self.window):
            return
        try:
            os.remove(tmpl["path"])
            from engine.pptx_mapper import _mapping_path
            map_path = _mapping_path(tmpl["filename"])
            if os.path.exists(map_path):
                os.remove(map_path)
            from engine.pptx_thumbs import _thumb_path
            thumb = _thumb_path(tmpl["filename"])
            if os.path.exists(thumb):
                os.remove(thumb)
        except OSError as e:
            messagebox.showerror("Error", f"Delete failed:\n{e}", parent=self.window)
        self._refresh_templates()

    # -- Debug tab --

    def _build_debug_tab(self, notebook, t):
        df = tk.Frame(notebook, bg=t["bg"]); notebook.add(df, text="Logs")
        tk.Label(df, text="Application Logs", font=("Segoe UI", 13, "bold"),
                 bg=t["bg"], fg=t["fg"]).pack(pady=(15, 5))
        log_frame = tk.Frame(df, bg=t["bg"])
        log_frame.pack(fill="both", expand=True, padx=15, pady=(0, 5))
        log_sb = tk.Scrollbar(log_frame)
        log_sb.pack(side="right", fill="y")
        self.log_text = tk.Text(log_frame, font=("Consolas", 8), bg="#1a1a1a",
                                 fg="#00ff00", height=12, wrap="none",
                                 yscrollcommand=log_sb.set, state="disabled")
        self.log_text.pack(fill="both", expand=True)
        log_sb.config(command=self.log_text.yview)
        self._load_log_tail()
        dbtn = tk.Frame(df, bg=t["bg"]); dbtn.pack(fill="x", padx=15, pady=5)
        tk.Button(dbtn, text="Refresh", font=("Segoe UI", 9), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=10, pady=4,
                  command=self._load_log_tail).pack(side="left", padx=(0, 5))
        tk.Button(dbtn, text="Open Logs Folder", font=("Segoe UI", 9), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=10, pady=4,
                  command=self._open_logs_folder).pack(side="left", padx=(0, 5))
        tk.Button(dbtn, text="Clear Logs", font=("Segoe UI", 9), bg=t["danger"], fg="white",
                  relief="flat", padx=10, pady=4,
                  command=self._clear_logs).pack(side="left")
        from config.logging_setup import LOG_FILE
        tk.Label(df, text=f"Log file: {LOG_FILE}", font=("Segoe UI", 8),
                 bg=t["bg"], fg=t["muted"]).pack(padx=15, pady=(0, 5), anchor="w")

    def _load_log_tail(self):
        from config.logging_setup import LOG_FILE
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        try:
            if os.path.exists(LOG_FILE):
                with open(LOG_FILE, "r", encoding="utf-8", errors="replace") as f:
                    lines = f.readlines()
                tail = lines[-200:] if len(lines) > 200 else lines
                self.log_text.insert("1.0", "".join(tail))
                self.log_text.see(tk.END)
            else:
                self.log_text.insert("1.0", "(No log file yet)")
        except Exception as e:
            self.log_text.insert("1.0", f"Error reading log: {e}")
        self.log_text.config(state="disabled")

    def _open_logs_folder(self):
        try:
            os.startfile(LOGS_DIR)
        except AttributeError:
            import subprocess
            subprocess.Popen(["xdg-open", LOGS_DIR])
        except Exception:
            messagebox.showinfo("Logs Folder", f"Logs are in:\n{LOGS_DIR}",
                                parent=self.window)

    def _clear_logs(self):
        if not messagebox.askyesno("Confirm", "Clear all log files?",
                                    parent=self.window):
            return
        try:
            for f in os.listdir(LOGS_DIR):
                fp = os.path.join(LOGS_DIR, f)
                if os.path.isfile(fp):
                    open(fp, "w").close()
            self._load_log_tail()
        except Exception as e:
            messagebox.showerror("Error", f"Could not clear logs:\n{e}",
                                  parent=self.window)

    # -- Existing methods --

    def _refresh_platforms(self):
        self.p_listbox.delete(0, tk.END)
        self._platform_names = list(self.platforms.keys())
        for p in self._platform_names:
            cfg = load_platform_config(p)
            sheets = len(cfg.get("sheets", [])) if cfg else 0
            self.p_listbox.insert(tk.END, f"{p}  ({sheets} sheets)")
        if not self._platform_names:
            self.p_listbox.insert(tk.END, "No platforms configured — choose Add Platform")

    def _add(self):
        PlatformSetupWindow(self.window, self.theme, self._on_saved)

    def _edit(self):
        sel = self.p_listbox.curselection()
        if not sel or sel[0] >= len(self._platform_names): return
        name = self._platform_names[sel[0]]
        PlatformSetupWindow(self.window, self.theme, self._on_saved,
                            existing_name=name, existing_config=load_platform_config(name))

    def _on_saved(self, name, config):
        self.platforms[name] = {"configured": True}; self._refresh_platforms()

    def _remove(self):
        sel = self.p_listbox.curselection()
        if not sel or sel[0] >= len(self._platform_names): return
        name = self._platform_names[sel[0]]
        if messagebox.askyesno("Confirm", f"Remove '{name}'?", parent=self.window):
            del self.platforms[name]
            try:
                delete_platform_config(name)
            except OSError:
                logger.warning("Could not delete platform config for %s", name,
                               exc_info=True)
            self._refresh_platforms()

    def _save_and_close(self):
        settings = load_settings()
        settings["platforms"] = self.platforms
        settings["theme"] = self.theme_var.get()
        settings["output_folder"] = self.output_folder_var.get()
        # Purge settings retired with the Search dashboard rework
        for stale in ("dashboard_metrics", "export_client_report",
                      "export_dashboard", "export_unified"):
            settings.pop(stale, None)
        try:
            save_settings(settings)
        except OSError as e:
            logger.exception("Could not save settings")
            messagebox.showerror("Save Failed",
                f"Could not save settings:\n{e}", parent=self.window)
            return
        self.on_save(self.theme_var.get()); self.window.destroy()
