"""Review rendering: KPI cards, campaign detail, and report generation.

``ReviewMixin`` is mixed into ``IngestionEngine`` (ui/main_window.py); it
renders the per-client review inside the main window and hands off to the
PowerPoint template selector/mapper.
"""

import logging
import os
import tkinter as tk
from ui.utils import (brand_header, bordered_card, fit_window, modern_button,
                      run_in_background)
from tkinter import ttk, filedialog, messagebox

from engine.kpi import KPI_METRICS, compute_kpis

logger = logging.getLogger(__name__)


class ReviewMixin:
    """Review-screen methods for IngestionEngine."""

    def _show_review_in_main(self, export_result, current_num, total_num):
        """Show review data directly in the main window."""
        t = self.t

        # Maximize the main window
        try:
            self.root.state('zoomed')
        except tk.TclError:
            logger.debug("Window zoom not supported", exc_info=True)

        # Clear main content
        for w in self.root.winfo_children(): w.destroy()

        brand_header(
            self.root, t, "Review & validate",
            f"{export_result['client_name']}  •  Client {current_num}/{total_num}  •  "
            f"{len(export_result['campaigns'])} campaigns",
            step=3)

        main = tk.Frame(self.root, bg=t["bg"])
        main.pack(fill="both", expand=True)

        # Bottom action bar
        bottom = tk.Frame(main, bg=t["bg"])
        bottom.pack(side="bottom", fill="x", padx=15, pady=8)

        modern_button(
            bottom, "Generate report", lambda: self._generate_report(export_result),
            t, kind="success", font_size=12, padx=20, pady=9).pack(side="left")

        if current_num < total_num:
            modern_button(
                bottom, "Next client  →", lambda: self._on_review_action("next"),
                t, kind="primary", font_size=11, padx=20, pady=9).pack(side="right")

        modern_button(
            bottom, "Finish", lambda: self._on_review_action("finish"),
            t, kind="secondary", padx=15, pady=8).pack(side="right", padx=(0, 10))

        modern_button(
            bottom, "← Back to import", self._build_ui, t, kind="secondary",
            padx=15, pady=8).pack(side="right", padx=(0, 10))

        # Scrollable content
        content_canvas = tk.Canvas(main, bg=t["bg"], highlightthickness=0)
        content_sb = ttk.Scrollbar(main, orient="vertical", command=content_canvas.yview)
        content_frame = tk.Frame(content_canvas, bg=t["bg"])
        content_frame.bind("<Configure>", lambda e: content_canvas.configure(scrollregion=content_canvas.bbox("all")))
        _cw = content_canvas.create_window((0, 0), window=content_frame, anchor="nw")
        content_canvas.bind(
            "<Configure>",
            lambda e: content_canvas.itemconfigure(_cw, width=e.width))
        content_canvas.configure(yscrollcommand=content_sb.set)
        content_canvas.pack(side="left", fill="both", expand=True)
        content_sb.pack(side="right", fill="y")

        # KPI values are precomputed in the background export pass
        # (main_window._on_wizard_complete); computing them here froze the
        # UI for large clients. Fallback stays for any caller without them.
        kpi_totals, campaign_details, flags = (
            export_result.get("kpi")
            or compute_kpis(export_result["client_data"]))

        # Which campaigns do the flags implicate? (mixed shapes: tuples
        # (campaign, metric, message) and plain strings)
        _flagged = set()
        for _f in flags:
            if isinstance(_f, (list, tuple)) and _f:
                _flagged.add(str(_f[0]))
            else:
                for _c in campaign_details:
                    if _c in str(_f):
                        _flagged.add(_c)

        # KPI Cards — show EVERY computed metric (KPI_METRICS is just a
        # preferred ordering, not a filter). New metric types from any
        # platform appear automatically.
        kpi_frame = tk.Frame(content_frame, bg=t["bg"])
        kpi_frame.pack(fill="x", padx=15, pady=10)
        preferred = [m for m in KPI_METRICS if kpi_totals.get(m) not in (None, 0)]
        rest = sorted(m for m, v in kpi_totals.items()
                      if m not in KPI_METRICS
                      and isinstance(v, (int, float)) and v != 0)
        shown_kpis = preferred + rest
        PER_ROW = 4  # wraps before running off narrower screens
        for i, metric in enumerate(shown_kpis):
            if i % PER_ROW == 0:
                kpi_row = tk.Frame(kpi_frame, bg=t["bg"])
                kpi_row.pack(fill="x", pady=2)
            val = kpi_totals.get(metric, 0)
            card = bordered_card(kpi_row, t, padx=12, pady=8)
            card.pack(side="left", padx=4, fill="x", expand=True)
            tk.Label(card, text=metric, font=("Segoe UI", 9), bg=t["card"], fg=t["muted"]).pack()
            if isinstance(val, float) and val != int(val):
                display = f"${val:,.2f}" if "cost" in metric.lower() else f"{val:,.2f}"
            elif isinstance(val, (int, float)):
                display = f"${int(val):,}" if "cost" in metric.lower() else f"{int(val):,}"
            else: display = str(val)
            tk.Label(card, text=display, font=("Segoe UI", 16, "bold"), bg=t["card"], fg=t["fg"]).pack()

        # Flags
        if flags:
            def _flag_lines():
                out = []
                for _f in flags:
                    if isinstance(_f, (list, tuple)):
                        _p = [str(x) for x in _f]
                        out.append("⚠  " + " — ".join(_p[:2]) + (": " + " ".join(_p[2:]) if len(_p) > 2 else ""))
                    else:
                        out.append("⚠  " + str(_f))
                return out

            def _show_flags():
                fw = tk.Toplevel(self.root)
                fw.title("Data Flags")
                fit_window(fw, 620, 380)
                fw.configure(bg=t["bg"]); fw.transient(self.root); fw.grab_set()
                tk.Label(fw, text=f"⚠  {len(flags)} data issue(s) found",
                         font=("Segoe UI", 13, "bold"), bg=t["bg"],
                         fg=t["warning"]).pack(anchor="w", padx=16, pady=(14, 6))
                body = tk.Frame(fw, bg=t["card"], relief="solid", borderwidth=1)
                body.pack(fill="both", expand=True, padx=16, pady=(0, 8))
                for _line in _flag_lines():
                    tk.Label(body, text=_line, font=("Segoe UI", 10), bg=t["card"],
                             fg=t["fg"], wraplength=540, justify="left",
                             anchor="w").pack(fill="x", padx=12, pady=6)
                tk.Label(fw, text="Affected campaigns are outlined in amber below.",
                         font=("Segoe UI", 9), bg=t["bg"], fg=t["muted"]).pack(anchor="w", padx=16)
                tk.Button(fw, text="Close", font=("Segoe UI", 10), bg=t["secondary"],
                          fg=t["secondary_fg"], relief="flat", padx=16, pady=4,
                          command=fw.destroy).pack(pady=(4, 12))

            _flags_btn = tk.Button(content_frame,
                      text=f"⚠ {len(flags)} data flag(s) — click for details",
                      font=("Segoe UI", 9, "bold"), bg=t["pill_warn_bg"],
                      fg=t["pill_warn_fg"], activebackground=t["pill_warn_bg"],
                      relief="flat", padx=15, pady=4, cursor="hand2",
                      command=_show_flags)
            _flags_btn.pack(fill="x", padx=15, pady=(0, 5))
            _tip = {}
            def _tip_show(_e):
                _t = tk.Toplevel(self.root); _t.overrideredirect(True)
                _t.geometry(f"+{_e.x_root + 12}+{_e.y_root + 14}")
                _first = _flag_lines()[0] if flags else ""
                tk.Label(_t, text=_first[:90] + ("…" if len(_first) > 90 else "") +
                         "\nClick for full details", font=("Segoe UI", 9),
                         bg="#332B00", fg="#FFE9B8", justify="left", padx=8,
                         pady=5).pack()
                _tip["w"] = _t
            def _tip_hide(_e):
                if _tip.get("w") is not None:
                    _tip["w"].destroy(); _tip["w"] = None
            _flags_btn.bind("<Enter>", _tip_show)
            _flags_btn.bind("<Leave>", _tip_hide)

        # Campaign Detail
        tk.Label(content_frame, text="Campaign Detail", font=("Segoe UI", 12, "bold"),
                 bg=t["bg"], fg=t["fg"]).pack(anchor="w", padx=15, pady=(5, 3))

        for camp_name in sorted(campaign_details.keys()):
            metrics = campaign_details[camp_name]
            cf = tk.Frame(content_frame, bg=t["card"], relief="solid", borderwidth=1)
            if camp_name in _flagged:
                cf.configure(highlightbackground=t["warning"],
                             highlightthickness=2, relief="flat", borderwidth=0)
            cf.pack(fill="x", padx=15, pady=2)
            ch = tk.Frame(cf, bg=t["card"], cursor="hand2")
            ch.pack(fill="x")
            summary = []
            camp_order = ([m for m in KPI_METRICS if m in metrics] +
                          sorted(m for m in metrics if m not in KPI_METRICS))
            for km in camp_order[:3]:
                v = metrics.get(km)
                if v is not None and isinstance(v, (int, float)):
                    vfmt = f"{v:,.2f}" if isinstance(v, float) and v != int(v) else f"{int(v):,}"
                    summary.append(f"{km}: {vfmt}")
            tk.Label(ch, text=f"▸  {camp_name}", font=("Segoe UI", 10, "bold"),
                     bg=t["card"], fg=t["fg"]).pack(side="left", padx=10, pady=5)
            if summary:
                tk.Label(ch, text="  |  ".join(summary), font=("Segoe UI", 9),
                         bg=t["card"], fg=t["muted"]).pack(side="right", padx=10)
            detail = tk.Frame(cf, bg=t["bg"])

            # Detail rows build lazily on first expand: creating them up
            # front for every campaign meant thousands of hidden widgets on
            # large imports and a visibly slow review screen.
            def _build_detail(d=detail, mets=metrics):
                for met, val in sorted(mets.items()):
                    row = tk.Frame(d, bg=t["bg"])
                    row.pack(fill="x", pady=1)
                    tk.Label(row, text=met, font=("Segoe UI", 9), bg=t["bg"],
                             fg=t["muted"], width=22, anchor="w").pack(side="left")
                    if isinstance(val, float) and val == int(val): vd = f"{int(val):,}"
                    elif isinstance(val, float): vd = f"{val:,.2f}"
                    elif isinstance(val, int): vd = f"{val:,}"
                    else: vd = str(val)
                    tk.Label(row, text=vd, font=("Segoe UI", 9, "bold"), bg=t["bg"],
                             fg=t["fg"]).pack(side="left")

            def toggle(event, d=detail, h=ch, build=_build_detail):
                if d.winfo_manager():
                    d.pack_forget()
                    for w in h.winfo_children():
                        if isinstance(w, tk.Label) and "▾" in str(w.cget("text")):
                            w.config(text=w.cget("text").replace("▾", "▸")); break
                else:
                    if not d.winfo_children():
                        build()
                    d.pack(fill="x", padx=15, pady=(0, 5))
                    for w in h.winfo_children():
                        if isinstance(w, tk.Label) and "▸" in str(w.cget("text")):
                            w.config(text=w.cget("text").replace("▸", "▾")); break
            ch.bind("<Button-1>", toggle)
            for child in ch.winfo_children(): child.bind("<Button-1>", toggle)



    def _on_review_action(self, action):
        """Handle review action: 'next' or 'finish'."""
        if action == "next":
            self._review_index += 1
            self._show_next_review()
        elif action == "finish":
            self._show_final_summary()

    def _show_final_summary(self):
        total_clients = len(self._export_results)
        total_rows = sum(r["rows"] for r in self._export_results)
        folders = set(os.path.dirname(r["folder"]) for r in self._export_results)
        output_root = list(folders)[0] if folders else ""

        # Rebuild the main import screen first — the review screen destroyed
        # all widgets including status_label, so we must recreate it before
        # referencing it.
        self._build_ui()
        self.status_label.config(text=f"Complete — {total_clients} client(s), {total_rows} rows")

        logger.info("Export complete: %d client(s), %d rows -> %s",
                    total_clients, total_rows, output_root)
        messagebox.showinfo("Export Complete",
            f"✅ {total_clients} client(s) exported\n"
            f"📊 {total_rows} total data rows\n"
            f"📁 Output: {output_root}")

    def _generate_report(self, export_result):
        """Smart template selector: if templates exist, show dropdown. Otherwise open mapper."""
        from engine.pptx_mapper import (fill_template, get_available_metrics,
                                        list_available_templates,
                                        load_template_mapping)
        from mapper.mapper_window import PPTXWizard

        templates = list_available_templates()
        mapped_templates = [t for t in templates if t["has_mapping"]]

        if not mapped_templates:
            # No templates saved — open mapper to build one
            PPTXWizard(self.root, export_result, self.t)
            return

        # Templates exist — show quick selector with thumbnail preview
        t = self.t
        win = tk.Toplevel(self.root)
        win.title("Select Template")
        fit_window(win, 600, 420)
        win.configure(bg=t["bg"]); win.transient(self.root)
        win.grab_set(); win.lift(); win.focus_force()

        win.columnconfigure(0, weight=1)
        win.rowconfigure(2, weight=1)

        tk.Label(win, text="📑 Select a Template", font=("Segoe UI", 14, "bold"),
                 bg=t["bg"], fg=t["fg"]).grid(row=0, column=0, pady=(15, 5))
        tk.Label(win, text="Choose a pre-mapped template to auto-fill, or create a new one.",
                 font=("Segoe UI", 10), bg=t["bg"], fg=t["muted"]).grid(
                     row=1, column=0, pady=(0, 8))

        # Main content: listbox left, thumbnail right
        content = tk.Frame(win, bg=t["bg"])
        content.grid(row=2, column=0, sticky="nsew", padx=20)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=0)
        content.rowconfigure(0, weight=1)

        # Template listbox
        lb_frame = tk.Frame(content, bg=t["bg"])
        lb_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        lb = tk.Listbox(lb_frame, font=("Segoe UI", 11), bg=t["input_bg"], fg=t["input_fg"],
                         selectbackground=t["accent"], selectforeground="white",
                         relief="solid", borderwidth=1)
        lb.pack(fill="both", expand=True)
        for tmpl in mapped_templates:
            lb.insert("end", tmpl["filename"])
        if mapped_templates:
            lb.select_set(0)
            lb.activate(0)

        # Thumbnail preview panel
        thumb_frame = tk.LabelFrame(content, text="  Preview  ", font=("Segoe UI", 9, "bold"),
                                     bg=t["card"], fg=t["card_fg"], width=220, height=180)
        thumb_frame.grid(row=0, column=1, sticky="ns")
        thumb_frame.pack_propagate(False)
        thumb_label = tk.Label(thumb_frame, bg="#333333", text="Select a template",
                                fg="#888888", font=("Segoe UI", 9))
        thumb_label.pack(fill="both", expand=True)
        _thumb_ref = [None]  # prevent GC
        preview_token = [0]

        def _show_thumbnail(event=None):
            sel = lb.curselection()
            if not sel:
                return
            tmpl = mapped_templates[sel[0]]
            preview_token[0] += 1
            token = preview_token[0]
            _thumb_ref[0] = None
            thumb_label.config(
                image="", text="Loading preview…", fg="#B8B8B8",
                font=("Segoe UI", 8), justify="center")

            def work():
                from engine.pptx_thumbs import get_template_preview
                return get_template_preview(tmpl["path"])

            def success(preview):
                if token != preview_token[0] or not win.winfo_exists():
                    return
                if preview.get("kind") == "image":
                    from PIL import Image, ImageTk
                    with Image.open(preview["value"]) as source:
                        img = source.copy()
                    img.thumbnail((210, 158), Image.LANCZOS)
                    _thumb_ref[0] = ImageTk.PhotoImage(img)
                    thumb_label.config(image=_thumb_ref[0], text="")
                else:
                    _thumb_ref[0] = None
                    thumb_label.config(
                        text=preview.get("value") or tmpl["filename"], image="",
                        fg="#B8B8B8", font=("Segoe UI", 8), justify="left",
                        wraplength=200)

            def error(exc):
                logger.debug(
                    "Thumbnail load failed for %s: %s", tmpl["filename"], exc,
                    exc_info=(type(exc), exc, exc.__traceback__))
                if token == preview_token[0] and win.winfo_exists():
                    thumb_label.config(
                        text=tmpl["filename"], image="", fg="#B8B8B8")

            run_in_background(win, work, success, error)

        lb.bind("<<ListboxSelect>>", _show_thumbnail)
        # Show first template info
        if mapped_templates:
            win.after_idle(_show_thumbnail)

        btn_frame = tk.Frame(win, bg=t["bg"])
        btn_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=12)

        def auto_fill():
            sel = lb.curselection()
            if not sel: return
            tmpl = mapped_templates[sel[0]]
            template_path = tmpl["path"]
            mapping = load_template_mapping(tmpl["filename"])
            if not mapping:
                messagebox.showerror("Error", "Mapping not found.", parent=win); return

            # Get client info
            client_name = export_result.get("client_name", "Report")
            start_date = ""
            end_date = ""
            for d in export_result.get("client_data", []):
                if d.get("start_date"): start_date = d["start_date"]
                if d.get("end_date"): end_date = d["end_date"]
                if start_date: break

            # Build metrics with all meta-keys the fill engine expects
            metrics, _ = get_available_metrics(
                export_result["client_data"], client_name, start_date, end_date)
            metrics["__client_data__"] = export_result["client_data"]
            metrics["__client_name__"] = client_name
            metrics["__start_date__"] = start_date
            metrics["__end_date__"] = end_date

            # Ask for save location
            save_path = filedialog.asksaveasfilename(
                title="Save Filled Report",
                parent=win,
                initialdir=export_result.get("folder", ""),
                initialfile=f"{client_name}_Report.pptx",
                defaultextension=".pptx",
                filetypes=[("PowerPoint", "*.pptx")])
            if not save_path:
                return

            try:
                from engine.fill_report import append_fill_history
                from engine.pptx_fill import fill_template_report
                _, report = fill_template_report(template_path, save_path,
                                                 mapping, metrics)
                append_fill_history(report)
                win.destroy()
                status = "✅ Report generated!" if report.ok else \
                    "⚠️ Report generated with gaps:"
                open_it = messagebox.askyesno("Report Generated",
                    f"{status}\n\n{save_path}\n\n{report.summary()}\n\n"
                    f"Would you like to open it?")
                if open_it:
                    try:
                        os.startfile(save_path)
                    except AttributeError:
                        import subprocess
                        subprocess.Popen(["xdg-open", save_path])
                    except Exception:
                        logger.debug("Could not open file: %s", save_path, exc_info=True)
            except Exception as e:
                logger.exception("Template auto-fill failed: %s", template_path)
                messagebox.showerror("Error",
                    f"Failed to generate the report:\n{e}", parent=win)

        tk.Button(btn_frame, text="⚡ Auto-Fill Report", font=("Segoe UI", 11, "bold"),
                  bg="#1E6E3E", fg="white", relief="flat", padx=20, pady=8,
                  command=auto_fill).pack(side="left")
        tk.Button(btn_frame, text="New Template", font=("Segoe UI", 10),
                  bg=t["accent"], fg="white", relief="flat", padx=15, pady=8,
                  command=lambda: [win.destroy(), PPTXWizard(self.root, export_result, self.t)]
                  ).pack(side="left", padx=(10, 0))
        tk.Button(btn_frame, text="Cancel", font=("Segoe UI", 10),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat", padx=15, pady=8,
                  command=win.destroy).pack(side="right")
