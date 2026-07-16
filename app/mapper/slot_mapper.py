"""Slot mapping GUI (template-first Phase B).

Maps data sources onto a template's named slots and writes mapping.json
into the template store. Sources are the SAME query objects the current
mapper uses (owner decision 2026-07-16): simple options built from the
client's data plus Advanced Query Builder queries — the builder window is
reused as-is (this window quacks like the wizard it expects).

With client data present (opened from a report flow), each row shows the
live resolved value and "Build Report…" produces the month's deck via
engine.template_ir.mapping.build_mapped_report.
"""

import logging
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from engine.query_resolver import build_simple_options
from engine.template_ir import (load_slot_mapping, load_template_ir,
                                new_slot_mapping, save_slot_mapping,
                                validate_slot_mapping)
from engine.template_ir.mapping import build_mapped_report, resolve_slot_values
from mapper.query_builder import show_query_builder
from ui.utils import fit_window

logger = logging.getLogger(__name__)

FORMATS = ["text", "number", "currency", "percentage", "date"]


class SlotMapperWindow:
    """One row per named slot: source picker, format, live preview."""

    def __init__(self, parent, theme, template_dir, export_result=None):
        self.t = theme
        self.template_dir = template_dir
        self.export_result = export_result or {"client_name": "",
                                               "client_data": [], "folder": ""}
        self.ir = load_template_ir(template_dir)
        self.mapping = (load_slot_mapping(template_dir)
                        or new_slot_mapping(self.ir.template_id))

        client_data = self.export_result.get("client_data", [])
        self.client_name = self.export_result.get("client_name", "")
        self.start_date = ""
        self.end_date = ""
        for d in client_data:
            if d.get("start_date"): self.start_date = d["start_date"]
            if d.get("end_date"): self.end_date = d["end_date"]
            if self.start_date: break

        from engine.pptx_mapper import get_available_metrics
        self.available_metrics, self.structured_metrics = \
            get_available_metrics(client_data, self.client_name,
                                  self.start_date, self.end_date)
        options = [o for o in build_simple_options(self.structured_metrics)
                   if o["key"] != "__image__"]   # image slots are not Phase B
        self._option_by_label = {o["label"]: o for o in options}

        # Query-builder host contract (mapper.query_builder duck-types the
        # wizard): applied queries land in named_queries/_pending_query and
        # _refresh_metrics() is called — we attach them to the active row.
        self.named_queries = {}
        self.selected_metric = None
        self._pending_query = None
        self._pending_image = None
        self._advanced_target = None

        self._rows = {}
        self._build(parent)

    # ── Window ────────────────────────────────────────────────────────────

    def _build(self, parent):
        t = self.t
        self.window = tk.Toplevel(parent)
        self.window.title(f"Map Slots — {self.ir.template_name}")
        fit_window(self.window, 980, 620)
        self.window.configure(bg=t["bg"])
        self.window.grab_set(); self.window.lift(); self.window.focus_force()

        has_data = bool(self.export_result.get("client_data"))
        hint = ("Pick a data source for each slot."
                if has_data else
                "No client data in this session — special fields are "
                "available now; open from a report to map metrics and "
                "preview values.")
        tk.Label(self.window, text=hint, font=("Segoe UI", 9), bg=t["bg"],
                 fg=t["muted"], wraplength=920, justify="left"
                 ).pack(padx=15, pady=(10, 5), anchor="w")

        holder = tk.Frame(self.window, bg=t["card"])
        holder.pack(fill="both", expand=True, padx=15)
        canvas = tk.Canvas(holder, bg=t["card"], highlightthickness=0)
        sb = ttk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        rows_frame = tk.Frame(canvas, bg=t["card"])
        rows_frame.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        _cw = canvas.create_window((0, 0), window=rows_frame, anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(_cw, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        if not self.ir.slot_registry:
            tk.Label(rows_frame, text="This template has no dynamic slots "
                     "yet — mark shapes dynamic in Template Review first.",
                     font=("Segoe UI", 10), bg=t["card"], fg=t["muted"]
                     ).pack(pady=30)
        for slot, spec in self.ir.slot_registry.items():
            self._render_row(rows_frame, slot, spec)

        bottom = tk.Frame(self.window, bg=t["bg"])
        bottom.pack(fill="x", padx=15, pady=10)
        tk.Button(bottom, text="Save Mapping", font=("Segoe UI", 10),
                  bg=t["accent"], fg="white", relief="flat", padx=15, pady=6,
                  command=self._save).pack(side="right", padx=6)
        if has_data:
            tk.Button(bottom, text="Build Report…",
                      font=("Segoe UI", 10, "bold"), bg="#1E6E3E",
                      fg="white", relief="flat", padx=15, pady=6,
                      command=self._build_report).pack(side="right")
        tk.Button(bottom, text="Close", font=("Segoe UI", 10),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
                  padx=15, pady=6,
                  command=self.window.destroy).pack(side="left")

    def _render_row(self, parent, slot, spec):
        t = self.t
        row = tk.Frame(parent, bg=t["card"])
        row.pack(fill="x", padx=8, pady=3)

        name_box = tk.Frame(row, bg=t["card"], width=240)
        name_box.pack(side="left")
        name_box.pack_propagate(False)
        name_box.configure(height=38)
        tk.Label(name_box, text=slot, font=("Segoe UI", 10, "bold"),
                 bg=t["card"], fg=t["card_fg"], anchor="w").pack(fill="x")
        tk.Label(name_box,
                 text=f"{spec['type']} · slide {spec['slide_index'] + 1} · "
                      f"“{spec['placeholder_text'][:28]}”",
                 font=("Segoe UI", 8), bg=t["card"], fg=t["muted"],
                 anchor="w").pack(fill="x")

        is_frame_slot = spec["type"] in ("chart_data", "table_data")
        source_var = tk.StringVar()
        combo = ttk.Combobox(row, textvariable=source_var, width=34,
                             state="disabled" if is_frame_slot
                             else "readonly",
                             values=[] if is_frame_slot
                             else list(self._option_by_label))
        if is_frame_slot:
            apply_as = ("Chart Data" if spec["type"] == "chart_data"
                        else "Table")
            source_var.set(f"(use Advanced… → Apply as {apply_as})")
        combo.pack(side="left", padx=6)

        format_var = tk.StringVar(value=spec["type"]
                                  if spec["type"] in FORMATS else "text")
        fmt = ttk.Combobox(row, textvariable=format_var, width=10,
                           state="disabled" if is_frame_slot
                           else "readonly", values=FORMATS)
        fmt.pack(side="left", padx=6)

        tk.Button(row, text="Advanced…", font=("Segoe UI", 8),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
                  padx=8, pady=2,
                  command=lambda: self._open_advanced(slot)
                  ).pack(side="left", padx=(0, 6))

        preview = tk.Label(row, text="", font=("Segoe UI", 9, "italic"),
                           bg=t["card"], fg="#2E8B57", anchor="w")
        preview.pack(side="left", fill="x", expand=True)

        state = {"source_var": source_var, "format_var": format_var,
                 "combo": combo, "preview": preview, "query": None}
        self._rows[slot] = state

        saved = self.mapping.get("slots", {}).get(slot)
        if saved and saved.get("query") is not None:
            state["query"] = saved["query"]
            if saved.get("format") in FORMATS:
                format_var.set(saved["format"])
            source_var.set(self._label_for_query(saved["query"], slot))
        combo.bind("<<ComboboxSelected>>",
                   lambda e, s=slot: self._on_source_change(s))
        fmt.bind("<<ComboboxSelected>>",
                 lambda e, s=slot: self._update_preview(s))
        self._update_preview(slot)

    # ── Row behavior ──────────────────────────────────────────────────────

    def _label_for_query(self, query, slot):
        for label, option in self._option_by_label.items():
            if option.get("query", option["key"]) == query:
                return label
        for name, entry in self.named_queries.items():
            if entry.get("query") == query:
                return f"⚡ {name}"
        # Saved builder queries have no session name — label per slot so
        # two saved advanced queries can never shadow each other
        label = f"⚡ saved query ({slot})"
        self._option_by_label[label] = {"key": None, "query": query}
        return label

    def _on_source_change(self, slot):
        state = self._rows[slot]
        label = state["source_var"].get()
        option = self._option_by_label.get(label)
        if option is None and label.startswith("⚡ "):
            entry = self.named_queries.get(label[2:])
            option = {"key": None, "query": entry["query"]} if entry else None
        if option is None:
            return
        state["query"] = option.get("query") or option["key"]
        self._update_preview(slot)

    def _update_preview(self, slot):
        state = self._rows[slot]
        client_data = self.export_result.get("client_data", [])
        if state["query"] is None or not client_data:
            state["preview"].config(text="")
            return
        mapping = {"slots": {slot: {"query": state["query"],
                                    "format": state["format_var"].get()}}}
        values, issues = resolve_slot_values(
            self.ir, mapping, client_data, self.client_name,
            self.start_date, self.end_date)
        if slot in values:
            value = values[slot]
            if isinstance(value, dict) and "rows" in value:
                shown = (f"{len(value['rows'])} row(s) × "
                         f"{len(value.get('headers', []))} column(s)")
            elif isinstance(value, dict) and "categories" in value:
                shown = (f"{len(value['categories'])} categories, "
                         f"{len(value.get('series', []))} series")
            else:
                shown = value
            state["preview"].config(text=f"→ {shown}")
        else:
            state["preview"].config(text=f"⚠ {dict(issues).get(slot, '')}")

    def _open_advanced(self, slot):
        if not self.export_result.get("client_data"):
            messagebox.showinfo(
                "Advanced Query", "The query builder needs client data — "
                "open the slot mapper from a report.", parent=self.window)
            return
        self._advanced_target = slot
        show_query_builder(self)

    def _refresh_metrics(self):
        """Query-builder host hook: an applied query arrives here. Attach
        it to the row whose Advanced… button opened the builder."""
        slot, self._advanced_target = self._advanced_target, None
        query = self._pending_query
        self._pending_query = None
        if not slot or not query or slot not in self._rows:
            return
        state = self._rows[slot]
        state["query"] = dict(query)
        name = self.selected_metric or "advanced query"
        label = f"⚡ {name}"
        self._option_by_label[label] = {"key": None, "query": state["query"]}
        for other in self._rows.values():
            other["combo"].configure(values=list(self._option_by_label))
        state["source_var"].set(label)
        self._update_preview(slot)

    # ── Save / build ──────────────────────────────────────────────────────

    def _collect_mapping(self):
        slots = {}
        for slot, state in self._rows.items():
            if state["query"] is None:
                continue
            slots[slot] = {"query": state["query"],
                           "format": state["format_var"].get()}
        self.mapping["slots"] = slots
        return self.mapping

    def _save(self, quiet=False):
        mapping = self._collect_mapping()
        warnings = validate_slot_mapping(self.ir, mapping)
        try:
            save_slot_mapping(mapping, self.template_dir)
        except OSError as e:
            logger.exception("Could not save slot mapping")
            messagebox.showerror("Save Failed", str(e), parent=self.window)
            return False
        if warnings and not quiet:
            messagebox.showwarning(
                "Mapping Saved — With Warnings",
                "Saved, but check these before building:\n\n• "
                + "\n• ".join(warnings), parent=self.window)
        elif not quiet:
            messagebox.showinfo("Saved",
                                f"Slot mapping saved:\n{self.template_dir}",
                                parent=self.window)
        return True

    def _build_report(self):
        if not self._save(quiet=True):
            return
        client_name = self.client_name or "Report"
        save_path = filedialog.asksaveasfilename(
            title="Save Template-First Report", parent=self.window,
            initialdir=self.export_result.get("folder", ""),
            initialfile=f"{client_name}_Report.pptx",
            defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")])
        if not save_path:
            return
        try:
            _, report = build_mapped_report(
                self.template_dir, save_path,
                self.export_result.get("client_data", []),
                self.client_name, self.start_date, self.end_date)
        except Exception as e:
            logger.exception("Template-first build failed")
            messagebox.showerror("Build Failed", str(e), parent=self.window)
            return
        summary = (f"{report['shapes_copied']} shape(s) copied, "
                   f"{report['slots_filled']} slot(s) filled")
        problems = [f"{s}: {r}" for s, r in report["slots_unfilled"]]
        problems += [f"{s}: {r}" for s, r in report["skipped"]]
        if problems:
            messagebox.showwarning(
                "Report Built — With Gaps",
                f"{save_path}\n\n{summary}\n\nNeeds attention:\n• "
                + "\n• ".join(problems), parent=self.window)
        else:
            messagebox.showinfo("Report Built",
                                f"{save_path}\n\n{summary}",
                                parent=self.window)
