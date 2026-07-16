"""Slot mapping GUI (template-first) — the old mapper's assignment feel.

Reworked from the flat per-slot rows after Windows field feedback
(2026-07-16): sources live in a sidebar like the classic mapper — click
one to ARM it, then highlight any part of a shape's text and Assign.
The slot targets exactly the selected substring, so "CLIENT NAME |
MONTH 1st, 2026" can carry a client slot and a date slot side by side
instead of one slot replacing the whole box. No selection = whole box.

Sources are the SAME query objects as ever: simple options built from
the client's data plus Advanced Query Builder queries (the builder
window is reused as-is; this window quacks like the wizard it expects).
mapping.json and the slot registry are written together on Save; with
client data present, each slot shows its live resolved value and
"Build Report…" produces the month's deck.
"""

import logging
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from engine.query_resolver import build_simple_options
from engine.template_ir import (load_slot_mapping, load_template_ir,
                                new_slot_mapping, save_slot_mapping,
                                save_template_ir, validate_slot_mapping)
from engine.template_ir.classify import add_slot, remove_slot
from engine.template_ir.mapping import build_mapped_report, resolve_slot_values
from mapper.query_builder import show_query_builder
from ui.utils import fit_window

logger = logging.getLogger(__name__)

FORMATS = ["text", "number", "currency", "percentage", "date"]
SLOT_COLOR = "#D2691E"
ARMED_COLOR = "#1E6E3E"

_SECTION_TITLES = {"special": "Quick Fill", "total": "Totals",
                   "queries": "Saved Queries"}


def _source_kind(query):
    """(slot_type, default_format) a source produces."""
    if isinstance(query, dict):
        return "number", "number"
    if isinstance(query, str) and query.startswith("__total_"):
        return "number", "number"
    if isinstance(query, str) and query in (
            "__date_range__", "__start_date__", "__end_date__",
            "__start_month__", "__end_month__", "__year__"):
        return "date", "date"
    return "text", "text"


class SlotMapperWindow:
    """Arm a source, highlight text, assign — slots target the selection."""

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
        self.options = [o for o in
                        build_simple_options(self.structured_metrics)
                        if o["key"] != "__image__"]

        # Query-builder host contract (mapper.query_builder duck-types the
        # wizard): applied queries land in named_queries/_pending_query and
        # _refresh_metrics() is called.
        self.named_queries = {}
        self.selected_metric = None
        self._pending_query = None
        self._pending_image = None
        self._advanced_target = None   # frame-slot name, or None = arm it

        self.armed = None              # {"label", "query"}
        self._source_buttons = {}
        self.current_slide = 0
        self._build(parent)

    # ── Window ────────────────────────────────────────────────────────────

    def _build(self, parent):
        t = self.t
        self.window = tk.Toplevel(parent)
        self.window.title(f"Map Slots — {self.ir.template_name}")
        fit_window(self.window, 1150, 700)
        self.window.configure(bg=t["bg"])
        self.window.grab_set(); self.window.lift(); self.window.focus_force()

        hint = ("Click a data source to arm it, highlight the exact text "
                "it should replace, then Assign Selection. No selection = "
                "the whole box.")
        if not self.export_result.get("client_data"):
            hint += ("  (No client data in this session — special fields "
                     "work now; open from a report for metrics, previews, "
                     "and building.)")
        tk.Label(self.window, text=hint, font=("Segoe UI", 9), bg=t["bg"],
                 fg=t["muted"], wraplength=1100, justify="left"
                 ).pack(padx=12, pady=(10, 4), anchor="w")

        main = tk.Frame(self.window, bg=t["bg"])
        main.pack(fill="both", expand=True, padx=12)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        self._build_sidebar(main)
        self._build_shape_pane(main)

        bottom = tk.Frame(self.window, bg=t["bg"])
        bottom.pack(fill="x", padx=12, pady=10)
        tk.Button(bottom, text="Close", font=("Segoe UI", 10),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
                  padx=15, pady=6,
                  command=self.window.destroy).pack(side="left")
        tk.Button(bottom, text="Save Mapping", font=("Segoe UI", 10),
                  bg=t["accent"], fg="white", relief="flat", padx=15, pady=6,
                  command=self._save).pack(side="right", padx=6)
        if self.export_result.get("client_data"):
            tk.Button(bottom, text="Build Report…",
                      font=("Segoe UI", 10, "bold"), bg=ARMED_COLOR,
                      fg="white", relief="flat", padx=15, pady=6,
                      command=self._build_report).pack(side="right")

        self._render_slide()

    def _build_sidebar(self, main):
        t = self.t
        side = tk.LabelFrame(main, text="  Data Sources  ",
                             font=("Segoe UI", 10, "bold"), bg=t["card"],
                             fg=t["card_fg"], padx=6, pady=4)
        side.grid(row=0, column=0, sticky="ns", padx=(0, 8))

        tk.Button(side, text="🔍 Advanced Query Builder…",
                  font=("Segoe UI", 9, "bold"), bg="#8B4513", fg="white",
                  relief="flat", padx=8, pady=4,
                  command=self._open_advanced_source).pack(fill="x",
                                                           pady=(0, 4))
        self.source_search = tk.StringVar()
        self.source_search.trace_add("write",
                                     lambda *_: self._render_sources())
        tk.Entry(side, textvariable=self.source_search, width=26,
                 font=("Segoe UI", 9), bg=t["input_bg"], fg=t["input_fg"],
                 insertbackground=t["insert"], relief="solid",
                 borderwidth=1).pack(fill="x", pady=(0, 4))

        canvas = tk.Canvas(side, bg=t["card"], highlightthickness=0,
                           width=230)
        sb = ttk.Scrollbar(side, orient="vertical", command=canvas.yview)
        self.source_frame = tk.Frame(canvas, bg=t["card"])
        self.source_frame.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        _cw = canvas.create_window((0, 0), window=self.source_frame,
                                   anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(_cw, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._render_sources()

    def _build_shape_pane(self, main):
        t = self.t
        pane = tk.Frame(main, bg=t["bg"])
        pane.grid(row=0, column=1, sticky="nsew")
        pane.rowconfigure(1, weight=1)
        pane.columnconfigure(0, weight=1)

        nav = tk.Frame(pane, bg=t["bg"])
        nav.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        tk.Button(nav, text="← Prev", font=("Segoe UI", 9),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
                  padx=10, pady=3,
                  command=lambda: self._go_slide(-1)).pack(side="left")
        self.slide_label = tk.Label(nav, font=("Segoe UI", 10, "bold"),
                                    bg=t["bg"], fg=t["fg"])
        self.slide_label.pack(side="left", padx=10)
        tk.Button(nav, text="Next →", font=("Segoe UI", 9),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
                  padx=10, pady=3,
                  command=lambda: self._go_slide(1)).pack(side="left")
        self.armed_label = tk.Label(nav, text="No source armed",
                                    font=("Segoe UI", 9, "italic"),
                                    bg=t["bg"], fg=t["muted"])
        self.armed_label.pack(side="right")

        holder = tk.Frame(pane, bg=t["card"])
        holder.grid(row=1, column=0, sticky="nsew")
        self.canvas = tk.Canvas(holder, bg=t["card"], highlightthickness=0)
        sb = ttk.Scrollbar(holder, orient="vertical",
                           command=self.canvas.yview)
        self.cards = tk.Frame(self.canvas, bg=t["card"])
        self.cards.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        _cw = self.canvas.create_window((0, 0), window=self.cards,
                                        anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(
            _cw, width=e.width))
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    # ── Sources ───────────────────────────────────────────────────────────

    def _source_catalog(self):
        catalog = list(self.options)
        for name, entry in self.named_queries.items():
            catalog.append({"key": None, "label": f"⚡ {name}",
                            "category": "queries",
                            "query": entry.get("query")})
        return catalog

    def _render_sources(self):
        t = self.t
        for child in self.source_frame.winfo_children():
            child.destroy()
        self._source_buttons = {}
        needle = self.source_search.get().strip().casefold()
        last_section = None
        for option in self._source_catalog():
            if needle and needle not in option["label"].casefold():
                continue
            section = option.get("category", "")
            section = _SECTION_TITLES.get(
                section, section.replace("breakdown_", "").title()
                if section.startswith("breakdown_") else section)
            if section != last_section:
                tk.Label(self.source_frame, text=section,
                         font=("Segoe UI", 8, "bold"), bg=t["card"],
                         fg=t["muted"]).pack(anchor="w", pady=(6, 1))
                last_section = section
            label = option["label"]
            armed = self.armed and self.armed["label"] == label
            btn = tk.Button(
                self.source_frame, text=label, font=("Segoe UI", 9),
                bg=ARMED_COLOR if armed else t["card"],
                fg="white" if armed else t["card_fg"],
                activebackground=ARMED_COLOR, relief="flat", anchor="w",
                padx=6, pady=2,
                command=lambda o=option: self._arm(o))
            btn.pack(fill="x")
            self._source_buttons[label] = btn

    def _arm(self, option):
        query = option.get("query") or option.get("key")
        self.armed = {"label": option["label"], "query": query}
        self.armed_label.config(
            text=f"Armed: {option['label']} — highlight text and Assign",
            fg=ARMED_COLOR)
        self._render_sources()

    def _open_advanced_source(self):
        if not self.export_result.get("client_data"):
            messagebox.showinfo(
                "Advanced Query", "The query builder needs client data — "
                "open the slot mapper from a report.", parent=self.window)
            return
        self._advanced_target = None
        show_query_builder(self)

    def _refresh_metrics(self):
        """Query-builder host hook. An applied query either attaches to
        the chart/table slot whose Advanced… opened the builder, or
        becomes an armed Saved Queries source."""
        target, self._advanced_target = self._advanced_target, None
        query = self._pending_query
        self._pending_query = None
        if not query:
            return
        if target and target in self.ir.slot_registry:
            self.mapping.setdefault("slots", {})[target] = {
                "query": dict(query), "format": "number"}
            self._render_slide()
            return
        name = self.selected_metric or "advanced query"
        self.named_queries.setdefault(name, {"query": dict(query)})
        self._arm({"label": f"⚡ {name}", "query": dict(query)})

    # ── Shape cards ───────────────────────────────────────────────────────

    def _go_slide(self, step):
        self.current_slide = max(
            0, min(len(self.ir.slides) - 1, self.current_slide + step))
        self._render_slide()

    def _render_slide(self):
        for child in self.cards.winfo_children():
            child.destroy()
        slide = self.ir.slides[self.current_slide]
        self.slide_label.config(
            text=f"Slide {slide.slide_index + 1} of {len(self.ir.slides)}")
        shown = 0
        for shape in sorted(slide.shapes, key=lambda s: s.z_order):
            if shape.excluded or shape.unsupported:
                continue
            if shape.shape_type in ("chart", "table"):
                self._render_frame_card(shape)
                shown += 1
            elif (shape.text or "").strip():
                self._render_text_card(shape)
                shown += 1
        if not shown:
            tk.Label(self.cards, text="No mappable shapes on this slide.",
                     font=("Segoe UI", 10), bg=self.t["card"],
                     fg=self.t["muted"]).pack(pady=30)

    def _shape_slots(self, shape):
        return [(name, spec) for name, spec in self.ir.slot_registry.items()
                if spec.get("shape_id") == shape.shape_id]

    def _render_text_card(self, shape):
        t = self.t
        card = tk.Frame(self.cards, bg=t["card"],
                        highlightbackground=SLOT_COLOR
                        if shape.classification == "dynamic" else "#2E8B57",
                        highlightthickness=1)
        card.pack(fill="x", padx=8, pady=4)
        head = tk.Frame(card, bg=t["card"])
        head.pack(fill="x", padx=8, pady=(5, 0))
        tk.Label(head, text=f"{shape.name}  ({shape.shape_type})",
                 font=("Segoe UI", 10, "bold"), bg=t["card"],
                 fg=t["card_fg"]).pack(side="left")
        tk.Button(head, text="Assign Selection →",
                  font=("Segoe UI", 9, "bold"), bg=ARMED_COLOR, fg="white",
                  relief="flat", padx=10, pady=2,
                  command=lambda: self._assign_selection(shape)
                  ).pack(side="right")

        lines = max(shape.text.count("\n") + 1, len(shape.text) // 80 + 1)
        text = tk.Text(card, height=min(4, lines),
                       font=("Segoe UI", 10), bg=t["input_bg"],
                       fg=t["input_fg"], relief="solid", borderwidth=1,
                       wrap="word", exportselection=False)
        text.insert("1.0", shape.text)
        for name, spec in self._shape_slots(shape):
            placeholder = spec.get("placeholder_text")
            if placeholder:
                where = text.search(placeholder, "1.0", tk.END)
                if where:
                    text.tag_add(name, where,
                                 f"{where}+{len(placeholder)}c")
                    text.tag_config(name, background="#F5CBA7",
                                    foreground="#5B2C06")
        text.configure(state="disabled")   # selectable, not editable
        text.pack(fill="x", padx=8, pady=4)
        card.text_widget = text

        for name, spec in self._shape_slots(shape):
            self._render_slot_row(card, name, spec)

    def _render_frame_card(self, shape):
        t = self.t
        card = tk.Frame(self.cards, bg=t["card"],
                        highlightbackground=SLOT_COLOR, highlightthickness=1)
        card.pack(fill="x", padx=8, pady=4)
        head = tk.Frame(card, bg=t["card"])
        head.pack(fill="x", padx=8, pady=(5, 0))
        tk.Label(head, text=f"{shape.name}  ({shape.shape_type})",
                 font=("Segoe UI", 10, "bold"), bg=t["card"],
                 fg=t["card_fg"]).pack(side="left")
        for name, spec in self._shape_slots(shape):
            apply_as = ("Chart Data" if spec["type"] == "chart_data"
                        else "Table")
            row = tk.Frame(card, bg=t["card"])
            row.pack(fill="x", padx=8, pady=(2, 5))
            tk.Label(row, text=f"● {name}", font=("Segoe UI", 9, "bold"),
                     bg=t["card"], fg=SLOT_COLOR).pack(side="left")
            tk.Button(row, text=f"Advanced… (Apply as {apply_as})",
                      font=("Segoe UI", 8), bg="#8B4513", fg="white",
                      relief="flat", padx=8, pady=2,
                      command=lambda n=name: self._open_advanced_frame(n)
                      ).pack(side="left", padx=8)
            entry = self.mapping.get("slots", {}).get(name)
            state = "mapped ✓" if entry and entry.get("query") \
                else "(unmapped)"
            tk.Label(row, text=state, font=("Segoe UI", 9), bg=t["card"],
                     fg=t["muted"]).pack(side="left")

    def _render_slot_row(self, card, name, spec):
        t = self.t
        row = tk.Frame(card, bg=t["card"])
        row.pack(fill="x", padx=8, pady=(0, 3))
        target = (f"“{spec['placeholder_text'][:32]}”"
                  if spec.get("placeholder_text") else "whole box")
        tk.Label(row, text=f"● {name} → {target}",
                 font=("Segoe UI", 9, "bold"), bg=t["card"], fg=SLOT_COLOR
                 ).pack(side="left")

        entry = self.mapping.get("slots", {}).get(name, {})
        source = tk.Label(row, text=self._source_label(entry.get("query")),
                          font=("Segoe UI", 9), bg=t["card"], fg=t["card_fg"])
        source.pack(side="left", padx=(8, 4))

        format_var = tk.StringVar(value=entry.get("format", spec["type"])
                                  if entry.get("format", spec["type"])
                                  in FORMATS else "text")
        fmt = ttk.Combobox(row, textvariable=format_var, width=9,
                           state="readonly", values=FORMATS)
        fmt.pack(side="left", padx=(0, 6))

        def on_format(_event, slot=name, var=format_var):
            slot_entry = self.mapping.setdefault("slots", {}).setdefault(
                slot, {"query": None})
            slot_entry["format"] = var.get()
            self._render_slide()
        fmt.bind("<<ComboboxSelected>>", on_format)

        preview = self._preview_text(name, entry)
        if preview:
            tk.Label(row, text=preview, font=("Segoe UI", 9, "italic"),
                     bg=t["card"], fg="#2E8B57").pack(side="left")
        tk.Button(row, text="✕", font=("Segoe UI", 8), bg=t["danger"],
                  fg="white", relief="flat", padx=4,
                  command=lambda: self._remove_slot(name)).pack(side="right")

    # ── Actions ───────────────────────────────────────────────────────────

    def _assign_selection(self, shape):
        if not self.armed:
            messagebox.showinfo(
                "Assign", "Arm a data source first (click one in the "
                "sidebar), then highlight the text it should replace.",
                parent=self.window)
            return
        selection = ""
        # Only a selection on THIS shape's own text widget counts
        for child in self.cards.winfo_children():
            text = getattr(child, "text_widget", None)
            if text is None or not self._card_holds(child, shape):
                continue
            try:
                ranges = text.tag_ranges("sel")
            except tk.TclError:
                ranges = None
            if ranges:
                selection = text.get(ranges[0], ranges[1])
            break
        selection = selection.strip()
        if not selection and not messagebox.askyesno(
                "Assign Whole Box",
                f"No text is highlighted in \"{shape.name}\" — assign "
                f"{self.armed['label']} to replace the WHOLE box?",
                parent=self.window):
            return

        slot_type, default_format = _source_kind(self.armed["query"])
        existing = next(
            (n for n, s in self._shape_slots_dict(shape).items()
             if (s.get("placeholder_text") or "") == selection), None)
        try:
            if existing:
                name = existing
            else:
                name = add_slot(self.ir, shape.shape_id, selection,
                                name=self.armed["label"],
                                slot_type=slot_type)
        except (ValueError, KeyError) as e:
            messagebox.showerror("Assign", str(e), parent=self.window)
            return
        self.mapping.setdefault("slots", {})[name] = {
            "query": self.armed["query"], "format": default_format}
        self._render_slide()

    def _shape_slots_dict(self, shape):
        return {name: spec for name, spec in self.ir.slot_registry.items()
                if spec.get("shape_id") == shape.shape_id}

    def _card_holds(self, card, shape):
        """True when this card renders this shape (the shape's own text
        widget holds the selection)."""
        text = getattr(card, "text_widget", None)
        return text is not None and \
            text.get("1.0", "end-1c") == (shape.text or "")

    def _remove_slot(self, name):
        try:
            remove_slot(self.ir, name)
        except KeyError:
            pass
        self.mapping.get("slots", {}).pop(name, None)
        self._render_slide()

    def _open_advanced_frame(self, slot):
        if not self.export_result.get("client_data"):
            messagebox.showinfo(
                "Advanced Query", "The query builder needs client data — "
                "open the slot mapper from a report.", parent=self.window)
            return
        self._advanced_target = slot
        show_query_builder(self)

    def _source_label(self, query):
        if query is None:
            return "(unmapped)"
        for option in self._source_catalog():
            if (option.get("query") or option.get("key")) == query:
                return option["label"]
        return "⚡ advanced query"

    def _preview_text(self, slot, entry):
        client_data = self.export_result.get("client_data", [])
        if not entry.get("query") or not client_data:
            return ""
        values, issues = resolve_slot_values(
            self.ir, {"slots": {slot: entry}}, client_data,
            self.client_name, self.start_date, self.end_date)
        if slot in values:
            value = values[slot]
            if isinstance(value, dict):
                return "→ (chart/table data)"
            return f"→ {value}"
        return f"⚠ {dict(issues).get(slot, '')}"

    # ── Save / build ──────────────────────────────────────────────────────

    def _save(self, quiet=False):
        # Slots and their mappings change together in this window — the
        # registry (template.json) and mapping.json save as one action.
        slots = {name: entry for name, entry in
                 self.mapping.get("slots", {}).items()
                 if name in self.ir.slot_registry}
        self.mapping["slots"] = slots
        warnings = validate_slot_mapping(self.ir, self.mapping)
        try:
            save_template_ir(self.ir, self.template_dir)
            save_slot_mapping(self.mapping, self.template_dir)
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
                                f"Slots and mapping saved:\n{self.template_dir}",
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
                   f"{report['charts_cloned']} chart(s) cloned, "
                   f"{report['slots_filled']} slot(s) filled")
        problems = [f"{s}: {r}" for s, r in report["slots_unfilled"]]
        problems += [f"{s}: {r}" for s, r in report["skipped"]]
        problems += list(report.get("notes", []))
        if problems:
            messagebox.showwarning(
                "Report Built — With Gaps",
                f"{save_path}\n\n{summary}\n\nNeeds attention:\n• "
                + "\n• ".join(problems), parent=self.window)
        else:
            messagebox.showinfo("Report Built",
                                f"{save_path}\n\n{summary}",
                                parent=self.window)
