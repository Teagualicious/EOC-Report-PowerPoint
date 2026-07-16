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
        if "custom_text" in query:
            return "text", "text"
        if "image_path" in query or "image_path_abs" in query:
            return "image", "text"
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
        pane.rowconfigure(2, weight=1)
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

        # Per-slide thumbnail (from the source deck kept in the store) —
        # size-locked like every preview pane in this app. Renders via
        # PowerPoint on Windows; elsewhere it shows a text summary.
        from engine.template_ir.ingest import SOURCE_PPTX_NAME
        self._source_pptx = os.path.join(self.template_dir,
                                         SOURCE_PPTX_NAME)
        self._thumb_ref = None
        self._thumb_token = 0
        self.preview_label = None
        if os.path.isfile(self._source_pptx):
            preview_frame = tk.LabelFrame(
                pane, text="  Slide Preview  ",
                font=("Segoe UI", 9, "bold"), bg=t["card"],
                fg=t["card_fg"], width=360, height=210)
            preview_frame.grid(row=1, column=0, sticky="w", pady=(0, 4))
            preview_frame.pack_propagate(False)
            self.preview_label = tk.Label(preview_frame, bg="#333333",
                                          text="Loading preview…",
                                          fg="#888888",
                                          font=("Segoe UI", 9))
            self.preview_label.pack(fill="both", expand=True)

        holder = tk.Frame(pane, bg=t["card"])
        holder.grid(row=2, column=0, sticky="nsew")
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
        # Quick Fill parity with the classic mapper: custom text and
        # images are first-class sources (both arm through a picker)
        catalog = [{"key": "__custom__", "label": "✏ Custom Text…",
                    "category": "special"},
                   {"key": "__browse_image__", "label": "🖼 Browse Image…",
                    "category": "special"}]
        catalog += list(self.options)
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
        if option.get("key") == "__custom__":
            text = self._ask_custom_text()
            if not text:
                return
            snippet = text.splitlines()[0][:24]
            option = {"label": f"✏ “{snippet}”",
                      "query": {"custom_text": text}}
        elif option.get("key") == "__browse_image__":
            path = filedialog.askopenfilename(
                parent=self.window, title="Choose the image to insert",
                filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp"),
                           ("All files", "*.*")])
            if not path:
                return
            option = {"label": f"🖼 {os.path.basename(path)}",
                      "query": {"image_path_abs": path}}
        query = option.get("query") or option.get("key")
        self.armed = {"label": option["label"], "query": query}
        hint = ("click Assign on the shape to replace its image"
                if _source_kind(query)[0] == "image"
                else "highlight text and Assign")
        self.armed_label.config(
            text=f"Armed: {option['label']} — {hint}", fg=ARMED_COLOR)
        self._render_sources()

    def _ask_custom_text(self):
        """Multi-line custom text prompt (the classic mapper's ✏ field,
        grown up: line breaks land as REAL breaks in the deck)."""
        t = self.t
        dialog = tk.Toplevel(self.window)
        dialog.title("Custom Text")
        fit_window(dialog, 420, 240)
        dialog.configure(bg=t["bg"]); dialog.transient(self.window)
        dialog.grab_set(); dialog.lift(); dialog.focus_force()
        tk.Label(dialog, text="Text to insert (line breaks are kept):",
                 font=("Segoe UI", 9), bg=t["bg"], fg=t["fg"]
                 ).pack(padx=12, pady=(10, 4), anchor="w")
        box = tk.Text(dialog, height=6, font=("Segoe UI", 10),
                      bg=t["input_bg"], fg=t["input_fg"],
                      insertbackground=t["insert"], relief="solid",
                      borderwidth=1)
        box.pack(fill="both", expand=True, padx=12)
        box.focus_set()
        result = {"text": ""}

        def ok():
            result["text"] = box.get("1.0", "end-1c").strip()
            dialog.destroy()

        row = tk.Frame(dialog, bg=t["bg"])
        row.pack(fill="x", padx=12, pady=8)
        tk.Button(row, text="Use This Text", font=("Segoe UI", 9, "bold"),
                  bg=ARMED_COLOR, fg="white", relief="flat", padx=12,
                  pady=4, command=ok).pack(side="left")
        tk.Button(row, text="Cancel", font=("Segoe UI", 9),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
                  padx=12, pady=4, command=dialog.destroy).pack(side="right")
        self.window.wait_window(dialog)
        return result["text"]

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
            elif shape.shape_type == "image":
                self._render_image_card(shape)
                shown += 1
            elif (shape.text or "").strip():
                self._render_text_card(shape)
                shown += 1
        if not shown:
            tk.Label(self.cards, text="No mappable shapes on this slide.",
                     font=("Segoe UI", 10), bg=self.t["card"],
                     fg=self.t["muted"]).pack(pady=30)
        self._load_preview()

    def _load_preview(self):
        if self.preview_label is None:
            return
        from ui.utils import run_in_background
        self._thumb_token += 1
        token = self._thumb_token
        self._thumb_ref = None
        self.preview_label.config(image="", text="Loading preview…",
                                  fg="#888888")
        slide_num = self.current_slide + 1
        source = self._source_pptx

        def work():
            from engine.pptx_thumbs import get_slide_preview
            return get_slide_preview(source, slide_num, width=360)

        def success(preview):
            if token != self._thumb_token or \
                    not self.window.winfo_exists():
                return
            if preview.get("kind") == "image":
                from PIL import Image, ImageTk
                with Image.open(preview["value"]) as img_file:
                    image = img_file.copy()
                image.thumbnail((350, 195), Image.LANCZOS)
                self._thumb_ref = ImageTk.PhotoImage(image)
                self.preview_label.config(image=self._thumb_ref, text="")
            else:
                self.preview_label.config(
                    image="", text=preview.get("value", ""), fg="#B8B8B8",
                    font=("Segoe UI", 8), justify="left", wraplength=340)

        def error(exc):
            logger.debug("Slide preview failed: %s", exc)
            if token == self._thumb_token and self.window.winfo_exists():
                self.preview_label.config(image="", text="No preview",
                                          fg="#888888")

        run_in_background(self.window, work, success, error)

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

    def _render_image_card(self, shape):
        t = self.t
        card = tk.Frame(self.cards, bg=t["card"],
                        highlightbackground=SLOT_COLOR
                        if self._shape_slots(shape) else "#2E8B57",
                        highlightthickness=1)
        card.pack(fill="x", padx=8, pady=4)
        head = tk.Frame(card, bg=t["card"])
        head.pack(fill="x", padx=8, pady=(5, 0))
        tk.Label(head, text=f"{shape.name}  (picture)",
                 font=("Segoe UI", 10, "bold"), bg=t["card"],
                 fg=t["card_fg"]).pack(side="left")
        tk.Button(head, text="Assign Image →",
                  font=("Segoe UI", 9, "bold"), bg=ARMED_COLOR, fg="white",
                  relief="flat", padx=10, pady=2,
                  command=lambda: self._assign_selection(shape)
                  ).pack(side="right")
        tk.Label(card, text="Arm 🖼 Browse Image… in the sidebar, then "
                            "Assign to swap this picture per client "
                            "(crop and effects are kept).",
                 font=("Segoe UI", 8), bg=t["card"], fg=t["muted"]
                 ).pack(anchor="w", padx=8, pady=(0, 2))
        for name, spec in self._shape_slots(shape):
            self._render_slot_row(card, name, spec)
        tk.Frame(card, bg=t["card"], height=4).pack()

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
        if spec.get("type") == "image":
            target = "replaces the picture"
        elif spec.get("placeholder_text"):
            target = f"“{spec['placeholder_text'][:32]}”"
        else:
            target = "whole box"
        tk.Label(row, text=f"● {name} → {target}",
                 font=("Segoe UI", 9, "bold"), bg=t["card"], fg=SLOT_COLOR
                 ).pack(side="left")

        entry = self.mapping.get("slots", {}).get(name, {})
        source = tk.Label(row, text=self._source_label(entry.get("query")),
                          font=("Segoe UI", 9), bg=t["card"], fg=t["card_fg"])
        source.pack(side="left", padx=(8, 4))

        tk.Button(row, text="✕", font=("Segoe UI", 8), bg=t["danger"],
                  fg="white", relief="flat", padx=4,
                  command=lambda: self._remove_slot(name)).pack(side="right")
        if spec.get("type") == "image":
            return   # images have no format/preview machinery

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
            if var.get() != "date":
                slot_entry.pop("format_details", None)
            self._render_slide()
        fmt.bind("<<ComboboxSelected>>", on_format)

        if format_var.get() == "date":
            self._render_date_style(row, name, entry)

        preview = self._preview_text(name, entry)
        if preview:
            tk.Label(row, text=preview, font=("Segoe UI", 9, "italic"),
                     bg=t["card"], fg="#2E8B57").pack(side="left")

    def _render_date_style(self, row, name, entry):
        """Date slots pick their style, like the classic format popup —
        stored as format_details, honored by the same formatter."""
        from engine.pptx_formats import DATE_STYLE_CHOICES
        label_by_key = dict(DATE_STYLE_CHOICES)
        key_by_label = {label: key for key, label in DATE_STYLE_CHOICES}
        details = entry.get("format_details") or {}
        current = details.get("date_style", "long_ordinal")
        style_var = tk.StringVar(
            value=label_by_key.get(current, label_by_key["long_ordinal"]))
        style = ttk.Combobox(row, textvariable=style_var, width=16,
                             state="readonly",
                             values=list(key_by_label))
        style.pack(side="left", padx=(0, 6))

        def on_style(_event, slot=name, var=style_var):
            key = key_by_label[var.get()]
            new_details = {"format": "date", "date_style": key}
            if key == "custom":
                from tkinter import simpledialog
                pattern = simpledialog.askstring(
                    "Custom Date Format",
                    "strftime pattern (e.g. %b %d, %Y):",
                    parent=self.window)
                if not pattern:
                    return
                new_details["custom_strftime"] = pattern
            slot_entry = self.mapping.setdefault("slots", {}).setdefault(
                slot, {"query": None})
            slot_entry["format"] = "date"
            slot_entry["format_details"] = new_details
            self._render_slide()
        style.bind("<<ComboboxSelected>>", on_style)

    # ── Actions ───────────────────────────────────────────────────────────

    def _assign_selection(self, shape):
        if not self.armed:
            messagebox.showinfo(
                "Assign", "Arm a data source first (click one in the "
                "sidebar), then highlight the text it should replace.",
                parent=self.window)
            return
        if _source_kind(self.armed["query"])[0] == "image":
            self._assign_image(shape)
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

    def _assign_image(self, shape):
        """An armed image replaces the shape's picture (whole shape —
        selections don't apply). The file is copied into the template
        store so the mapping stays portable."""
        import shutil
        source_path = self.armed["query"].get("image_path_abs", "")
        if not os.path.isfile(source_path):
            messagebox.showerror("Assign Image",
                                 f"Image file not found:\n{source_path}",
                                 parent=self.window)
            return
        existing = next(
            (n for n, s in self._shape_slots_dict(shape).items()
             if s.get("type") == "image"), None)
        try:
            name = existing or add_slot(self.ir, shape.shape_id, "",
                                        name=f"{shape.name} image",
                                        slot_type="image")
        except (ValueError, KeyError) as e:
            messagebox.showerror("Assign Image", str(e), parent=self.window)
            return
        assets_dir = os.path.join(self.template_dir, "assets")
        os.makedirs(assets_dir, exist_ok=True)
        ext = os.path.splitext(source_path)[1] or ".png"
        rel = f"assets/{name}{ext}"
        try:
            shutil.copy2(source_path, os.path.join(self.template_dir, rel))
        except OSError as e:
            messagebox.showerror("Assign Image",
                                 f"Could not copy the image into the "
                                 f"template store:\n{e}", parent=self.window)
            return
        self.mapping.setdefault("slots", {})[name] = {
            "query": {"image_path": rel, "image_path_abs": source_path},
            "format": "text"}
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
        if isinstance(query, dict) and "custom_text" in query:
            return f"✏ “{query['custom_text'].splitlines()[0][:24]}”"
        if isinstance(query, dict) and ("image_path" in query
                                        or "image_path_abs" in query):
            path = query.get("image_path") or query.get("image_path_abs")
            return f"🖼 {os.path.basename(path)}"
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
