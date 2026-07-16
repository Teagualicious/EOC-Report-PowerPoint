"""Template review GUI (template-first Phase B).

Presents the auto-classification suggestions for an ingested template —
green static / orange dynamic, per proposal — for mandatory human review:
override classifications, rename slots, exclude shapes. All mutations go
through the pure helpers in engine.template_ir.classify (headlessly
tested); this window is a thin shell over them.

``open_template_first`` is the entry point shared by Settings and the
report-generation flow: pick an ingested template store (or ingest a new
deck), then review or map it.
"""

import logging
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from engine.template_ir import (classify_template, load_template_ir,
                                save_template_ir)
from engine.template_ir.classify import (rename_slot, set_classification,
                                         set_excluded, shape_slots)
from engine.template_ir.ingest import ingest_template, list_template_stores
from ui.utils import fit_window, grow_to_content

logger = logging.getLogger(__name__)

STATIC_COLOR = "#2E8B57"     # green — reproduced verbatim
DYNAMIC_COLOR = "#D2691E"    # orange — receives mapped data
EXCLUDED_COLOR = "#777777"


def open_template_first(parent, theme, export_result=None):
    """Template-first launcher: list ingested template stores, ingest a
    new deck, and open the review or slot-mapping window."""
    t = theme
    win = tk.Toplevel(parent)
    win.title("Template-First (Beta)")
    fit_window(win, 560, 400)
    win.configure(bg=t["bg"]); win.transient(parent)
    win.grab_set(); win.lift(); win.focus_force()

    tk.Label(win, text="🧪 Template-First Reports (Beta)",
             font=("Segoe UI", 13, "bold"), bg=t["bg"], fg=t["fg"]
             ).pack(pady=(15, 2))
    tk.Label(win, text="Ingest a client deck once, review what's dynamic, "
                       "map named slots —\nthen every month builds a fresh "
                       "deck from the stored template.",
             font=("Segoe UI", 9), bg=t["bg"], fg=t["muted"]).pack(pady=(0, 8))

    lb = tk.Listbox(win, font=("Segoe UI", 10), bg=t["input_bg"],
                    fg=t["input_fg"], selectbackground=t["accent"],
                    selectforeground="white", relief="solid", borderwidth=1)
    lb.pack(fill="both", expand=True, padx=20)

    stores = []

    def refresh():
        stores[:] = list_template_stores()
        lb.delete(0, tk.END)
        from engine.template_ir.mapping import load_slot_mapping
        for store in stores:
            mapped = "mapped" if load_slot_mapping(store) else "needs mapping"
            lb.insert(tk.END, f"{os.path.basename(store)}  ({mapped})")
        if not stores:
            lb.insert(tk.END, "No templates ingested yet — start below")
        else:
            lb.selection_set(0)

    refresh()

    def selected_store():
        sel = lb.curselection()
        if not sel or sel[0] >= len(stores):
            return None
        return stores[sel[0]]

    def ingest_new():
        path = filedialog.askopenfilename(
            parent=win, title="Select client PowerPoint to ingest",
            filetypes=[("PowerPoint", "*.pptx")])
        if not path:
            return
        try:
            template_dir = ingest_template(path)
            ir = classify_template(load_template_ir(template_dir))
            save_template_ir(ir, template_dir)
        except Exception as e:
            logger.exception("Template ingest failed: %s", path)
            messagebox.showerror("Ingest Failed", str(e), parent=win)
            return
        refresh()
        TemplateReviewWindow(win, t, template_dir, export_result)

    def review():
        store = selected_store()
        if store:
            TemplateReviewWindow(win, t, store, export_result)

    def update_from_deck():
        store = selected_store()
        if not store:
            return
        path = filedialog.askopenfilename(
            parent=win, title="Select the client's UPDATED PowerPoint",
            filetypes=[("PowerPoint", "*.pptx")])
        if not path:
            return
        from engine.template_ir import reingest_template
        try:
            _, deltas = reingest_template(path, store)
        except Exception as e:
            logger.exception("Template re-ingest failed: %s", path)
            messagebox.showerror("Update Failed", str(e), parent=win)
            return
        refresh()
        details = [d["detail"] for d in deltas]
        messagebox.showinfo(
            "Template Updated",
            "Review work carried forward.\n\nWhat changed:\n• "
            + "\n• ".join(details) if details else
            "Template updated — nothing changed that needs review.",
            parent=win)
        TemplateReviewWindow(win, t, store, export_result)

    def map_slots():
        store = selected_store()
        if store:
            from mapper.slot_mapper import SlotMapperWindow
            SlotMapperWindow(win, t, store, export_result)

    # Two logical rows so this compact dialog never mashes its buttons:
    # template actions first, then what to open for the selected store
    btns = tk.Frame(win, bg=t["bg"])
    btns.pack(fill="x", padx=20, pady=(10, 2))
    tk.Button(btns, text="Ingest New Template…", font=("Segoe UI", 10),
              bg=t["success"], fg="white", relief="flat", padx=12, pady=6,
              command=ingest_new).pack(side="left")
    tk.Button(btns, text="Update from New Deck…", font=("Segoe UI", 10),
              bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
              padx=12, pady=6, command=update_from_deck).pack(side="left",
                                                              padx=8)
    open_row = tk.Frame(win, bg=t["bg"])
    open_row.pack(fill="x", padx=20, pady=(2, 12))
    tk.Button(open_row, text="Review Shapes", font=("Segoe UI", 10),
              bg=t["accent"], fg="white", relief="flat", padx=12, pady=6,
              command=review).pack(side="left")
    tk.Button(open_row, text="Map Slots →", font=("Segoe UI", 10, "bold"),
              bg="#1E6E3E", fg="white", relief="flat", padx=12, pady=6,
              command=map_slots).pack(side="left", padx=8)
    tk.Button(open_row, text="Close", font=("Segoe UI", 10),
              bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
              padx=12, pady=6, command=win.destroy).pack(side="right")
    grow_to_content(win)


class TemplateReviewWindow:
    """Slide-by-slide classification review for one template store."""

    def __init__(self, parent, theme, template_dir, export_result=None):
        self.t = theme
        self.template_dir = template_dir
        self.export_result = export_result
        self.ir = load_template_ir(template_dir)
        if not self.ir.slot_registry and not any(
                sh.classify_reason
                for sl in self.ir.slides for sh in sl.shapes):
            classify_template(self.ir)   # first open of a Phase A store
        self.current_slide = 0
        self._build(parent)

    def _build(self, parent):
        t = self.t
        self.window = tk.Toplevel(parent)
        self.window.title(f"Template Review — {self.ir.template_name}")
        fit_window(self.window, 960, 640)
        self.window.configure(bg=t["bg"])
        self.window.grab_set(); self.window.lift(); self.window.focus_force()

        tk.Label(self.window,
                 text="Review what the engine found: 🟢 static branding is "
                      "copied exactly; 🟠 dynamic shapes receive mapped "
                      "data. Fix anything it got wrong, then name the slots.",
                 font=("Segoe UI", 9), bg=t["bg"], fg=t["muted"],
                 wraplength=900, justify="left").pack(padx=15, pady=(10, 5),
                                                      anchor="w")

        main = tk.Frame(self.window, bg=t["bg"])
        main.pack(fill="both", expand=True, padx=15)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        self.slide_list = tk.Listbox(main, font=("Segoe UI", 10), width=18,
                                     bg=t["input_bg"], fg=t["input_fg"],
                                     selectbackground=t["accent"],
                                     selectforeground="white",
                                     relief="solid", borderwidth=1,
                                     exportselection=False)
        self.slide_list.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        for slide in self.ir.slides:
            dynamic = sum(1 for s in slide.shapes
                          if s.classification == "dynamic")
            self.slide_list.insert(
                tk.END, f"Slide {slide.slide_index + 1}  ({dynamic} dyn)")
        self.slide_list.selection_set(0)
        self.slide_list.bind("<<ListboxSelect>>", self._on_slide_select)

        card_holder = tk.Frame(main, bg=t["card"])
        card_holder.grid(row=0, column=1, sticky="nsew")
        self.canvas = tk.Canvas(card_holder, bg=t["card"],
                                highlightthickness=0)
        sb = ttk.Scrollbar(card_holder, orient="vertical",
                           command=self.canvas.yview)
        self.cards = tk.Frame(self.canvas, bg=t["card"])
        self.cards.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        _cw = self.canvas.create_window((0, 0), window=self.cards, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(
            _cw, width=e.width))
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        bottom = tk.Frame(self.window, bg=t["bg"])
        bottom.pack(fill="x", padx=15, pady=10)
        tk.Button(bottom, text="Re-run Auto-Classify", font=("Segoe UI", 9),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
                  padx=10, pady=6, command=self._reclassify).pack(side="left")
        tk.Button(bottom, text="Save & Map Slots →",
                  font=("Segoe UI", 10, "bold"), bg="#1E6E3E", fg="white",
                  relief="flat", padx=15, pady=6,
                  command=self._save_and_map).pack(side="right")
        tk.Button(bottom, text="Save", font=("Segoe UI", 10), bg=t["accent"],
                  fg="white", relief="flat", padx=15, pady=6,
                  command=self._save).pack(side="right", padx=6)

        grow_to_content(self.window)
        self._render_slide()

    # ── Rendering ─────────────────────────────────────────────────────────

    def _on_slide_select(self, _event=None):
        sel = self.slide_list.curselection()
        if sel:
            self.current_slide = sel[0]
            self._render_slide()

    def _render_slide(self):
        for child in self.cards.winfo_children():
            child.destroy()
        slide = self.ir.slides[self.current_slide]
        for shape in sorted(slide.shapes, key=lambda s: s.z_order):
            self._render_card(shape)

    def _render_card(self, shape):
        t = self.t
        color = (EXCLUDED_COLOR if shape.excluded
                 else DYNAMIC_COLOR if shape.classification == "dynamic"
                 else STATIC_COLOR)
        card = tk.Frame(self.cards, bg=t["card"], highlightbackground=color,
                        highlightthickness=2)
        card.pack(fill="x", padx=8, pady=4)

        head = tk.Frame(card, bg=t["card"])
        head.pack(fill="x", padx=8, pady=(6, 0))
        state = ("excluded" if shape.excluded else shape.classification)
        tk.Label(head, text=f"{shape.name}  ({shape.shape_type})",
                 font=("Segoe UI", 10, "bold"), bg=t["card"], fg=t["card_fg"]
                 ).pack(side="left")
        tk.Label(head, text=state.upper(), font=("Segoe UI", 9, "bold"),
                 bg=t["card"], fg=color).pack(side="right")

        snippet = shape.text.replace("\n", " · ")[:90]
        detail = shape.classify_reason + (f"  —  “{snippet}”" if snippet
                                          else "")
        tk.Label(card, text=detail, font=("Segoe UI", 8), bg=t["card"],
                 fg=t["muted"], wraplength=640, justify="left"
                 ).pack(anchor="w", padx=8)

        for slot in shape_slots(self.ir, shape.shape_id):
            self._render_slot_row(card, shape, slot)

        row = tk.Frame(card, bg=t["card"])
        row.pack(fill="x", padx=8, pady=(2, 6))
        other = "static" if shape.classification == "dynamic" else "dynamic"
        tk.Button(row, text=f"Make {other.title()}", font=("Segoe UI", 8),
                  bg=t["secondary"], fg=t["secondary_fg"], relief="flat",
                  padx=8, pady=2,
                  command=lambda: self._set_classification(shape, other)
                  ).pack(side="left")
        excl_var = tk.BooleanVar(value=shape.excluded)
        tk.Checkbutton(row, text="Exclude from build", variable=excl_var,
                       font=("Segoe UI", 8), bg=t["card"], fg=t["card_fg"],
                       selectcolor=t["card"], activebackground=t["card"],
                       command=lambda: self._set_excluded(shape,
                                                          excl_var.get())
                       ).pack(side="left", padx=10)

    def _render_slot_row(self, card, shape, slot):
        t = self.t
        spec = self.ir.slot_registry[slot]
        row = tk.Frame(card, bg=t["card"])
        row.pack(fill="x", padx=8, pady=1)
        tk.Label(row, text="Slot:", font=("Segoe UI", 8, "bold"),
                 bg=t["card"], fg=DYNAMIC_COLOR).pack(side="left")
        var = tk.StringVar(value=slot)
        entry = tk.Entry(row, textvariable=var, width=26,
                         font=("Segoe UI", 9), bg=t["input_bg"],
                         fg=t["input_fg"], insertbackground=t["insert"],
                         relief="solid", borderwidth=1)
        entry.pack(side="left", padx=5)
        tk.Label(row, text=f"{spec['type']}  ·  fills "
                           f"“{spec['placeholder_text'][:40]}”",
                 font=("Segoe UI", 8), bg=t["card"], fg=t["muted"]
                 ).pack(side="left")

        def commit(_event=None, current=slot):
            new = var.get().strip()
            if new == current:
                return
            try:
                final = rename_slot(self.ir, current, new)
            except (ValueError, KeyError) as e:
                messagebox.showerror("Rename Slot", str(e),
                                     parent=self.window)
                var.set(current)
                return
            var.set(final)
            self._render_slide()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)

    # ── Mutations (thin wrappers over the tested helpers) ─────────────────

    def _set_classification(self, shape, classification):
        set_classification(self.ir, shape.shape_id, classification)
        self._render_slide()

    def _set_excluded(self, shape, excluded):
        set_excluded(self.ir, shape.shape_id, excluded)
        self._render_slide()

    def _reclassify(self):
        if not messagebox.askyesno(
                "Re-run Auto-Classify",
                "This rebuilds every suggestion and slot name, discarding "
                "your manual overrides for this template. Continue?",
                parent=self.window):
            return
        classify_template(self.ir)
        self._render_slide()

    def _save(self, quiet=False):
        try:
            save_template_ir(self.ir, self.template_dir)
        except OSError as e:
            logger.exception("Could not save template IR")
            messagebox.showerror("Save Failed", str(e), parent=self.window)
            return False
        if not quiet:
            messagebox.showinfo(
                "Saved", f"Template saved:\n{self.template_dir}",
                parent=self.window)
        return True

    def _save_and_map(self):
        if not self._save(quiet=True):
            return
        from mapper.slot_mapper import SlotMapperWindow
        SlotMapperWindow(self.window, self.t, self.template_dir,
                         self.export_result)
