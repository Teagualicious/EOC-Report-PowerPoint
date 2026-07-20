"""Sidebar (metrics & values) methods for the template mapper."""

import logging
import os
import tkinter as tk
from tkinter import filedialog

from engine.query_resolver import build_simple_options, resolve_query
from mapper.format_popup import show_format_popup

logger = logging.getLogger(__name__)


class SidebarMixin:
    """Metric sidebar methods for PPTXWizard."""

    def _refresh_metrics(self):
        """Rebuild sidebar with simple options + advanced query toggle."""
        for w in self.metric_frame.winfo_children(): w.destroy()
        self.metric_buttons = {}
        search = self.metric_search_var.get().lower() if hasattr(self, 'metric_search_var') else ""

        options = build_simple_options(self.structured_metrics)
        # Saved/applied queries — rendered like any other metric row. The
        # stored query dict rides along so re-clicking the entry re-arms it
        # for assignment (and re-fills recompute it instead of going stale).
        # They slot in directly below the Quick Fill block so fresh applies
        # are visible without scrolling past every breakdown section.
        query_opts = [{'category': 'queries', 'label': _qn, 'key': _qn,
                       'query': self.named_queries[_qn].get('query')}
                      for _qn in sorted(getattr(self, 'named_queries', {}))]
        if query_opts:
            insert_at = 0
            for _i, _opt in enumerate(options):
                if _opt['category'] == 'special':
                    insert_at = _i + 1
            options[insert_at:insert_at] = query_opts

        current_category = ""
        category_labels = {"special": "Quick Fill", "total": "KPI Totals",
                           "queries": "Saved Queries"}

        for opt in options:
            cat = opt["category"]
            label = opt["label"]
            key = opt["key"]
            if search and search not in label.lower(): continue
            if cat != current_category:
                current_category = cat
                if cat.startswith("breakdown_"):
                    source = cat.replace("breakdown_", "").title()
                    self._add_separator(f"{source} Breakdown")
                elif cat in category_labels:
                    # Insert custom text and advanced query at end of Quick Fill
                    if cat == "total" and (not search or "custom" in search or "text" in search):
                        self._add_custom_text_field()
                    if cat == "total" and not search:
                        self._add_advanced_query_btn()
                    self._add_separator(category_labels[cat])

            if key in self.available_metrics:
                val = self.available_metrics.get(key, "")
            elif opt.get("query"):
                val = resolve_query(opt["query"], self.export_result["client_data"],
                    self.export_result.get("client_name", ""),
                    self._get_start_date(), self._get_end_date())
                # Cache so the format popup preview (and re-apply on slide
                # navigation) sees the real value instead of defaulting to 0
                self.available_metrics[key] = val
            else:
                val = ""
            if isinstance(val, (int, float)):
                vd = self._format_for_sidebar(key, val)
            else:
                vd = str(val)
            self._add_metric_btn(key, label, vd, special=(cat == "special"), query=opt.get("query"))

            # Right after the Browse Image entry, list this session's
            # browsed pictures as reusable one-click entries
            if key == "__image__" and getattr(self, "_session_images", None):
                for img_key, img_path in self._session_images.items():
                    base = os.path.basename(img_path)
                    if search and search not in base.lower():
                        continue
                    self._add_metric_btn(img_key, f"  🖼 Image: {base}", "",
                                          special=True)

        # If no KPI Totals section existed, still add custom text
        if "total" not in [o["category"] for o in options]:
            if not search or "custom" in search or "text" in search:
                self._add_custom_text_field()

    def _add_advanced_query_btn(self):
        t = self.t
        adv = tk.Frame(self.metric_frame, bg=t["card"], cursor="hand2")
        adv.pack(fill="x", pady=2, padx=3)
        tk.Label(adv, text="\U0001f527 Advanced Query Builder...", font=("Calibri", 9, "bold"),
                 bg=t["card"], fg=t["accent"], cursor="hand2").pack(fill="x", padx=5, pady=5)
        adv.bind("<Button-1>", lambda e: self._show_query_builder())
        for child in adv.winfo_children():
            child.bind("<Button-1>", lambda e: self._show_query_builder())

    def _format_for_sidebar(self, metric_key, val):
        """Format a value using the stored format details for that metric."""
        from engine.pptx_formats import format_with_details
        details = getattr(self, '_metric_format_details', {}).get(metric_key)
        return format_with_details(val, details)

    def _get_start_date(self):
        for d in self.export_result.get("client_data", []):
            if d.get("start_date"): return d["start_date"]
        return ""

    def _get_end_date(self):
        for d in self.export_result.get("client_data", []):
            if d.get("end_date"): return d["end_date"]
        return ""

    def _add_separator(self, label):
        t = self.t
        sep = tk.Frame(self.metric_frame, bg=t["accent"])
        sep.pack(fill="x", padx=3, pady=(8, 3))
        tk.Label(sep, text=f"  {label}", font=("Calibri", 8, "bold"),
                 bg=t["accent"], fg="white").pack(anchor="w", pady=2)

    def _add_custom_text_field(self):
        t = self.t
        custom_frame = tk.Frame(self.metric_frame, bg=t["card"])
        custom_frame.pack(fill="x", pady=2, padx=3)
        tk.Label(custom_frame, text="✏ Custom Text:", font=("Calibri", 8, "bold"),
                 bg=t["card"], fg="#1E6E3E").pack(anchor="w", padx=5, pady=(3, 0))
        row = tk.Frame(custom_frame, bg=t["card"])
        row.pack(fill="x", padx=5, pady=(2, 5))
        self._custom_text_var = tk.StringVar()
        tk.Entry(row, textvariable=self._custom_text_var, width=22,
                 font=("Calibri", 9), bg=t["input_bg"], fg=t["input_fg"],
                 insertbackground=t["insert"], relief="solid", borderwidth=1
                 ).pack(side="left", fill="x", expand=True)
        def select_custom():
            text = self._custom_text_var.get().strip()
            if text:
                key = f"__custom_{hash(text)}__"
                self.available_metrics[key] = text
                self.selected_metric = key
                # The text rides IN the assignment as a query so it
                # persists in the saved mapping — Auto-Fill used to find
                # only the session-local __custom key and fill nothing
                self._pending_query = {"custom_text": text}
                self._pending_image = None
        tk.Button(row, text="Use", font=("Calibri", 8, "bold"),
                  bg=t["accent"], fg="white", relief="flat", padx=8,
                  command=select_custom).pack(side="right", padx=(3, 0))

    def _add_metric_btn(self, key, name, value_display, special=False, query=None):
        t = self.t
        btn = tk.Frame(self.metric_frame, bg=t["card"], cursor="hand2")
        btn.pack(fill="x", pady=1)
        fg = "#1E6E3E" if special else t["fg"]

        # Show format indicator if set
        fmt_indicator = self._metric_formats.get(key, "")
        display_name = name
        if fmt_indicator and fmt_indicator != "text":
            fmt_labels = {"number": "  #", "currency": "  $", "percentage": "  %",
                          "custom": "  ✎"}
            display_name += fmt_labels.get(fmt_indicator, "")

        tk.Label(btn, text=display_name, font=("Calibri", 9, "bold"),
                 bg=t["card"], fg=fg, anchor="w", cursor="hand2"
                 ).pack(fill="x", padx=5, pady=(2, 0))
        tk.Label(btn, text=value_display, font=("Calibri", 8),
                 bg=t["card"], fg=t["muted"], anchor="w", cursor="hand2"
                 ).pack(fill="x", padx=5, pady=(0, 2))
        def select(event, k=key, b=btn, q=query):
            self._select_metric(k, b, q)
        def right_click(event, k=key):
            self._show_format_popup(event, k)
        btn.bind("<Button-1>", select)
        btn.bind("<Button-3>", right_click)
        for child in btn.winfo_children():
            child.bind("<Button-1>", select)
            child.bind("<Button-3>", right_click)
        # Keep the armed selection visibly highlighted across sidebar
        # rebuilds (e.g. right after an Advanced Query "Apply as ...")
        if key == getattr(self, "selected_metric", None):
            btn.configure(bg=t["highlight"])
            for child in btn.winfo_children():
                child.configure(bg=t["highlight"])
        self.metric_buttons[key] = btn

    # Date-flavored quick fills get the compact date-only format popup
    _DATE_QUICKFILL_KEYS = {"__date_range__", "__start_date__", "__end_date__"}

    def _show_format_popup(self, event, metric_key):
        """Right-click format popup (delegates to mapper.format_popup)."""
        show_format_popup(self, event, metric_key,
                          date_only=metric_key in self._DATE_QUICKFILL_KEYS)

    def _select_metric(self, metric_key, btn, query=None):
        t = self.t
        for b in self.metric_buttons.values():
            b.configure(bg=t["card"])
            for child in b.winfo_children(): child.configure(bg=t["card"])
        if metric_key == "__image__":
            img_path = filedialog.askopenfilename(title="Select Image",
                parent=self.window,
                filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp")])
            if img_path:
                # Register the picture as a reusable session quick-fill entry
                # (one entry per browsed file — no more shared pending state
                # bleeding between assignments)
                base = os.path.basename(img_path)
                img_key = f"__image_{base}__"
                if not hasattr(self, "_session_images"):
                    self._session_images = {}
                self._session_images[img_key] = img_path
                self.selected_metric = img_key
                self._pending_image = img_path
                self._pending_query = None
                self._refresh_metrics()
            return
        # Selecting a previously browsed image re-arms it directly
        if metric_key.startswith("__image_") and hasattr(self, "_session_images") \
                and metric_key in self._session_images:
            self.selected_metric = metric_key
            self._pending_image = self._session_images[metric_key]
            self._pending_query = None
            btn.configure(bg=t["highlight"])
            for child in btn.winfo_children(): child.configure(bg=t["highlight"])
            return
        self.selected_metric = metric_key
        self._pending_image = None
        self._pending_query = query
        btn.configure(bg=t["highlight"])
        for child in btn.winfo_children(): child.configure(bg=t["highlight"])
