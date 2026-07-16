"""Query search popup: the Excel search's term grammar over a live pivot.

DEMO REWORK (v1.32 branch): the multi-select filter panels are replaced
by a search box speaking the exported workbook's Search-sheet grammar —
comma-separated metric names, breakdown types, level values, campaign
names, and the "campaign" keyword (parsed by engine.search_terms). Only
ordering controls remain as widgets (Order + Show, also typeable as
"top 5" / "bottom 3" / "a-z"). Everything downstream is unchanged: the
same build_pivot preview, Apply as Value/Table/Chart, metric naming, and
saved-query behavior.
"""

import logging
import tkinter as tk
from ui.utils import fit_window
from tkinter import ttk, messagebox

import pandas as pd

from engine.pivot import build_pivot, pivot_total  # noqa: F401 — re-exported;
# moved to engine/ so the query resolver can reuse it without tkinter
from engine.query_resolver import get_available_breakdowns
from engine.search_terms import (SORT_LABELS, apply_sort, default_metric,
                                 parse_search_terms)

logger = logging.getLogger(__name__)


def _auto_query_name(query, existing):
    """Readable sidebar name for an unnamed applied query, unique among
    ``existing`` names."""
    base = f"Query: {query.get('metric', 'Value')} ({query.get('agg', 'sum')})"
    name = base
    counter = 2
    while name in existing:
        name = f"{base} {counter}"
        counter += 1
    return name


def show_query_builder(wizard):
    """Open the query search; writes the chosen query/metric back
    onto the wizard (selected_metric, _pending_query, _pending_table_data,
    _pending_chart_data)."""
    t = wizard.t

    win = tk.Toplevel(wizard.window)
    win.title("Query Search")
    fit_window(win, 950, 700)
    win.configure(bg=t["bg"]); win.transient(wizard.window); win.grab_set()
    win.lift(); win.focus_force()

    # Collect raw data into a DataFrame for filtering. The scan is
    # O(all level rows), so it is cached per mapper session — client_data
    # never changes for an open wizard, and rebuilding it on every open of
    # the query builder made large imports feel frozen.
    cache = getattr(wizard, "_query_builder_cache", None)
    if cache is None:
        raw_rows = []
        all_campaigns = set()
        for data in wizard.export_result["client_data"]:
            for mdata in data.get("campaign_metrics", {}).values():
                mn = mdata.get("universal_name", "")
                mv = mdata.get("value", 0)
                camp = mdata.get("campaign_name", "")
                if mn and isinstance(mv, (int, float)):
                    raw_rows.append({"campaign": camp, "metric": mn, "value": mv,
                                     "source": "campaign", "level_value": ""})
                    all_campaigns.add(camp)
            for ld in data.get("level_data", []):
                mn = ld.get("metric_name", "")
                mv = ld.get("metric_value", 0)
                ml = ld.get("metric_level", "")
                camp = ld.get("_campaign", "")
                prefix = ml.split(":")[0] if ":" in ml else ml
                lv = ml.split(":", 1)[1] if ":" in ml else ""
                if mn and isinstance(mv, (int, float)):
                    raw_rows.append({"campaign": camp, "metric": mn, "value": mv,
                                     "source": prefix, "level_value": lv})
                    all_campaigns.add(camp)
        cache = {
            "df": pd.DataFrame(raw_rows) if raw_rows else None,
            "campaigns": sorted(all_campaigns),
            "breakdowns": get_available_breakdowns(
                wizard.export_result["client_data"]),
        }
        wizard._query_builder_cache = cache

    if cache["df"] is None:
        messagebox.showwarning("No Data", "No metric data available.")
        win.destroy(); return

    full_df = cache["df"]
    metrics_list = sorted(full_df["metric"].unique())
    campaigns_list = cache["campaigns"]
    breakdowns = cache["breakdowns"]

    # ── Search row (the Search-sheet grammar, same as the Excel export) ──
    search_row = tk.Frame(win, bg=t["bg"])
    search_row.pack(fill="x", padx=10, pady=(10, 2))
    tk.Label(search_row, text="🔍 Search:", font=("Calibri", 11, "bold"),
             bg=t["bg"], fg=t["fg"]).pack(side="left", padx=(0, 5))
    search_var = tk.StringVar()
    search_entry = tk.Entry(search_row, textvariable=search_var, width=64,
                            font=("Calibri", 11), bg=t["input_bg"],
                            fg=t["input_fg"], insertbackground=t["insert"],
                            relief="solid", borderwidth=1)
    search_entry.pack(side="left", fill="x", expand=True)
    search_entry.focus_set()

    tk.Label(win, text="Comma-separated terms, same as the workbook's "
                       "Search sheet — e.g.   28167, impressions   ·   "
                       "campaign, network, impressions   ·   "
                       "top 5, zip, impressions.  Blank = campaign totals.",
             font=("Calibri", 8), bg=t["bg"], fg=t["muted"], wraplength=920,
             justify="left").pack(padx=10, anchor="w")

    # ── Order controls (the only remaining widgets; also typeable) ──
    ctrl = tk.Frame(win, bg=t["bg"])
    ctrl.pack(fill="x", padx=10, pady=(4, 2))
    tk.Label(ctrl, text="Agg:", font=("Calibri", 10, "bold"), bg=t["bg"],
             fg=t["fg"]).pack(side="left", padx=(0, 3))
    agg_var = tk.StringVar(value="sum")
    agg_box = ttk.Combobox(ctrl, textvariable=agg_var, state="readonly",
                           values=["sum", "avg", "max", "min", "count"],
                           width=7)
    agg_box.pack(side="left", padx=(0, 12))

    tk.Label(ctrl, text="Order:", font=("Calibri", 10, "bold"), bg=t["bg"],
             fg=t["fg"]).pack(side="left", padx=(0, 3))
    sort_by_label = {label: key for key, label in SORT_LABELS}
    label_by_sort = {key: label for key, label in SORT_LABELS}
    sort_var = tk.StringVar(value=label_by_sort["largest"])
    sort_box = ttk.Combobox(ctrl, textvariable=sort_var, state="readonly",
                            values=list(sort_by_label), width=13)
    sort_box.pack(side="left", padx=(0, 12))

    tk.Label(ctrl, text="Show:", font=("Calibri", 10, "bold"), bg=t["bg"],
             fg=t["fg"]).pack(side="left", padx=(0, 3))
    show_var = tk.StringVar(value="all")
    show_box = ttk.Combobox(ctrl, textvariable=show_var,
                            values=["all", "3", "5", "10", "15", "20"],
                            width=5)
    show_box.pack(side="left")

    ignored_label = tk.Label(win, text="", font=("Calibri", 8, "italic"),
                             bg=t["bg"], fg="#B8860B", wraplength=920,
                             justify="left")
    ignored_label.pack(padx=10, anchor="w")

    # Output kind is chosen by the Apply buttons at the bottom
    output_var = tk.StringVar(value="value")

    # Result label
    result_label = tk.Label(win, text="", font=("Calibri", 13, "bold"),
                             bg=t["bg"], fg=t["fg"])
    result_label.pack(padx=10, anchor="w")

    def refresh_pivot():
        parsed, ignored = parse_search_terms(
            search_var.get(), metrics_list, breakdowns, campaigns_list)
        ignored_label.config(
            text="  ·  ".join(f"“{term}” ignored — {why}"
                              for term, why in ignored))
        # Typed ordering terms ("top 5", "a-z") drive the controls, so the
        # widgets always show what the table is doing
        if parsed["sort"]:
            sort_var.set(label_by_sort[parsed["sort"]])
        if parsed["top_n"] != "all":
            show_var.set(parsed["top_n"])

        metric = parsed["metric"] or default_metric(metrics_list)
        agg = agg_var.get()
        sort = sort_by_label.get(sort_var.get(), "largest")
        top_n = show_var.get().strip() or "all"
        # Campaigns: typed names narrow the search; none typed = all (the
        # old panel's select-all default)
        sel_camps = parsed["campaigns"] or campaigns_list

        pivot, note = build_pivot(full_df, metric, agg, "all",
                                  sel_camps, parsed["sources"],
                                  parsed["values"])
        pivot = apply_sort(pivot, sort, top_n)
        if pivot is None or pivot.empty:
            result_label.config(text=note or "No data matches the search")
            tree.delete(*tree.get_children())
            pivot_result["df"] = None
            return

        total = pivot["Total"].sum() if "Total" in pivot.columns else pivot.iloc[:, -1].sum()
        vd = f"{total:,.0f}" if isinstance(total, (int, float)) else str(total)
        result_label.config(text=f"Total: {vd}  ({len(pivot)} rows)" + note)

        tree.delete(*tree.get_children())
        cols = list(pivot.columns)
        tree["columns"] = cols
        tree.heading("#0", text=pivot.index.name or "Label")
        tree.column("#0", width=200)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=110, anchor="e", stretch=False)
        for idx, row in pivot.iterrows():
            vals = [f"{v:,.0f}" if isinstance(v, (int, float)) else str(v) for v in row]
            tree.insert("", "end", text=str(idx), values=vals)

        pivot_result["df"] = pivot
        pivot_result["query"] = {"metric": metric,
                                  "campaigns": sel_camps,
                                  "sources": parsed["sources"],
                                  "values": parsed["values"],
                                  "agg": agg, "top_n": top_n, "sort": sort,
                                  "search": search_var.get()}

        # Refresh the export-column selector with the pivot's columns,
        # preserving any current selection
        prev_sel = {col_listbox.get(i) for i in col_listbox.curselection()}
        col_listbox.delete(0, "end")
        label_col = pivot.index.name or "Label"
        export_cols = [label_col] + list(pivot.columns)
        for c in export_cols:
            col_listbox.insert("end", c)
        for i, c in enumerate(export_cols):
            if not prev_sel or c in prev_sel:
                col_listbox.select_set(i)

    # Bottom buttons — packed side="bottom" so they can never be squeezed
    # out by the expanding preview (they vanished after the v16.31 reorder)
    bottom = tk.Frame(win, bg=t["bg"])
    bottom.pack(side="bottom", fill="x", padx=10, pady=(5, 8))

    # R3: name the query — the result becomes a reusable metric under this
    # name (assignable in the mapper; recomputed automatically on refill)
    tk.Label(bottom, text="Metric name:", font=("Calibri", 9),
             bg=t["bg"], fg=t["fg"]).pack(side="left", padx=(0, 4))
    qname_var = tk.StringVar()
    tk.Entry(bottom, textvariable=qname_var, width=18,
             font=("Calibri", 9)).pack(side="left", padx=(0, 12))

    def apply_value():
        q = pivot_result.get("query", {})
        if not q:
            refresh_pivot()
            q = pivot_result.get("query", {})
        if not q: return
        # Every applied query becomes a visible sidebar metric: the typed
        # name if given, otherwise a readable auto-name. Apply used to arm
        # an invisible selection when the name was blank — nothing appeared
        # in Metrics & Values and users assumed the button did nothing.
        if not hasattr(wizard, "named_queries"):
            wizard.named_queries = {}
        key = qname_var.get().strip() or _auto_query_name(q, wizard.named_queries)
        # Resolve value
        piv = pivot_result.get("df")
        if piv is not None:
            wizard.available_metrics[key] = pivot_total(piv)

        # Which columns did the user pick for export?
        sel_cols = [col_listbox.get(i) for i in col_listbox.curselection()]
        label_col = (piv.index.name or "Label") if piv is not None else "Label"
        include_label = (not sel_cols) or (label_col in sel_cols)
        if piv is not None:
            data_cols = [c for c in piv.columns
                         if (not sel_cols) or (c in sel_cols)]
        else:
            data_cols = []

        # Custom headers override (comma-separated)
        custom = [h.strip() for h in custom_headers_var.get().split(",") if h.strip()]

        out = output_var.get()
        if out == "table" and piv is not None:
            headers = ([label_col] if include_label else []) + data_cols
            if custom:
                headers = custom[:len(headers)] + headers[len(custom):]
            rows_data = []
            for idx, row in piv.iterrows():
                cells = [str(idx)] if include_label else []
                cells += [f"{row[c]:,.0f}" if isinstance(row[c], (int, float)) else str(row[c])
                          for c in data_cols]
                rows_data.append(cells)
            wizard._pending_table_data = {
                "headers": headers, "rows": rows_data,
                "keep_header_row": keep_header_var.get(),
            }
            q["output"] = "table"
        elif out == "chart" and piv is not None:
            categories = [str(i) for i in piv.index]
            series = [{"name": col, "values": piv[col].tolist()}
                      for col in data_cols if col != "Total"]
            if custom:
                for si, s in enumerate(series):
                    if si < len(custom):
                        s["name"] = custom[si]
            wizard._pending_chart_data = {"categories": categories, "series": series}
            q["output"] = "chart"
        else:
            q["output"] = "value"
        wizard.named_queries[key] = {"query": dict(q)}
        wizard.selected_metric = key
        wizard._pending_query = q
        wizard._pending_image = None
        # The applied query appears in the sidebar (Saved Queries) armed for
        # assignment — click a shape to place it, or re-click it later.
        try:
            if hasattr(wizard, "_refresh_metrics"):
                wizard._refresh_metrics()
        except Exception:
            logger.warning("Sidebar refresh after query apply failed",
                           exc_info=True)
        win.destroy()

    tk.Button(bottom, text="Apply as Value", font=("Calibri", 10),
              bg=t["accent"], fg="white", relief="flat", padx=15, pady=6,
              command=apply_value).pack(side="left", padx=(0, 5))
    tk.Button(bottom, text="Apply as Table", font=("Calibri", 10),
              bg="#1E6E3E", fg="white", relief="flat", padx=15, pady=6,
              command=lambda: [output_var.set("table"), apply_value()]).pack(side="left", padx=(0, 5))
    tk.Button(bottom, text="Apply as Chart Data", font=("Calibri", 10),
              bg="#8B4513", fg="white", relief="flat", padx=15, pady=6,
              command=lambda: [output_var.set("chart"), apply_value()]).pack(side="left")
    tk.Button(bottom, text="Cancel", font=("Calibri", 10),
              bg=t["secondary"], fg=t["secondary_fg"], relief="flat", padx=15, pady=6,
              command=win.destroy).pack(side="right")

    # Bottom controls pack FIRST (side=bottom) so they can never be
    # squeezed off short screens; the scrollable preview absorbs the
    # deficit instead (same fix class as the review action bar).
    refresh_btn = tk.Button(win, text="🔄 Refresh Preview", font=("Calibri", 10, "bold"),
                             bg=t["accent"], fg="white", relief="flat", padx=15, pady=4,
                             command=refresh_pivot)
    refresh_btn.pack(side="bottom", padx=10, pady=3, anchor="w")
    export_frame = tk.LabelFrame(win, text="  Export Options  ", font=("Calibri", 9, "bold"),
                                  bg=t["card"], fg=t["card_fg"])
    export_frame.pack(side="bottom", fill="x", padx=10, pady=(0, 3))

    col_wrap = tk.Frame(export_frame, bg=t["card"])
    col_wrap.pack(side="left", padx=5, pady=3)
    tk.Label(col_wrap, text="Columns to export:", font=("Calibri", 8, "bold"),
             bg=t["card"], fg=t["fg"]).pack(anchor="w")
    col_listbox = tk.Listbox(col_wrap, selectmode="multiple", height=4, width=28,
                              font=("Calibri", 8), bg=t["input_bg"], fg=t["input_fg"],
                              exportselection=False)
    col_listbox.pack()
    col_btns = tk.Frame(col_wrap, bg=t["card"])
    col_btns.pack(fill="x")
    tk.Button(col_btns, text="All", font=("Calibri", 7), bg=t["secondary"],
              fg=t["secondary_fg"], relief="flat", padx=4,
              command=lambda: col_listbox.select_set(0, "end")).pack(side="left", padx=1)
    tk.Button(col_btns, text="None", font=("Calibri", 7), bg=t["secondary"],
              fg=t["secondary_fg"], relief="flat", padx=4,
              command=lambda: col_listbox.select_clear(0, "end")).pack(side="left", padx=1)

    hdr_wrap = tk.Frame(export_frame, bg=t["card"])
    hdr_wrap.pack(side="left", padx=15, pady=3, anchor="n")
    keep_header_var = tk.BooleanVar(value=True)
    tk.Checkbutton(hdr_wrap, text="Table already has headers\n(write data below row 1)",
                   variable=keep_header_var, font=("Calibri", 8), justify="left",
                   bg=t["card"], fg=t["fg"], selectcolor=t["input_bg"]
                   ).pack(anchor="w")
    tk.Label(hdr_wrap, text="Custom headers (comma-separated, optional):",
             font=("Calibri", 8), bg=t["card"], fg=t["muted"]).pack(anchor="w", pady=(6, 0))
    custom_headers_var = tk.StringVar()
    tk.Entry(hdr_wrap, textvariable=custom_headers_var, width=40, font=("Calibri", 8),
             bg=t["input_bg"], fg=t["input_fg"], insertbackground=t["insert"],
             relief="solid", borderwidth=1).pack(anchor="w")
    tk.Label(hdr_wrap, text="Applies to Table/Chart output. Column picks apply to all outputs.",
             font=("Calibri", 7), bg=t["card"], fg=t["muted"]).pack(anchor="w", pady=(2, 0))

    # Pivot table preview
    pivot_frame = tk.LabelFrame(win, text="  Data Preview  ", font=("Calibri", 10, "bold"),
                                 bg=t["card"], fg=t["card_fg"])
    pivot_frame.pack(fill="both", expand=True, padx=10, pady=5)
    tree_sy = ttk.Scrollbar(pivot_frame, orient="vertical")
    tree_sx = ttk.Scrollbar(pivot_frame, orient="horizontal")
    tree = ttk.Treeview(pivot_frame, yscrollcommand=tree_sy.set, xscrollcommand=tree_sx.set)
    tree_sy.config(command=tree.yview); tree_sx.config(command=tree.xview)
    tree_sy.pack(side="right", fill="y"); tree_sx.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)

    pivot_result = {"df": None, "query": None}

    # Live refresh: debounced typing (250 ms), Enter for immediate, and
    # the order controls
    _debounce = [None]

    def _search_changed(_event=None):
        if _debounce[0] is not None:
            try:
                win.after_cancel(_debounce[0])
            except Exception:
                logger.debug("Search debounce cancel failed", exc_info=True)
        _debounce[0] = win.after(250, refresh_pivot)

    search_entry.bind("<KeyRelease>", _search_changed)
    search_entry.bind("<Return>", lambda e: refresh_pivot())
    for box in (agg_box, sort_box, show_box):
        box.bind("<<ComboboxSelected>>", lambda e: refresh_pivot())
    show_box.bind("<Return>", lambda e: refresh_pivot())
    refresh_pivot()
