"""Advanced query builder popup: multi-select filters + visual pivot table."""

import logging
import tkinter as tk
from ui.utils import fit_window
from tkinter import ttk, messagebox

import pandas as pd

from engine.query_resolver import get_available_breakdowns

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


def build_pivot(full_df, metric, agg, topn, sel_camps, sel_sources, sel_vals):
    """Pure pivot computation behind the preview and Table/Chart outputs.

    Returns ``(pivot_df_or_None, note)``. Data-correctness rules — these
    tables ship in client reports:

    - With NO breakdown type selected, only campaign-total rows are used.
      Pooling every breakdown type counts the same delivery once per type,
      which surfaced as a giant bogus "Other" top row.
    - A level value that exists in more than one selected breakdown type
      ("Other" appears in zip AND dow AND network) stays as separate
      disambiguated rows — e.g. "Other (zip)" — instead of being silently
      summed across types.
    """
    df = full_df[full_df["metric"] == metric]
    if sel_camps:
        df = df[df["campaign"].isin(sel_camps)]
    if sel_sources:
        df = df[df["source"].isin(sel_sources)]
    else:
        # "Clean campaign-totals preview" — the documented default for an
        # empty Breakdown Type selection.
        df = df[df["source"] == "campaign"]
    if sel_vals:
        df = df[(df["level_value"].isin(sel_vals)) | (df["level_value"] == "")]

    if df.empty:
        return None, "No data matches filters"

    agg_func = {"sum": "sum", "avg": "mean", "max": "max", "min": "min",
                "count": "count"}.get(agg, "sum")

    note = ""
    has_levels = bool(df["level_value"].any())
    if has_levels:
        level_df = df[df["level_value"] != ""].copy()
        source_counts = level_df.groupby("level_value")["source"].nunique()
        ambiguous = set(source_counts[source_counts > 1].index)
        if ambiguous:
            mask = level_df["level_value"].isin(ambiguous)
            level_df.loc[mask, "level_value"] = (
                level_df.loc[mask, "level_value"] + " ("
                + level_df.loc[mask, "source"] + ")")
        pivot = level_df.pivot_table(
            index="level_value", columns="campaign",
            values="value", aggfunc=agg_func, fill_value=0)
        pivot["Total"] = pivot.sum(axis=1)
        pivot = pivot[pivot["Total"] != 0]   # all-zero rows are noise
        pivot = pivot.sort_values("Total", ascending=False)
        if pivot.empty:
            has_levels = False
            note = ("  (no " + metric + " at the selected level - "
                    "showing campaign totals)")
    if not has_levels:
        pivot = df.groupby("campaign")["value"].agg(agg_func).reset_index()
        pivot.columns = ["Campaign", metric]
        pivot = pivot.sort_values(metric, ascending=False).set_index("Campaign")

    if topn != "all":
        try:
            pivot = pivot.head(int(topn))
        except ValueError:
            logger.debug("Invalid Top N value: %r", topn)
    return pivot, note


def show_query_builder(wizard):
    """Open the advanced query builder; writes the chosen query/metric back
    onto the wizard (selected_metric, _pending_query, _pending_table_data,
    _pending_chart_data)."""
    t = wizard.t

    win = tk.Toplevel(wizard.window)
    win.title("Advanced Query Builder")
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
    sources_list = sorted(full_df["source"].unique())
    campaigns_list = cache["campaigns"]
    breakdowns = cache["breakdowns"]

    # ── Top controls row ──
    ctrl = tk.Frame(win, bg=t["bg"])
    ctrl.pack(fill="x", padx=10, pady=(10, 5))

    # Metric selector
    tk.Label(ctrl, text="Metric:", font=("Calibri", 10, "bold"), bg=t["bg"],
             fg=t["fg"]).pack(side="left", padx=(0, 3))
    metric_var = tk.StringVar(value=metrics_list[0] if metrics_list else "")
    ttk.Combobox(ctrl, textvariable=metric_var, values=metrics_list,
                 state="readonly", width=18).pack(side="left", padx=(0, 10))

    # Aggregation
    tk.Label(ctrl, text="Agg:", font=("Calibri", 10, "bold"), bg=t["bg"],
             fg=t["fg"]).pack(side="left", padx=(0, 3))
    agg_var = tk.StringVar(value="sum")
    ttk.Combobox(ctrl, textvariable=agg_var, values=["sum", "avg", "max", "min", "count"],
                 state="readonly", width=8).pack(side="left", padx=(0, 10))

    # Top N
    tk.Label(ctrl, text="Top N:", font=("Calibri", 10, "bold"), bg=t["bg"],
             fg=t["fg"]).pack(side="left", padx=(0, 3))
    topn_var = tk.StringVar(value="all")
    ttk.Combobox(ctrl, textvariable=topn_var, values=["all", "3", "5", "10", "15", "20"],
                 width=5).pack(side="left", padx=(0, 10))

    # Output kind is chosen by the Apply buttons at the bottom - the old
    # dropdown duplicated them and confused the flow
    output_var = tk.StringVar(value="value")

    # ── Filter panels (3 columns) ──
    filter_frame = tk.Frame(win, bg=t["bg"])
    filter_frame.pack(fill="x", padx=10, pady=5)
    filter_frame.columnconfigure(0, weight=1)
    filter_frame.columnconfigure(1, weight=1)
    filter_frame.columnconfigure(2, weight=1)

    # Campaign filter
    cf = tk.LabelFrame(filter_frame, text="  Campaigns  ", font=("Calibri", 9, "bold"),
                        bg=t["card"], fg=t["card_fg"])
    cf.grid(row=0, column=0, sticky="nsew", padx=3)
    camp_listbox = tk.Listbox(cf, selectmode="extended", height=6, font=("Calibri", 9),
                               bg=t["input_bg"], fg=t["input_fg"], exportselection=False)
    camp_listbox.pack(fill="both", expand=True, padx=3, pady=3)
    for c in campaigns_list:
        camp_listbox.insert("end", c)
    # Select all by default
    camp_listbox.select_set(0, "end")
    camp_btns = tk.Frame(cf, bg=t["card"])
    camp_btns.pack(fill="x", padx=3, pady=(0, 3))
    tk.Button(camp_btns, text="All", font=("Calibri", 8), bg=t["secondary"],
              fg=t["secondary_fg"], relief="flat", padx=5,
              command=lambda: camp_listbox.select_set(0, "end")).pack(side="left", padx=1)
    tk.Button(camp_btns, text="None", font=("Calibri", 8), bg=t["secondary"],
              fg=t["secondary_fg"], relief="flat", padx=5,
              command=lambda: camp_listbox.select_clear(0, "end")).pack(side="left", padx=1)

    # Breakdown source filter
    bf = tk.LabelFrame(filter_frame, text="  Breakdown Type  ", font=("Calibri", 9, "bold"),
                        bg=t["card"], fg=t["card_fg"])
    bf.grid(row=0, column=1, sticky="nsew", padx=3)
    bd_listbox = tk.Listbox(bf, selectmode="extended", height=6, font=("Calibri", 9),
                             bg=t["input_bg"], fg=t["input_fg"], exportselection=False)
    bd_listbox.pack(fill="both", expand=True, padx=3, pady=3)
    for s in sources_list:
        bd_listbox.insert("end", s)
    # default: no breakdown selected -> clean campaign-totals preview
    bd_btns = tk.Frame(bf, bg=t["card"])
    bd_btns.pack(fill="x", padx=3, pady=(0, 3))
    tk.Button(bd_btns, text="All", font=("Calibri", 8), bg=t["secondary"],
              fg=t["secondary_fg"], relief="flat", padx=5,
              command=lambda: bd_listbox.select_set(0, "end")).pack(side="left", padx=1)
    tk.Button(bd_btns, text="None", font=("Calibri", 8), bg=t["secondary"],
              fg=t["secondary_fg"], relief="flat", padx=5,
              command=lambda: bd_listbox.select_clear(0, "end")).pack(side="left", padx=1)

    # Value filter (level values)
    vf = tk.LabelFrame(filter_frame, text="  Values  ", font=("Calibri", 9, "bold"),
                        bg=t["card"], fg=t["card_fg"])
    vf.grid(row=0, column=2, sticky="nsew", padx=3)
    val_listbox = tk.Listbox(vf, selectmode="extended", height=6, font=("Calibri", 9),
                              bg=t["input_bg"], fg=t["input_fg"], exportselection=False)
    val_listbox.pack(fill="both", expand=True, padx=3, pady=3)
    val_btns = tk.Frame(vf, bg=t["card"])
    val_btns.pack(fill="x", padx=3, pady=(0, 3))
    tk.Button(val_btns, text="All", font=("Calibri", 8), bg=t["secondary"],
              fg=t["secondary_fg"], relief="flat", padx=5,
              command=lambda: val_listbox.select_set(0, "end")).pack(side="left", padx=1)
    tk.Button(val_btns, text="None", font=("Calibri", 8), bg=t["secondary"],
              fg=t["secondary_fg"], relief="flat", padx=5,
              command=lambda: val_listbox.select_clear(0, "end")).pack(side="left", padx=1)

    def update_values(*args):
        val_listbox.delete(0, "end")
        sel_sources = [bd_listbox.get(i) for i in bd_listbox.curselection()]
        vals = set()
        for src in sel_sources:
            for v in breakdowns.get(src, []):
                vals.add(v)
        for v in sorted(vals):
            val_listbox.insert("end", v)
        val_listbox.select_set(0, "end")

    bd_listbox.bind("<<ListboxSelect>>", update_values)
    update_values()

    # Result label
    result_label = tk.Label(win, text="", font=("Calibri", 13, "bold"),
                             bg=t["bg"], fg=t["fg"])
    result_label.pack(padx=10, anchor="w")

    def refresh_pivot():
        metric = metric_var.get()
        agg = agg_var.get()
        topn = topn_var.get()
        sel_camps = [camp_listbox.get(i) for i in camp_listbox.curselection()]
        sel_sources = [bd_listbox.get(i) for i in bd_listbox.curselection()]
        sel_vals = [val_listbox.get(i) for i in val_listbox.curselection()]

        pivot, note = build_pivot(full_df, metric, agg, topn,
                                  sel_camps, sel_sources, sel_vals)
        if pivot is None:
            result_label.config(text=note or "No data matches filters")
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
        pivot_result["query"] = {"metric": metric, "campaigns": sel_camps,
                                  "sources": sel_sources, "values": sel_vals,
                                  "agg": agg, "top_n": topn}

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
            total = piv["Total"].sum() if "Total" in piv.columns else piv.iloc[:, -1].sum()
            wizard.available_metrics[key] = total

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

    # Refresh button


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

    # ── Export options: column selection + header handling ──


    # Auto-refresh on metric/agg change
    for var in [metric_var, agg_var, topn_var]:
        var.trace_add("write", lambda *a: refresh_pivot())
    refresh_pivot()


