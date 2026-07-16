"""Pure pivot computation shared by the Advanced Query Builder (preview,
Table/Chart outputs) and the query resolver (re-resolving saved builder
queries). Lives in engine/ — NOT mapper/ — so headless callers (fill
engine, CLI, MCP server) never import tkinter.

The long-format input DataFrame has columns:
    campaign, metric, value, source, level_value
("source" is the breakdown type prefix, e.g. "zip"; "campaign" rows carry
source == "campaign" and an empty level_value.)
"""

import logging

logger = logging.getLogger(__name__)


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


def pivot_total(pivot):
    """The single value an applied builder query represents — identical to
    what the builder shows/caches at apply time (query_builder.apply_value)."""
    total = (pivot["Total"].sum() if "Total" in pivot.columns
             else pivot.iloc[:, -1].sum())
    item = getattr(total, "item", None)
    return item() if callable(item) else total


def _cell(value):
    """One table cell, formatted the way the builder's Apply as Table
    renders it (query_builder.apply_value)."""
    from engine.pptx_formats import _coerce_number
    value = _coerce_number(value)   # numpy scalars fail isinstance(int/float)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{value:,.0f}"
    return str(value)


def pivot_to_table(pivot):
    """A pivot as table-fill payload: {"headers": [...], "rows": [[...]]} —
    label column first, then every pivot column, cells formatted like the
    builder's Apply as Table output."""
    label = pivot.index.name or "Label"
    headers = [str(label)] + [str(c) for c in pivot.columns]
    rows = [[str(idx)] + [_cell(row[c]) for c in pivot.columns]
            for idx, row in pivot.iterrows()]
    return {"headers": headers, "rows": rows}


def pivot_to_chart(pivot):
    """A pivot as chart-data payload: {"categories": [...], "series":
    [{"name", "values"}, ...]} — one series per pivot column, "Total"
    excluded, matching the builder's Apply as Chart Data output."""
    categories = [str(i) for i in pivot.index]
    columns = [c for c in pivot.columns if c != "Total"] or list(pivot.columns)
    series = [{"name": str(c), "values": [float(v) for v in pivot[c]]}
              for c in columns]
    return {"categories": categories, "series": series}
