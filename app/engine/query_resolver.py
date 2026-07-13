"""
Metric Query Resolver

Resolves metric queries against client data.
A query defines: what metric, from what breakdown, with what filter, what aggregation.

Simple mode: pre-built options like "Total Impressions", "Client Name", "Date Range"
Advanced mode: custom queries with metric + breakdown + filter + aggregation

All queries are generic and work across any client.
"""

from datetime import datetime
from numbers import Real

import pandas as pd


# Pre-built simple options that work for any client
SIMPLE_OPTIONS = [
    # Special
    {"key": "__client_name__", "label": "📋 Client Name", "category": "special"},
    {"key": "__date_range__", "label": "📅 Date Range", "category": "special"},
    {"key": "__start_date__", "label": "📅 Start Date", "category": "special"},
    {"key": "__end_date__", "label": "📅 End Date", "category": "special"},
    {"key": "__start_month__", "label": "📅 Start Month", "category": "special"},
    {"key": "__end_month__", "label": "📅 End Month", "category": "special"},
    {"key": "__year__", "label": "📅 Year", "category": "special"},
    {"key": "__image__", "label": "🖼 Browse Image", "category": "special"},
]


def build_simple_options(structured_metrics):
    """Build the list of simple dropdown options from available data."""
    options = list(SIMPLE_OPTIONS)

    totals = structured_metrics.get("totals", {})
    display_aliases = {"100% Completions": "Completions", "100% Complete": "Completions"}
    from parsers.dictionary import get_metric_aggregation

    # Add total (or average, for ratio metrics) for each metric found
    for metric in sorted(totals.keys()):
        display = display_aliases.get(metric, metric)
        agg = get_metric_aggregation(metric)
        word = "Avg" if agg == "avg" else "Total"
        options.append({
            "key": f"__total_{metric}__",
            "label": f"📊 {word} {display}",
            "category": "total",
            "query": {"metric": metric, "breakdown": "all", "filter": "all", "agg": agg},
        })

    # Add breakdown totals
    breakdowns = structured_metrics.get("breakdowns", {})
    for source in sorted(breakdowns.keys()):
        source_title = source.title()
        for metric in sorted(totals.keys()):
            display = display_aliases.get(metric, metric)
            agg = get_metric_aggregation(metric)
            word = "Avg" if agg == "avg" else "Total"
            options.append({
                "key": f"__total_{source}_{metric}__",
                "label": f"  {source_title} {word}: {display}",
                "category": f"breakdown_{source}",
                "query": {"metric": metric, "breakdown": source, "filter": "all", "agg": agg},
            })

    return options


def resolve_query(query, client_data, client_name="", start_date="", end_date=""):
    """Resolve a metric query against client data.

    Args:
        query: dict with keys: metric, breakdown, filter, agg
               OR a string key like "__client_name__", "__total_Impressions__"
        client_data: list of parsed data dicts
        client_name, start_date, end_date: client info

    Returns: the resolved value (number, string, etc.)
    """
    # Handle string keys (simple options)
    if isinstance(query, str):
        if query == "__client_name__": return client_name
        if query == "__date_range__":
            return f"{start_date} - {end_date}" if start_date and end_date else start_date or end_date
        if query == "__start_date__": return start_date
        if query == "__end_date__": return end_date
        if query == "__start_month__":
            try:
                return datetime.strptime(start_date.strip(), "%Y-%m-%d").strftime("%B")
            except (ValueError, AttributeError):
                return start_date
        if query == "__end_month__":
            try:
                return datetime.strptime(end_date.strip(), "%Y-%m-%d").strftime("%B")
            except (ValueError, AttributeError):
                return end_date
        if query == "__year__":
            try:
                return str(datetime.strptime(start_date.strip(), "%Y-%m-%d").year)
            except (ValueError, AttributeError):
                return ""
        if query == "__image__": return ""
        # Parse total keys: __total_Impressions__ or
        # __total_device_Impressions__. Match the longest known source prefix
        # so underscores inside source or metric names remain unambiguous.
        if query.startswith("__total_") and query.endswith("__"):
            inner = query[8:-2]
            breakdown = "all"
            metric = inner
            for source in sorted(_get_sources(client_data), key=len, reverse=True):
                prefix = source + "_"
                if inner.startswith(prefix):
                    breakdown = source
                    metric = inner[len(prefix):]
                    break
            from parsers.dictionary import get_metric_aggregation
            return resolve_query(
                {"metric": metric, "breakdown": breakdown, "filter": "all",
                 "agg": get_metric_aggregation(metric)},
                client_data, client_name, start_date, end_date)
        return ""

    # Handle dict queries
    if not isinstance(query, dict):
        return ""

    metric = query.get("metric", "")
    breakdown = query.get("breakdown", "all")
    filt = query.get("filter", "all")
    agg = query.get("agg", "sum")

    # Collect data into DataFrame
    rows = []
    for data in client_data:
        for mdata in data.get("campaign_metrics", {}).values():
            mn = mdata.get("universal_name", "")
            mv = mdata.get("value", 0)
            camp = mdata.get("campaign_name", "")
            if mn and isinstance(mv, Real) and not isinstance(mv, bool):
                rows.append({"campaign": camp, "metric": mn, "value": mv,
                             "source": "campaign", "level_value": ""})
        for ld in data.get("level_data", []):
            mn = ld.get("metric_name", "")
            mv = ld.get("metric_value", 0)
            ml = ld.get("metric_level", "")
            camp = ld.get("_campaign", "")
            prefix = ml.split(":")[0] if ":" in ml else ml
            lv = ml.split(":", 1)[1] if ":" in ml else ""
            if mn and isinstance(mv, Real) and not isinstance(mv, bool):
                rows.append({"campaign": camp, "metric": mn, "value": mv,
                             "source": prefix, "level_value": lv})

    if not rows:
        return 0

    df = pd.DataFrame(rows)

    # Filter to requested metric
    df = df[df["metric"] == metric]
    if df.empty:
        return 0

    # Filter to a requested source first. For "all", source selection is
    # performed per campaign below so a campaign is not dropped merely because
    # another campaign has a larger total from a different export type.
    if breakdown != "all":
        df = df[df["source"] == breakdown]

    if df.empty:
        return 0

    # Apply a campaign or level-value filter before choosing the best source.
    if filt != "all" and filt:
        wanted = str(filt).casefold()
        mask = (df["level_value"].astype(str).str.casefold() == wanted) | \
               (df["campaign"].astype(str).str.casefold() == wanted)
        df = df[mask]

    if df.empty:
        return 0

    if breakdown == "all":
        if agg in ("avg", "mean"):
            # Prefer a campaign-summary row when present. Otherwise use the
            # source with the most observations for that campaign, then average
            # one value per campaign. This avoids averaging duplicate exports.
            grouped = df.groupby(["campaign", "source"], as_index=False).agg(
                value=("value", "mean"), observations=("value", "size"))
            grouped["preferred"] = (grouped["source"] == "campaign").astype(int)
            grouped = grouped.sort_values(
                ["campaign", "preferred", "observations"],
                ascending=[True, False, False])
            df = grouped.drop_duplicates("campaign", keep="first")
        else:
            source_totals = df.groupby(["campaign", "source"], as_index=False)["value"].sum()
            best = source_totals.loc[
                source_totals.groupby("campaign")["value"].idxmax()]
            if agg == "sum":
                df = best
            else:
                chosen = best[["campaign", "source"]]
                df = df.merge(chosen, on=["campaign", "source"], how="inner")

    if df.empty:
        return 0

    if agg == "sum":
        result = df["value"].sum()
    elif agg in ("avg", "mean"):
        result = df["value"].mean()
    elif agg == "max":
        result = df["value"].max()
    elif agg == "min":
        result = df["value"].min()
    elif agg == "count":
        result = len(df)
    elif agg == "first":
        result = df["value"].iloc[0]
    else:
        result = df["value"].sum()

    item = getattr(result, "item", None)
    return item() if callable(item) else result


def _get_sources(client_data):
    """Get all breakdown source types from client data."""
    sources = set()
    for data in client_data:
        for ld in data.get("level_data", []):
            ml = ld.get("metric_level", "")
            if ":" in ml:
                sources.add(ml.split(":")[0])
    return sources


def get_available_breakdowns(client_data):
    """Get available breakdown types and their unique values."""
    breakdowns = {}
    for data in client_data:
        for ld in data.get("level_data", []):
            ml = ld.get("metric_level", "")
            if ":" in ml:
                prefix = ml.split(":")[0]
                lv = ml.split(":", 1)[1]
                breakdowns.setdefault(prefix, set()).add(lv)
    return {k: sorted(v) for k, v in breakdowns.items()}


