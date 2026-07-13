"""Build the catalog of available metrics for template mapping.

Aggregates client data with pandas into flat metric keys (for filling)
plus a structured view (for the mapper sidebar).
"""

import logging
from datetime import datetime
from numbers import Real

logger = logging.getLogger(__name__)


def _to_native(val):
    """numpy scalar -> native Python; collapse float noise to int.

    pandas aggregations return numpy.int64/float64. numpy.int64 does not
    subclass Python int, so isinstance(v, (int, float)) checks fail
    downstream and formatting silently degrades to raw strings.
    """
    item = getattr(val, "item", None)
    if callable(item):
        try:
            val = val.item()
        except Exception:
            pass
    if isinstance(val, float) and abs(val - round(val)) < 0.0001:
        return int(round(val))
    return val

# Display names for special metrics
SPECIAL_METRICS = {
    "__client_name__": "📋 Client Name",
    "__date_range__": "📅 Date Range",
    "__start_date__": "📅 Start Date",
    "__end_date__": "📅 End Date",
    "__start_month__": "📅 Start Month",
    "__end_month__": "📅 End Month",
    "__year__": "📅 Year",
    "__image__": "🖼 Browse Image...",
}


def get_available_metrics(client_data, client_name="", start_date="", end_date=""):
    """Build structured metrics with generic totals and grouped breakdowns using pandas.

    Returns:
        metrics: flat dict of {key: value} for template filling
        structured: dict with categories for UI display
    """
    import pandas as pd

    metrics = {}

    # Special entries
    metrics["__client_name__"] = client_name
    if start_date and end_date:
        metrics["__date_range__"] = f"{start_date} - {end_date}"
    if start_date:
        metrics["__start_date__"] = start_date
        # Month and year from start date
        try:
            sd = datetime.strptime(start_date.strip(), "%Y-%m-%d")
            metrics["__start_month__"] = sd.strftime("%B")
            metrics["__year__"] = str(sd.year)
        except ValueError:
            logger.debug("Unparseable start date: %r", start_date)
    if end_date:
        metrics["__end_date__"] = end_date
        try:
            ed = datetime.strptime(end_date.strip(), "%Y-%m-%d")
            metrics["__end_month__"] = ed.strftime("%B")
            if "__year__" not in metrics:
                metrics["__year__"] = str(ed.year)
        except ValueError:
            logger.debug("Unparseable end date: %r", end_date)

    # Collect all data into a DataFrame
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
        return metrics, {"special": metrics, "totals": {}, "breakdowns": {}}

    df = pd.DataFrame(rows)
    from parsers.dictionary import get_metric_aggregation
    df["agg"] = df["metric"].map(get_metric_aggregation)
    add_df = df[df["agg"] == "sum"]
    ratio_df = df[df["agg"] == "avg"]

    kpi_totals = {}

    # GENERIC TOTALS (additive) — best source PER CAMPAIGN, then sum across
    # campaigns. (Matches engine.kpi.compute_kpis. The old global-max
    # approach dropped any campaign whose best source differed from the
    # group's best source — e.g. a campaign with no summary sheet vanished
    # from the deck total.)
    if not add_df.empty:
        camp_source = add_df.groupby(["campaign", "metric", "source"])["value"].sum().round(10).reset_index()
        best_per_campaign = camp_source.loc[
            camp_source.groupby(["campaign", "metric"])["value"].idxmax()]
        best_totals = best_per_campaign.groupby("metric")["value"].sum().reset_index()
        for _, row in best_totals.iterrows():
            total_key = f"Total {row['metric']}"
            val = _to_native(row["value"])
            metrics[total_key] = val
            kpi_totals[row["metric"]] = val

    # GENERIC AVERAGES (ratio metrics: Frequency, CTR, CPM...) — summing is
    # meaningless, so average per campaign then across campaigns, keyed
    # "Avg X" so the label is honest in templates.
    if not ratio_df.empty:
        camp_avg = ratio_df.groupby(["campaign", "metric"])["value"].mean().reset_index()
        metric_avg = camp_avg.groupby("metric")["value"].mean().reset_index()
        for _, row in metric_avg.iterrows():
            val = round(_to_native(row["value"]), 2)
            metrics[f"Avg {row['metric']}"] = val
            kpi_totals[row["metric"]] = val

    # BREAKDOWN TOTALS by source type (sum for additive, mean for ratio)
    for (metric, source), group in df.groupby(["metric", "source"]):
        if source == "campaign":
            continue
        if get_metric_aggregation(metric) == "avg":
            val = round(_to_native(group["value"].mean()), 2)
            key = f"{source.title()} Avg: {metric}"
        else:
            val = _to_native(group["value"].sum())
            key = f"{source.title()} Total: {metric}"
        metrics[key] = val

    # INDIVIDUAL BREAKDOWN VALUES (sum for additive, mean for ratio)
    if not df[df["source"] != "campaign"].empty:
        bd = df[df["source"] != "campaign"]
        for (source, level_value, metric), group in bd.groupby(
                ["source", "level_value", "metric"]):
            key = f"{source}:{level_value}:{metric}"
            if get_metric_aggregation(metric) == "avg":
                metrics[key] = round(_to_native(group["value"].mean()), 2)
            else:
                metrics[key] = _to_native(group["value"].sum())

    # Build structured data for UI
    structured = {
        "special": {k: v for k, v in metrics.items() if k.startswith("__")},
        "totals": kpi_totals,
        "breakdowns": {},
    }

    # Group breakdown data by source type. Ratio metrics must remain averages
    # in the mapper preview; summing CTR/Frequency here made the preview disagree
    # with the value inserted into PowerPoint.
    breakdown_df = df[df["source"] != "campaign"]
    if not breakdown_df.empty:
        for source in breakdown_df["source"].unique():
            source_df = breakdown_df[breakdown_df["source"] == source]
            additive = source_df[source_df["agg"] == "sum"].groupby(
                ["level_value", "metric"], as_index=False)["value"].sum()
            ratios = source_df[source_df["agg"] == "avg"].groupby(
                ["level_value", "metric"], as_index=False)["value"].mean()
            bd_agg = pd.concat([additive, ratios], ignore_index=True)
            pivoted = bd_agg.pivot_table(index="level_value", columns="metric",
                                          values="value", aggfunc="first").reset_index()
            structured["breakdowns"][source] = pivoted.to_dict("records")

    return metrics, structured
