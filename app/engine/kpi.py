"""KPI calculation for the review screen.

The KPI rule: sum each metric per campaign per source type, then take the
MAX across sources. This handles exports that have no campaign summary
sheet — the most complete breakdown becomes the campaign total.
"""

import pandas as pd

KPI_METRICS = ["Impressions", "Clicks", "Completions", "Cost"]
KPI_ALIASES = {
    "100% Completions": "Completions",
    "100% Complete": "Completions",
    "Contributions": "Completions",
    "Contribution": "Completions",
}


def compute_kpis(client_data):
    """Compute KPI totals, per-campaign details, and data-quality flags.

    Args:
        client_data: list of parsed data dicts for one client.

    Returns:
        (kpi_totals, campaign_details, flags) where
        kpi_totals is {display_metric: value} summed over campaigns,
        campaign_details is {campaign: {display_metric: value}}, and
        flags is a list of (campaign, metric, reason) tuples.
    """
    all_rows = []
    flags = []
    for data in client_data:
        for mdata in data.get("campaign_metrics", {}).values():
            camp = mdata.get("campaign_name", "")
            metric = mdata.get("universal_name", "")
            val = mdata.get("value", 0)
            display = KPI_ALIASES.get(metric, metric)
            if isinstance(val, (int, float)):
                all_rows.append({"campaign": camp, "metric": display, "value": val, "source": "campaign"})
            if isinstance(val, (int, float)) and val == 0:
                flags.append((camp, metric, "Zero value"))
        for ld in data.get("level_data", []):
            camp = ld.get("_campaign", "")
            ml = ld.get("metric_level", "")
            mn = ld.get("metric_name", "")
            mv = ld.get("metric_value", 0)
            prefix = ml.split(":")[0] if ":" in ml else ml
            display = KPI_ALIASES.get(mn, mn)
            if isinstance(mv, (int, float)):
                all_rows.append({"campaign": camp, "metric": display, "value": mv, "source": prefix})

    kpi_totals = {}
    campaign_details = {}
    if all_rows:
        from parsers.dictionary import get_metric_aggregation
        df = pd.DataFrame(all_rows)
        df["agg"] = df["metric"].map(get_metric_aggregation)

        # Additive metrics: sum per campaign per source, best source wins
        add_df = df[df["agg"] == "sum"]
        if not add_df.empty:
            # Suspicion check: an additive metric whose row values are ALL
            # tiny (≤1.5) across many rows is probably a rate/percentage
            # mis-aliased into an additive column — summing it produces
            # nonsense totals like "Completions: 1.00"
            for metric, grp in add_df.groupby("metric"):
                if len(grp) >= 5 and grp["value"].abs().max() <= 1.5:
                    flags.append(
                        f"'{metric}' values are all ≤ 1.5 across {len(grp)} rows — "
                        f"this looks like a rate/percentage, not a count. Check the "
                        f"column aliases in the metric dictionary; its total may be "
                        f"meaningless.")
            source_totals = add_df.groupby(["campaign", "metric", "source"])["value"].sum().reset_index()
            best = source_totals.loc[source_totals.groupby(["campaign", "metric"])["value"].idxmax()]
            for _, row in best.iterrows():
                campaign_details.setdefault(row["campaign"], {})[row["metric"]] = row["value"]
            for metric in best["metric"].unique():
                kpi_totals[metric] = best[best["metric"] == metric]["value"].sum()

        # Ratio metrics (Frequency, CTR, CPM...): summing is meaningless —
        # average per campaign, then average across campaigns
        ratio_df = df[df["agg"] == "avg"]
        if not ratio_df.empty:
            camp_avg = ratio_df.groupby(["campaign", "metric"])["value"].mean().round(2).reset_index()
            for _, row in camp_avg.iterrows():
                campaign_details.setdefault(row["campaign"], {})[row["metric"]] = row["value"]
            for metric in camp_avg["metric"].unique():
                kpi_totals[metric] = round(
                    camp_avg[camp_avg["metric"] == metric]["value"].mean(), 2)

        _derive_completion_rate(kpi_totals, campaign_details)
        _relabel_non_dedup_totals(kpi_totals, flags)

    return kpi_totals, campaign_details, flags


# Cross-campaign Reach/Frequency need household-level deduplication only
# the vendor can do (campaign aggregates carry no overlap information, so a
# 3x gap against vendor dashboards is possible). Approved 2026-07-13: keep
# the values but label them honestly and flag the caveat. Per-campaign
# values in campaign_details are vendor-computed and stay as-is.
_NON_DEDUP_RELABELS = {
    "Reach": "Combined Reach (not deduplicated)",
    "Frequency": "Avg Campaign Frequency",
}
_NON_DEDUP_FLAGS = {
    "Reach": ("Reach total is campaign reach summed WITHOUT household "
              "deduplication — the vendor's true order reach will be lower. "
              "Take order-level reach from the vendor dashboard."),
    "Frequency": ("Frequency total is the average of per-campaign "
                  "frequencies — the vendor's order frequency (impressions ÷ "
                  "deduplicated reach) will be higher. Take order-level "
                  "frequency from the vendor dashboard."),
}


def _relabel_non_dedup_totals(kpi_totals, flags):
    for metric, label in _NON_DEDUP_RELABELS.items():
        if metric in kpi_totals:
            kpi_totals[label] = kpi_totals.pop(metric)
            flags.append(_NON_DEDUP_FLAGS[metric])


_RATE_NAMES = ("Completion Rate", "Completion Percent")
_NUMERATORS = ("Completions", "Contributions", "100% Completions")
# Vendor dashboards compute VCR against video starts; approved rule
# (2026-07-13): use starts when the export carries them, impressions
# otherwise. Completions/impressions read ~6 points lower than the vendor
# because not every impression starts the video.
_DENOMINATORS = ("Video Starts", "Impressions")


def _derive_completion_rate(kpi_totals, campaign_details):
    """Approved rule: rate = 100 * Σcompletions / Σvideo starts, falling
    back to Σimpressions when the export has no starts metric. Overrides
    the naive average (which overweights small campaigns) for totals and
    per-campaign details whenever the inputs exist."""
    def rate_from(d):
        denom = next((d[n] for n in _DENOMINATORS if d.get(n)), None)
        if not denom:
            return None
        for n in _NUMERATORS:
            if d.get(n) is not None:
                return round(100.0 * d[n] / denom, 2)
        return None
    total = rate_from(kpi_totals)
    if total is not None:
        target = next((r for r in _RATE_NAMES if r in kpi_totals),
                      "Completion Rate")
        kpi_totals[target] = total
        for det in campaign_details.values():
            r = rate_from(det)
            if r is not None:
                det[target] = r
