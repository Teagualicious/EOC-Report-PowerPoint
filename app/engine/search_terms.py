"""The Excel search's term grammar, for the query builder UI.

The exported workbook's Search sheet takes comma-separated terms — metric
names, breakdown types, level values, campaign names, the "campaign"
keyword — and renders the matching table. This module parses the SAME
grammar against the client's parsed data and maps it onto the query
builder's existing query fields (metric/campaigns/sources/values), so the
builder keeps its logic (build_pivot, apply-as-value/table/chart, saved
queries) behind a search box instead of multi-select filter panels.

Extra display terms the search box adds over the Search sheet: "top N" /
"bottom N" and "a-z" / "z-a" ordering. tk-free — fully covered by the
automated suite.
"""

import re

_TOP_RE = re.compile(r"^(top|bottom)\s*(\d+)$", re.IGNORECASE)

# sort keys: "largest" (build_pivot's native order), "smallest", "az", "za"
SORT_LABELS = [("largest", "Largest first"), ("smallest", "Smallest first"),
               ("az", "A → Z"), ("za", "Z → A")]


def parse_search_terms(raw, metrics, breakdowns, campaigns):
    """Parse a comma-separated search string against the client's data.

    ``metrics``: metric names; ``breakdowns``: {source: [level values]};
    ``campaigns``: campaign names. Matching is case-insensitive, exact
    first, then unique-prefix (typing "netw" finds "network" as long as
    nothing else starts with it).

    Returns (parsed, ignored): parsed = {"metric", "sources", "values",
    "campaigns", "top_n", "sort"} in the builder's existing query
    vocabulary; ignored = [(term, why), ...] — surfaced in the UI, never
    silently dropped.
    """
    metric_map = {str(m).casefold(): str(m) for m in metrics}
    source_map = {str(s).casefold(): str(s) for s in breakdowns}
    value_map = {}
    for values in breakdowns.values():
        for value in values:
            value_map.setdefault(str(value).casefold(), str(value))
    campaign_map = {str(c).casefold(): str(c) for c in campaigns}

    parsed = {"metric": None, "sources": [], "values": [], "campaigns": [],
              "top_n": "all", "sort": None}
    ignored = []

    for term in [p.strip() for p in (raw or "").split(",")]:
        if not term:
            continue
        fold = term.casefold()

        top = _TOP_RE.match(fold)
        if top:
            parsed["top_n"] = top.group(2)
            parsed["sort"] = ("smallest" if top.group(1).casefold() == "bottom"
                              else "largest")
            continue
        if fold in ("a-z", "a→z", "az"):
            parsed["sort"] = "az"
            continue
        if fold in ("z-a", "z→a", "za"):
            parsed["sort"] = "za"
            continue
        if fold in ("campaign", "campaigns"):
            continue  # campaign grouping IS the no-breakdown default table

        kind, resolved = _match_term(fold, metric_map, source_map,
                                     value_map, campaign_map)
        if kind == "metric":
            if parsed["metric"] is None:
                parsed["metric"] = resolved
            else:
                ignored.append((term, "one metric per query — "
                                      f"already using {parsed['metric']}"))
        elif kind == "source":
            parsed["sources"].append(resolved)
        elif kind == "value":
            parsed["values"].append(resolved)
        elif kind == "campaign":
            parsed["campaigns"].append(resolved)
        else:
            ignored.append((term, "not a metric, breakdown type, value, "
                                  "or campaign"))

    # A typed level value implies its breakdown type(s) — "28167" alone
    # must surface the zip rows, exactly like the Search sheet.
    if parsed["values"] and not parsed["sources"]:
        wanted = {v.casefold() for v in parsed["values"]}
        parsed["sources"] = sorted(
            source for source, values in breakdowns.items()
            if any(str(v).casefold() in wanted for v in values))
    return parsed, ignored


def _match_term(fold, metric_map, source_map, value_map, campaign_map):
    """(kind, resolved_name) for one term — exact match first, then a
    prefix that is unique across ALL vocabularies; else (None, None)."""
    for kind, mapping in (("metric", metric_map), ("source", source_map),
                          ("value", value_map), ("campaign", campaign_map)):
        if fold in mapping:
            return kind, mapping[fold]
    prefix_hits = [(kind, name)
                   for kind, mapping in (("metric", metric_map),
                                         ("source", source_map),
                                         ("value", value_map),
                                         ("campaign", campaign_map))
                   for key, name in mapping.items() if key.startswith(fold)]
    if len(prefix_hits) == 1:
        return prefix_hits[0]
    return None, None


def default_metric(metrics):
    """The metric a blank/metric-less search shows — Impressions when the
    data has it (the Search sheet's primary default column), else the
    first metric alphabetically."""
    metrics = sorted(str(m) for m in metrics)
    for m in metrics:
        if m.casefold() == "impressions":
            return m
    return metrics[0] if metrics else ""


def apply_sort(pivot, sort, top_n="all"):
    """Order and slice a build_pivot result for display/apply.

    "largest" is build_pivot's native order (Total descending);
    "smallest" inverts it; "az"/"za" order by the label column. ``top_n``
    slices AFTER ordering, so "bottom 3" really is the three smallest.
    """
    if pivot is None:
        return None
    if sort == "smallest":
        column = "Total" if "Total" in pivot.columns else pivot.columns[-1]
        pivot = pivot.sort_values(column, ascending=True)
    elif sort in ("az", "za"):
        pivot = pivot.sort_index(
            ascending=(sort == "az"),
            key=lambda idx: idx.astype(str).str.casefold())
    if top_n != "all":
        try:
            pivot = pivot.head(int(top_n))
        except (TypeError, ValueError):
            pass
    return pivot
