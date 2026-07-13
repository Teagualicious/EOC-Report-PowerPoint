"""Pure data helpers for the Excel writer: normalization, collection, pivot.

Everything here is UI-free and file-free (pandas in, pandas out), which is
what makes the Excel pipeline unit-testable.
"""

import numpy as np
import pandas as pd

SCHEMA_COLUMNS = [
    "client", "campaign", "campaign_type", "source",
    "metric_level", "metric_name", "metric_value", "start_date", "end_date"
]
KEY_COLUMNS = ["client", "campaign", "campaign_type", "source",
               "metric_level", "metric_name", "start_date", "end_date"]

DISPLAY_NAMES = {"100% Completions": "Completions", "100% Complete": "Completions"}
TOP_N = 10


def _display_name(m):
    return DISPLAY_NAMES.get(m, m)


def _normalize_value(raw):
    if raw is None or raw == "":
        return np.nan
    if isinstance(raw, (int, float)):
        return raw
    s = str(raw).strip()
    is_pct = "%" in s
    cleaned = s.replace("$", "").replace(",", "").replace("%", "").strip()
    try:
        val = float(cleaned)
        if is_pct and val > 1:
            val = val / 100.0
        return val
    except (ValueError, TypeError):
        return raw


def _normalize_text(text, style="title"):
    if not text or not isinstance(text, str):
        return text or ""
    text = text.strip()
    if style == "title":
        return text.title() if text == text.upper() or text == text.lower() else text
    return text


def _format_cell(cell, val):
    if isinstance(val, float) and not np.isnan(val):
        cell.number_format = "0.00%" if val < 1 else "#,##0.00"
    elif isinstance(val, (int, np.integer)):
        cell.number_format = "#,##0"


def _collect_rows(parsed_data_list, platform_name=""):
    """Collect all data into a flat list of row dicts."""
    rows = []

    for data in parsed_data_list:
        sd = data.get("start_date", "")
        ed = data.get("end_date", "")
        client = _normalize_text(data.get("client_name", ""), "title")
        ctype = _normalize_text(data.get("campaign_type", ""), "title")
        source = _normalize_text(data.get("source_platform", platform_name), "title")
        cname = _normalize_text(data.get("campaign_name", ""), "title")

        # Campaign-level metrics
        cm = data.get("campaign_metrics", {})
        if cm:
            for md in cm.values():
                cc = md.get("campaign_name", "")
                cc = _normalize_text(str(cc), "title") if cc else cname
                u = md.get("universal_name", "")
                v = _normalize_value(md.get("value"))
                if pd.isna(v) or not u:
                    continue
                rows.append({
                    "client": client, "campaign": cc, "campaign_type": ctype,
                    "source": source, "metric_level": "",
                    "metric_name": str(u).strip(),
                    "metric_value": v, "start_date": sd, "end_date": ed,
                })
        elif data.get("mapped_metrics"):
            for mn, rv in data["mapped_metrics"].items():
                nv = _normalize_value(rv)
                if pd.isna(nv):
                    continue
                rows.append({
                    "client": client, "campaign": cname, "campaign_type": ctype,
                    "source": source, "metric_level": "",
                    "metric_name": str(mn).strip(),
                    "metric_value": nv, "start_date": sd, "end_date": ed,
                })

        # Level breakdown data
        for ld in data.get("level_data", []):
            ml = ld.get("metric_level", "")
            mn = ld.get("metric_name", "")
            mv = _normalize_value(ld.get("metric_value"))
            if pd.isna(mv) or not ml or not mn:
                continue
            lc = ld.get("_campaign", "")
            lc = _normalize_text(str(lc), "title") if lc else cname
            rows.append({
                "client": client, "campaign": lc, "campaign_type": ctype,
                "source": source, "metric_level": ml,
                "metric_name": str(mn).strip(),
                "metric_value": mv, "start_date": sd, "end_date": ed,
            })

    return rows


def _collapse_rows(rows):
    """Collapse duplicate rows within one import batch.

    Additive metrics are summed; ratio metrics are averaged. Text values keep
    the last occurrence. This is deliberately batch-local so a re-export can
    replace the prior batch instead of doubling it.
    """
    if not rows:
        return pd.DataFrame(columns=SCHEMA_COLUMNS)

    df = pd.DataFrame(rows, columns=SCHEMA_COLUMNS)
    for col in KEY_COLUMNS:
        df[col] = df[col].fillna("").astype(str)

    numeric_mask = df["metric_value"].map(
        lambda value: isinstance(value, (int, float, np.integer, np.floating))
        and not isinstance(value, (bool, np.bool_)))
    numeric = df[numeric_mask].copy()
    text = df[~numeric_mask].copy()

    collapsed = []
    if not numeric.empty:
        from parsers.dictionary import get_metric_aggregation
        numeric["metric_value"] = pd.to_numeric(numeric["metric_value"], errors="coerce")
        numeric["_aggregation"] = numeric["metric_name"].map(get_metric_aggregation)
        additive = numeric[numeric["_aggregation"] != "avg"]
        ratios = numeric[numeric["_aggregation"] == "avg"]
        if not additive.empty:
            collapsed.append(additive.groupby(KEY_COLUMNS, as_index=False, sort=False).agg(
                {"metric_value": "sum"}))
        if not ratios.empty:
            collapsed.append(ratios.groupby(KEY_COLUMNS, as_index=False, sort=False).agg(
                {"metric_value": "mean"}))

    if not text.empty:
        collapsed.append(text.drop_duplicates(subset=KEY_COLUMNS, keep="last"))

    if not collapsed:
        return pd.DataFrame(columns=SCHEMA_COLUMNS)
    return pd.concat(collapsed, ignore_index=True)[SCHEMA_COLUMNS]


def _build_dataframe(new_rows, existing_rows=None):
    """Build the report DataFrame with re-export replacement semantics.

    Exact keys inside the new import batch are aggregated. If the workbook
    already contains the same key, the newly imported value replaces it. Rows
    from different date ranges or sources remain appended.
    """
    new_df = _collapse_rows(new_rows)
    old_df = _collapse_rows(existing_rows or [])
    if old_df.empty:
        return new_df
    if new_df.empty:
        return old_df

    new_keys = pd.MultiIndex.from_frame(new_df[KEY_COLUMNS])
    old_keys = pd.MultiIndex.from_frame(old_df[KEY_COLUMNS])
    old_df = old_df[~old_keys.isin(new_keys)]
    return pd.concat([old_df, new_df], ignore_index=True)[SCHEMA_COLUMNS]

