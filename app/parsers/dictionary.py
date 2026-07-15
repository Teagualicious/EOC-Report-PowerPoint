"""
Shared metric dictionary — the universal translation layer.

All parsers (HTML, CSV, Excel) use this module for:
- Matching column/metric names to universal names via aliases
- Identifying context columns (campaign name, order)
- Identifying level identifier columns (device, geo, daypart, etc.)
- Identifying skip columns
- Classifying a set of column headers into context/level/metric/skip

This ensures consistent behavior regardless of file format.
"""

import json
import os

from config.paths import DICT_PATH

_cached_dict = None


def load_dictionary():
    global _cached_dict
    if _cached_dict is None:
        if os.path.exists(DICT_PATH):
            with open(DICT_PATH, "r") as f:
                _cached_dict = json.load(f)
        else:
            _cached_dict = {
                "metric_aliases": {},
                "level_definitions": {},
                "skip_columns": [],
                "context_columns": {
                    "campaign_name": ["Campaign Name", "Campaign", "Order"],
                },
            }
    return _cached_dict



def get_metric_aggregation(metric_name):
    """Aggregation type for a metric: 'sum' (additive, default) or 'avg'
    (ratio metrics like Frequency, CTR, CPM — summing them is meaningless)."""
    return load_dictionary().get("metric_aggregation", {}).get(metric_name, "sum")

def match_metric_alias(col_name, dictionary=None):
    """Match a column/metric name to its universal name. Returns universal name or None."""
    if dictionary is None:
        dictionary = load_dictionary()
    col_lower = col_name.strip().lower()
    for universal_name, aliases in dictionary.get("metric_aliases", {}).items():
        for alias in aliases:
            if alias.lower() == col_lower:
                return universal_name
    return None


def match_level(col_name, dictionary=None):
    """Match a column name to a level definition. Returns level_def dict or None."""
    if dictionary is None:
        dictionary = load_dictionary()
    col_lower = col_name.strip().lower()
    for level_def in dictionary.get("level_definitions", {}).values():
        for alias in level_def.get("id_column_aliases", []):
            if alias.lower() == col_lower:
                return level_def
    return None


def match_context(col_name, dictionary=None):
    """Match a column name to a context type (campaign_name, etc). Returns context key or None."""
    if dictionary is None:
        dictionary = load_dictionary()
    col_lower = col_name.strip().lower()
    for ctx_key, aliases in dictionary.get("context_columns", {}).items():
        for alias in aliases:
            if alias.lower() == col_lower:
                return ctx_key
    return None


def is_skip_column(col_name, dictionary=None):
    """Check if a column should be skipped."""
    if dictionary is None:
        dictionary = load_dictionary()
    col_lower = col_name.strip().lower()
    return any(s.lower() == col_lower for s in dictionary.get("skip_columns", []))


def classify_columns(header_names, dictionary=None):
    """
    Classify a list of column headers into categories.
    
    Returns:
        {
            "context": {"campaign_name": "Order", ...},  # ctx_key -> actual col name
            "levels": [("Device", level_def), ...],
            "metrics": [("Impressions", "Impressions"), ("Contribution", "Contributions"), ...],
                        # (original_col_name, universal_name)
            "skip": ["Flight Name", ...],
        }
    """
    if dictionary is None:
        dictionary = load_dictionary()

    result = {
        "context": {},
        "levels": [],
        "metrics": [],
        "skip": [],
    }

    for col_name in header_names:
        # Check context first
        ctx = match_context(col_name, dictionary)
        if ctx:
            result["context"][ctx] = col_name
            continue

        # Check level identifier
        ldef = match_level(col_name, dictionary)
        if ldef:
            result["levels"].append((col_name, ldef))
            continue

        # Check skip
        if is_skip_column(col_name, dictionary):
            result["skip"].append(col_name)
            continue

        # Check metric alias
        universal = match_metric_alias(col_name, dictionary)
        if universal:
            result["metrics"].append((col_name, universal))
        else:
            # Unknown column — include with original name as potential metric
            result["metrics"].append((col_name, col_name))

    return result


def extract_row_context(row_dict, context_map):
    """
    Extract context values from a data row.
    
    Args:
        row_dict: {"Order": "Acme Motors - Used VLA", "Impressions": 45230, ...}
        context_map: {"campaign_name": "Order"} from classify_columns
    
    Returns:
        {"campaign_name": "Acme Motors - Used VLA"}
    """
    result = {}
    for ctx_key, col_name in context_map.items():
        val = row_dict.get(col_name, "")
        if val is not None:
            result[ctx_key] = str(val).strip()
        else:
            result[ctx_key] = ""
    return result


def build_level_rows(row_dict, levels, metrics, context_values):
    """
    Build long-format level_data entries from a single data row.
    
    For each level column × each metric column, creates one entry.
    
    Returns list of dicts with metric_level, metric_name, metric_value, _campaign.
    """
    rows = []
    campaign = context_values.get("campaign_name", "")

    for level_col_name, ldef in levels:
        id_value = row_dict.get(level_col_name, "")
        if not id_value or id_value == "":
            continue
        prefix = ldef["prefix"]
        level_str = f"{prefix}:{id_value}"

        for metric_col_name, universal_name in metrics:
            raw_val = row_dict.get(metric_col_name, "")
            if raw_val == "" or raw_val is None:
                continue
            rows.append({
                "metric_level": level_str,
                "metric_name": universal_name,
                "metric_value": raw_val,
                "_campaign": campaign,
            })
    return rows


def build_campaign_rows(row_dict, metrics, context_values):
    """
    Build campaign-level metric entries from a single data row.
    
    Returns dict of {key: {value, universal_name, campaign_name}}.
    """
    campaign = context_values.get("campaign_name", "")
    result = {}
    for metric_col_name, universal_name in metrics:
        raw_val = row_dict.get(metric_col_name, "")
        if raw_val == "" or raw_val is None:
            continue
        mkey = f"{campaign}|{universal_name}" if campaign else universal_name
        result[mkey] = {
            "value": raw_val,
            "universal_name": universal_name,
            "campaign_name": campaign,
        }
    return result
