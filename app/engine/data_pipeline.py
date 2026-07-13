"""Headless data pipeline: file scanning, platform configs, campaign filtering.

These functions take parsed-data dicts (from the parsers package) and are
UI-free, which makes them the primary unit-test surface of the app.
"""

import logging
import os

from parsers.csv_parser import CSVParser
from parsers.dictionary import classify_columns, match_metric_alias
from parsers.excel_parser import ExcelParser
from parsers.html_parser import HTMLParser

logger = logging.getLogger(__name__)


def scan_file_structure(filepath):
    """Parse a sample file and describe its sheets/tables for platform setup."""
    ext = os.path.splitext(filepath)[1].lower()
    sheets = []
    if ext == ".csv":
        data = CSVParser(filepath).parse(build_outputs=False)
        for tbl in data.get("detected_tables", []):
            classification = classify_columns(tbl["headers"])
            sheets.append({"name": tbl.get("sheet_name", os.path.basename(filepath)),
                           "headers": tbl["headers"], "row_count": tbl["row_count"],
                           "classification": classification,
                           "sample_row": tbl["rows"][0] if tbl["rows"] else {}})
    elif ext in (".xlsx", ".xlsm"):
        data = ExcelParser(filepath).parse(build_outputs=False)
        for tbl in data.get("detected_tables", []):
            classification = classify_columns(tbl["headers"])
            sheets.append({"name": tbl.get("sheet_name", "Sheet"),
                           "headers": tbl["headers"], "row_count": tbl["row_count"],
                           "classification": classification,
                           "sample_row": tbl["rows"][0] if tbl["rows"] else {}})
    elif ext in (".html", ".htm"):
        data = HTMLParser(filepath, "setup").parse()
        if data.get("metrics"):
            headers = list(data["metrics"].keys())
            classification = classify_columns(headers)
            sheets.append({"name": "HTML Report", "headers": headers, "row_count": 1,
                           "classification": classification, "sample_row": data["metrics"]})
        for tbl in data.get("detected_tables", []):
            classification = classify_columns(tbl["headers"])
            sheets.append({"name": f"Table ({', '.join(tbl['headers'][:3])}...)",
                           "headers": tbl["headers"], "row_count": tbl["row_count"],
                           "classification": classification,
                           "sample_row": tbl["rows"][0] if tbl["rows"] else {}})
    return sheets


def apply_platform_config(parsed_data, platform_config):
    """Rebuild campaign_metrics/level_data on parsed_data using column roles."""
    if not platform_config: return
    sheet_configs = {sc["sheet_name"]: sc for sc in platform_config.get("sheets", [])}
    new_campaign_metrics = {}; new_level_data = []
    detected = list(parsed_data.get("detected_tables", []))
    # Metric-only HTML dashboards may not expose a conventional table. Treat
    # the parser's flat metric scan as a one-row pseudo sheet so the platform
    # mapping can still be applied.
    if parsed_data.get("metrics"):
        detected.append({
            "sheet_name": "HTML Report",
            "rows": [dict(parsed_data["metrics"])],
            "headers": list(parsed_data["metrics"]),
        })

    matched_config = False
    default_campaign = parsed_data.get("campaign_name", "")
    for tbl in detected:
        sn = tbl.get("sheet_name", "")
        cfg = sheet_configs.get(sn)
        if not cfg:
            for key in sheet_configs:
                if key.lower() in sn.lower() or sn.lower() in key.lower():
                    cfg = sheet_configs[key]; break
        if not cfg:
            if len(sheet_configs) == 1 and len(detected) == 1:
                cfg = list(sheet_configs.values())[0]
        if not cfg: continue
        matched_config = True

        campaign_cols = []; level_cols = []; metric_cols = []
        for col_cfg in cfg["columns"]:
            cn = col_cfg["column"]; role = col_cfg.get("role", "skip")
            sel = col_cfg.get("selected", False)
            if not sel: continue
            if role == "campaign_id": campaign_cols.append(cn)
            elif role.startswith("level: "):
                level_cols.append((cn, role.replace("level: ", "")))
            elif role == "metric":
                universal = match_metric_alias(cn)
                metric_cols.append((cn, universal or cn))

        skip_labels = ("total", "totals", "grand total", "summary", "subtotal")
        for row in tbl.get("rows", []):
            parts = [str(row.get(cc, "")).strip() for cc in campaign_cols if row.get(cc)]
            row_campaign = " - ".join(parts) if parts else default_campaign
            if row_campaign and row_campaign.lower().strip() in skip_labels: continue
            is_total = any(str(row.get(cc, "")).strip().lower() in skip_labels for cc in campaign_cols)
            if is_total: continue

            if level_cols:
                for lc, prefix in level_cols:
                    idv = row.get(lc, "")
                    if not idv or str(idv).strip() == "": continue
                    ls = f"{prefix}:{str(idv).strip()}"
                    for mc, universal in metric_cols:
                        rv = row.get(mc, "")
                        if rv == "" or rv is None: continue
                        new_level_data.append({"metric_level": ls, "metric_name": universal,
                                               "metric_value": rv, "_campaign": row_campaign})
            else:
                for mc, universal in metric_cols:
                    rv = row.get(mc, "")
                    if rv == "" or rv is None: continue
                    mk = f"{row_campaign}|{universal}" if row_campaign else universal
                    new_campaign_metrics[mk] = {"value": rv, "universal_name": universal,
                                                "campaign_name": row_campaign}

    if not matched_config:
        logger.warning("No platform-config sheet matched %s; preserving parser output",
                       parsed_data.get("source_file", "the source file"))
        return

    parsed_data["campaign_metrics"] = new_campaign_metrics
    parsed_data["level_data"] = new_level_data
    flat = {v["universal_name"]: v["value"] for v in new_campaign_metrics.values()}
    parsed_data["metrics"] = flat
    parsed_data["mapped_metrics"] = flat.copy()



def filter_data_by_campaigns(all_parsed, campaign_names):
    """Create a filtered copy of parsed data containing only specified campaigns."""
    filtered = []
    target = set(campaign_names)
    for data in all_parsed:
        new_data = dict(data)
        new_cm = {k: v for k, v in data.get("campaign_metrics", {}).items()
                  if v.get("campaign_name", "") in target}
        new_ld = [ld for ld in data.get("level_data", [])
                  if ld.get("_campaign", "") in target]
        if new_cm or new_ld:
            new_data["campaign_metrics"] = new_cm
            new_data["level_data"] = new_ld
            filtered.append(new_data)
    return filtered
