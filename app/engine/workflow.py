"""Headless end-to-end workflow service — the engine behind the AI-facing
interfaces (``app/cli.py`` and ``app/mcp_server.py``).

Mirrors the interactive Quick Run flow (ui.main_window / ui.client_wizard /
ui.review_view) step for step, with no Tk imports anywhere on the path:

    parse_files -> list_campaigns -> client_dataset -> kpis_for
                                   -> export_workbook / fill_deck / query

Every function takes and returns plain JSON-friendly Python data so callers
can serialize results directly. Anything Windows/Office-specific keeps its
existing graceful degradation (VBA injection falls back to plain .xlsx with
a warning; deck filling always uses the python-pptx engine).
"""

import logging
import os

from config.paths import OUTPUT_DIR, TEMPLATES_DIR
from config.naming import unique_component
from config.settings import load_platform_config, load_settings
from engine.data_pipeline import (apply_platform_config,
                                  filter_data_by_campaigns,
                                  scan_file_structure)
from engine.errors import IngestionError

logger = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = (".csv", ".xlsx", ".xlsm", ".html", ".htm")


def _native(value):
    """Recursively convert numpy scalars to native Python so results
    serialize as JSON numbers, not strings."""
    if isinstance(value, dict):
        return {k: _native(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_native(v) for v in value]
    item = getattr(value, "item", None)
    if callable(item) and not isinstance(value, (str, bytes)):
        try:
            return value.item()
        except Exception:
            return value
    return value


def list_platforms():
    """Configured platforms and their sheet/column counts."""
    settings = load_settings()
    out = []
    for name in sorted(settings.get("platforms", {})):
        config = load_platform_config(name) or {}
        sheets = config.get("sheets", [])
        out.append({
            "name": name,
            "sheets": len(sheets),
            "columns": sum(len(s.get("columns", [])) for s in sheets),
        })
    return out


def list_templates():
    """Saved PowerPoint templates and whether each has a mapping."""
    from engine.pptx_mapper import list_available_templates
    return [{"filename": t["filename"], "has_mapping": t["has_mapping"]}
            for t in list_available_templates()]


def scan_export(path):
    """Inspect an export's structure (sheets, headers, sample row) —
    what Platform Setup shows when importing a sample file."""
    if not os.path.isfile(path):
        raise FileNotFoundError(f"No such file: {path}")
    return scan_file_structure(path)


def parse_files(file_platform_pairs):
    """Parse exports with their platform configs, exactly like Quick Run.

    Args:
        file_platform_pairs: list of (filepath, platform_name).

    Returns:
        (parsed_list, failures) where failures is [(basename, reason)].
    """
    from parsers.csv_parser import CSVParser
    from parsers.excel_parser import ExcelParser
    from parsers.html_parser import HTMLParser

    parsed = []
    failures = []
    for filepath, platform_name in file_platform_pairs:
        name = os.path.basename(filepath)
        if not os.path.isfile(filepath):
            failures.append((name, "File not found"))
            continue
        config = load_platform_config(platform_name)
        if not config:
            failures.append(
                (name, f"Platform '{platform_name}' has no saved config"))
            continue
        ext = os.path.splitext(filepath)[1].lower()
        try:
            if ext == ".csv":
                data = CSVParser(filepath).parse(build_outputs=False)
            elif ext in (".xlsx", ".xlsm"):
                data = ExcelParser(filepath).parse(build_outputs=False)
            elif ext in (".html", ".htm"):
                data = HTMLParser(filepath, platform_name).parse()
            else:
                failures.append((name, "Unsupported file type"))
                continue
            apply_platform_config(data, config)
            data["source_platform"] = platform_name
            parsed.append(data)
        except IngestionError as exc:
            logger.warning("Parse failed for %s: %s", filepath, exc)
            failures.append((name, exc.user_message))
        except Exception as exc:
            logger.exception("Parse failed for %s", filepath)
            failures.append((name, str(exc)))
    return parsed, failures


def list_campaigns(parsed):
    """Campaigns found in parsed data — what the client wizard offers."""
    campaigns = []
    seen = set()
    for data in parsed:
        plat = data.get("source_platform", "Unknown")
        for mdata in data.get("campaign_metrics", {}).values():
            cn = mdata.get("campaign_name", "")
            if cn and cn not in seen:
                campaigns.append({"name": cn, "platform": plat})
                seen.add(cn)
        for ld in data.get("level_data", []):
            cn = ld.get("_campaign", "")
            if cn and cn not in seen:
                campaigns.append({"name": cn, "platform": plat})
                seen.add(cn)
    campaigns.sort(key=lambda c: (c["platform"], c["name"]))
    return campaigns


def client_dataset(parsed, client_name, campaigns=None, start_date="",
                   end_date=""):
    """Filter parsed data to one client's campaigns and stamp identity,
    exactly like the client wizard hand-off. ``campaigns=None`` -> all."""
    if campaigns is None:
        campaigns = [c["name"] for c in list_campaigns(parsed)]
    client_data = filter_data_by_campaigns(parsed, campaigns)
    for data in client_data:
        data["client_name"] = client_name
        data["campaign_type"] = ""
        data["start_date"] = start_date
        data["end_date"] = end_date
    return client_data


def kpis_for(client_data):
    """Review-screen numbers: KPI totals, per-campaign detail, data flags."""
    from engine.kpi import compute_kpis
    totals, details, flags = compute_kpis(client_data)
    return _native({
        "totals": totals,
        "campaign_details": details,
        "flags": [list(f) if isinstance(f, tuple) else f for f in flags],
    })


def export_workbook(client_data, client_name, output_dir="",
                    inject_vba=True):
    """Write the unified Search/Data workbook for one client.

    Same layout as the interactive export: <output>/<safe client>/<safe
    client>_unified_report.xlsx (upgraded to .xlsm when VBA injection is
    available). Returns path, row count, and any VBA warning.
    """
    from engine.excel_writer import write_to_excel
    output_dir = output_dir or load_settings().get("output_folder", OUTPUT_DIR)
    safe_client = unique_component(client_name, set(), fallback="Client")
    client_folder = os.path.join(output_dir, safe_client)
    os.makedirs(client_folder, exist_ok=True)
    output_path = os.path.join(client_folder,
                               f"{safe_client}_unified_report.xlsx")
    final_path, vba_error = write_to_excel(client_data, output_path, "",
                                           inject_vba=inject_vba)
    rows = sum(len(d.get("campaign_metrics", {})) + len(d.get("level_data", []))
               for d in client_data)
    return {"path": final_path, "rows": rows, "vba_warning": vba_error}


def fill_deck(client_data, template_filename, output_path, client_name="",
              start_date="", end_date=""):
    """Fill a mapped PowerPoint template (python-pptx engine) and return
    the saved path plus the FillReport dict — same as review Auto-Fill."""
    from engine.fill_report import append_fill_history
    from engine.metrics_catalog import get_available_metrics
    from engine.pptx_fill import fill_template_report
    from engine.pptx_mapper import load_template_mapping

    template_path = os.path.join(TEMPLATES_DIR, template_filename)
    if not os.path.isfile(template_path):
        raise FileNotFoundError(
            f"Template not found in workspace/templates: {template_filename}")
    mapping = load_template_mapping(template_filename)
    if not mapping:
        raise ValueError(
            f"No saved mapping for {template_filename} — map it in the "
            "Template Mapper first.")

    if not start_date or not end_date:
        for d in client_data:
            start_date = start_date or d.get("start_date", "")
            end_date = end_date or d.get("end_date", "")
    metrics, _ = get_available_metrics(client_data, client_name,
                                       start_date, end_date)
    metrics["__client_data__"] = client_data
    metrics["__client_name__"] = client_name
    metrics["__start_date__"] = start_date
    metrics["__end_date__"] = end_date

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    path, report = fill_template_report(template_path, output_path,
                                        mapping, metrics)
    append_fill_history(report)
    return {"path": path, "report": report.to_dict()}


def query_metric(client_data, metric, breakdown="all", agg="sum",
                 metric_filter="all", client_name="", start_date="",
                 end_date=""):
    """Resolve an advanced query (same engine as mapped deck queries)."""
    from engine.query_resolver import resolve_query
    value = resolve_query(
        {"metric": metric, "breakdown": breakdown, "filter": metric_filter,
         "agg": agg},
        client_data, client_name, start_date, end_date)
    return _native({"metric": metric, "breakdown": breakdown, "agg": agg,
                    "filter": metric_filter, "value": value})
