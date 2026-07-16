"""Phase 3 of the template-first pipeline: data sources → named slots.

A slot mapping reuses the EXACT query schema the current mapper persists
(owner decision 2026-07-16): each slot maps to either a simple string key
("__client_name__", "__total_Impressions__", ...) or a query dict —
including Advanced Query Builder queries, which re-resolve through the
same pivot the builder displayed (engine.query_resolver, fidelity fixed
in v1.32.0). Values format through engine.pptx_formats with the slot's
placeholder text as context, byte-for-byte the pptx_fill display logic.

mapping.json lives next to template.json in the template store directory.
"""

import json
import logging
import os

from engine.pptx_formats import _coerce_number, _format_value, format_with_details
from engine.query_resolver import resolve_query, resolve_query_payload
from engine.template_ir.schema import load_template_ir

logger = logging.getLogger(__name__)

MAPPING_JSON_NAME = "mapping.json"
MAPPING_SCHEMA_VERSION = "1.0"

# What kind of value a query produces, for map-time type validation
_TEXT_KEYS = {"__client_name__"}
_DATE_KEYS = {"__date_range__", "__start_date__", "__end_date__",
              "__start_month__", "__end_month__", "__year__"}


def new_slot_mapping(template_id):
    return {"schema_version": MAPPING_SCHEMA_VERSION,
            "template_id": template_id, "slots": {}}


def save_slot_mapping(mapping, template_dir):
    path = os.path.join(template_dir, MAPPING_JSON_NAME)
    from config.settings import _atomic_json_write
    _atomic_json_write(path, mapping)
    return path


def load_slot_mapping(template_dir):
    """The stored slot mapping, or None when the template has none yet."""
    path = os.path.join(template_dir, MAPPING_JSON_NAME)
    if not os.path.isfile(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _query_value_kind(query):
    """'text' | 'date' | 'number' | 'chart_data' | 'table_data' — what a
    query resolves to. Builder queries carry the output the user applied
    them as (value/table/chart)."""
    if isinstance(query, dict):
        output = query.get("output", "value")
        if output == "table":
            return "table_data"
        if output == "chart":
            return "chart_data"
        return "number"
    if query in _TEXT_KEYS:
        return "text"
    if query in _DATE_KEYS:
        return "date"
    if isinstance(query, str) and query.startswith("__total_"):
        return "number"
    return "text"


def validate_slot_mapping(ir, mapping):
    """Map-time validation (the slot registry's whole point): returns a
    list of human-readable warnings — unknown slots, unmapped slots, and
    type mismatches ('text value mapped to a number slot')."""
    warnings = []
    registry = ir.slot_registry
    slots = mapping.get("slots", {})
    for name in slots:
        if name not in registry:
            warnings.append(f"'{name}' is not a slot in this template "
                            "(renamed or re-classified?)")
    for name, spec in registry.items():
        entry = slots.get(name)
        if not entry or not entry.get("query"):
            warnings.append(f"'{name}' has no data source mapped")
            continue
        expected = spec.get("type", "text")
        actual = _query_value_kind(entry["query"])
        if expected in ("number", "date") and actual != expected:
            warnings.append(f"'{name}' expects a {expected} but the mapped "
                            f"source produces {actual}")
        elif expected == "chart_data" and actual != "chart_data":
            warnings.append(f"'{name}' is a chart — map an Advanced Query "
                            "Builder query applied as Chart Data")
        elif expected == "table_data" and actual != "table_data":
            warnings.append(f"'{name}' is a table — map an Advanced Query "
                            "Builder query applied as Table")
    return warnings


def resolve_slot_values(ir, mapping, client_data, client_name="",
                        start_date="", end_date=""):
    """Resolve every mapped slot to its display string.

    Returns (values, issues): values is {slot_name: formatted string} for
    slots that resolved; issues is [(slot_name, reason)] for slots that
    could not resolve — callers surface them, never drop them.
    """
    values, issues = {}, []
    registry = ir.slot_registry
    for name, entry in mapping.get("slots", {}).items():
        query = entry.get("query")
        if not query:
            issues.append((name, "no data source mapped"))
            continue
        slot_type = registry.get(name, {}).get("type", "text")
        try:
            if slot_type in ("chart_data", "table_data"):
                kind = "chart" if slot_type == "chart_data" else "table"
                payload = resolve_query_payload(query, client_data, kind)
                if payload is None:
                    issues.append((name, "query matched no data rows"))
                else:
                    values[name] = payload
                continue
            value = resolve_query(query, client_data, client_name,
                                  start_date, end_date)
        except Exception as e:
            logger.exception("Slot query failed: %s", name)
            issues.append((name, f"query failed: {e}"))
            continue
        context = registry.get(name, {}).get("placeholder_text", "")
        values[name] = _display_value(value, entry.get("format", "text"),
                                      entry.get("format_details"), context)
    return values, issues


def _display_value(value, fmt, details, context):
    """Format a resolved value exactly the way pptx_fill displays
    assignments (date details first, numeric details next, else the
    context-matching default path)."""
    coerced = _coerce_number(value)
    if details and details.get("format") == "date":
        return format_with_details(coerced, details)
    if (details and isinstance(coerced, (int, float))
            and not isinstance(coerced, bool)):
        return format_with_details(coerced, details)
    return _format_value(value, fmt, context)


def build_mapped_report(template_dir, output_path, client_data,
                        client_name="", start_date="", end_date=""):
    """One-call monthly build: load template + mapping, resolve slots,
    build the deck. Returns (output_path, report) — the build report plus
    resolution issues folded into slots_unfilled with their real reasons.
    """
    from engine.template_ir.build import build_from_template
    ir = load_template_ir(template_dir)
    mapping = load_slot_mapping(template_dir)
    if mapping is None:
        raise FileNotFoundError(
            f"No slot mapping saved for this template yet ({template_dir})")
    values, issues = resolve_slot_values(ir, mapping, client_data,
                                         client_name, start_date, end_date)
    path, report = build_from_template(template_dir, output_path,
                                       slot_values=values)
    # A slot that failed resolution shows its actual failure, not the
    # builder's generic "no value provided" — and a resolution issue the
    # builder never saw (e.g. a mapped slot no longer in the template) is
    # still reported, never dropped.
    reasons = dict(issues)
    unfilled = []
    for slot, reason in report["slots_unfilled"]:
        if reason == "no value provided" and slot in reasons:
            reason = reasons[slot]
        reasons.pop(slot, None)
        unfilled.append((slot, reason))
    unfilled.extend(reasons.items())
    report["slots_unfilled"] = unfilled
    return path, report
