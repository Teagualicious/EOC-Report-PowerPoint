"""Headless Deck Engine workflow service.

The CLI and future Tk shell call only these functions. Fork Stage 0 establishes
this dependency boundary; later stages implement parse, stage, and build behind
these stable verbs. No function in this module imports Tk.
"""

from __future__ import annotations

import os
from typing import Any, Mapping

from config.paths import (OUTPUT_DIR, PROJECT_ROOT, STAGING_DIR, TEMPLATES_DIR,
                          ensure_dirs)
from config.settings import load_settings, save_settings

SUPPORTED_EXTENSIONS = (".csv", ".xlsx", ".xlsm", ".html", ".htm")


def _native(value: Any) -> Any:
    """Recursively convert numpy/pandas scalars to JSON-native values."""
    if isinstance(value, dict):
        return {str(key): _native(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_native(item) for item in value]
    item = getattr(value, "item", None)
    if callable(item) and not isinstance(value, (str, bytes)):
        try:
            return item()
        except Exception:
            return value
    return value


def _version() -> str:
    path = os.path.join(PROJECT_ROOT, "VERSION")
    try:
        with open(path, "r", encoding="utf-8") as handle:
            return handle.read().strip()
    except OSError:
        return "0.0.0-unknown"


def list_templates() -> list[dict[str, Any]]:
    """Return saved PowerPoint templates and mapping availability."""
    from engine.pptx_mapper import list_available_templates
    return [
        {"filename": item["filename"], "has_mapping": bool(item["has_mapping"])}
        for item in list_available_templates()
    ]


def get_settings() -> dict[str, Any]:
    """Return normalized application settings."""
    return load_settings()


def set_settings(updates: Mapping[str, Any]) -> dict[str, Any]:
    """Merge and atomically persist settings."""
    current = load_settings()
    current.update(dict(updates))
    return save_settings(current)


def describe_state() -> dict[str, Any]:
    """Return the JSON-friendly state used by support and the future UI."""
    ensure_dirs()
    return {
        "version": _version(),
        "phase": "Fork Stage 0",
        "analyst_ui_available": False,
        "paths": {
            "templates": TEMPLATES_DIR,
            "staging": STAGING_DIR,
            "output": OUTPUT_DIR,
        },
        "settings": get_settings(),
        "templates": list_templates(),
    }


def parse_dump(path: str, *, profile: Mapping[str, Any] | None = None) -> dict[str, Any]:
    """Parse one campaign export into Unified Data (implemented Stage 1)."""
    del path, profile
    raise NotImplementedError("Dump parsing is implemented in Fork Stage 1")


def generate_staging(dump_path: str, template_name: str | None = None) -> dict[str, Any]:
    """Generate the analyst staging workbook (implemented Stage 2)."""
    del dump_path, template_name
    raise NotImplementedError("Staging generation is implemented in Fork Stage 2")


def build_deck(staging_path: str | None = None) -> dict[str, Any]:
    """Build a deck only from a saved staging workbook (implemented Stage 4)."""
    del staging_path
    raise NotImplementedError("Deck building is implemented in Fork Stage 4")

# ── Developer-only template authoring surface ────────────────────────────────
# These commands are intentionally absent from the analyst UI but remain as the
# tested maintenance path for app/mapper and engine/template_ir.


def ingest_template_store(pptx_path: str, template_id: str | None = None) -> dict[str, Any]:
    """Ingest or reconcile a deck in the developer template-first store."""
    from config.naming import storage_key
    from engine.template_ir import (classify_template, ingest_template,
                                    load_template_ir, reingest_template,
                                    save_template_ir)
    from engine.template_ir.ingest import TEMPLATE_STORE_DIR
    from engine.template_ir.schema import TEMPLATE_JSON_NAME

    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"No such file: {pptx_path}")
    template_id = template_id or storage_key(
        os.path.splitext(os.path.basename(pptx_path))[0], replace_dots=True)
    template_dir = os.path.join(TEMPLATE_STORE_DIR, template_id)

    if os.path.isfile(os.path.join(template_dir, TEMPLATE_JSON_NAME)):
        template_dir, deltas = reingest_template(pptx_path, template_dir)
    else:
        template_dir = ingest_template(pptx_path, template_id=template_id)
        ir = classify_template(load_template_ir(template_dir))
        save_template_ir(ir, template_dir)
        deltas = []

    ir = load_template_ir(template_dir)
    shapes = [
        {
            "shape_id": shape.shape_id,
            "name": shape.name,
            "type": shape.shape_type,
            "slide": slide.slide_index + 1,
            "classification": shape.classification,
            "reason": shape.classify_reason,
            "excluded": shape.excluded,
        }
        for slide in ir.slides
        for shape in slide.shapes
    ]
    return _native({
        "template_id": ir.template_id,
        "template_dir": template_dir,
        "shapes": shapes,
        "slots": ir.slot_registry,
        "deltas": deltas,
    })


def list_template_stores() -> list[dict[str, Any]]:
    """List developer template-first stores and mapping readiness."""
    from engine.template_ir import load_slot_mapping, load_template_ir
    from engine.template_ir.ingest import list_template_stores as _stores

    out = []
    for store in _stores():
        try:
            ir = load_template_ir(store)
        except Exception:
            continue
        out.append({
            "template_id": ir.template_id,
            "template_dir": store,
            "slots": len(ir.slot_registry),
            "has_mapping": load_slot_mapping(store) is not None,
        })
    return out


def build_template_report(client_data: list[dict[str, Any]], template: str,
                          output_path: str, client_name: str = "",
                          start_date: str = "", end_date: str = "") -> dict[str, Any]:
    """Developer-only template-first build retained from the donor."""
    from engine.template_ir.ingest import TEMPLATE_STORE_DIR
    from engine.template_ir.mapping import build_mapped_report

    template_dir = template if os.path.isdir(template) \
        else os.path.join(TEMPLATE_STORE_DIR, template)
    if not os.path.isdir(template_dir):
        raise FileNotFoundError(f"No ingested template store named '{template}'")

    if not start_date or not end_date:
        for item in client_data:
            start_date = start_date or item.get("start_date", "")
            end_date = end_date or item.get("end_date", "")
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    path, report = build_mapped_report(
        template_dir, output_path, client_data, client_name, start_date, end_date)
    return _native({"path": path, "report": report})
