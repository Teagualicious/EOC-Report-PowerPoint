"""The template IR schema — the contract between ingest, classify, map,
and build. Plain dataclasses + JSON (stdlib; no new dependencies).

Identity rules (mapper roadmap Phase 4 lessons apply here from day 1):
``shape_uid`` is PowerPoint's persistent shape id — the key future slot
names and re-ingest reconciliation ride on. ``z_order`` is positional
metadata only, never identity.
"""

import json
import os
from dataclasses import asdict, dataclass, field

SCHEMA_VERSION = "1.1"   # 1.1: classification reasons, excluded flag, slot
                         # registry, chart part assets (all Phase B/C; 1.0
                         # stores load unchanged)

TEMPLATE_JSON_NAME = "template.json"
ASSETS_DIR_NAME = "assets"


@dataclass
class ShapeIR:
    """One shape. ``element_xml`` is the verbatim original XML — the render
    source at build time. Parsed fields are preview/classify/diff metadata,
    deliberately NOT what the builder draws from (schema gaps must never
    become brand drift)."""
    shape_id: str                    # "s{slide_index}_u{shape_uid}"
    shape_uid: int | None
    name: str
    shape_type: str                  # text_box|image|table|chart|shape|group|other
    z_order: int
    geometry: dict                   # left/top/width/height (EMU), rotation (deg)
    element_xml: str
    text: str = ""                   # plain-text snapshot
    image_rels: dict = field(default_factory=dict)  # embed rId -> asset rel path
    # Chart shapes: extracted part assets for build-time cloning —
    # {"xml": rel path, "workbook": rel path|None,
    #  "extra": [[reltype, content_type, rel path], ...]} (colors/style)
    chart_part: dict = field(default_factory=dict)
    classification: str = "static"   # static|dynamic (suggested, user-reviewed)
    classify_reason: str = ""        # why the heuristic suggested it
    excluded: bool = False           # user removed it entirely — builder skips
    slot_name: str | None = None     # first slot on this shape (display helper)
    unsupported: str | None = None   # e.g. "chart" — builder skips + reports


@dataclass
class SlideIR:
    slide_index: int
    layout_name: str
    shapes: list


@dataclass
class TemplateIR:
    schema_version: str
    template_id: str
    template_name: str
    source_file: str
    created_at: str
    slide_width_emu: int
    slide_height_emu: int
    slides: list
    # slot_name -> {type, description, slide_index, shape_id, shape_uid,
    #               placeholder_text, paragraph, run}. Slots pin to
    #               (shape identity + placeholder text); paragraph/run
    #               indices are reconciliation hints only, never identity.
    slot_registry: dict = field(default_factory=dict)

    def to_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, data):
        slides = [SlideIR(slide_index=s["slide_index"],
                          layout_name=s.get("layout_name", ""),
                          shapes=[ShapeIR(**sh) for sh in s.get("shapes", [])])
                  for s in data.get("slides", [])]
        return cls(schema_version=data["schema_version"],
                   template_id=data["template_id"],
                   template_name=data.get("template_name", ""),
                   source_file=data.get("source_file", ""),
                   created_at=data.get("created_at", ""),
                   slide_width_emu=data["slide_width_emu"],
                   slide_height_emu=data["slide_height_emu"],
                   slides=slides,
                   slot_registry=data.get("slot_registry", {}))


def save_template_ir(ir, template_dir):
    """Write template.json into the template's store directory."""
    os.makedirs(template_dir, exist_ok=True)
    path = os.path.join(template_dir, TEMPLATE_JSON_NAME)
    from config.settings import _atomic_json_write
    _atomic_json_write(path, ir.to_dict())
    return path


def load_template_ir(template_dir_or_json):
    """Load a TemplateIR from a store directory or a template.json path."""
    path = template_dir_or_json
    if os.path.isdir(path):
        path = os.path.join(path, TEMPLATE_JSON_NAME)
    with open(path, "r", encoding="utf-8") as f:
        return TemplateIR.from_dict(json.load(f))
