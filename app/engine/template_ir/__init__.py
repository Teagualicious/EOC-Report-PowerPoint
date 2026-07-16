"""Template-first mapper pipeline (Phases A+B: ingest → classify → map →
build).

Design authority: documentation/proposals/TEMPLATE_FIRST_MAPPER_2026-07-15.md
plus the required amendments in documentation/reviews/
TEMPLATE_FIRST_MAPPER_REVIEW_2026-07-15.md — most importantly: static
shapes are rebuilt from their VERBATIM original XML (pixel-exact by
construction), never re-authored from parsed properties. Dynamic text is
the only thing the builder changes: slot values replace placeholder run
TEXT inside those verbatim copies.

Phase A: ingest a deck into a JSON IR + assets, rebuild from verbatim
copies, round-trip fidelity proven in CI. Phase B: classification
suggestions + slot registry (classify), slot mappings on the current
mapper's query schema (mapping), dynamic text build. Chart-part cloning
is Phase C; re-ingest reconciliation is Phase D.
"""

from engine.template_ir.schema import (SCHEMA_VERSION, ShapeIR, SlideIR,  # noqa: F401
                                       TemplateIR, load_template_ir,
                                       save_template_ir)
from engine.template_ir.ingest import ingest_template  # noqa: F401
from engine.template_ir.build import build_from_template  # noqa: F401
from engine.template_ir.classify import classify_template  # noqa: F401
from engine.template_ir.mapping import (build_mapped_report,  # noqa: F401
                                        load_slot_mapping, new_slot_mapping,
                                        resolve_slot_values, save_slot_mapping,
                                        validate_slot_mapping)
