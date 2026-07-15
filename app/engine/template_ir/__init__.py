"""Template-first mapper pipeline (Phase A: schema + ingest + verbatim
rebuild).

Design authority: documentation/proposals/TEMPLATE_FIRST_MAPPER_2026-07-15.md
plus the required amendments in documentation/reviews/
TEMPLATE_FIRST_MAPPER_REVIEW_2026-07-15.md — most importantly: static
shapes are rebuilt from their VERBATIM original XML (pixel-exact by
construction), never re-authored from parsed properties. The parsed
schema fields exist for preview, classification (Phase B), and diffing.

Phase A scope: ingest a deck into a JSON IR + assets, rebuild a deck from
verbatim copies, prove round-trip fidelity in CI. Classification, slots,
and data mapping are Phase B; chart-part cloning is Phase C.
"""

from engine.template_ir.schema import (SCHEMA_VERSION, ShapeIR, SlideIR,  # noqa: F401
                                       TemplateIR, load_template_ir)
from engine.template_ir.ingest import ingest_template  # noqa: F401
from engine.template_ir.build import build_from_template  # noqa: F401
