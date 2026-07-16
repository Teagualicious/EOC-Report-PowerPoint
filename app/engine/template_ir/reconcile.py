"""Phase D of the template-first pipeline: template evolution.

When a client updates their deck, ``reingest_template`` re-ingests it
into the SAME store and carries the human review work forward — the
review doc's Critique 3: without this, every template revision costs a
full re-classification session. Only the deltas are presented for
review.

Matching follows the Phase 4 identity rules: a shape carries forward
when its persistent PowerPoint id (`shape_uid`) matches uniquely, else
when its NAME matches uniquely on both sides. Duplicates never match —
no guessing. Slots re-anchor by their stored placeholder text inside the
matched shape; a placeholder that no longer exists drops its slot loudly
(a delta, plus a mapping warning if it was mapped) rather than guessing
a new target.
"""

import logging
import os
import shutil

from engine.template_ir.classify import classify_template, find_anchor
from engine.template_ir.ingest import ingest_template
from engine.template_ir.schema import (TEMPLATE_JSON_NAME, load_template_ir,
                                       save_template_ir)

logger = logging.getLogger(__name__)

PREV_JSON_NAME = "template.prev.json"


def reingest_template(pptx_path, template_dir):
    """Re-ingest an updated deck into an existing template store.

    Returns (template_dir, deltas). The previous template.json is kept as
    template.prev.json so a bad re-ingest never destroys review work.
    """
    old_ir = load_template_ir(template_dir)
    json_path = os.path.join(template_dir, TEMPLATE_JSON_NAME)
    try:
        shutil.copy2(json_path, os.path.join(template_dir, PREV_JSON_NAME))
    except OSError:
        logger.warning("Could not write %s", PREV_JSON_NAME, exc_info=True)

    ingest_template(pptx_path, template_id=old_ir.template_id,
                    store_dir=template_dir)
    new_ir = classify_template(load_template_ir(template_dir))
    deltas = reconcile(old_ir, new_ir)
    _note_mapping_impact(template_dir, deltas)
    save_template_ir(new_ir, template_dir)
    logger.info("Template re-ingested: %s -> %s (%d delta(s))",
                pptx_path, template_dir, len(deltas))
    return template_dir, deltas


def reconcile(old_ir, new_ir):
    """Carry review state from ``old_ir`` onto ``new_ir`` (a freshly
    ingested + auto-classified IR of the updated deck), in place.

    Returns the deltas — [{"kind", "detail", ...}] — everything a human
    should look at: shapes added/removed, text changes on carried shapes,
    slots dropped (shape or placeholder gone) and newly suggested.
    """
    if not _was_reviewed(old_ir):
        # A Phase A store has no review work to preserve — fresh
        # suggestions stand, nothing to diff against.
        return [{"kind": "no_prior_review",
                 "detail": "previous store had no classifications or slots "
                           "— fresh auto-classification applied"}]

    deltas = []
    old_shapes = [s for slide in old_ir.slides for s in slide.shapes]
    new_shapes = [s for slide in new_ir.slides for s in slide.shapes]
    old_by_uid = _unique_index(old_shapes, lambda s: s.shape_uid)
    new_by_uid = _unique_index(new_shapes, lambda s: s.shape_uid)
    old_by_name = _unique_index(old_shapes, lambda s: s.name)
    new_by_name = _unique_index(new_shapes, lambda s: s.name)

    def match_old(new_shape):
        old = old_by_uid.get(new_shape.shape_uid)
        if old is None and new_shape.name in old_by_name \
                and new_shape.name in new_by_name:
            old = old_by_name[new_shape.name]
        return old

    matched_old_ids = set()
    matches = {}   # new shape_id -> old shape
    for shape in new_shapes:
        old = match_old(shape)
        if old is None:
            deltas.append({"kind": "shape_added", "shape": shape.shape_id,
                           "detail": f"new shape \"{shape.name}\" — "
                                     f"suggested {shape.classification} "
                                     f"({shape.classify_reason})"})
            continue
        matches[shape.shape_id] = old
        matched_old_ids.add(old.shape_id)
        shape.classification = old.classification
        shape.classify_reason = old.classify_reason
        shape.excluded = old.excluded
        if (shape.text or "") != (old.text or ""):
            deltas.append({"kind": "shape_text_changed",
                           "shape": shape.shape_id,
                           "detail": f"\"{shape.name}\" text changed — "
                                     "review its classification/slots"})

    for old in old_shapes:
        if old.shape_id not in matched_old_ids:
            deltas.append({"kind": "shape_removed", "shape": old.shape_id,
                           "detail": f"shape \"{old.name}\" is gone from "
                                     "the updated deck"})

    new_ir.slot_registry = _reconcile_slots(old_ir, new_ir, matches, deltas)
    for slide in new_ir.slides:
        for shape in slide.shapes:
            shape.slot_name = next(
                (name for name, spec in new_ir.slot_registry.items()
                 if spec.get("shape_id") == shape.shape_id), None)
    return deltas


def _reconcile_slots(old_ir, new_ir, matches, deltas):
    """The final registry: carried old slots (names preserved,
    re-anchored) plus fresh suggestions that don't overlap them."""
    old_shape_to_new = {old.shape_id: new_id
                        for new_id, old in matches.items()}
    new_by_id = {s.shape_id: (slide, s) for slide in new_ir.slides
                 for s in slide.shapes}
    final = {}

    for name, spec in old_ir.slot_registry.items():
        new_id = old_shape_to_new.get(spec.get("shape_id"))
        if new_id is None:
            deltas.append({"kind": "slot_dropped", "slot": name,
                           "detail": f"slot '{name}' dropped — its shape "
                                     "is gone from the updated deck"})
            continue
        slide, shape = new_by_id[new_id]
        carried = dict(spec, slide_index=slide.slide_index,
                       shape_id=shape.shape_id, shape_uid=shape.shape_uid)
        if spec.get("type") in ("chart_data", "table_data"):
            if shape.unsupported:
                deltas.append({"kind": "slot_dropped", "slot": name,
                               "detail": f"slot '{name}' dropped — the "
                                         f"shape is now unsupported "
                                         f"({shape.unsupported})"})
                continue
        elif spec.get("placeholder_text"):
            anchor = find_anchor(shape, spec["placeholder_text"])
            if anchor is None:
                deltas.append({"kind": "slot_dropped", "slot": name,
                               "detail": f"slot '{name}' dropped — its "
                                         "placeholder text no longer "
                                         "appears in the shape"})
                continue
            carried["paragraph"], carried["run"] = anchor
        final[name] = carried

    # Fresh suggestions: only where they don't duplicate a carried slot,
    # and never on shapes the reviewer marked static/excluded.
    covered = {(spec.get("shape_id"), spec.get("placeholder_text"))
               for spec in final.values()}
    for name, spec in new_ir.slot_registry.items():
        shape_id = spec.get("shape_id")
        _, shape = new_by_id[shape_id]
        if shape.classification != "dynamic" or shape.excluded:
            continue
        if (shape_id, spec.get("placeholder_text")) in covered:
            continue
        unique = name
        n = 2
        while unique in final:
            unique = f"{name}_{n}"
            n += 1
        final[unique] = spec
        deltas.append({"kind": "slot_new", "slot": unique,
                       "detail": f"new slot suggested: '{unique}' "
                                 f"({spec.get('type')}) — "
                                 f"{spec.get('description', '')}"})
    return final


def _was_reviewed(ir):
    return bool(ir.slot_registry) or any(
        s.classify_reason for slide in ir.slides for s in slide.shapes)


def _unique_index(shapes, key):
    """{key: shape} for keys that appear exactly once (None excluded) —
    ambiguous identities never match."""
    counts = {}
    for shape in shapes:
        k = key(shape)
        if k is not None:
            counts[k] = counts.get(k, 0) + 1
    return {key(s): s for s in shapes
            if key(s) is not None and counts[key(s)] == 1}


def _note_mapping_impact(template_dir, deltas):
    """A dropped slot that was mapped to a data source needs remapping —
    say so in its delta."""
    from engine.template_ir.mapping import load_slot_mapping
    mapping = load_slot_mapping(template_dir)
    if not mapping:
        return
    mapped = set(mapping.get("slots", {}))
    for delta in deltas:
        if delta["kind"] == "slot_dropped" and delta.get("slot") in mapped:
            delta["detail"] += " (it WAS mapped to a data source — remap "
            delta["detail"] += "or the next build reports it unfilled)"
