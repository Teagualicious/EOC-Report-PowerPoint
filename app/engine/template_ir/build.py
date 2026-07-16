"""Phase 4 of the template-first pipeline: template IR → new PPTX.

The builder never re-authors shapes from parsed properties: every
supported shape is a VERBATIM copy of its original XML (pixel-exact by
construction), with image relationships re-targeted to the extracted
assets. Unsupported shapes (charts until Phase C) are skipped and
REPORTED — silent partial output is the failure mode this project has
been burned by before.

Phase B: dynamic text. After the verbatim copies land, slot values are
written into their placeholder runs via engine.pptx_fill's replace
machinery (same multi-line/<a:br/> handling and outcome-based reporting
the production fill path uses). Only run TEXT changes — formatting stays
verbatim by construction.
"""

import logging
import os

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu

from engine.template_ir.schema import load_template_ir

logger = logging.getLogger(__name__)

_BLIP_TAG = qn("a:blip")
_EMBED_ATTR = qn("r:embed")


def build_from_template(template_dir, output_path, slot_values=None):
    """Rebuild a deck from a template store directory.

    ``slot_values`` maps slot names (from the IR's slot registry) to
    display strings; each is written into its placeholder run after the
    verbatim copy. Returns (output_path, report) where report is a
    JSON-friendly dict: shapes_copied, images_relinked,
    skipped: [(shape_id, reason)], slots_filled,
    slots_unfilled: [(slot_name, reason)].
    """
    ir = load_template_ir(template_dir)
    report = {"template_id": ir.template_id, "shapes_copied": 0,
              "images_relinked": 0, "skipped": [],
              "slots_filled": 0, "slots_unfilled": []}

    prs = Presentation()
    prs.slide_width = Emu(ir.slide_width_emu)
    prs.slide_height = Emu(ir.slide_height_emu)
    blank = prs.slide_layouts[6]

    for slide_ir in ir.slides:
        slide = prs.slides.add_slide(blank)
        sp_tree = slide.shapes._spTree
        appended = {}
        for shape_ir in sorted(slide_ir.shapes, key=lambda s: s.z_order):
            if shape_ir.excluded:
                report["skipped"].append((shape_ir.shape_id,
                                          "excluded in template review"))
                continue
            if shape_ir.unsupported:
                report["skipped"].append((shape_ir.shape_id,
                                          shape_ir.unsupported))
                logger.warning("Build skipped %s: %s", shape_ir.shape_id,
                               shape_ir.unsupported)
                continue
            # python-pptx's typed parser (not raw lxml) so the appended
            # element can be wrapped by slide.shapes for the slot fill
            element = parse_xml(shape_ir.element_xml)
            report["images_relinked"] += _relink_images(
                element, shape_ir, slide, template_dir, report)
            sp_tree.append(element)
            appended[shape_ir.shape_id] = element
            report["shapes_copied"] += 1
        if slot_values:
            _fill_slide_slots(slide, slide_ir, ir.slot_registry,
                              slot_values, appended, report)

    if slot_values is not None:
        registry = ir.slot_registry
        for slot in slot_values:
            if slot not in registry:
                report["slots_unfilled"].append(
                    (slot, "not a slot in this template"))
        for slot in registry:
            if slot not in slot_values:
                report["slots_unfilled"].append((slot, "no value provided"))

    prs.save(output_path)
    logger.info("Template rebuilt: %s -> %s | copied=%d relinked=%d skipped=%d "
                "slots=%d/%d",
                ir.template_id, output_path, report["shapes_copied"],
                report["images_relinked"], len(report["skipped"]),
                report["slots_filled"],
                report["slots_filled"] + len(report["slots_unfilled"]))
    return output_path, report


def _fill_slide_slots(slide, slide_ir, registry, slot_values, appended,
                      report):
    """Write slot values into their placeholder runs on one built slide.

    Resolution is by the slot's stored placeholder text inside its own
    shape (identity + text, per the review doc) — the shape was located
    by shape_id, so identical placeholders in different shapes can never
    cross. Failed matches are reported, never silent.
    """
    from engine.pptx_fill import _replace_in_text_frame
    shapes_by_id = {s.shape_id: s for s in slide_ir.shapes}
    for slot, spec in registry.items():
        if spec.get("slide_index") != slide_ir.slide_index \
                or slot not in slot_values:
            continue
        shape_ir = shapes_by_id.get(spec.get("shape_id"))
        if shape_ir is None:
            report["slots_unfilled"].append((slot, "shape not in template"))
            continue
        element = appended.get(shape_ir.shape_id)
        if element is None:
            reason = ("shape excluded in template review" if shape_ir.excluded
                      else f"shape skipped: {shape_ir.unsupported}")
            report["slots_unfilled"].append((slot, reason))
            continue
        display = str(slot_values[slot])
        placeholder = spec.get("placeholder_text") or None
        wrote = False
        for text_frame in _iter_text_frames(_wrap_shape(slide, element)):
            if _replace_in_text_frame(text_frame, display, placeholder):
                wrote = True
                break
        if wrote:
            report["slots_filled"] += 1
        else:
            report["slots_unfilled"].append(
                (slot, "placeholder text not found in shape"))


def _wrap_shape(slide, element):
    """The python-pptx wrapper for an element just appended to the slide's
    shape tree (identity match — the wrappers are built over the same
    lxml elements)."""
    for shape in slide.shapes:
        if shape._element is element:
            return shape
    return None


def _iter_text_frames(shape):
    """Yield every text frame under a shape, descending into groups."""
    if shape is None:
        return
    try:
        is_group = shape.shape_type == MSO_SHAPE_TYPE.GROUP
    except NotImplementedError:   # python-pptx: unrecognized shape kinds
        is_group = False
    if is_group:
        for member in shape.shapes:
            yield from _iter_text_frames(member)
        return
    if getattr(shape, "has_text_frame", False):
        yield shape.text_frame


def _relink_images(element, shape_ir, slide, template_dir, report):
    """Point every image reference in the copied XML at a freshly added
    image part on the new slide. Returns how many were re-targeted."""
    if not shape_ir.image_rels:
        return 0
    relinked = 0
    new_rids = {}
    for old_rid, asset_rel in shape_ir.image_rels.items():
        asset_path = os.path.join(template_dir, asset_rel)
        try:
            _image_part, rid = slide.part.get_or_add_image_part(asset_path)
            new_rids[old_rid] = rid
        except Exception:
            logger.warning("Could not add image asset %s for %s", asset_rel,
                           shape_ir.shape_id, exc_info=True)
            report["skipped"].append(
                (shape_ir.shape_id, f"image asset missing: {asset_rel}"))
    for blip in element.iter(_BLIP_TAG):
        old = blip.get(_EMBED_ATTR)
        if old and old in new_rids:
            blip.set(_EMBED_ATTR, new_rids[old])
            relinked += 1
    return relinked
