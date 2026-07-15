"""Phase 4 of the template-first pipeline: template IR → new PPTX.

The builder never re-authors shapes from parsed properties: every
supported shape is a VERBATIM copy of its original XML (pixel-exact by
construction), with image relationships re-targeted to the extracted
assets. Unsupported shapes (charts until Phase C) are skipped and
REPORTED — silent partial output is the failure mode this project has
been burned by before.
"""

import logging
import os

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

from engine.template_ir.schema import load_template_ir

logger = logging.getLogger(__name__)

_BLIP_TAG = qn("a:blip")
_EMBED_ATTR = qn("r:embed")


def build_from_template(template_dir, output_path):
    """Rebuild a deck from a template store directory.

    Returns (output_path, report) where report is a JSON-friendly dict:
    shapes_copied, images_relinked, skipped: [(shape_id, reason), ...].
    """
    ir = load_template_ir(template_dir)
    report = {"template_id": ir.template_id, "shapes_copied": 0,
              "images_relinked": 0, "skipped": []}

    prs = Presentation()
    prs.slide_width = Emu(ir.slide_width_emu)
    prs.slide_height = Emu(ir.slide_height_emu)
    blank = prs.slide_layouts[6]

    for slide_ir in ir.slides:
        slide = prs.slides.add_slide(blank)
        sp_tree = slide.shapes._spTree
        for shape_ir in sorted(slide_ir.shapes, key=lambda s: s.z_order):
            if shape_ir.unsupported:
                report["skipped"].append((shape_ir.shape_id,
                                          shape_ir.unsupported))
                logger.warning("Build skipped %s: %s", shape_ir.shape_id,
                               shape_ir.unsupported)
                continue
            element = etree.fromstring(shape_ir.element_xml.encode("utf-8"))
            report["images_relinked"] += _relink_images(
                element, shape_ir, slide, template_dir, report)
            sp_tree.append(element)
            report["shapes_copied"] += 1

    prs.save(output_path)
    logger.info("Template rebuilt: %s -> %s | copied=%d relinked=%d skipped=%d",
                ir.template_id, output_path, report["shapes_copied"],
                report["images_relinked"], len(report["skipped"]))
    return output_path, report


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
