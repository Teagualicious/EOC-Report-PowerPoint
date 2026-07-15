"""Single owner of PowerPoint mapping state for the template mapper.

Every mutation of a template mapping goes through ``MappingModel``; the Tk
panels and the live COM preview are observers that re-render from it after
each change. The model is pure Python (no Tk, no COM) so it is fully
exercised by the automated suite.

The internal ``data`` dict is kept in the exact on-disk schema that
``engine.pptx_mapper.save_template_mapping`` persists and
``engine.pptx_fill`` consumes, including the legacy shape-level
single-assignment form — serialization is the identity function and old
mappings load unchanged.
"""

import logging

logger = logging.getLogger(__name__)

# Outcomes of assign_metric()
CREATED = "created"            # first assignment on the shape
APPENDED = "appended"          # added alongside existing assignments
UPDATED = "updated"            # same metric already present — updated in place
REPLACED = "replaced"          # confirmed full-text replace of all assignments
NEEDS_CONFIRM = "needs_confirm"  # would replace existing assignments — ask first


class MappingModel:
    """Owns a template mapping dict plus per-metric format preferences."""

    def __init__(self, mapping=None):
        self.data = mapping if mapping else {"slides": {}}
        self.metric_formats = {}         # metric key -> format string
        self.metric_format_details = {}  # metric key -> format-details dict
        self._observers = []
        self._scan_identity = {}         # (slide_num, shape_id) -> (uid, name)

    # ── Observers ─────────────────────────────────────────────────────────

    def subscribe(self, callback):
        """Register callback(event, slide_num=None, shape_id=None, metric=None).

        Called synchronously after every mutation. Bookkeeping writes
        (``note_rendered``) do not notify.
        """
        self._observers.append(callback)

    def _notify(self, event, slide_num=None, shape_id=None, metric=None):
        for cb in self._observers:
            try:
                cb(event, slide_num=slide_num, shape_id=shape_id, metric=metric)
            except Exception:
                logger.warning("Mapping observer failed on %s", event,
                               exc_info=True)

    # ── Reads ─────────────────────────────────────────────────────────────

    def slide_map(self, slide_num):
        return self.data.get("slides", {}).get(str(slide_num), {})

    def shape_map(self, slide_num, shape_id):
        return self.slide_map(slide_num).get("shape_mappings", {}).get(
            str(shape_id), {})

    def assignments(self, slide_num, shape_id):
        """Normalized read view: the assignments list, or the legacy
        shape-level single assignment wrapped in a list. The returned dicts
        are the live stored objects (same as the pre-model UI relied on)."""
        smap = self.shape_map(slide_num, shape_id)
        asgns = smap.get("assignments", [])
        if not asgns and smap.get("metric"):
            return [smap]
        return asgns

    def to_dict(self):
        """The canonical mapping dict in the persisted schema."""
        return self.data

    # ── Mutations ─────────────────────────────────────────────────────────

    def _shape_mappings(self, slide_num):
        slides = self.data.setdefault("slides", {})
        return slides.setdefault(str(slide_num), {"shape_mappings": {}}) \
                     .setdefault("shape_mappings", {})

    def set_scan_identity(self, slides):
        """Record each scanned shape's persistent identity (PowerPoint shape
        id + name, from scan_template's structure) so mutations stamp it
        onto the entries they touch. Shapes without a ``shape_uid`` stay
        unstamped — those entries keep resolving positionally, and callers
        that never attach a scan (headless tests, old flows) see the exact
        pre-uid schema."""
        for slide in slides or []:
            for shape in slide.get("shapes", []):
                uid = shape.get("shape_uid")
                if uid is None:
                    continue
                key = (str(slide["slide_num"]), str(shape["shape_id"]))
                self._scan_identity[key] = (uid, shape.get("name", ""))

    def _stamp(self, slide_num, shape_id):
        """Copy the scanned identity onto the stored entry (no-op when the
        scan carried none). Legacy entries in a loaded mapping upgrade
        lazily, the first time the user touches them."""
        ident = self._scan_identity.get((str(slide_num), str(shape_id)))
        if not ident:
            return
        smap = self._shape_mappings(slide_num).get(str(shape_id))
        if smap is not None:
            smap["shape_uid"], smap["shape_name"] = ident

    def assign_metric(self, slide_num, shape_id, metric, fmt="text",
                      replace_text="", format_details=None, query=None,
                      confirm_replace=False):
        """Assign a metric to a shape. Returns one of the outcome constants.

        Semantics preserved from the pre-model mapper:
        - the same metric already on the shape is updated in place (new
          format/details/query; the replace target is kept unless a new
          non-empty one is given);
        - a full-text assignment (no replace_text) onto a shape that already
          has assignments replaces ALL of them — returns NEEDS_CONFIRM and
          changes nothing until called again with confirm_replace=True;
        - otherwise the assignment is appended (partial) or becomes the
          shape's first assignment (stored legacy-style at shape level).
        """
        sid = str(shape_id)
        shape_mappings = self._shape_mappings(slide_num)

        new_assignment = {"metric": metric, "format": fmt,
                          "replace_text": replace_text}
        if format_details:
            new_assignment["format_details"] = format_details
        if query:
            new_assignment["query"] = query

        existing = shape_mappings.get(sid, {})
        existing_asgns = existing.get("assignments", [])
        if not existing_asgns and existing.get("metric"):
            # Normalize the legacy single-assignment form into the list form
            # so every path below appends/updates instead of overwriting.
            first = {"metric": existing["metric"],
                     "format": existing.get("format", "text"),
                     "replace_text": existing.get("replace_text", ""),
                     "format_details": existing.get("format_details"),
                     "last_rendered": existing.get("last_rendered")}
            if existing.get("query"):
                first["query"] = existing["query"]
            existing_asgns = [first]

        # Same metric already assigned? Update in place.
        for asgn in existing_asgns:
            if asgn.get("metric") == metric:
                asgn["format"] = fmt
                if format_details:
                    asgn["format_details"] = format_details
                if query:
                    asgn["query"] = query
                if replace_text:
                    asgn["replace_text"] = replace_text
                shape_mappings[sid] = {"skip": False,
                                       "assignments": existing_asgns}
                self._stamp(slide_num, sid)
                self._notify("assign", slide_num, shape_id, metric)
                return UPDATED

        if existing_asgns and not replace_text:
            # Full replace onto a shape that already has assignments wipes
            # them all — the caller must confirm.
            if not confirm_replace:
                return NEEDS_CONFIRM
            shape_map = dict(new_assignment)
            shape_map["skip"] = False
            shape_map["assignments"] = []
            shape_mappings[sid] = shape_map
            self._stamp(slide_num, sid)
            self._notify("assign", slide_num, shape_id, metric)
            return REPLACED

        if existing_asgns:
            existing_asgns.append(new_assignment)
            shape_mappings[sid] = {"skip": False,
                                   "assignments": existing_asgns}
            self._stamp(slide_num, sid)
            self._notify("assign", slide_num, shape_id, metric)
            return APPENDED

        # First assignment on this shape (legacy-compatible shape-level form)
        shape_map = dict(new_assignment)
        shape_map["skip"] = False
        shape_map["assignments"] = []
        shape_mappings[sid] = shape_map
        self._stamp(slide_num, sid)
        self._notify("assign", slide_num, shape_id, metric)
        return CREATED

    def assign_image(self, slide_num, shape_id, image_path, image_path_abs=""):
        """Map a shape to an image (replaces any text assignments on it)."""
        self._shape_mappings(slide_num)[str(shape_id)] = {
            "metric": "",
            "image_path": image_path,
            "image_path_abs": image_path_abs,
            "skip": False,
            "assignments": [],
        }
        self._stamp(slide_num, shape_id)
        self._notify("image", slide_num, shape_id)

    def set_skip(self, slide_num, shape_id, skip):
        """Set the skip flag, PRESERVING the shape's assignments.

        (Changed 2026-07-15, owner decision — the pre-model behavior
        replaced the whole mapping, so an exploratory Skip click silently
        discarded the shape's assignments. Toggling skip off now restores
        exactly what was there.)"""
        smap = self._shape_mappings(slide_num).setdefault(str(shape_id), {})
        smap["skip"] = skip
        self._stamp(slide_num, shape_id)
        self._notify("skip", slide_num, shape_id)

    def clear_shape(self, slide_num, shape_id):
        shape_mappings = self.slide_map(slide_num).get("shape_mappings", {})
        if str(shape_id) in shape_mappings:
            del shape_mappings[str(shape_id)]
        self._notify("clear", slide_num, shape_id)

    def set_metric_format(self, metric, fmt, details=None):
        """Store per-metric format preferences and push the details into
        every existing assignment of that metric (re-formatting never
        requires re-assigning). Returns the number of assignments touched."""
        self.metric_formats[metric] = fmt
        if details is not None:
            self.metric_format_details[metric] = details
        touched = 0
        if details:
            for slide in self.data.get("slides", {}).values():
                for smap in slide.get("shape_mappings", {}).values():
                    targets = smap.get("assignments") or (
                        [smap] if smap.get("metric") else [])
                    for asgn in targets:
                        if asgn.get("metric") == metric:
                            asgn["format_details"] = details
                            asgn["format"] = details.get(
                                "format", asgn.get("format", "text"))
                            touched += 1
            if touched:
                logger.info("Format change propagated to %d assignment(s) "
                            "of %s", touched, metric)
        self._notify("format", metric=metric)
        return touched

    def note_rendered(self, slide_num, shape_id, metric, display):
        """Record the text last rendered into the preview for a metric's
        assignment(s) on a shape, so re-applies can replace it in place.
        Bookkeeping only — observers are not notified."""
        smap = self.shape_map(slide_num, shape_id)
        if smap.get("metric") == metric:
            smap["last_rendered"] = display
        for asgn in smap.get("assignments", []):
            if asgn.get("metric") == metric:
                asgn["last_rendered"] = display
