"""PowerPoint Live Preview via COM automation (Windows only).

Opens PowerPoint, displays the template, and updates shapes in real time
as the user assigns metrics in the mapper UI.

Falls back gracefully if PowerPoint or pywin32 is not available: every
public method no-ops when inactive, and failures are logged (DEBUG for
per-call COM flakiness, WARNING for setup problems) instead of printed
or silently swallowed.
"""

import functools
import logging
import os
import shutil
import tempfile

logger = logging.getLogger(__name__)

# COM automation — Windows only
_com_available = False
try:
    import pythoncom
    import win32com.client
    _com_available = True
except ImportError:
    logger.info("pywin32 not available — live preview disabled")


def _tracked(op_name):
    """Health-tracking wrapper for public COM methods.

    Counts every call; a call is a failure iff the method's own exception
    handler called _record_failure during it (early no-op returns count as
    success — the COM reads that got us there worked). After
    MAX_CONSECUTIVE_FAILURES in a row the preview disables itself LOUDLY:
    active goes False, on_disabled fires once, and Save & Fill falls back
    to the tested python-pptx engine automatically.
    """
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(self, *args, **kwargs):
            if not self.active:
                return fn(self, *args, **kwargs)  # existing no-op paths
            self.health["calls"] += 1
            self._call_failed = False
            result = fn(self, *args, **kwargs)
            if self._call_failed:
                self.health["failures"] += 1
                self.health["consecutive_failures"] += 1
                self.health["last_error"] = op_name
                if (self.health["consecutive_failures"]
                        >= self.MAX_CONSECUTIVE_FAILURES):
                    self._disable(
                        f"{self.health['consecutive_failures']} consecutive "
                        f"PowerPoint errors (last: {op_name})")
            else:
                self.health["consecutive_failures"] = 0
            return result
        return wrapper
    return deco


class PPTXLivePreview:
    """Manages a live PowerPoint instance for real-time preview."""

    MAX_CONSECUTIVE_FAILURES = 3

    def __init__(self, template_path):
        self.template_path = os.path.abspath(template_path)
        self.pptx_app = None
        self.presentation = None
        self.active = False
        self._working_copy = None
        # Health tracking — read via get_health(); set on_disabled to be
        # notified (once) when the preview turns itself off.
        self.health = {"calls": 0, "failures": 0, "consecutive_failures": 0,
                       "last_error": None, "disabled_reason": None}
        self.on_disabled = None
        self._call_failed = False
        # (slide_num, shape_index) -> list of original paragraph texts,
        # captured before a shape's first modification (used by Clear All)
        self._original_paragraphs = {}

        if not _com_available:
            return

        try:
            # Create a working copy so we don't modify the original
            work_dir = os.path.dirname(self.template_path)
            base = os.path.splitext(os.path.basename(self.template_path))[0]
            self._working_copy = os.path.join(work_dir, f"_preview_{base}.pptx")
            shutil.copy2(self.template_path, self._working_copy)

            # Initialize COM
            pythoncom.CoInitialize()

            # Open PowerPoint (must be Visible for COM; minimized instead)
            self.pptx_app = win32com.client.Dispatch("PowerPoint.Application")
            self.pptx_app.Visible = True

            # Open the working copy
            self.presentation = self.pptx_app.Presentations.Open(
                self._working_copy,
                ReadOnly=False,
                Untitled=False,
                WithWindow=True
            )

            # Minimize PowerPoint window for embedded preview mode
            try:
                self.pptx_app.WindowState = 2  # ppWindowMinimized
            except Exception:
                logger.debug("Could not minimize PowerPoint window", exc_info=True)

            self.active = True
            logger.info("Live preview started for %s",
                        os.path.basename(self.template_path))

        except Exception:
            logger.warning("Live preview COM init failed for %s",
                           self.template_path, exc_info=True)
            self.active = False
            self.cleanup()

    def _record_failure(self, message, *args):
        """Called from a method's exception handler: log loudly and flag
        this call as failed for the _tracked wrapper."""
        logger.warning(message, *args, exc_info=True)
        self._call_failed = True

    def _disable(self, reason):
        """Turn the preview off and say so — no more silent half-working."""
        self.active = False
        self.health["disabled_reason"] = reason
        logger.error("LIVE PREVIEW DISABLED: %s — mapping is unaffected; "
                     "Save & Fill will use the built-in engine.", reason)
        cb, self.on_disabled = self.on_disabled, None  # fire once
        if cb:
            try:
                cb(reason)
            except Exception:
                logger.debug("on_disabled callback failed", exc_info=True)

    def get_health(self):
        return dict(self.health, active=self.active)

    def is_active(self):
        """True when PowerPoint is running and the working copy is open."""
        return self.active and self.pptx_app is not None

    @_tracked("export_slide_image")
    def export_slide_image(self, slide_num, output_path=None, width=800):
        """Export a slide as a PNG image for embedded preview.

        Returns the path to the exported image, or None on failure.
        """
        if not self.is_active():
            return None
        try:
            if output_path is None:
                output_path = os.path.join(tempfile.gettempdir(),
                                           f"_slide_preview_{slide_num}.png")

            slide = self.presentation.Slides(slide_num)
            slide.Export(output_path, "PNG", width)
            return output_path
        except Exception:
            self._record_failure("Slide %s export failed", slide_num)
            return None

    @_tracked("go_to_slide")
    def go_to_slide(self, slide_num):
        """Navigate to a specific slide number (1-based)."""
        if not self.is_active():
            return
        try:
            if slide_num <= self.presentation.Slides.Count:
                self.pptx_app.ActiveWindow.View.GotoSlide(slide_num)
        except Exception:
            self._record_failure("GotoSlide %s failed", slide_num)

    def _snapshot_shape(self, slide_num, shape_index, shape, tf):
        """Cache the shape's per-paragraph text AND geometry before its
        first modification.

        Enables restore_shape_text() to put the shape back exactly — per
        paragraph (so each keeps its own run formatting) and at its
        original left/width (in case auto-widen grew it).
        """
        key = (slide_num, shape_index)
        if key in self._original_paragraphs:
            return
        try:
            paras = []
            count = tf.TextRange.Paragraphs().Count
            for pi in range(1, count + 1):
                paras.append(tf.TextRange.Paragraphs(pi).Text)
            snap = {
                "paras": paras,
                "left": shape.Left,
                "width": shape.Width,
            }
            try:
                snap["autosize"] = shape.TextFrame2.AutoSize
            except Exception:
                snap["autosize"] = None
            self._original_paragraphs[key] = snap
        except Exception:
            logger.debug("Paragraph snapshot failed (slide %s, shape %s)",
                         slide_num, shape_index, exc_info=True)

    @staticmethod
    def _set_paragraph_text(para, new_text):
        """Replace a paragraph's content WITHOUT touching its paragraph mark.

        In PowerPoint COM, Paragraph.Text includes the trailing '\\r'.
        Assigning the whole Text is hazardous both ways: omitting the mark
        merges the paragraph into the next one; including it can insert an
        extra break (splitting). Writing through a Characters range that
        excludes the mark sidesteps both failure modes entirely.
        """
        old = para.Text
        has_mark = old.endswith("\r")
        content_len = len(old) - (1 if has_mark else 0)
        new_text = new_text.rstrip("\r")
        if content_len > 0:
            para.Characters(1, content_len).Text = new_text
        elif has_mark:
            # Empty paragraph (mark only) — insert content before the mark
            para.InsertBefore(new_text)
        else:
            para.Text = new_text

    @_tracked("update_shape_text")
    def update_shape_text(self, slide_num, shape_index, new_text,
                          replace_portion=None, expected_name=None):
        """Update shape text using direct .Text assignment (VBA approach).

        Preserves paragraph structure (trailing '\\r' marks) so multi-line
        text boxes keep their per-line formatting. Overflow is handled by
        auto-widen (with shrink-to-fit fallback).

        expected_name, when provided, is checked against the shape actually
        found at the index — a mismatch means indices drifted (e.g. shapes
        were added/deleted in PowerPoint) and is logged loudly.
        """
        if not self.is_active():
            return
        try:
            slide = self.presentation.Slides(slide_num)
            shape = slide.Shapes(shape_index + 1)

            if expected_name:
                try:
                    if shape.Name != expected_name:
                        logger.warning(
                            "SHAPE INDEX DRIFT: slide %s index %s is now '%s' "
                            "(expected '%s') — assignment may hit the wrong "
                            "object", slide_num, shape_index, shape.Name,
                            expected_name)
                except Exception:
                    logger.debug("Name check failed", exc_info=True)

            if not shape.HasTextFrame:
                return

            tf = shape.TextFrame

            # Snapshot original state once, before the first modification —
            # this is what Clear All restores from.
            self._snapshot_shape(slide_num, shape_index, shape, tf)

            grew = False
            if replace_portion:
                # Accept a single portion or a list of candidates (e.g. the
                # assignment's last rendered output first, then the original
                # template placeholder) — first one found in the live text
                # wins. This makes re-formatting idempotent: the new render
                # replaces the previous render, not some unrelated text.
                candidates = (replace_portion if isinstance(replace_portion, list)
                              else [replace_portion])
                current = tf.TextRange.Text
                hit = next((c for c in candidates if c and c in current), None)
                if hit is None:
                    logger.debug("No replace candidate found in shape text "
                                 "(slide %s, shape %s): %r",
                                 slide_num, shape_index, candidates)
                    return
                idx = current.index(hit)
                tf.TextRange.Characters(idx + 1, len(hit)).Text = new_text
                grew = len(new_text) > len(hit)
            else:
                para_count = tf.TextRange.Paragraphs().Count
                if para_count == 1:
                    para = tf.TextRange.Paragraphs(1)
                    grew = len(new_text) > len(para.Text.rstrip("\r"))
                    self._set_paragraph_text(para, new_text)
                else:
                    for pi in range(1, para_count + 1):
                        para = tf.TextRange.Paragraphs(pi)
                        if para.Text.strip():
                            grew = len(new_text) > len(para.Text.rstrip("\r"))
                            self._set_paragraph_text(para, new_text)
                            break

            # If the new text is longer than what it replaced, widen the box
            # so the value stays on one line (grown symmetrically to keep
            # centered stat blocks centered). Falls back to shrink-to-fit —
            # never to disabling WordWrap, which is what used to mangle
            # formatting.
            if grew:
                self._widen_to_fit(shape, tf)

        except Exception:
            self._record_failure("Live text update failed (slide %s, shape %s)",
                                 slide_num, shape_index)

    def _widen_to_fit(self, shape, tf):
        """Widen a shape so its longest line fits without wrapping.

        Measures the true one-line width by toggling WordWrap off just for
        the measurement, then grows the shape SYMMETRICALLY (left shifts by
        half the delta) so centered stat blocks stay visually centered.
        Growth is clamped to the slide bounds. If measurement fails, falls
        back to shrink-text-to-fit — never to permanently disabling
        WordWrap (the old failure mode that mangled formatting).
        """
        try:
            old_wrap = tf.WordWrap
            tf.WordWrap = False
            needed = tf.TextRange.BoundWidth + 12  # breathing room
            tf.WordWrap = old_wrap

            if needed <= shape.Width:
                return  # already fits

            slide_w = self.presentation.PageSetup.SlideWidth
            margin = 6
            max_width = slide_w - 2 * margin
            new_width = min(needed, max_width)
            delta = new_width - shape.Width
            if delta <= 0:
                return

            # Grow symmetrically, then clamp within the slide
            new_left = shape.Left - delta / 2
            if new_left < margin:
                new_left = margin
            if new_left + new_width > slide_w - margin:
                new_left = max(margin, slide_w - margin - new_width)

            shape.Left = new_left
            shape.Width = new_width
            logger.debug("Auto-widened shape to %.0f (needed %.0f)",
                         new_width, needed)
        except Exception:
            logger.debug("Auto-widen failed — shrink-to-fit fallback",
                         exc_info=True)
            try:
                shape.TextFrame2.AutoSize = 2  # msoAutoSizeTextToFitShape
            except Exception:
                logger.debug("Shrink-to-fit fallback failed", exc_info=True)

    @_tracked("restore_shape_text")
    def restore_shape_text(self, slide_num, shape_index):
        """Restore a shape's text to its pre-modification snapshot.

        Per-paragraph restore preserves each paragraph's own formatting.
        The snapshot strings include trailing '\\r' paragraph marks (COM's
        Paragraph.Text includes them) — writing those back verbatim would
        INSERT an extra paragraph break, shift every index mid-loop, and
        cascade paragraph 1's formatting over the whole box. So the mark
        is stripped and _set_paragraph_text re-appends it only when the
        live paragraph still ends with one.
        Returns True if a snapshot existed and was applied.
        """
        if not self.is_active():
            return False
        key = (slide_num, shape_index)
        snap = self._original_paragraphs.get(key)
        if snap is None:
            return False
        paras = snap["paras"]
        try:
            slide = self.presentation.Slides(slide_num)
            shape = slide.Shapes(shape_index + 1)
            if not shape.HasTextFrame:
                return False
            tf = shape.TextFrame
            count = tf.TextRange.Paragraphs().Count
            for pi in range(1, min(count, len(paras)) + 1):
                para = tf.TextRange.Paragraphs(pi)
                original = paras[pi - 1].rstrip("\r")
                self._set_paragraph_text(para, original)
            # Restore original geometry (in case auto-widen grew the shape)
            try:
                shape.Left = snap["left"]
                shape.Width = snap["width"]
            except Exception:
                logger.debug("Geometry restore failed", exc_info=True)
            # Restore the shape's original autosize (templates may use
            # shape-to-fit natively — don't stomp it to None)
            try:
                if snap.get("autosize") is not None:
                    shape.TextFrame2.AutoSize = snap["autosize"]
            except Exception:
                logger.debug("AutoSize restore failed", exc_info=True)
            return True
        except Exception:
            self._record_failure("Shape restore failed (slide %s, shape %s)",
                                 slide_num, shape_index)
            return False

    @_tracked("replace_shape_with_image")
    def replace_shape_with_image(self, slide_num, shape_index, image_path):
        """Put an image into a shape WITHOUT disturbing shape indices.

        The old delete-then-AddPicture approach appended the new picture at
        the END of the Shapes collection, shifting every later shape down
        one index — after one image swap, all subsequent assignments (text,
        custom, images) silently hit the WRONG shapes.

        Now: regular shapes get a picture FILL (identity preserved, index
        unchanged); real picture shapes are swapped and the replacement is
        sent backward to the deleted shape's exact z-order position, which
        in PowerPoint COM is the same thing as its collection index.
        """
        if not self.is_active():
            return
        try:
            slide = self.presentation.Slides(slide_num)
            shape = slide.Shapes(shape_index + 1)
            abs_path = os.path.abspath(image_path)

            is_picture = False
            try:
                is_picture = shape.Type in (13, 11)  # msoPicture, msoLinkedPicture
            except Exception:
                logger.debug("Shape Type check failed", exc_info=True)

            if not is_picture:
                # Identity-preserving path: picture fill on the existing shape
                try:
                    if shape.HasTextFrame:
                        # Snapshot for Clear All, then blank the label text so
                        # it doesn't sit on top of the image
                        self._snapshot_shape(slide_num, shape_index, shape,
                                             shape.TextFrame)
                        tf = shape.TextFrame
                        count = tf.TextRange.Paragraphs().Count
                        for pi in range(1, count + 1):
                            self._set_paragraph_text(tf.TextRange.Paragraphs(pi), "")
                    shape.Fill.UserPicture(abs_path)
                    return
                except Exception:
                    logger.debug("Picture fill failed — falling back to swap",
                                 exc_info=True)

            # Swap path (real pictures, or fill failure): preserve z-order
            left, top = shape.Left, shape.Top
            width, height = shape.Width, shape.Height
            name = shape.Name
            old_pos = shape.ZOrderPosition
            shape.Delete()

            pic = slide.Shapes.AddPicture(
                abs_path, LinkToFile=False, SaveWithDocument=True,
                Left=left, Top=top, Width=width, Height=height)
            try:
                pic.Name = name
            except Exception:
                logger.debug("Could not carry shape name over", exc_info=True)
            try:
                # Shapes collection order == z-order; walking the new picture
                # back to the old slot restores index mapping for everything
                guard = 0
                while pic.ZOrderPosition > old_pos and guard < 500:
                    pic.ZOrder(3)  # msoSendBackward
                    guard += 1
            except Exception:
                logger.warning("Z-order restore failed — shape indices may "
                               "have shifted", exc_info=True)

        except Exception:
            self._record_failure("Live image replace failed (slide %s, shape %s)",
                                 slide_num, shape_index)

    @_tracked("update_chart_data")
    def update_chart_data(self, slide_num, shape_index, data_dict):
        """Update an existing chart's data via COM.

        Args:
            slide_num: 1-based slide number
            shape_index: 0-based shape index
            data_dict: {"categories": ["Roku", "CTV", ...],
                        "series": [{"name": "Impressions", "values": [20000, ...]}, ...]}
        """
        if not self.is_active():
            return
        try:
            slide = self.presentation.Slides(slide_num)
            shape = slide.Shapes(shape_index + 1)

            if not shape.HasChart:
                logger.warning("Shape %s on slide %s is not a chart",
                               shape_index, slide_num)
                return

            chart = shape.Chart
            chart_data = chart.ChartData

            # Activate the chart data workbook
            chart_data.Activate()
            wb = chart_data.Workbook
            ws = wb.Worksheets(1)

            categories = data_dict.get("categories", [])
            series_list = data_dict.get("series", [])

            # Write categories (column A, starting row 2)
            for i, cat in enumerate(categories):
                ws.Cells(i + 2, 1).Value = str(cat)
            # Clear extra rows
            for i in range(len(categories) + 2, ws.UsedRange.Rows.Count + 1):
                ws.Cells(i, 1).Value = ""

            # Write series data
            for si, series in enumerate(series_list):
                col = si + 2  # Column B, C, D...
                ws.Cells(1, col).Value = series.get("name", f"Series {si+1}")
                for ri, val in enumerate(series.get("values", [])):
                    ws.Cells(ri + 2, col).Value = val
                # Clear extra rows
                for ri in range(len(series.get("values", [])) + 2, ws.UsedRange.Rows.Count + 1):
                    ws.Cells(ri, col).Value = ""

            # Update the chart range
            wb.Close(True)

        except Exception:
            self._record_failure("Live chart update failed (slide %s, shape %s)",
                                 slide_num, shape_index)

    @_tracked("update_table_data")
    def update_table_data(self, slide_num, shape_index, data_dict):
        """Update an existing PowerPoint table shape via COM.

        Args:
            slide_num: 1-based slide number
            shape_index: 0-based shape index
            data_dict: {"headers": ["Device", "Impressions", ...],
                        "rows": [["Roku", 20000], ["CTV", 15000], ...]}
        """
        if not self.is_active():
            return
        try:
            slide = self.presentation.Slides(slide_num)
            shape = slide.Shapes(shape_index + 1)

            if not shape.HasTable:
                logger.warning("Shape %s on slide %s is not a table",
                               shape_index, slide_num)
                return

            table = shape.Table
            headers = data_dict.get("headers", [])
            rows = data_dict.get("rows", [])
            keep_header = data_dict.get("keep_header_row", False)

            # Write headers — unless the table's existing header row should
            # be preserved (data then starts at row 2 as usual)
            if not keep_header:
                for ci, h in enumerate(headers):
                    if ci < table.Columns.Count:
                        table.Cell(1, ci + 1).Shape.TextFrame.TextRange.Text = str(h)

            # Write data rows
            for ri, row in enumerate(rows):
                if ri + 1 < table.Rows.Count:  # +1 for header row
                    for ci, val in enumerate(row):
                        if ci < table.Columns.Count:
                            if isinstance(val, (int, float)):
                                display = f"{val:,}" if isinstance(val, int) else f"{val:,.2f}"
                            else:
                                display = str(val)
                            table.Cell(ri + 2, ci + 1).Shape.TextFrame.TextRange.Text = display

        except Exception:
            self._record_failure("Live table update failed (slide %s, shape %s)",
                                 slide_num, shape_index)

    @_tracked("save_as")
    def save_as(self, output_path):
        """Save the current state to a new file. Returns True on success."""
        if not self.is_active():
            return False
        try:
            abs_path = os.path.abspath(output_path)
            self.presentation.SaveAs(abs_path)
            logger.info("Live preview saved to %s", abs_path)
            return True
        except Exception:
            self._record_failure("Live preview save failed: %s", output_path)
            return False

    def cleanup(self):
        """Close PowerPoint (if we own the only presentation) and clean up."""
        if self.presentation:
            try:
                self.presentation.Close()
            except Exception:
                logger.debug("Presentation close failed", exc_info=True)
            self.presentation = None

        # Don't quit PowerPoint if user has other presentations open
        if self.pptx_app:
            try:
                if self.pptx_app.Presentations.Count == 0:
                    self.pptx_app.Quit()
            except Exception:
                logger.debug("PowerPoint quit failed", exc_info=True)
            self.pptx_app = None

        # Clean up working copy
        if self._working_copy and os.path.exists(self._working_copy):
            try:
                os.remove(self._working_copy)
            except OSError:
                logger.debug("Could not remove working copy %s",
                             self._working_copy, exc_info=True)

        if _com_available:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                logger.debug("CoUninitialize failed", exc_info=True)

        self.active = False

    def __del__(self):
        try:
            self.cleanup()
        except Exception:
            pass  # Never raise from a destructor
