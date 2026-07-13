"""Slide/shape panel methods for the template mapper."""

import logging
import os
import shutil
import tkinter as tk
from tkinter import messagebox

from config.paths import APP_ROOT, IMAGES_DIR
from engine.pptx_formats import _format_value
from engine.query_resolver import resolve_query

logger = logging.getLogger(__name__)


class SlideViewMixin:
    """Slide rendering and shape-assignment methods for PPTXWizard."""

    def _propagate_format_details(self, metric_key):
        """Push new format details into every existing assignment that uses
        this metric, then re-render the current slide in place. Re-formatting
        never requires re-assigning (which used to wipe assignment lists)."""
        details = getattr(self, '_metric_format_details', {}).get(metric_key)
        if not details:
            return
        touched = 0
        for smap_slides in self.mapping.get("slides", {}).values():
            for smap in smap_slides.get("shape_mappings", {}).values():
                targets = smap.get("assignments") or (
                    [smap] if smap.get("metric") else [])
                for asgn in targets:
                    if asgn.get("metric") == metric_key:
                        asgn["format_details"] = details
                        asgn["format"] = details.get("format",
                                                     asgn.get("format", "text"))
                        touched += 1
        if touched:
            logger.info("Format change propagated to %d assignment(s) of %s",
                        touched, metric_key)
            self._show_slide()  # re-applies with last_rendered replacement

    @staticmethod
    def _display_for(val, fmt, fmt_details, existing_text):
        """Format a value for insertion: detailed settings for numbers when
        present, otherwise the standard formatter (dates, case matching)."""
        from engine.pptx_formats import _coerce_number, format_with_details
        coerced = _coerce_number(val)
        if fmt_details and fmt_details.get("format") == "date":
            return format_with_details(coerced, fmt_details)
        if fmt_details and isinstance(coerced, (int, float)) and not isinstance(coerced, bool):
            return format_with_details(coerced, fmt_details)
        return _format_value(val, fmt, existing_text)

    def _show_slide(self):
        t = self.t
        for w in self.shapes_frame.winfo_children(): w.destroy()
        if not self.slides: return

        slide = self.slides[self.current_slide]
        snum = slide["slide_num"]
        title = slide["slide_title"] or f"Slide {snum}"
        self.slide_label.config(text=f"Slide {snum}/{len(self.slides)}: {title}")
        self.slide_frame.config(text=f"  Slide {snum}: {title}  ")

        slide_map = self.mapping.get("slides", {}).get(str(snum), {})
        shape_maps = slide_map.get("shape_mappings", {})

        # Navigate live preview to this slide and re-apply all assignments
        if self.live_preview:
            self.live_preview.go_to_slide(snum)
            # Re-apply all existing assignments for this slide
            for shape in slide["shapes"]:
                s_id = str(shape["shape_id"])
                smap = shape_maps.get(s_id, {})
                if smap.get("skip", False) or not smap:
                    continue
                if smap.get("image_path"):
                    continue

                # Get all assignments (multi or single)
                assignments = smap.get("assignments", [])
                if not assignments and smap.get("metric"):
                    # Use the mapping dict itself so last_rendered persists
                    assignments = [smap]

                existing = self.original_texts.get((snum, int(s_id)), "")
                for asgn in assignments:
                    mk = asgn.get("metric", "")
                    val = self.available_metrics.get(mk)

                    # Resolve query if value not in available_metrics
                    if val is None and (asgn.get("query") or mk.startswith("__total_")):
                        try:
                            q = asgn.get("query", mk)
                            val = resolve_query(q, self.export_result["client_data"],
                                self.export_result.get("client_name", ""),
                                self._get_start_date(), self._get_end_date())
                            self.available_metrics[mk] = val
                        except Exception:
                            logger.warning("Query resolution failed for %s", mk,
                                           exc_info=True)

                    if val is not None and mk:
                        fmt = asgn.get("format", "text")
                        rt = asgn.get("replace_text", "")
                        display = self._display_for(val, fmt,
                                                    asgn.get("format_details"),
                                                    rt or existing)
                        candidates = [c for c in
                                      [asgn.get("last_rendered"), rt] if c] or None
                        self.live_preview.update_shape_text(
                            snum, int(s_id), display,
                            replace_portion=candidates,
                            expected_name=next(
                                (sh.get("name") for sh in slide["shapes"]
                                 if str(sh["shape_id"]) == str(s_id)), None))
                        asgn["last_rendered"] = display

        if not slide["shapes"]:
            tk.Label(self.shapes_frame, text="No shapes with text on this slide",
                     font=("Calibri", 10), bg=t["card"], fg=t["muted"]).pack(pady=20)
            return

        for shape in slide["shapes"]:
            sid = str(shape["shape_id"])
            smap = shape_maps.get(sid, {})
            assigned = smap.get("metric", "")
            is_skipped = smap.get("skip", False)
            fmt = smap.get("format", "number")
            replace_text = smap.get("replace_text", "")
            img = smap.get("image_path", "")

            sf = tk.Frame(self.shapes_frame, bg=t["bg"] if not is_skipped else "#F0F0F0",
                          relief="solid", borderwidth=1)
            sf.pack(fill="x", padx=5, pady=3)

            # Header row: name + controls
            hdr = tk.Frame(sf, bg=sf.cget("bg"))
            hdr.pack(fill="x", padx=10, pady=(5, 0))

            tk.Label(hdr, text=shape["name"], font=("Calibri", 10, "bold"),
                     bg=sf.cget("bg"), fg=t["fg"] if not is_skipped else t["muted"]
                     ).pack(side="left")

            # Skip checkbox
            skip_var = tk.BooleanVar(value=is_skipped)
            def toggle_skip(sv=skip_var, sh=shape):
                self._toggle_skip(sh, sv.get())
            tk.Checkbutton(hdr, text="Skip", variable=skip_var, command=toggle_skip,
                           font=("Calibri", 8), bg=sf.cget("bg"),
                           fg=t["muted"]).pack(side="right")

            if not is_skipped:
                # Text preview with selectable text for partial replacement
                preview = tk.Frame(sf, bg=sf.cget("bg"))
                preview.pack(fill="x", padx=10, pady=(2, 5))

                current_text = shape["text"][:100] if shape["text"] else "(empty)"

                text_widget = tk.Text(preview, height=2, width=50, font=("Calibri", 9),
                                       bg=t["input_bg"], fg=t["fg"], relief="solid",
                                       borderwidth=1, wrap="word")
                text_widget.insert("1.0", current_text)
                text_widget.config(state="normal")
                text_widget.pack(side="left", fill="x", expand=True)

                def assign(sh=shape, tw=text_widget):
                    self._assign_to_shape(sh, tw)
                tk.Button(preview, text="Assign ←", font=("Calibri", 8, "bold"),
                          bg=t["accent"], fg="white", relief="flat", padx=8, pady=2,
                          command=assign).pack(side="right", padx=(5, 0))

                # Show current assignment(s)
                assignments_list = smap.get("assignments", [])
                if not assignments_list and assigned:
                    # Single assignment format
                    assignments_list = [{"metric": assigned, "format": fmt,
                                          "replace_text": replace_text}]

                for asgn in assignments_list:
                    a_metric = asgn.get("metric", "")
                    a_replace = asgn.get("replace_text", "")
                    display_name = self.SPECIAL_METRICS.get(a_metric, a_metric)
                    val = self.available_metrics.get(a_metric, "?")
                    vd = self._format_for_sidebar(a_metric, val)

                    assign_frame = tk.Frame(sf, bg=sf.cget("bg"))
                    assign_frame.pack(fill="x", padx=10, pady=(0, 2))
                    info = f"→  {display_name} = {vd}"
                    if a_replace:
                        info += f'  (replaces: "{a_replace}")'
                    tk.Label(assign_frame, text=info, font=("Calibri", 9, "bold"),
                             bg=sf.cget("bg"), fg="#1E6E3E", wraplength=330,
                             justify="left").pack(side="left")

                if assignments_list or assigned:
                    clear_frame = tk.Frame(sf, bg=sf.cget("bg"))
                    clear_frame.pack(anchor="e", padx=10, pady=(0, 5))
                    tk.Label(clear_frame, text="✕ Clear All", font=("Calibri", 8),
                             bg=sf.cget("bg"), fg=t["danger"], cursor="hand2"
                             ).pack(side="right")
                    clear_frame.winfo_children()[-1].bind("<Button-1>", lambda e, s=shape: self._clear_shape(s))

                elif img:
                    img_frame = tk.Frame(sf, bg=sf.cget("bg"))
                    img_frame.pack(fill="x", padx=10, pady=(0, 5))
                    tk.Label(img_frame, text=f"→  Image: {os.path.basename(img)}",
                             font=("Calibri", 9, "bold"), bg=sf.cget("bg"), fg="#1E6E3E"
                             ).pack(side="left")
                    tk.Label(img_frame, text="✕ Clear", font=("Calibri", 8),
                             bg=sf.cget("bg"), fg=t["danger"], cursor="hand2"
                             ).pack(side="right")
                    img_frame.winfo_children()[-1].bind("<Button-1>", lambda e, s=shape: self._clear_shape(s))

        # Refresh slide preview thumbnail
        self.window.after(200, self._refresh_preview)

    def _assign_to_shape(self, shape, text_widget):
        if not self.selected_metric:
            messagebox.showinfo("Select", "Click a metric in the sidebar first.",
                                parent=self.window)
            return

        snum_str = str(self.slides[self.current_slide]["slide_num"])
        snum_int = self.slides[self.current_slide]["slide_num"]
        sid = str(shape["shape_id"])

        if "slides" not in self.mapping: self.mapping["slides"] = {}
        if snum_str not in self.mapping["slides"]:
            self.mapping["slides"][snum_str] = {"shape_mappings": {}}

        # Check for highlighted text (partial replacement)
        replace_text = ""
        try:
            sel = text_widget.selection_get()
            if sel and sel.strip():
                replace_text = sel.strip()
        except tk.TclError:
            pass  # No selection — full replacement

        # Get format from metric's stored format (set via right-click)
        metric_fmt = self._metric_formats.get(self.selected_metric, "text")

        if (self.selected_metric or "").startswith("__image") and getattr(self, '_pending_image', None):
            # Copy image to internal templates/images folder
            os.makedirs(IMAGES_DIR, exist_ok=True)
            img_filename = os.path.basename(self._pending_image)
            internal_img = os.path.join(IMAGES_DIR, img_filename)
            if os.path.abspath(self._pending_image) != os.path.abspath(internal_img):
                try:
                    shutil.copy2(self._pending_image, internal_img)
                except OSError:
                    logger.warning("Could not copy image to internal folder: %s",
                                   self._pending_image, exc_info=True)

            # Store relative path for portability
            rel_path = os.path.relpath(internal_img, APP_ROOT)

            self.mapping["slides"][snum_str]["shape_mappings"][sid] = {
                "metric": "",
                "image_path": rel_path,
                "image_path_abs": self._pending_image,  # Keep absolute for current session
                "skip": False,
                "assignments": [],
            }
            # Live preview: use absolute path for current session
            if self.live_preview:
                self.live_preview.replace_shape_with_image(
                    snum_int, int(sid), self._pending_image)
        else:
            new_assignment = {
                "metric": self.selected_metric,
                "format": metric_fmt,
                "replace_text": replace_text,
            }
            # Detailed format settings (decimals/commas/prefix/suffix) travel
            # with the assignment so fill paths can apply them
            fmt_details = getattr(self, '_metric_format_details', {}).get(self.selected_metric)
            if fmt_details:
                new_assignment["format_details"] = fmt_details
            if getattr(self, '_pending_query', None):
                new_assignment["query"] = self._pending_query

            # Get or create existing shape mapping
            existing = self.mapping["slides"][snum_str]["shape_mappings"].get(sid, {})
            existing_asgns = existing.get("assignments", [])
            if not existing_asgns and existing.get("metric"):
                # normalize legacy single-assignment into the list form so
                # every path below appends/updates instead of overwriting
                first = {"metric": existing["metric"],
                         "format": existing.get("format", "text"),
                         "replace_text": existing.get("replace_text", ""),
                         "format_details": existing.get("format_details"),
                         "last_rendered": existing.get("last_rendered")}
                if existing.get("query"):
                    first["query"] = existing["query"]
                existing_asgns = [first]

            # Same metric already assigned on this shape? Update it in place
            # (new format/details/query, keep its replace target) instead of
            # wiping the shape's assignment list.
            updated_in_place = None
            for asgn in existing_asgns:
                if asgn.get("metric") == self.selected_metric:
                    asgn["format"] = metric_fmt
                    if fmt_details:
                        asgn["format_details"] = fmt_details
                    if getattr(self, '_pending_query', None):
                        asgn["query"] = self._pending_query
                    if replace_text:
                        asgn["replace_text"] = replace_text
                    updated_in_place = asgn
                    break

            _lr_targets = []
            if updated_in_place is not None:
                self.mapping["slides"][snum_str]["shape_mappings"][sid] = {
                    "skip": False, "assignments": existing_asgns}
                new_assignment = updated_in_place
                _lr_targets = [updated_in_place]
            elif existing_asgns and not replace_text:
                # Full replace onto a shape that already has assignments
                # would previously nuke them all silently — ask first.
                if not messagebox.askyesno(
                        "Replace assignments?",
                        f"This shape already has {len(existing_asgns)} "
                        f"assignment(s).\n\nAssign without highlighting text "
                        f"replaces ALL of them with a full-text assignment.\n\n"
                        f"Replace them?\n(Tip: highlight the text to replace "
                        f"for a partial assignment instead.)",
                        parent=self.window):
                    return
                shape_map = dict(new_assignment)
                shape_map["skip"] = False
                shape_map["assignments"] = []
                self.mapping["slides"][snum_str]["shape_mappings"][sid] = shape_map
                _lr_targets = [new_assignment, shape_map]
            elif existing_asgns:
                existing_asgns.append(new_assignment)
                self.mapping["slides"][snum_str]["shape_mappings"][sid] = {
                    "skip": False, "assignments": existing_asgns}
                _lr_targets = [new_assignment]
            else:
                # First assignment on this shape
                shape_map = dict(new_assignment)
                shape_map["skip"] = False
                shape_map["assignments"] = []
                self.mapping["slides"][snum_str]["shape_mappings"][sid] = shape_map
                _lr_targets = [new_assignment, shape_map]

            # Live preview: update based on output type
            if self.live_preview:
                output_type = self._pending_query.get("output", "value") if self._pending_query else "value"

                if output_type == "chart" and getattr(self, '_pending_chart_data', None):
                    self.live_preview.update_chart_data(
                        snum_int, int(sid), self._pending_chart_data)
                    self._pending_chart_data = None  # consume — no stale re-use
                elif output_type == "table" and getattr(self, '_pending_table_data', None):
                    self.live_preview.update_table_data(
                        snum_int, int(sid), self._pending_table_data)
                    self._pending_table_data = None  # consume — no stale re-use
                else:
                    val = self.available_metrics.get(self.selected_metric, "")
                    existing_text = self.original_texts.get((snum_int, int(sid)), "")
                    eff_fmt = new_assignment.get("format", metric_fmt)
                    eff_details = new_assignment.get("format_details") or fmt_details
                    eff_rt = new_assignment.get("replace_text", "") or replace_text
                    # Partial replace: detect date style / case from the target
                    display = self._display_for(val, eff_fmt, eff_details,
                                                eff_rt or existing_text)
                    candidates = [c for c in
                                  [new_assignment.get("last_rendered"), eff_rt]
                                  if c] or None
                    self.live_preview.update_shape_text(
                        snum_int, int(sid), display,
                        replace_portion=candidates,
                        expected_name=shape.get("name"))
                    for tgt in _lr_targets:
                        tgt["last_rendered"] = display

        self._show_slide()

    def _toggle_skip(self, shape, skip):
        snum = str(self.slides[self.current_slide]["slide_num"])
        sid = str(shape["shape_id"])
        if "slides" not in self.mapping: self.mapping["slides"] = {}
        if snum not in self.mapping["slides"]:
            self.mapping["slides"][snum] = {"shape_mappings": {}}
        self.mapping["slides"][snum]["shape_mappings"][sid] = {"skip": skip}
        self._show_slide()

    def _clear_shape(self, shape):
        snum_str = str(self.slides[self.current_slide]["slide_num"])
        snum_int = self.slides[self.current_slide]["slide_num"]
        sid = str(shape["shape_id"])
        sm = self.mapping.get("slides", {}).get(snum_str, {}).get("shape_mappings", {})
        if sid in sm:
            del sm[sid]
        # Restore the shape's original text from the pre-modification
        # snapshot (global Undo was unreliable: re-applied assignments pile
        # up COM operations that a single undo cannot unwind).
        if self.live_preview and self.live_preview.is_active():
            restored = self.live_preview.restore_shape_text(snum_int, int(sid))
            if not restored:
                logger.debug("No snapshot to restore for shape %s", sid)
        self._show_slide()

    def _prev_slide(self):
        if self.current_slide > 0:
            self.current_slide -= 1; self._show_slide()

    def _next_slide(self):
        if self.current_slide < len(self.slides) - 1:
            self.current_slide += 1; self._show_slide()

    # Fixed thumbnail bounds — keeps preview from growing with each metric assignment
    _PREVIEW_MAX_W = 660
    _PREVIEW_MAX_H = 490

    def _refresh_preview(self):
        """Export current slide as PNG and display in the preview panel."""
        if not self.live_preview or not self.live_preview.is_active():
            return
        try:
            snum = self.slides[self.current_slide]["slide_num"]
            img_path = self.live_preview.export_slide_image(snum, width=960)
            if img_path and os.path.exists(img_path):
                from PIL import Image, ImageTk
                img = Image.open(img_path)
                # Adaptive-but-CAPPED bounds: the fixed 660x490 didn't fit
                # windowed/scaled laptop screens and clipped the preview
                # (looked like the old growth bug). Computed fresh from the
                # window at refresh time — no <Configure> binding, so the
                # historical resize feedback loop cannot return. On large
                # screens the caps make behavior identical to before.
                try:
                    # Lock to the actual (size-locked) container — the true
                    # "dynamically locked to the window" behavior. The frame
                    # cannot grow (propagate off), so no feedback loop.
                    avail_w = max(280, min(self._PREVIEW_MAX_W,
                                           self.preview_label.winfo_width() - 12))
                    avail_h = max(220, min(self._PREVIEW_MAX_H,
                                           self.preview_label.winfo_height() - 12))
                except Exception:
                    avail_w, avail_h = self._PREVIEW_MAX_W, self._PREVIEW_MAX_H
                img.thumbnail((avail_w, avail_h), Image.LANCZOS)
                self._preview_image = ImageTk.PhotoImage(img)
                self.preview_label.config(
                    image=self._preview_image, text="",
                    width=avail_w, height=avail_h)
        except Exception:
            logger.debug("Slide preview refresh failed", exc_info=True)
            self.preview_label.config(text="Preview error", image="")
