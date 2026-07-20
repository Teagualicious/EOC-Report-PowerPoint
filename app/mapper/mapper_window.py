"""PowerPoint template mapper window (PPTXWizard).

Three-column layout: metric sidebar (SidebarMixin), COM-rendered slide
preview, and shape list (SlideViewMixin). The advanced query builder and
number-format popup live in mapper.query_builder / mapper.format_popup.
"""

import logging
import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from config.paths import TEMPLATES_DIR
from engine.pptx_mapper import (SPECIAL_METRICS, fill_template,
                                get_available_metrics, load_template_mapping,
                                save_template_mapping, scan_template)
from mapper.mapping_model import MappingModel
from mapper.query_builder import show_query_builder
from mapper.sidebar import SidebarMixin
from mapper.slide_view import SlideViewMixin

logger = logging.getLogger(__name__)


class PPTXWizard(SidebarMixin, SlideViewMixin):
    """Slide-by-slide template mapper with live PowerPoint preview."""

    def __init__(self, parent, export_result, theme, template_path=None):
        self.scan_template = scan_template
        self.load_mapping = load_template_mapping
        self.save_mapping = save_template_mapping
        self.fill_template = fill_template
        self.SPECIAL_METRICS = SPECIAL_METRICS

        self.t = theme
        self.export_result = export_result
        self.template_path = template_path
        self.slides = []
        self.current_slide = 0
        # Single owner of mapping state — the Tk panels and the live COM
        # preview re-render from it after every mutation (see _on_model_change)
        self.model = MappingModel()
        self.selected_metric = None
        self.original_texts = {}
        self._pending_query = None
        self._pending_image = None
        self.live_preview = None  # COM-based live preview

        # Get available metrics including special ones
        client_name = export_result.get("client_name", "")
        start_date = ""
        end_date = ""
        for d in export_result.get("client_data", []):
            if d.get("start_date"): start_date = d["start_date"]
            if d.get("end_date"): end_date = d["end_date"]
            if start_date: break
        self.available_metrics, self.structured_metrics = get_available_metrics(
            export_result["client_data"], client_name, start_date, end_date)

        # Ask for template (unless a path was passed in, e.g. Settings Edit)
        if not self.template_path:
            self.template_path = filedialog.askopenfilename(
                title="Select PowerPoint Template",
                filetypes=[("PowerPoint", "*.pptx")])
        if not self.template_path: return

        # Copy template to internal templates folder for portability
        os.makedirs(TEMPLATES_DIR, exist_ok=True)
        internal_path = os.path.join(TEMPLATES_DIR, os.path.basename(self.template_path))
        if os.path.abspath(self.template_path) != os.path.abspath(internal_path):
            try:
                shutil.copy2(self.template_path, internal_path)
            except OSError:
                logger.warning("Could not copy template to internal folder",
                               exc_info=True)

        try:
            self.slides = self.scan_template(self.template_path)
        except Exception as e:
            logger.exception("Template scan failed: %s", self.template_path)
            messagebox.showerror("Error",
                f"Failed to read the template:\n{e}")
            return
        if not self.slides:
            messagebox.showwarning("Empty", "No slides found."); return

        # Save original text for undo on Clear
        for slide in self.slides:
            for shape in slide["shapes"]:
                key = (slide["slide_num"], shape["shape_id"])
                self.original_texts[key] = shape.get("text", "")

        saved = self.load_mapping(os.path.basename(self.template_path))
        if saved: self.model = MappingModel(saved)
        # Let mutations stamp each entry with the scanned persistent shape
        # identity (uid + name) so fills survive template edits
        self.model.set_scan_identity(self.slides)
        self.model.subscribe(self._on_model_change)

        # Launch live PowerPoint preview
        try:
            from engine.pptx_live import PPTXLivePreview
            self.live_preview = PPTXLivePreview(self.template_path)
            if not self.live_preview.is_active():
                self.live_preview = None
            else:
                self.live_preview.on_disabled = self._on_preview_disabled
        except Exception:
            logger.warning("Live preview unavailable — falling back to "
                           "static mode", exc_info=True)
            self.live_preview = None

        self._build_window(parent)

    def _build_window(self, parent):
        t = self.t
        self.window = tk.Toplevel(parent)
        self.window.title(f"Template Mapper — {os.path.basename(self.template_path)}")
        # Clamp to the actual screen (laptops/DPI scaling) and center —
        # a fixed 1400x800 hung off-screen on smaller displays and took
        # the slide preview with it. Zoomed state below still applies on
        # screens that support it; when windowed/restored, this size
        # always fits. The preview bounds (slide_view) derive from the
        # window, so they scale with this automatically.
        sw = self.window.winfo_screenwidth()
        sh = self.window.winfo_screenheight()
        w = min(1400, sw - 40)
        h = min(800, sh - 90)
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        self.window.geometry(f"{w}x{h}+{x}+{y}")
        try:
            self.window.state("zoomed")
        except tk.TclError:
            logger.debug("Window zoom not supported", exc_info=True)
        self.window.configure(bg=t["bg"])
        self.window.grab_set()
        self.window.lift()
        self.window.focus_force()
        self.window.protocol("WM_DELETE_WINDOW", self._on_window_close)

        # Top bar
        top = tk.Frame(self.window, bg=t["accent"])
        top.pack(fill="x")
        tk.Label(top, text=f"📑 {os.path.basename(self.template_path)}",
                 font=("Calibri", 12, "bold"), bg=t["accent"], fg="white"
                 ).pack(side="left", padx=15, pady=8)
        # Live preview indicator
        if self.live_preview:
            tk.Label(top, text="🟢 Live Preview Active", font=("Calibri", 9),
                     bg=t["accent"], fg="#90EE90").pack(side="left", padx=10, pady=8)
        else:
            tk.Label(top, text="⚪ Static Mode", font=("Calibri", 9),
                     bg=t["accent"], fg="#888888").pack(side="left", padx=10, pady=8)
        self.slide_label = tk.Label(top, font=("Calibri", 10), bg=t["accent"], fg="#B0C4DE")
        self.slide_label.pack(side="right", padx=15, pady=8)

        # Main: sidebar + preview + shapes (3 columns)
        main = tk.Frame(self.window, bg=t["bg"])
        main.pack(fill="both", expand=True, padx=10, pady=10)
        main.columnconfigure(0, weight=0)  # sidebar fixed
        main.columnconfigure(1, weight=0)  # preview fixed size (locked thumbnail)
        main.columnconfigure(2, weight=1)  # shapes gets remaining space for editing
        main.rowconfigure(0, weight=1)

        # ── Left: Sidebar ──
        sidebar = tk.LabelFrame(main, text="  Metrics & Values  ", font=("Calibri", 10, "bold"),
                                 bg=t["card"], fg=t["card_fg"], padx=8, pady=5)
        sidebar.grid(row=0, column=0, sticky="ns", padx=(0, 5))

        tk.Label(sidebar, text="Click metric, then click shape.\nHighlight text for partial replace.",
                 font=("Calibri", 8), bg=t["card"], fg=t["muted"], justify="left").pack(pady=(0, 5))

        self.metric_search_var = tk.StringVar()
        self.metric_search_var.trace_add("write", lambda *_: self._refresh_metrics())
        tk.Entry(sidebar, textvariable=self.metric_search_var, width=24,
                 font=("Calibri", 9), bg=t["input_bg"], fg=t["input_fg"],
                 insertbackground=t["insert"], relief="solid", borderwidth=1).pack(pady=(0, 5))

        self.metric_canvas = tk.Canvas(sidebar, bg=t["card"], highlightthickness=0, width=220)
        msb = ttk.Scrollbar(sidebar, orient="vertical", command=self.metric_canvas.yview)
        self.metric_frame = tk.Frame(self.metric_canvas, bg=t["card"])
        self.metric_frame.bind("<Configure>",
            lambda e: self.metric_canvas.configure(scrollregion=self.metric_canvas.bbox("all")))
        _mcw = self.metric_canvas.create_window((0, 0), window=self.metric_frame, anchor="nw")
        self.metric_canvas.bind(
            "<Configure>",
            lambda e: self.metric_canvas.itemconfigure(_mcw, width=e.width))
        self.metric_canvas.configure(yscrollcommand=msb.set)
        self.metric_canvas.pack(side="left", fill="both", expand=True)
        msb.pack(side="right", fill="y")

        self.metric_buttons = {}
        self._refresh_metrics()

        # ── Center: Slide Preview ──
        # Size-locked to prevent growth (the classic bug), but the LOCKED
        # size is computed from the screen: the shapes panel (col 2) is
        # guaranteed ~430px so its cards and assign targets always fit;
        # the preview absorbs the squeeze on small screens. Caps keep big
        # monitors identical to the old fixed 680x520.
        _sw = self.window.winfo_screenwidth()
        _sh = self.window.winfo_screenheight()
        _win_w = min(1400, _sw - 40)
        _win_h = min(800, _sh - 90)
        _pw = min(680, max(400, _win_w - 720))   # 720 = sidebar + shapes + pads
        _ph = min(520, max(320, _win_h - 210))
        preview_frame = tk.LabelFrame(main, text="  Slide Preview  ", font=("Calibri", 10, "bold"),
                                       bg=t["card"], fg=t["card_fg"], padx=5, pady=5,
                                       width=_pw, height=_ph)
        preview_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        preview_frame.pack_propagate(False)  # Lock frame size — thumbnail must not grow
        preview_frame.grid_propagate(False)

        self.preview_label = tk.Label(preview_frame, bg="#333333", text="Loading preview...",
                                       fg="#888888", font=("Calibri", 10))
        self.preview_label.pack(fill="both", expand=True)
        self._preview_image = None  # Hold reference to prevent GC

        # ── Right: Shapes ──
        self.slide_frame = tk.LabelFrame(main, text="  Shapes  ", font=("Calibri", 10, "bold"),
                                          bg=t["card"], fg=t["card_fg"], padx=10, pady=5)
        self.slide_frame.grid(row=0, column=2, sticky="nsew", padx=(5, 0))

        self.shapes_canvas = tk.Canvas(self.slide_frame, bg=t["card"], highlightthickness=0)
        ssb = ttk.Scrollbar(self.slide_frame, orient="vertical", command=self.shapes_canvas.yview)
        self.shapes_frame = tk.Frame(self.shapes_canvas, bg=t["card"])
        self.shapes_frame.bind("<Configure>",
            lambda e: self.shapes_canvas.configure(scrollregion=self.shapes_canvas.bbox("all")))
        _shw = self.shapes_canvas.create_window((0, 0), window=self.shapes_frame, anchor="nw")
        # Width-lock the inner frame to the canvas viewport: long assignment
        # labels can no longer widen the panel and push Assign buttons out
        # of frame (sets item width only — no resize feedback possible)
        self.shapes_canvas.bind(
            "<Configure>",
            lambda e: self.shapes_canvas.itemconfigure(_shw, width=e.width))
        self.shapes_canvas.configure(yscrollcommand=ssb.set)
        self.shapes_canvas.pack(side="left", fill="both", expand=True)
        ssb.pack(side="right", fill="y")

        # ── Bottom ──
        bottom = tk.Frame(self.window, bg=t["bg"])
        bottom.pack(fill="x", padx=10, pady=(0, 10))

        tk.Button(bottom, text="← Prev", font=("Calibri", 10), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=15, pady=6,
                  command=self._prev_slide).pack(side="left")
        tk.Button(bottom, text="Next →", font=("Calibri", 10), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=15, pady=6,
                  command=self._next_slide).pack(side="left", padx=5)

        tk.Button(bottom, text="Save & Fill Template", font=("Calibri", 11, "bold"),
                  bg="#1E6E3E", fg="white", relief="flat", padx=20, pady=8,
                  command=self._save_and_fill).pack(side="right")
        tk.Button(bottom, text="Save Mapping", font=("Calibri", 10),
                  bg=t["accent"], fg="white", relief="flat", padx=15, pady=8,
                  command=self._save_only).pack(side="right", padx=5)

        self._show_slide()

    # Compatibility views onto the model — every read the mixins and the
    # save/fill paths do goes to the same single source of truth.
    @property
    def mapping(self):
        return self.model.data

    @property
    def _metric_formats(self):
        return self.model.metric_formats

    @property
    def _metric_format_details(self):
        return self.model.metric_format_details

    def _on_model_change(self, event, slide_num=None, shape_id=None, metric=None):
        """Model observer: re-render the current slide (shape panel + live
        preview re-apply) after every mapping mutation."""
        if hasattr(self, "window") and self.window.winfo_exists():
            self._show_slide()

    def _on_preview_disabled(self, reason):
        """Live preview turned itself off after repeated COM errors —
        tell the user once, without interrupting whatever they're doing."""
        def notify():
            try:
                messagebox.showwarning(
                    "Live Preview Off",
                    "PowerPoint live preview was turned off after repeated "
                    f"errors ({reason}).\n\nYour mapping is unaffected — "
                    "keep assigning as normal, and Save & Fill will use the "
                    "built-in engine.", parent=self.window)
            except Exception:
                logger.warning("Could not show preview-disabled notice")
        try:
            self.window.after(0, notify)
        except Exception:
            logger.warning("Preview disabled: %s", reason)

    def _show_query_builder(self):
        """Advanced query builder (delegates to mapper.query_builder)."""
        self._chart_target = None   # sidebar open — no chart is targeted
        show_query_builder(self)

    def _save_only(self):
        try:
            self.save_mapping(os.path.basename(self.template_path), self.mapping)
        except OSError as e:
            logger.exception("Could not save mapping")
            messagebox.showerror("Save Failed",
                f"Could not save the mapping:\n{e}", parent=self.window)
            return
        messagebox.showinfo("Saved", f"Mapping saved for: {os.path.basename(self.template_path)}",
                            parent=self.window)

    def _save_and_fill(self):
        try:
            self.save_mapping(os.path.basename(self.template_path), self.mapping)
        except OSError:
            logger.exception("Could not save mapping before fill")
        client_name = self.export_result.get("client_name", "Report")
        save_path = filedialog.asksaveasfilename(
            title="Save Filled Report",
            parent=self.window,
            initialdir=self.export_result.get("folder", ""),
            initialfile=f"{client_name}_Report.pptx",
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx")])
        if not save_path: return

        # If live preview is active, save from there (already has the changes)
        if self.live_preview and self.live_preview.is_active():
            if self.live_preview.save_as(save_path):
                messagebox.showinfo("Success", f"Report generated (from live preview):\n{save_path}")
                self._cleanup_and_close()
                return

        # Fallback: use python-pptx to fill
        try:
            # Add client_data to metrics for query resolution during fill
            fill_metrics = dict(self.available_metrics)
            fill_metrics["__client_data__"] = self.export_result.get("client_data", [])
            from engine.fill_report import append_fill_history
            from engine.pptx_fill import fill_template_report
            _, report = fill_template_report(self.template_path, save_path,
                                             self.mapping, fill_metrics)
            append_fill_history(report)
            if report.ok:
                messagebox.showinfo("Success",
                    f"Report generated:\n{save_path}\n\n{report.summary()}")
            else:
                messagebox.showwarning("Report Generated — With Gaps",
                    f"Report saved to:\n{save_path}\n\n{report.summary()}",
                    parent=self.window)
            self._cleanup_and_close()
        except Exception as e:
            logger.exception("Template fill failed")
            messagebox.showerror("Error",
                f"Could not generate the report:\n{e}", parent=self.window)

    def _cleanup_and_close(self):
        """Clean up live preview and close the mapper window."""
        if self.live_preview:
            self.live_preview.cleanup()
            self.live_preview = None
        if hasattr(self, 'window') and self.window:
            self.window.destroy()

    def _on_window_close(self):
        """Handle window X button — clean up COM."""
        if self.live_preview:
            if messagebox.askyesno("Close", "Close template mapper?\nUnsaved changes to the live preview will be lost."):
                self._cleanup_and_close()
        else:
            self.window.destroy()
