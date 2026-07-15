"""Shared tkinter helpers and modern visual primitives."""

from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import ttk

_BASE_TK_SCALING = 96 / 72  # Tk's scaling value at 100% Windows DPI.


def _scroll_nearest(widget, delta):
    """Scroll the nearest scrollable ancestor of ``widget`` by ``delta``."""
    target = widget
    while target is not None:
        if isinstance(target, (tk.Canvas, ttk.Treeview, tk.Listbox, tk.Text)):
            try:
                units = -1 if delta > 0 else 1
                if abs(delta) >= 120:
                    units = -1 * (delta // 120)
                target.yview_scroll(units, "units")
            except tk.TclError:
                pass
            return
        target = getattr(target, "master", None)


def enable_mousewheel(window):
    """Route mouse-wheel events to the scrollable widget under the cursor.

    Also neutralizes the ttk.Combobox class binding that spins a closed
    combobox's VALUE on wheel — hovering a role/platform dropdown while
    scrolling a pane silently corrupted selections. The wheel still scrolls
    the surrounding pane (and an open dropdown list scrolls normally); it
    just never changes a closed combobox's selection.
    """
    def _on_mousewheel(event):
        _scroll_nearest(event.widget, event.delta)

    window.bind_all("<MouseWheel>", _on_mousewheel)
    window.bind_all("<Button-4>", lambda e: _scroll_nearest(e.widget, 120))
    window.bind_all("<Button-5>", lambda e: _scroll_nearest(e.widget, -120))

    # Class bindings override ttk's default combobox wheel behavior globally.
    # "break" stops both the value spin and the bind_all handler above, so
    # the guard forwards the scroll to the pane itself.
    def _combo_guard(event):
        _scroll_nearest(event.widget, getattr(event, "delta", 0) or 0)
        return "break"

    window.bind_class("TCombobox", "<MouseWheel>", _combo_guard)
    window.bind_class("TCombobox", "<Button-4>",
                      lambda e: (_scroll_nearest(e.widget, 120), "break")[1])
    window.bind_class("TCombobox", "<Button-5>",
                      lambda e: (_scroll_nearest(e.widget, -120), "break")[1])


def _dpi_scale_factor(win):
    """Return a conservative Windows DPI multiplier for pixel-sized windows.

    Tk scales fonts and many ttk measurements automatically, but geometry
    strings are still pixel based.  Without compensating, dialogs designed at
    100% scaling become too small at 125-175% Windows display scaling and their
    controls are clipped.  Restricting the adjustment to Windows avoids
    surprising sizing changes on macOS/Linux window managers.
    """
    if sys.platform != "win32":
        return 1.0
    try:
        tk_scaling = float(win.tk.call("tk", "scaling"))
    except (tk.TclError, TypeError, ValueError, AttributeError):
        return 1.0
    return max(1.0, min(2.0, tk_scaling / _BASE_TK_SCALING))


def _work_area(win):
    """Return the usable screen rectangle as ``(x, y, width, height)``."""
    if sys.platform == "win32":
        try:
            import ctypes

            class RECT(ctypes.Structure):
                _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                            ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

            rect = RECT()
            # SPI_GETWORKAREA excludes the Windows taskbar.
            if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
                return (rect.left, rect.top,
                        max(1, rect.right - rect.left),
                        max(1, rect.bottom - rect.top))
        except (AttributeError, OSError, ValueError):
            pass
    return 0, 0, win.winfo_screenwidth(), win.winfo_screenheight()


def fit_window(win, want_w, want_h, center=True):
    """Size a window for the current DPI and keep it within the work area.

    ``want_w`` and ``want_h`` are design dimensions at 100% Windows scaling.
    On high-DPI Windows displays they are enlarged in step with Tk's font
    scaling, then clamped to the usable desktop so taskbars and title bars do
    not hide action buttons.
    """
    try:
        scale = _dpi_scale_factor(win)
        left, top, avail_w, avail_h = _work_area(win)
        margin = max(12, round(18 * scale))
        w = min(round(want_w * scale), max(320, avail_w - margin * 2))
        h = min(round(want_h * scale), max(240, avail_h - margin * 2))
        if center:
            x = left + max(margin, (avail_w - w) // 2)
            # Slightly above true center leaves more visual room below modal
            # dialogs without ever crossing the usable work-area boundary.
            y = top + max(margin, min((avail_h - h) // 3, avail_h - h - margin))
            win.geometry(f"{w}x{h}+{x}+{y}")
        else:
            win.geometry(f"{w}x{h}")
    except (tk.TclError, ValueError, AttributeError):
        win.geometry(f"{want_w}x{want_h}")


def configure_ttk_styles(root, theme):
    """Apply a consistent ttk skin without changing widget behavior."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("TCombobox", fieldbackground=theme["input_bg"],
                    background=theme["input_bg"], foreground=theme["input_fg"],
                    bordercolor=theme["border"], lightcolor=theme["border"],
                    darkcolor=theme["border"], padding=5, arrowsize=14)
    style.map("TCombobox",
              fieldbackground=[("readonly", theme["input_bg"])],
              foreground=[("readonly", theme["input_fg"])],
              selectbackground=[("readonly", theme["input_bg"])],
              selectforeground=[("readonly", theme["input_fg"])])
    style.configure("TNotebook", background=theme["bg"], borderwidth=0)
    style.configure("TNotebook.Tab", background=theme["secondary"],
                    foreground=theme["secondary_fg"], padding=(10, 7), borderwidth=0)
    style.map("TNotebook.Tab",
              background=[("selected", theme["card"])],
              foreground=[("selected", theme["accent"])])
    style.configure("Modern.Horizontal.TProgressbar", troughcolor=theme["secondary"],
                    background=theme["accent"], borderwidth=0, thickness=5)
    return style


def modern_button(parent, text, command, theme, *, kind="primary", font_size=10,
                  padx=14, pady=7, width=None):
    """Create a flat button with hover and keyboard-focus affordances."""
    palettes = {
        "primary": (theme["accent"], "#FFFFFF", theme["hover"]),
        "success": (theme["success"], "#FFFFFF", theme["success"]),
        "danger": (theme["danger"], "#FFFFFF", theme["danger"]),
        "secondary": (theme["secondary"], theme["secondary_fg"], theme["highlight"]),
        "brand": (theme["brand"], theme["brand_fg"], theme["hover"]),
    }
    bg, fg, hover = palettes[kind]
    button = tk.Button(parent, text=text, command=command, font=("Segoe UI", font_size, "bold"),
                       bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
                       relief="flat", borderwidth=0, highlightthickness=0,
                       cursor="hand2", padx=padx, pady=pady, width=width)
    button.bind("<Enter>", lambda _e: button.configure(bg=hover))
    button.bind("<Leave>", lambda _e: button.configure(bg=bg))
    return button


def bordered_card(parent, theme, **kwargs):
    """Return a white/navy card with a one-pixel theme border."""
    return tk.Frame(parent, bg=theme["card"], highlightbackground=theme["border"],
                    highlightcolor=theme["focus"], highlightthickness=1, bd=0, **kwargs)


def brand_header(parent, theme, title, subtitle="", step=None, action=None):
    """Build the application header and return its frame.

    The header intentionally sizes to its content rather than using a fixed
    height.  Fixed-height headers clipped the subtitle and step navigation on
    Windows laptops using 125-175% display scaling.
    """
    frame = tk.Frame(parent, bg=theme["brand"])
    frame.pack(fill="x")

    inner = tk.Frame(frame, bg=theme["brand"])
    inner.pack(fill="both", expand=True, padx=26, pady=(14, 13))
    inner.columnconfigure(1, weight=1)

    logo_label = None
    try:
        from PIL import Image, ImageTk
        from config.paths import RESOURCE_DIR
        logo_path = os.path.join(RESOURCE_DIR, "assets", "logo_white.png")
        if os.path.exists(logo_path):
            image = Image.open(logo_path)
            image.thumbnail((165, 48), Image.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            logo_label = tk.Label(inner, image=photo, bg=theme["brand"])
            logo_label.image = photo
            logo_label.grid(row=0, column=0, sticky="w", padx=(0, 18))
    except (ImportError, OSError, tk.TclError):
        logo_label = None
    if logo_label is None:
        tk.Label(inner, text="SPECTRUM REACH", font=("Segoe UI", 13, "bold"),
                 bg=theme["brand"], fg=theme["brand_fg"]).grid(
                     row=0, column=0, sticky="w", padx=(0, 18))

    text_frame = tk.Frame(inner, bg=theme["brand"])
    text_frame.grid(row=0, column=1, sticky="w")
    tk.Label(text_frame, text=title, font=("Segoe UI", 16, "bold"),
             bg=theme["brand"], fg=theme["brand_fg"]).pack(anchor="w")
    if subtitle:
        tk.Label(text_frame, text=subtitle, font=("Segoe UI", 9),
                 bg=theme["brand"], fg=theme["brand_muted"],
                 justify="left").pack(anchor="w", pady=(2, 0))

    right = tk.Frame(inner, bg=theme["brand"])
    right.grid(row=0, column=2, sticky="e", padx=(16, 0))

    if step is not None:
        steps = tk.Frame(right, bg=theme["brand"])
        steps.pack(side="left", padx=(0, 12))
        labels = ("Files", "Clients", "Review")
        for index, label in enumerate(labels, 1):
            active = index == step
            bg = theme["accent"] if active else theme["brand"]
            fg = theme["brand_fg"] if active else theme["brand_muted"]
            tk.Label(steps, text=f"{index}  {label}", font=("Segoe UI", 8, "bold"),
                     bg=bg, fg=fg, padx=8, pady=4).pack(side="left", padx=2)

    if action:
        modern_button(right, action[0], action[1], theme, kind="secondary",
                      font_size=9, padx=11, pady=6).pack(side="left")
    return frame


def run_in_background(root, work, on_success, on_error, poll_ms=80,
                      on_progress=None):
    """Run non-Tk work on a daemon thread and deliver callbacks on Tk's thread.

    With on_progress set, work is called as work(progress) where
    progress(message) may be called freely from the worker thread; each
    message is delivered to on_progress(message) on Tk's thread, in order,
    before the final success/error callback.
    """
    result_queue = queue.Queue(maxsize=1)
    progress_queue = queue.Queue() if on_progress is not None else None

    def runner():
        try:
            if progress_queue is not None:
                result_queue.put((True, work(progress_queue.put)))
            else:
                result_queue.put((True, work()))
        except Exception as exc:  # delivered to the UI callback with traceback intact in logs
            result_queue.put((False, exc))

    threading.Thread(target=runner, daemon=True).start()

    def poll():
        if progress_queue is not None:
            while True:
                try:
                    on_progress(progress_queue.get_nowait())
                except queue.Empty:
                    break
        try:
            ok, payload = result_queue.get_nowait()
        except queue.Empty:
            try:
                root.after(poll_ms, poll)
            except tk.TclError:
                pass
            return
        if ok:
            on_success(payload)
        else:
            on_error(payload)

    root.after(poll_ms, poll)
