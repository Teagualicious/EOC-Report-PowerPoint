"""Excel-style number-format popup for sidebar metrics (right-click)."""

import tkinter as tk
from ui.utils import fit_window


def show_format_popup(wizard, event, metric_key, date_only=False):
    """Open the format popup; writes results into the wizard's format dicts.

    date_only: compact variant for date quick-fills (Date Range, Start/End
    Date) — only the date style chooser, no number formatting noise.
    """
    t = wizard.t
    win = tk.Toplevel(wizard.window)
    win.title("Format Date" if date_only else "Format Number")
    fit_window(win, 360, 300 if date_only else 540)
    win.configure(bg=t["bg"]); win.transient(wizard.window)
    win.grab_set()
    win.attributes('-topmost', True)
    win.lift(); win.focus_force()

    # Center on screen
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - 350) // 2
    h = 300 if date_only else 540
    y = (sh - h) // 2
    win.geometry(f"360x{h}+{x}+{y}")

    tk.Label(win, text="Format Date" if date_only else "Format Cells",
             font=("Calibri", 13, "bold"),
             bg=t["bg"], fg=t["fg"]).pack(pady=(10, 5))

    # Category list
    categories = tk.Frame(win, bg=t["bg"])
    categories.pack(fill="both", expand=True, padx=15, pady=5)

    current_fmt = wizard._metric_formats.get(metric_key,
                                              "date" if date_only else "text")
    fmt_var = tk.StringVar(value="date" if date_only else current_fmt)
    decimals_var = tk.IntVar(value=0)
    commas_var = tk.BooleanVar(value=True)
    prefix_var = tk.StringVar(value="")
    suffix_var = tk.StringVar(value="")

    # Format options
    options = [
        ("General (text)", "text"),
        ("Number (1,234)", "number"),
        ("Currency ($1,234)", "currency"),
        ("Percentage (85.0%)", "percentage"),
        ("Date", "date"),
        ("Custom", "custom"),
    ]
    if not date_only:
        for label, val_key in options:
            tk.Radiobutton(categories, text=label, variable=fmt_var, value=val_key,
                           font=("Calibri", 10), bg=t["bg"], fg=t["fg"],
                           selectcolor=t["input_bg"], anchor="w"
                           ).pack(fill="x", pady=1)

    # ── Date style (Excel-style date formatting) ──
    from engine.pptx_formats import DATE_STYLE_CHOICES
    existing_details = getattr(wizard, '_metric_format_details', {}).get(metric_key, {})
    date_style_var = tk.StringVar(value=existing_details.get("date_style", "long_ordinal"))
    custom_strftime_var = tk.StringVar(value=existing_details.get("custom_strftime", "%m/%d/%Y"))

    date_frame = tk.LabelFrame(win, text="  Date style  ", font=("Calibri", 8, "bold"),
                                bg=t["bg"], fg=t["muted"])
    date_frame.pack(fill="x", padx=15, pady=3)
    style_labels = {key: f"{example}" for key, example in DATE_STYLE_CHOICES}
    label_to_key = {v: k for k, v in style_labels.items()}
    from tkinter import ttk as _ttk
    date_combo = _ttk.Combobox(date_frame, state="readonly", font=("Calibri", 9),
                                values=list(style_labels.values()), width=24)
    date_combo.set(style_labels.get(date_style_var.get(), style_labels["long_ordinal"]))
    date_combo.pack(side="left", padx=5, pady=4)
    date_combo.bind("<<ComboboxSelected>>",
                    lambda e: date_style_var.set(label_to_key[date_combo.get()]))
    strf_row = tk.Frame(date_frame, bg=t["bg"])
    strf_row.pack(fill="x", padx=5, pady=(0, 4))
    tk.Label(strf_row, text="strftime:", font=("Calibri", 8),
             bg=t["bg"], fg=t["muted"]).pack(side="left")
    tk.Entry(strf_row, textvariable=custom_strftime_var, width=16,
             font=("Calibri", 9), bg=t["input_bg"], fg=t["input_fg"]
             ).pack(side="left", padx=4)
    tk.Label(strf_row, text="(used with 'Custom (strftime)')", font=("Calibri", 7),
             bg=t["bg"], fg=t["muted"]).pack(side="left")

    # Decimal places (hidden in date-only mode)
    dec_frame = tk.Frame(win, bg=t["bg"])
    if not date_only:
        dec_frame.pack(fill="x", padx=15, pady=3)
    tk.Label(dec_frame, text="Decimal places:", font=("Calibri", 9),
             bg=t["bg"], fg=t["fg"]).pack(side="left")
    tk.Spinbox(dec_frame, from_=0, to=6, textvariable=decimals_var, width=5,
               font=("Calibri", 9)).pack(side="left", padx=5)
    tk.Checkbutton(dec_frame, text="Use commas", variable=commas_var,
                   font=("Calibri", 9), bg=t["bg"], fg=t["fg"],
                   selectcolor=t["input_bg"]).pack(side="left", padx=10)

    # Custom prefix/suffix (hidden in date-only mode)
    custom_frame = tk.Frame(win, bg=t["bg"])
    if not date_only:
        custom_frame.pack(fill="x", padx=15, pady=3)
    tk.Label(custom_frame, text="Prefix:", font=("Calibri", 9),
             bg=t["bg"], fg=t["fg"]).pack(side="left")
    tk.Entry(custom_frame, textvariable=prefix_var, width=5, font=("Calibri", 9),
             bg=t["input_bg"], fg=t["input_fg"]).pack(side="left", padx=3)
    tk.Label(custom_frame, text="Suffix:", font=("Calibri", 9),
             bg=t["bg"], fg=t["fg"]).pack(side="left", padx=(10, 0))
    tk.Entry(custom_frame, textvariable=suffix_var, width=5, font=("Calibri", 9),
             bg=t["input_bg"], fg=t["input_fg"]).pack(side="left", padx=3)

    # Preview
    preview_label = tk.Label(win, text="", font=("Calibri", 14, "bold"),
                              bg=t["bg"], fg=t["fg"])
    preview_label.pack(pady=5)

    def _collect_details():
        try:
            dec = decimals_var.get()
        except tk.TclError:
            dec = 0  # Spinbox mid-edit (empty/non-numeric)
        return {
            "format": fmt_var.get(),
            "decimals": dec,
            "commas": commas_var.get(),
            "prefix": prefix_var.get(),
            "suffix": suffix_var.get(),
            "date_style": date_style_var.get(),
            "custom_strftime": custom_strftime_var.get(),
        }

    def update_preview(*args):
        from engine.pptx_formats import _coerce_number, format_with_details
        raw = _coerce_number(wizard.available_metrics.get(metric_key))
        if raw is None:
            preview_label.config(text="Preview: (no value yet)")
            return
        details = _collect_details()
        if details["format"] == "date" or (
                isinstance(raw, (int, float)) and not isinstance(raw, bool)):
            preview_label.config(text=f"Preview: {format_with_details(raw, details)}")
        else:
            preview_label.config(text=f"Preview: {raw}")

    for var in [fmt_var, decimals_var, commas_var, prefix_var, suffix_var,
                date_style_var, custom_strftime_var]:
        var.trace_add("write", update_preview)
    update_preview()

    def apply():
        wizard._metric_formats[metric_key] = fmt_var.get()
        # Store custom format details
        wizard._metric_format_details = getattr(wizard, '_metric_format_details', {})
        wizard._metric_format_details[metric_key] = _collect_details()
        win.destroy()
        wizard._refresh_metrics()
        # Push the new format into existing assignments and re-render the
        # live slide — re-formatting never requires re-assigning
        if hasattr(wizard, "_propagate_format_details"):
            wizard._propagate_format_details(metric_key)

    tk.Button(win, text="OK", font=("Calibri", 10, "bold"), bg=t["accent"],
              fg="white", relief="flat", padx=20, pady=5,
              command=apply).pack(pady=(5, 10))
