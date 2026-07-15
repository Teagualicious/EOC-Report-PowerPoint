"""Client assignment wizard: assign campaigns to clients before export."""

import logging
import tkinter as tk
from ui.utils import brand_header, bordered_card, fit_window, modern_button
from tkinter import ttk, messagebox

from config.paths import OUTPUT_DIR

logger = logging.getLogger(__name__)


class ClientWizard:
    """Loops a campaign-assignment step per client, then calls on_complete.

    ``on_complete(client_assignments, start_date, end_date, output_folder,
)`` receives {client_name: [campaign, ...]}.
    """

    def __init__(self, parent, all_parsed, platforms, theme, on_complete,
                 start_date="", end_date="", output_folder=""):
        self.parent = parent
        self.all_parsed = all_parsed
        self.platforms = platforms
        self.t = theme
        self.on_complete = on_complete
        self.client_assignments = {}
        self.start_date = start_date
        self.end_date = end_date
        self.output_folder = output_folder if output_folder else OUTPUT_DIR
        self.window = None
        # Export options from settings

        # Build campaign list with platform info
        self.all_campaigns = []  # list of {"name": str, "platform": str}
        seen = set()
        for data in all_parsed:
            plat = data.get("source_platform", "Unknown")
            for mdata in data.get("campaign_metrics", {}).values():
                cn = mdata.get("campaign_name", "")
                if cn and cn not in seen:
                    self.all_campaigns.append({"name": cn, "platform": plat})
                    seen.add(cn)
            for ld in data.get("level_data", []):
                cn = ld.get("_campaign", "")
                if cn and cn not in seen:
                    self.all_campaigns.append({"name": cn, "platform": plat})
                    seen.add(cn)

        # Sort by platform then name
        self.all_campaigns.sort(key=lambda c: (c["platform"], c["name"]))
        self.unassigned = list(self.all_campaigns)

        self._show_client_step()

    def _close_window(self):
        if self.window:
            self.window.destroy()
            self.window = None

    def _show_client_step(self):
        if not self.unassigned:
            self._finish(); return

        self._close_window()
        t = self.t
        self.window = tk.Toplevel(self.parent)
        assigned_count = len(self.all_campaigns) - len(self.unassigned)
        self.window.title(f"Quick Run — Assign Client ({assigned_count}/{len(self.all_campaigns)} assigned)")
        # No state("zoomed") here: zooming a TRANSIENT toplevel on Windows
        # maximizes over the taskbar, hiding the bottom Next/Back bar on
        # smaller screens. fit_window clamps to the usable work area
        # (excludes the taskbar), so every control stays visible; the large
        # design size fills most of the screen on any display.
        fit_window(self.window, 1280, 940)
        self.window.configure(bg=t["bg"])
        self.window.transient(self.parent)
        self.window.grab_set()
        self.window.lift()
        self.window.focus_force()

        brand_header(
            self.window, t, "Assign campaigns to clients",
            f"{len(self.unassigned)} campaign(s) remain unassigned", step=2)

        # Top card: client name + search
        top = bordered_card(self.window, t)
        top.pack(fill="x", padx=24, pady=(18, 8))
        top.columnconfigure(1, weight=1)

        tk.Label(top, text="Client name", font=("Segoe UI", 11, "bold"),
                 bg=t["card"], fg=t["fg"]).grid(
                     row=0, column=0, sticky="w", padx=(16, 10), pady=(14, 7))
        self.client_name_var = tk.StringVar()
        tk.Entry(top, textvariable=self.client_name_var, width=30, font=("Segoe UI", 12),
                 bg=t["input_bg"], fg=t["input_fg"], insertbackground=t["insert"],
                 relief="solid", borderwidth=1).grid(
                     row=0, column=1, sticky="ew", padx=(0, 16), pady=(14, 7))

        # Search bar
        search_frame = tk.Frame(top, bg=t["card"])
        search_frame.grid(row=1, column=0, columnspan=2, sticky="ew",
                          padx=16, pady=(0, 14))
        tk.Label(search_frame, text="🔍", font=("Segoe UI", 12), bg=t["card"],
                 fg=t["muted"]).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._filter_list())
        tk.Entry(search_frame, textvariable=self.search_var, width=40, font=("Segoe UI", 10),
                 bg=t["input_bg"], fg=t["input_fg"], insertbackground=t["insert"],
                 relief="solid", borderwidth=1).pack(side="left", padx=5, fill="x", expand=True)

        # Info + buttons
        info_frame = tk.Frame(self.window, bg=t["bg"])
        info_frame.pack(fill="x", padx=20, pady=(0, 5))
        self.info_label = tk.Label(info_frame,
            text=f"{len(self.unassigned)} campaigns remaining",
            font=("Segoe UI", 10), bg=t["bg"], fg=t["muted"])
        self.info_label.pack(side="left")
        tk.Button(info_frame, text="Select All", font=("Segoe UI", 9), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=8,
                  command=self._select_all).pack(side="right", padx=(5, 0))
        tk.Button(info_frame, text="Select None", font=("Segoe UI", 9), bg=t["secondary"],
                  fg=t["secondary_fg"], relief="flat", padx=8,
                  command=self._select_none).pack(side="right")

        # Bottom actions must be reserved before the expandable campaign area.
        # Packing the full-height canvas first caused the Next button to float
        # beside the list on Windows instead of remaining at the bottom.
        bottom = tk.Frame(self.window, bg=t["bg"])
        bottom.pack(side="bottom", fill="x", padx=24, pady=(8, 14))
        if self.client_assignments:  # Show back only if we've assigned at least one client
            modern_button(bottom, "← Back", self._go_back, t, kind="secondary",
                          padx=15, pady=8).pack(side="left")
        modern_button(bottom, "Next  →", self._assign_client, t, kind="primary",
                      font_size=12, padx=25, pady=9).pack(side="right")

        # Campaign list — grouped by platform
        list_shell = bordered_card(self.window, t)
        list_shell.pack(fill="both", expand=True, padx=24, pady=(2, 0))
        list_shell.columnconfigure(0, weight=1)
        list_shell.rowconfigure(0, weight=1)

        self.list_canvas = tk.Canvas(list_shell, bg=t["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_shell, orient="vertical", command=self.list_canvas.yview)
        self.list_frame = tk.Frame(self.list_canvas, bg=t["bg"])
        self.list_frame.bind("<Configure>",
            lambda e: self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all")))
        self._list_window = self.list_canvas.create_window(
            (0, 0), window=self.list_frame, anchor="nw")
        self.list_canvas.bind(
            "<Configure>",
            lambda e: self.list_canvas.itemconfigure(self._list_window, width=e.width))
        self.list_canvas.configure(yscrollcommand=scrollbar.set)
        self.list_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Build campaign checkboxes
        self.campaign_vars = {}
        self.campaign_widgets = {}  # for filtering
        self.selection_state = {c["name"]: False for c in self.unassigned}
        self._build_campaign_list()

    def _build_campaign_list(self):
        t = self.t
        for w in self.list_frame.winfo_children(): w.destroy()
        self.campaign_vars = {}
        self.campaign_widgets = {}

        # Group by platform
        platforms = {}
        for c in self.unassigned:
            plat = c["platform"]
            if plat not in platforms: platforms[plat] = []
            platforms[plat].append(c)

        search = self.search_var.get().lower() if hasattr(self, 'search_var') else ""

        for plat, campaigns in platforms.items():
            # Filter by search
            filtered = [c for c in campaigns if search in c["name"].lower()] if search else campaigns
            if not filtered: continue

            # Platform header
            phdr = tk.Frame(self.list_frame, bg=t["accent"])
            phdr.pack(fill="x", padx=5, pady=(8, 2))
            tk.Label(phdr, text=f"  {plat}  ({len(filtered)} campaigns)",
                     font=("Segoe UI", 10, "bold"), bg=t["accent"], fg="white",
                     anchor="w").pack(fill="x", padx=5, pady=3)

            for c in filtered:
                name = c["name"]
                var = tk.BooleanVar(value=self.selection_state.get(name, False))
                var.trace_add(
                    "write",
                    lambda *_args, n=name, v=var: self.selection_state.__setitem__(n, v.get()))
                self.campaign_vars[name] = var

                row = tk.Frame(self.list_frame, bg=t["card"], cursor="hand2")
                row.pack(fill="x", padx=10, pady=1)

                cb = tk.Checkbutton(row, variable=var, bg=t["card"],
                                    selectcolor=t["input_bg"])
                cb.pack(side="left", padx=(8, 5))

                lbl = tk.Label(row, text=name, font=("Segoe UI", 10), bg=t["card"],
                               fg=t["card_fg"], anchor="w", cursor="hand2")
                lbl.pack(side="left", padx=5, fill="x", expand=True)

                self.campaign_widgets[name] = {"frame": row, "var": var, "cb": cb, "label": lbl}

                # Click anywhere on row to toggle. (Drag-select was removed
                # as dead code — Tk's implicit grab during a button press
                # eats <Enter> events, so the drag handlers never fired;
                # roadmap Phase 5 decision.)
                def toggle(event, v=var):
                    v.set(not v.get())
                row.bind("<Button-1>", toggle)
                lbl.bind("<Button-1>", toggle)

    def _filter_list(self):
        """Debounced: rebuilding hundreds of checkbox rows on every keystroke
        made searching feel sluggish with large imports. One rebuild fires
        250ms after typing pauses."""
        pending = getattr(self, "_filter_after_id", None)
        if pending:
            try:
                self.window.after_cancel(pending)
            except tk.TclError:
                pass
        self._filter_after_id = self.window.after(250, self._build_campaign_list)

    def _select_all(self):
        for var in self.campaign_vars.values(): var.set(True)

    def _select_none(self):
        for var in self.campaign_vars.values(): var.set(False)

    def _go_back(self):
        """Undo last client assignment and go back."""
        if self.client_assignments:
            last_client = list(self.client_assignments.keys())[-1]
            last_campaigns = self.client_assignments.pop(last_client)
            # Add campaigns back to unassigned
            for cn in last_campaigns:
                # Find the campaign info
                match = next((c for c in self.all_campaigns if c["name"] == cn), None)
                if match and match not in self.unassigned:
                    self.unassigned.append(match)
            self.unassigned.sort(key=lambda c: (c["platform"], c["name"]))
        self._close_window()
        self._show_client_step()

    def _assign_client(self):
        client_name = self.client_name_var.get().strip()
        selected = [cn for cn, is_selected in self.selection_state.items()
                    if is_selected]
        if not client_name:
            messagebox.showwarning("No Name", "Enter a client name.",
                                   parent=self.window)
            return
        if not selected:
            messagebox.showwarning("None Selected", "Check at least one campaign.",
                                   parent=self.window)
            return
        if any(existing.casefold() == client_name.casefold()
               for existing in self.client_assignments):
            messagebox.showwarning(
                "Client Already Assigned",
                "That client name has already been used. Go back to edit the earlier "
                "assignment or use a distinct client name.",
                parent=self.window)
            return

        self.client_assignments[client_name] = selected
        self.unassigned = [c for c in self.unassigned if c["name"] not in selected]
        self._close_window()
        self._show_client_step()

    def _finish(self):
        self._close_window()
        logger.info("Client assignment complete: %d client(s)",
                    len(self.client_assignments))
        self.on_complete(self.client_assignments, self.start_date,
                         self.end_date, self.output_folder)
