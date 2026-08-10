"""Developer-only launcher for the inherited PowerPoint template mapper."""

from __future__ import annotations

import os
import sys
import tkinter as tk

_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)

from config.paths import ensure_dirs
from config.themes import get_theme
from mapper.mapper_window import PPTXWizard


def main() -> None:
    ensure_dirs()
    root = tk.Tk()
    root.withdraw()
    wizard = PPTXWizard(
        root,
        {"client_name": "", "client_data": [], "folder": ""},
        get_theme("light"),
    )
    if getattr(wizard, "template_path", None):
        root.mainloop()
    else:
        root.destroy()


if __name__ == "__main__":
    main()
