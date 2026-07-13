# ── Portable dependencies ──
# If a ./vendor folder exists beside the app, packages inside it are used
# (populated via make_portable.bat); otherwise the machine's installed
# packages are used. Additive: with no vendor folder this is a no-op.
import os as _os
import sys as _sys
_VENDOR = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "vendor")
if _os.path.isdir(_VENDOR):
    _sys.path.insert(0, _VENDOR)
    for _sub in ("win32", _os.path.join("win32", "lib"), "Pythonwin"):
        _p = _os.path.join(_VENDOR, _sub)
        if _os.path.isdir(_p):
            _sys.path.insert(1, _p)
    _dll = _os.path.join(_VENDOR, "pywin32_system32")
    if _os.path.isdir(_dll):
        try:
            _os.add_dll_directory(_dll)
            _os.environ["PATH"] = _dll + _os.pathsep + _os.environ.get("PATH", "")
        except Exception:
            pass

"""Entry point for the Spectrum Reach Reporting Ingestion Engine."""

import logging
import sys
from tkinter import messagebox

from config.logging_setup import setup_logging
from config.paths import ensure_dirs


def _handle_exception(exc_type, exc_value, exc_traceback):
    """Last-resort handler: log the crash and tell the user where to look."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.getLogger("crash").error(
        "Unhandled exception", exc_info=(exc_type, exc_value, exc_traceback))
    try:
        messagebox.showerror(
            "Unexpected Error",
            "Something went wrong. Details were written to the logs folder.\n\n"
            f"{exc_type.__name__}: {exc_value}")
    except Exception:
        pass  # No display available (e.g. crash before Tk started)


def _enable_dpi_awareness():
    """Ask Windows to render Tk at native DPI before the root is created."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def main():
    _enable_dpi_awareness()
    setup_logging()
    ensure_dirs()
    logger = logging.getLogger(__name__)
    logger.info("Application starting (Python %s)", sys.version.split()[0])

    sys.excepthook = _handle_exception

    from ui.main_window import IngestionEngine
    app = IngestionEngine()

    # Route tkinter callback exceptions (button handlers etc.) through the
    # same log-and-notify path — one hook on root covers all Toplevels.
    def _tk_exception(exc_type, exc_value, exc_traceback):
        _handle_exception(exc_type, exc_value, exc_traceback)
    app.root.report_callback_exception = _tk_exception

    app.run()
    logger.info("Application closed")


if __name__ == "__main__":
    main()
