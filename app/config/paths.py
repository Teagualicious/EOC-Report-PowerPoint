"""Central path constants for the application.

The cleaned project layout separates program files from user-generated data:

* ``APP_DIR`` - Python code and bundled resources under ``app/``.
* ``PROJECT_ROOT`` - the folder containing the launcher and README.
* ``RESOURCE_DIR`` - read-only files bundled with the application.
* ``WORKSPACE_DIR`` - writable settings, mappings, templates, and logs.
* ``INPUT_DIR`` / ``OUTPUT_DIR`` - user-facing folders at the project root.

Under PyInstaller, bundled resources are read from ``sys._MEIPASS`` while the
workspace remains beside the executable. ``APP_ROOT`` is retained as a
backward-compatible alias for ``PROJECT_ROOT`` because older template mappings
may contain paths relative to the project folder.
"""

from __future__ import annotations

import logging
import os
import shutil
import sys

logger = logging.getLogger(__name__)

if getattr(sys, "frozen", False):  # PyInstaller bundle
    PROJECT_ROOT = os.path.dirname(sys.executable)
    APP_DIR = PROJECT_ROOT
    RESOURCE_DIR = os.path.join(
        getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)), "resources")
else:
    APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    PROJECT_ROOT = os.path.dirname(APP_DIR)
    RESOURCE_DIR = os.path.join(APP_DIR, "resources")

# Compatibility alias used by existing mapping/image resolution code.
APP_ROOT = PROJECT_ROOT

WORKSPACE_DIR = os.path.join(PROJECT_ROOT, "workspace")

# Read-only resources
DICT_PATH = os.path.join(RESOURCE_DIR, "metric_dictionary.json")
LOGO_PATH = os.path.join(RESOURCE_DIR, "assets", "logo.png")

# Writable application state
SETTINGS_PATH = os.path.join(WORKSPACE_DIR, "settings.json")
MAPPINGS_DIR = os.path.join(WORKSPACE_DIR, "mappings")
TEMPLATES_DIR = os.path.join(WORKSPACE_DIR, "templates")
IMAGES_DIR = os.path.join(TEMPLATES_DIR, "images")
LOGS_DIR = os.path.join(WORKSPACE_DIR, "logs")

# User-facing source and report folders live at the project root for easier
# navigation in File Explorer.
INPUT_DIR = os.path.join(PROJECT_ROOT, "input")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")


def _copy_missing_tree(source: str, destination: str) -> int:
    """Copy files from a legacy directory without replacing newer files."""
    if not os.path.isdir(source):
        return 0
    copied = 0
    for root, _dirs, files in os.walk(source):
        rel = os.path.relpath(root, source)
        target_root = destination if rel == "." else os.path.join(destination, rel)
        os.makedirs(target_root, exist_ok=True)
        for filename in files:
            src = os.path.join(root, filename)
            dst = os.path.join(target_root, filename)
            if not os.path.exists(dst):
                shutil.copy2(src, dst)
                copied += 1
    return copied


def migrate_legacy_layout() -> None:
    """Copy data from older layouts into the current folder structure.

    The migration is deliberately non-destructive: old files are left in place
    until the user has verified the new build. Existing destination files
    always win, so running this function repeatedly is safe.
    """
    legacy_pairs = (
        # Original root-level application state.
        (os.path.join(PROJECT_ROOT, "mappings"), MAPPINGS_DIR),
        (os.path.join(PROJECT_ROOT, "templates"), TEMPLATES_DIR),
        (os.path.join(PROJECT_ROOT, "logs"), LOGS_DIR),
        # Original source staging name and the previous cleaned layout.
        (os.path.join(PROJECT_ROOT, "input_files"), INPUT_DIR),
        (os.path.join(WORKSPACE_DIR, "input"), INPUT_DIR),
        (os.path.join(WORKSPACE_DIR, "output"), OUTPUT_DIR),
    )
    total = 0
    for source, destination in legacy_pairs:
        # Do not mistake the new workspace directory for a legacy source.
        if os.path.abspath(source) == os.path.abspath(destination):
            continue
        try:
            total += _copy_missing_tree(source, destination)
        except OSError:
            logger.warning("Could not migrate legacy folder: %s", source,
                           exc_info=True)

    legacy_settings = os.path.join(PROJECT_ROOT, "settings.json")
    if os.path.isfile(legacy_settings) and not os.path.exists(SETTINGS_PATH):
        try:
            shutil.copy2(legacy_settings, SETTINGS_PATH)
            total += 1
        except OSError:
            logger.warning("Could not migrate legacy settings.json", exc_info=True)

    if total:
        logger.info("Copied %d legacy data file(s) into the current layout", total)


def ensure_dirs() -> None:
    """Create writable directories and import legacy data if present."""
    for path in (
        WORKSPACE_DIR,
        MAPPINGS_DIR,
        TEMPLATES_DIR,
        IMAGES_DIR,
        INPUT_DIR,
        OUTPUT_DIR,
        LOGS_DIR,
    ):
        os.makedirs(path, exist_ok=True)
    migrate_legacy_layout()
