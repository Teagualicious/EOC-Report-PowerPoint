"""Central, project-anchored paths for Deck Engine.

All application-managed paths derive from :data:`PROJECT_ROOT`; business code
must never depend on the shell's current working directory.  Under PyInstaller,
read-only resources come from ``sys._MEIPASS`` while the writable workspace
remains beside the executable.
"""

from __future__ import annotations

import logging
import os
import shutil
import sys

logger = logging.getLogger(__name__)

if getattr(sys, "frozen", False):
    PROJECT_ROOT = os.path.dirname(sys.executable)
    APP_DIR = PROJECT_ROOT
    RESOURCE_DIR = os.path.join(
        getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)), "resources")
else:
    APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    PROJECT_ROOT = os.path.dirname(APP_DIR)
    RESOURCE_DIR = os.path.join(APP_DIR, "resources")

# Backward-compatible alias used by inherited mapping/image resolution code.
APP_ROOT = PROJECT_ROOT
WORKSPACE_DIR = os.path.join(PROJECT_ROOT, "workspace")

# Read-only resources.
DICT_PATH = os.path.join(RESOURCE_DIR, "metric_dictionary.json")
LOGO_PATH = os.path.join(RESOURCE_DIR, "assets", "logo.png")

# Writable application state.
SETTINGS_PATH = os.path.join(WORKSPACE_DIR, "settings.json")
MAPPINGS_DIR = os.path.join(WORKSPACE_DIR, "mappings")
TEMPLATES_DIR = os.path.join(WORKSPACE_DIR, "templates")
IMAGES_DIR = os.path.join(TEMPLATES_DIR, "images")
STAGING_DIR = os.path.join(WORKSPACE_DIR, "staging")
DICTIONARY_DIR = os.path.join(WORKSPACE_DIR, "dictionary")
LOGS_DIR = os.path.join(WORKSPACE_DIR, "logs")

# User-facing source and report folders.
INPUT_DIR = os.path.join(PROJECT_ROOT, "input")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")
QUARANTINE_DIR = os.path.join(OUTPUT_DIR, "_quarantine")


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
    """Copy inherited Jughead state into the fork layout non-destructively."""
    legacy_pairs = (
        (os.path.join(PROJECT_ROOT, "mappings"), MAPPINGS_DIR),
        (os.path.join(PROJECT_ROOT, "templates"), TEMPLATES_DIR),
        (os.path.join(PROJECT_ROOT, "logs"), LOGS_DIR),
        (os.path.join(PROJECT_ROOT, "input_files"), INPUT_DIR),
        (os.path.join(WORKSPACE_DIR, "input"), INPUT_DIR),
        (os.path.join(WORKSPACE_DIR, "output"), OUTPUT_DIR),
    )
    total = 0
    for source, destination in legacy_pairs:
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
        logger.info("Copied %d legacy data file(s) into the Deck Engine layout", total)


def ensure_dirs() -> None:
    """Create every writable directory and import legacy data if present."""
    for path in (
        WORKSPACE_DIR,
        MAPPINGS_DIR,
        TEMPLATES_DIR,
        IMAGES_DIR,
        STAGING_DIR,
        DICTIONARY_DIR,
        INPUT_DIR,
        OUTPUT_DIR,
        QUARANTINE_DIR,
        LOGS_DIR,
    ):
        os.makedirs(path, exist_ok=True)
    migrate_legacy_layout()
