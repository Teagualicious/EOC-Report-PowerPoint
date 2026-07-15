"""Settings and platform-config persistence.

App-wide settings live in ``workspace/settings.json``; per-platform
column-role configs are stored as ``workspace/mappings/platform_{name}.json``.
"""

import json
import logging
import os
import tempfile

from config.paths import MAPPINGS_DIR, OUTPUT_DIR, SETTINGS_PATH, WORKSPACE_DIR
from config.naming import storage_key

logger = logging.getLogger(__name__)

# Metrics pre-unchecked in platform setup (partial completion quartiles)
DEFAULT_UNCHECKED = ["25% Completions", "50% Completions", "75% Completions",
                     "25% Complete", "50% Complete", "75% Complete"]


def _normalize_settings(data):
    """Translate settings that referenced defaults from older layouts.

    A stored output_folder that (a) was the old workspace/output default or
    (b) is a default-shaped path (…/output) that no longer exists — the
    signature of a project folder that was renamed or moved — falls back to
    the CURRENT project's output folder. Otherwise exports would silently
    recreate the old folder name and write there.
    """
    if not isinstance(data, dict):
        return {"platforms": {}, "theme": "light"}

    output_folder = data.get("output_folder")
    previous_default = os.path.join(WORKSPACE_DIR, "output")
    if output_folder:
        try:
            stored = os.path.abspath(output_folder)
            stale = (
                os.path.normcase(stored)
                == os.path.normcase(os.path.abspath(previous_default))
            ) or (
                os.path.basename(os.path.normpath(stored)).lower() == "output"
                and not os.path.isdir(stored)
            )
        except (OSError, TypeError, ValueError):
            stale = False
        if stale and os.path.normcase(stored) != os.path.normcase(
                os.path.abspath(OUTPUT_DIR)):
            logger.info("Stored output folder %s is stale — using the "
                        "current default %s", output_folder, OUTPUT_DIR)
            data = dict(data)
            data["output_folder"] = OUTPUT_DIR
    return data


def load_settings():
    """Load settings.json, returning defaults if missing or unreadable."""
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                return _normalize_settings(json.load(f))
        except (OSError, json.JSONDecodeError):
            logger.warning("settings.json unreadable — using defaults", exc_info=True)
    return {"platforms": {}, "theme": "light"}


def _atomic_json_write(path, data):
    """Write JSON atomically so a crash cannot leave a half-written config."""
    directory = os.path.dirname(path) or "."
    os.makedirs(directory, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", suffix=".json", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


def save_settings(data):
    """Persist the settings dict to settings.json."""
    _atomic_json_write(SETTINGS_PATH, data)


def _platform_config_path(platform_name):
    return os.path.join(MAPPINGS_DIR,
                        f"platform_{storage_key(platform_name)}.json")


def load_platform_config(platform_name):
    """Load a platform's column-role config, or None if not configured."""
    path = _platform_config_path(platform_name)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (OSError, json.JSONDecodeError):
            logger.warning("Platform config unreadable: %s", path, exc_info=True)
    return None


def save_platform_config(platform_name, config):
    """Persist a platform's column-role config."""
    os.makedirs(MAPPINGS_DIR, exist_ok=True)
    _atomic_json_write(_platform_config_path(platform_name), config)


def delete_platform_config(platform_name):
    """Remove a platform's stored config file if it exists."""
    path = _platform_config_path(platform_name)
    if os.path.exists(path):
        os.remove(path)
