"""Atomic settings and import-profile persistence for Deck Engine."""

from __future__ import annotations

import json
import logging
import os
import shutil
import tempfile
from typing import Any

from config.naming import storage_key
from config.paths import (IMAGES_DIR, MAPPINGS_DIR, OUTPUT_DIR, SETTINGS_PATH,
                          WORKSPACE_DIR)

logger = logging.getLogger(__name__)

DEFAULT_SETTINGS = {
    "schema_version": 1,
    "active_template": "",
    "output_folder": "",
    "images_root": "",
    "active_profile": "",
    "validation": {"enabled": True},
    # Kept only so an inherited platform config is not destroyed during the
    # fork. Stage 1 import profiles replace this in the analyst workflow.
    "platforms": {},
    "theme": "light",
}

# Metrics pre-unchecked in the inherited profile editor. Stage 1 reduces that
# editor but the constant remains part of its compatibility surface.
DEFAULT_UNCHECKED = ["25% Completions", "50% Completions", "75% Completions",
                     "25% Complete", "50% Complete", "75% Complete"]


def _defaults() -> dict[str, Any]:
    return {
        **DEFAULT_SETTINGS,
        "validation": dict(DEFAULT_SETTINGS["validation"]),
        "platforms": {},
    }


def _normalize_settings(data: Any) -> dict[str, Any]:
    """Merge settings with defaults and repair inherited default paths."""
    normalized = _defaults()
    if isinstance(data, dict):
        normalized.update(data)
        if isinstance(data.get("validation"), dict):
            normalized["validation"] = {
                **DEFAULT_SETTINGS["validation"], **data["validation"]}
        if not isinstance(normalized.get("platforms"), dict):
            normalized["platforms"] = {}

    # Dark mode was intentionally removed from the fork's analyst surface.
    normalized["theme"] = "light"

    output_folder = normalized.get("output_folder")
    previous_default = os.path.join(WORKSPACE_DIR, "output")
    if output_folder:
        try:
            stored = os.path.abspath(os.fspath(output_folder))
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
            logger.info("Stored output folder %s is stale — using %s",
                        output_folder, OUTPUT_DIR)
            normalized["output_folder"] = OUTPUT_DIR

    # Empty values deliberately mean "follow the project default" so moving
    # the folder cannot pin a stale absolute path.
    if normalized.get("images_root") == IMAGES_DIR:
        normalized["images_root"] = ""
    return normalized


def _backup_corrupt(path: str) -> str | None:
    """Move an unreadable JSON file aside so it never blocks startup."""
    if not os.path.isfile(path):
        return None
    backup = f"{path}.corrupt"
    suffix = 2
    while os.path.exists(backup):
        backup = f"{path}.corrupt-{suffix}"
        suffix += 1
    try:
        shutil.move(path, backup)
        return backup
    except OSError:
        logger.warning("Could not back up corrupt settings file %s", path,
                       exc_info=True)
        return None


def load_settings() -> dict[str, Any]:
    """Load settings, backing up corrupt JSON and returning safe defaults."""
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as handle:
                return _normalize_settings(json.load(handle))
        except (OSError, json.JSONDecodeError):
            backup = _backup_corrupt(SETTINGS_PATH)
            logger.warning("settings.json unreadable — using defaults%s",
                           f"; backed up to {backup}" if backup else "",
                           exc_info=True)
    return _defaults()


def _atomic_json_write(path: str, data: Any) -> None:
    """Write JSON atomically so a crash cannot leave half-written state."""
    directory = os.path.dirname(path) or os.curdir
    os.makedirs(directory, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_", suffix=".json", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(data, handle, indent=2, ensure_ascii=False, sort_keys=True)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


def save_settings(data: dict[str, Any]) -> dict[str, Any]:
    """Normalize, persist, and return the settings that were written."""
    normalized = _normalize_settings(data)
    _atomic_json_write(SETTINGS_PATH, normalized)
    return normalized


def _platform_config_path(platform_name: str) -> str:
    """Inherited config path retained for migration and developer tooling."""
    return os.path.join(MAPPINGS_DIR,
                        f"platform_{storage_key(platform_name)}.json")


def load_platform_config(platform_name: str) -> dict[str, Any] | None:
    path = _platform_config_path(platform_name)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                return json.load(handle)
        except (OSError, json.JSONDecodeError):
            logger.warning("Platform config unreadable: %s", path, exc_info=True)
    return None


def save_platform_config(platform_name: str, config: dict[str, Any]) -> None:
    os.makedirs(MAPPINGS_DIR, exist_ok=True)
    _atomic_json_write(_platform_config_path(platform_name), config)


def delete_platform_config(platform_name: str) -> None:
    path = _platform_config_path(platform_name)
    if os.path.exists(path):
        os.remove(path)


def import_profile_path(fingerprint: str) -> str:
    """Return the profile path for a stable structure fingerprint."""
    return os.path.join(MAPPINGS_DIR,
                        f"profile_{storage_key(fingerprint, replace_dots=True)}.json")


def load_import_profile(fingerprint: str) -> dict[str, Any] | None:
    path = import_profile_path(fingerprint)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError):
        logger.warning("Import profile unreadable: %s", path, exc_info=True)
        return None
    return data if isinstance(data, dict) else None


def save_import_profile(fingerprint: str,
                        profile: dict[str, Any]) -> dict[str, Any]:
    os.makedirs(MAPPINGS_DIR, exist_ok=True)
    payload = dict(profile)
    payload.setdefault("schema_version", 1)
    payload["fingerprint"] = fingerprint
    _atomic_json_write(import_profile_path(fingerprint), payload)
    return payload


def list_import_profiles() -> list[dict[str, Any]]:
    if not os.path.isdir(MAPPINGS_DIR):
        return []
    profiles = []
    for filename in sorted(os.listdir(MAPPINGS_DIR)):
        if not (filename.startswith("profile_") and filename.endswith(".json")):
            continue
        path = os.path.join(MAPPINGS_DIR, filename)
        try:
            with open(path, "r", encoding="utf-8") as handle:
                profile = json.load(handle)
        except (OSError, json.JSONDecodeError):
            logger.warning("Import profile unreadable: %s", path, exc_info=True)
            continue
        if isinstance(profile, dict):
            profiles.append(profile)
    return profiles
