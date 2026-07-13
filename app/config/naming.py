"""Safe, deterministic names for user-controlled files and folders.

The application stores platform configs, fingerprints, mappings, templates,
and client exports beneath program-managed directories.  These helpers keep
names portable on Windows and prevent path traversal while preserving the
legacy spelling for ordinary names.
"""

from __future__ import annotations

import os
import re

_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_MULTI_UNDERSCORE = re.compile(r"_+")
_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def safe_component(value: object, fallback: str = "untitled", max_length: int = 120) -> str:
    """Return a Windows-safe single path component.

    Directory separators, control characters, trailing dots/spaces, and
    reserved device names are neutralized.  The result never contains a path.
    """
    name = str(value or "").strip()
    name = _INVALID.sub("_", name)
    name = name.replace("..", "_")
    name = name.rstrip(". ")
    if not name:
        name = fallback

    stem = os.path.splitext(name)[0].upper()
    if stem in _RESERVED:
        name = f"_{name}"

    if len(name) > max_length:
        root, ext = os.path.splitext(name)
        keep = max(1, max_length - len(ext))
        name = root[:keep].rstrip(". ") + ext
    return name or fallback


def storage_key(value: object, *, replace_dots: bool = False) -> str:
    """Return a lowercase underscore key suitable for internal JSON names."""
    name = safe_component(value)
    name = name.replace(" ", "_")
    if replace_dots:
        name = name.replace(".", "_")
    name = _MULTI_UNDERSCORE.sub("_", name).strip("_").lower()
    return name or "untitled"


def safe_child(base_dir: str, component: object, *, fallback: str = "untitled") -> str:
    """Join a safe component beneath *base_dir* and verify containment."""
    base = os.path.abspath(base_dir)
    child = os.path.abspath(os.path.join(base, safe_component(component, fallback=fallback)))
    if os.path.commonpath([base, child]) != base:
        raise ValueError("Unsafe path component")
    return child


def unique_component(value: object, used: set[str], *, fallback: str = "untitled",
                     max_length: int = 120) -> str:
    """Return a safe component that is unique within a case-insensitive set.

    ``used`` is updated in place. This matters on Windows, where names that
    differ only by case—or that sanitize to the same text—address the same
    folder.
    """
    base = safe_component(value, fallback=fallback, max_length=max_length)
    candidate = base
    root, ext = os.path.splitext(base)
    index = 2
    while candidate.casefold() in used:
        suffix = f"_{index}"
        keep = max(1, max_length - len(ext) - len(suffix))
        candidate = f"{root[:keep].rstrip('. ')}{suffix}{ext}"
        index += 1
    used.add(candidate.casefold())
    return candidate
