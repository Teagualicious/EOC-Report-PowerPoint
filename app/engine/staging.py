"""Staging-workbook contract boundary.

The implementation lands in Fork Stage 2. Public dataclasses and signatures
exist now so workflow/UI code cannot invent competing representations.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping, Sequence

SCHEMA_VERSION = 1
REPORT_VALUES_SHEET = "Report Values"
IMAGES_SHEET = "Images"
UNIFIED_DATA_SHEET = "Unified Data"
META_SHEET = "_Meta"


@dataclass(frozen=True)
class StagingEdit:
    key: str
    generated: Any
    current: Any


@dataclass
class StagingPayload:
    values: dict[str, Any]
    images: dict[str, str]
    edits: list[StagingEdit] = field(default_factory=list)
    warnings: list[Any] = field(default_factory=list)
    meta: dict[str, Any] = field(default_factory=dict)


def inventory_mapping(mapping: Mapping[str, Any]) -> dict[str, Any]:
    """Return the mapped value keys and image slots required by a template."""
    raise NotImplementedError("Mapping inventory is implemented in Stage 2")


def write_staging_workbook(path: str, *, values: Sequence[Mapping[str, Any]],
                           images: Sequence[Mapping[str, Any]],
                           unified_rows: Sequence[Mapping[str, Any]],
                           meta: Mapping[str, Any],
                           previous_path: str | None = None) -> str:
    """Atomically write one workbook conforming to the v1 staging schema."""
    raise NotImplementedError("Staging writer is implemented in Stage 2")


def read_staging_workbook(path: str, *, expected_mapping_hash: str | None = None,
                          required_keys: Sequence[str] = ()) -> StagingPayload:
    """Read, validate, coerce, and diff a staging workbook by key."""
    raise NotImplementedError("Staging reader is implemented in Stage 2")
