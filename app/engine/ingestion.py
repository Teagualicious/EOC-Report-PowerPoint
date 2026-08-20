"""Stage 1 dump ingestion contracts and deterministic source inspection.

This module owns the boundary between one source export and the headless
workflow.  It deliberately separates *structure discovery* from *data
interpretation*:

* :func:`inspect_source` reads one CSV, XLSX, XLSM, or HTML export and returns
  parser tables plus a structure descriptor with no user data samples;
* :func:`structure_fingerprint` hashes normalized source type, sheet names,
  and header multisets, so column order and workbook sheet order do not alter
  the fingerprint;
* :func:`suggest_import_profile` creates the reviewable Stage 1 profile shape;
* :func:`validate_import_profile` turns an approved profile into the existing
  platform-config shape used by the parsers;
* :func:`flatten_unified_rows` and :func:`reconcile_source` create deterministic
  audit data for later staging and validation stages.

Profile schema version 1 is intentionally small and JSON-only::

    {
      "schema_version": 1,
      "fingerprint": "<64 lowercase hex characters>",
      "source_type": "csv|excel|html",
      "sheets": [{
        "sheet_name": "source sheet name",
        "columns": [{
          "column": "source header",
          "role": "campaign_id|level: <prefix>|metric|skip",
          "selected": true
        }]
      }]
    }

The profile is the approval to interpret a known structure.  A missing or
mismatched profile is never replaced with automatic parser classification;
the workflow returns ``status == "profile_required"`` instead.
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
from collections import Counter, defaultdict
from typing import Any, Mapping

from config.settings import import_profile_path
from engine.errors import ConfigError, ParserError, UserInputError
from parsers.csv_parser import CSVParser
from parsers.dictionary import classify_columns, match_metric_alias
from parsers.excel_parser import ExcelParser
from parsers.html_parser import HTMLParser


PROFILE_SCHEMA_VERSION = 1
SUPPORTED_EXTENSIONS = (".csv", ".xlsx", ".xlsm", ".html", ".htm")
_SOURCE_TYPES = {
    ".csv": "csv",
    ".xlsx": "excel",
    ".xlsm": "excel",
    ".html": "html",
    ".htm": "html",
}
_SPACE = re.compile(r"\s+")


def normalize_structure_label(value: object) -> str:
    """Normalize a sheet/header label for structure identity comparisons."""
    return _SPACE.sub(" ", str(value or "").strip()).casefold()


def source_type_for_path(path: str | os.PathLike[str]) -> str:
    """Return the parser family for a supported path or raise an actionable error."""
    extension = os.path.splitext(os.fspath(path))[1].lower()
    source_type = _SOURCE_TYPES.get(extension)
    if source_type is None:
        raise UserInputError(
            f"Unsupported dump extension: {extension or '<none>'}",
            user_message=(
                "Choose one campaign export in CSV, XLSX, XLSM, or HTML format."
            ),
            code="UNSUPPORTED_EXTENSION",
        )
    return source_type


def _read_parser(path: str, source_type: str) -> dict[str, Any]:
    if source_type == "csv":
        return CSVParser(path).parse(build_outputs=False)
    if source_type == "excel":
        return ExcelParser(path).parse(build_outputs=False)
    return HTMLParser(path, "default").parse()


def _table_name(table: Mapping[str, Any], source_type: str, index: int) -> str:
    name = str(table.get("sheet_name") or "").strip()
    if name:
        return name
    if source_type == "html":
        return f"Table {index + 1}"
    if source_type == "csv":
        return "CSV"
    return f"Sheet {index + 1}"


def _fingerprint_name(name: str, source_type: str, index: int) -> str:
    normalized = normalize_structure_label(name)
    if source_type == "csv":
        # A CSV filename is not a structural fact.  This keeps profile reuse
        # stable when an analyst saves the same export under a new name.
        return "csv"
    if source_type == "html" and normalized in {"html report", "report"}:
        return "html report"
    if source_type == "html" and normalized.startswith("table "):
        return f"table {index + 1}"
    return normalized


def _duplicate_labels(labels: list[str]) -> list[str]:
    counts = Counter(normalize_structure_label(label) for label in labels)
    return sorted(label for label, count in counts.items() if count > 1)


def _role_suggestions(headers: list[str]) -> dict[str, str]:
    classification = classify_columns(headers)
    roles: dict[str, str] = {}
    for column in classification["context"].values():
        roles[normalize_structure_label(column)] = "campaign_id"
    for column, level_def in classification["levels"]:
        roles[normalize_structure_label(column)] = f"level: {level_def['prefix']}"
    for column in classification["skip"]:
        roles[normalize_structure_label(column)] = "skip"
    for column, _universal in classification["metrics"]:
        roles[normalize_structure_label(column)] = "metric"
    return roles


def _metric_columns(profile_sheet: Mapping[str, Any] | None,
                    table: Mapping[str, Any]) -> list[tuple[str, str]]:
    """Return selected source-column/universal-metric pairs for reconciliation."""
    if profile_sheet is not None:
        pairs = []
        for item in profile_sheet.get("columns", []):
            if item.get("role") != "metric" or not item.get("selected", True):
                continue
            column = str(item.get("column", ""))
            universal = match_metric_alias(column) or column
            pairs.append((column, universal))
        return pairs
    return list(table.get("found_metrics", []))


def inspect_source(path: str | os.PathLike[str]) -> tuple[dict[str, Any], dict[str, Any]]:
    """Read one source once and return ``(parser_data, structure_descriptor)``.

    The descriptor contains row counts and headers but never stores source row
    values.  Empty sources and duplicate normalized headers fail before a
    profile can be approved, so malformed input cannot be mistaken for a new
    valid structure.
    """
    raw_path = os.fspath(path)
    if not raw_path:
        raise UserInputError(
            "No dump path was supplied.",
            user_message="Select one campaign export before parsing it.",
            code="SOURCE_PATH_REQUIRED",
        )
    source_type = source_type_for_path(raw_path)
    if not os.path.isfile(raw_path):
        raise ParserError(
            f"File not found: {raw_path}",
            user_message=f"The selected export does not exist: {raw_path}",
            code="SOURCE_NOT_FOUND",
        )
    parsed = _read_parser(raw_path, source_type)
    structure = build_structure_descriptor(parsed, source_type, raw_path)
    return parsed, structure


def build_structure_descriptor(parsed: Mapping[str, Any], source_type: str,
                               path: str | os.PathLike[str] | None = None
                               ) -> dict[str, Any]:
    """Build the JSON structure descriptor used by profiles and fingerprints."""
    tables = list(parsed.get("detected_tables", []))
    if not tables and source_type == "html" and parsed.get("metrics"):
        tables = [{
            "sheet_name": "HTML Report",
            "headers": list(parsed["metrics"].keys()),
            "rows": [dict(parsed["metrics"])],
            "row_count": 1,
        }]

    sheets: list[dict[str, Any]] = []
    for index, table in enumerate(tables):
        headers = [str(header).strip() for header in table.get("headers", [])]
        headers = [header for header in headers if header]
        if not headers:
            continue
        duplicate_labels = _duplicate_labels(headers)
        if duplicate_labels:
            display = ", ".join(duplicate_labels)
            name = _table_name(table, source_type, index)
            raise ParserError(
                f"Duplicate headers in {name}: {display}",
                user_message=(
                    f"The export has duplicate column names in {name} ({display}). "
                    "Rename the duplicate columns and export it again."
                ),
                code="DUPLICATE_HEADERS",
            )
        normalized_headers = sorted(normalize_structure_label(header)
                                    for header in headers)
        name = _table_name(table, source_type, index)
        role_map = _role_suggestions(headers)
        sheets.append({
            "index": index,
            "sheet_name": name,
            "normalized_name": normalize_structure_label(name),
            "fingerprint_name": _fingerprint_name(name, source_type, index),
            "headers": headers,
            "normalized_headers": normalized_headers,
            "row_count": int(table.get("row_count", len(table.get("rows", [])))),
            "columns": [
                {
                    "column": header,
                    "normalized": normalize_structure_label(header),
                    "suggested_role": role_map.get(
                        normalize_structure_label(header), "metric"),
                }
                for header in headers
            ],
        })

    if not sheets:
        source_name = os.path.basename(os.fspath(path)) if path else "the export"
        raise ParserError(
            f"No usable table found in {source_name}",
            user_message=(
                "The export has no usable header and data rows. Check that the "
                "file contains a report table and export it again."
            ),
            code="EMPTY_SOURCE",
        )

    return {
        "schema_version": PROFILE_SCHEMA_VERSION,
        "source_type": source_type,
        "sheets": sheets,
    }


def structure_fingerprint(structure: Mapping[str, Any]) -> str:
    """Return a SHA-256 fingerprint of structure, not row values or filenames."""
    canonical_sheets = sorted(
        (
            {
                "name": str(sheet.get("fingerprint_name") or
                            sheet.get("normalized_name") or ""),
                "headers": sorted(str(header) for header in
                                   sheet.get("normalized_headers", [])),
            }
            for sheet in structure.get("sheets", [])
        ),
        key=lambda item: (item["name"], item["headers"]),
    )
    payload = {
        "schema_version": PROFILE_SCHEMA_VERSION,
        "source_type": structure.get("source_type", ""),
        "sheets": canonical_sheets,
    }
    encoded = json.dumps(payload, ensure_ascii=False, sort_keys=True,
                          separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


# Descriptive aliases keep the contract easy to discover for later stages.
compute_structure_fingerprint = structure_fingerprint


def suggest_import_profile(structure: Mapping[str, Any], fingerprint: str | None = None
                           ) -> dict[str, Any]:
    """Return the deterministic, reviewable profile suggestion for a structure."""
    fingerprint = fingerprint or structure_fingerprint(structure)
    sheets = []
    for sheet in structure.get("sheets", []):
        sheets.append({
            "sheet_name": sheet["sheet_name"],
            "columns": [
                {
                    "column": column["column"],
                    "role": column["suggested_role"],
                    "selected": column["suggested_role"] != "skip",
                }
                for column in sheet.get("columns", [])
            ],
        })
    return {
        "schema_version": PROFILE_SCHEMA_VERSION,
        "fingerprint": fingerprint,
        "source_type": structure.get("source_type", ""),
        "sheets": sheets,
    }


build_import_profile = suggest_import_profile


def _resolve_profile_sheet(raw_sheet: Mapping[str, Any], structure_sheets: list[dict[str, Any]],
                           used: set[int]) -> tuple[int, dict[str, Any]]:
    raw_name = raw_sheet.get("sheet_name", raw_sheet.get("name", ""))
    normalized = normalize_structure_label(raw_name)
    candidates = [
        (index, sheet) for index, sheet in enumerate(structure_sheets)
        if index not in used and normalized
        and sheet["normalized_name"] == normalized
    ]
    if not candidates and isinstance(raw_sheet.get("index"), int):
        index = raw_sheet["index"]
        if 0 <= index < len(structure_sheets) and index not in used:
            candidates = [(index, structure_sheets[index])]
    if not candidates and len(structure_sheets) == 1 and not used:
        candidates = [(0, structure_sheets[0])]
    if not candidates:
        raise ConfigError(
            f"Import profile sheet does not match the source: {raw_name or '<unnamed>'}",
            user_message=(
                "The saved import profile does not match this export's sheet names. "
                "Review the profile for the current structure."
            ),
            code="PROFILE_SHEET_MISMATCH",
        )
    return candidates[0]


def validate_import_profile(profile: Mapping[str, Any], structure: Mapping[str, Any],
                            fingerprint: str | None = None) -> dict[str, Any]:
    """Validate and canonicalize an approved profile for the current source."""
    if not isinstance(profile, Mapping):
        raise ConfigError(
            "Import profile must be a JSON object.",
            user_message="The import profile is not valid JSON object data.",
            code="PROFILE_INVALID",
        )
    expected = fingerprint or structure_fingerprint(structure)
    supplied = str(profile.get("fingerprint") or
                   profile.get("structure_fingerprint") or "").lower()
    if supplied != expected:
        raise ConfigError(
            f"Import profile fingerprint {supplied or '<missing>'} does not match "
            f"{expected}",
            user_message=(
                "The import profile belongs to a different export structure. "
                "Choose or create the profile for this source."
            ),
            code="PROFILE_FINGERPRINT_MISMATCH",
        )
    try:
        schema_version = int(profile.get("schema_version", PROFILE_SCHEMA_VERSION))
    except (TypeError, ValueError) as exc:
        raise ConfigError(
            "Import profile schema_version is not an integer.",
            user_message="The import profile version is invalid; create it again.",
            code="PROFILE_SCHEMA_INVALID",
        ) from exc
    if schema_version != PROFILE_SCHEMA_VERSION:
        raise ConfigError(
            f"Unsupported import profile schema version: {schema_version}",
            user_message=(
                f"This import profile uses version {schema_version}; expected "
                f"version {PROFILE_SCHEMA_VERSION}. Create it again."
            ),
            code="PROFILE_SCHEMA_UNSUPPORTED",
        )
    supplied_source_type = str(profile.get("source_type", "")).strip().casefold()
    expected_source_type = str(structure.get("source_type", "")).strip().casefold()
    if supplied_source_type and supplied_source_type != expected_source_type:
        raise ConfigError(
            f"Import profile source type {supplied_source_type} does not match "
            f"{expected_source_type}",
            user_message=(
                "The import profile belongs to a different export format. "
                "Choose the profile for this source."
            ),
            code="PROFILE_SOURCE_TYPE_MISMATCH",
        )

    raw_sheets = profile.get("sheets", profile.get("tables"))
    if not isinstance(raw_sheets, list) or not raw_sheets:
        raise ConfigError(
            "Import profile has no sheets.",
            user_message=(
                "The import profile has no sheet and column roles. Review the "
                "profile before parsing the export."
            ),
            code="PROFILE_SHEETS_REQUIRED",
        )

    structure_sheets = list(structure.get("sheets", []))
    canonical_sheets = []
    used: set[int] = set()
    allowed_roles = {"campaign_id", "metric", "skip"}
    for raw_sheet in raw_sheets:
        if not isinstance(raw_sheet, Mapping):
            raise ConfigError(
                "Import profile contains a non-object sheet.",
                user_message="Each import-profile sheet must be an object.",
                code="PROFILE_SHEET_INVALID",
            )
        index, source_sheet = _resolve_profile_sheet(raw_sheet, structure_sheets, used)
        used.add(index)
        source_headers = {
            normalize_structure_label(header): header
            for header in source_sheet.get("headers", [])
        }
        raw_columns = raw_sheet.get("columns")
        if not isinstance(raw_columns, list):
            raise ConfigError(
                f"Import profile sheet {source_sheet['sheet_name']} has no columns.",
                user_message=(
                    f"Select the columns and roles for {source_sheet['sheet_name']}."
                ),
                code="PROFILE_COLUMNS_REQUIRED",
            )
        columns = []
        seen_columns: set[str] = set()
        for raw_column in raw_columns:
            if not isinstance(raw_column, Mapping):
                raise ConfigError(
                    "Import profile contains a non-object column.",
                    user_message="Each import-profile column must be an object.",
                    code="PROFILE_COLUMN_INVALID",
                )
            requested_name = str(raw_column.get("column", raw_column.get("name", "")))
            normalized_name = normalize_structure_label(requested_name)
            actual_name = source_headers.get(normalized_name)
            if not actual_name:
                raise ConfigError(
                    f"Import profile column is not in the source: {requested_name}",
                    user_message=(
                        f"The profile names a missing column ({requested_name}) in "
                        f"{source_sheet['sheet_name']}."
                    ),
                    code="PROFILE_COLUMN_MISMATCH",
                )
            if normalized_name in seen_columns:
                raise ConfigError(
                    f"Import profile repeats column: {requested_name}",
                    user_message=f"The profile repeats column {requested_name}.",
                    code="PROFILE_COLUMN_DUPLICATE",
                )
            seen_columns.add(normalized_name)
            role = str(raw_column.get("role", "")).strip().casefold()
            if role in {"campaign_name", "context: campaign_name"}:
                role = "campaign_id"
            if role.startswith("level:"):
                role = "level: " + role.split(":", 1)[1].strip()
            if role not in allowed_roles and not role.startswith("level: "):
                raise ConfigError(
                    f"Unsupported import profile role: {role or '<missing>'}",
                    user_message=(
                        "Use campaign_id, metric, skip, or level: <prefix> as an "
                        "import-profile role."
                    ),
                    code="PROFILE_ROLE_INVALID",
                )
            columns.append({
                "column": actual_name,
                "role": role,
                "selected": bool(raw_column.get("selected", role != "skip")),
            })
        canonical_sheets.append({
            "sheet_name": source_sheet["sheet_name"],
            "columns": columns,
        })

    return {
        "schema_version": PROFILE_SCHEMA_VERSION,
        "fingerprint": expected,
        "source_type": structure.get("source_type", ""),
        "sheets": canonical_sheets,
    }


def flatten_unified_rows(parsed: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Flatten parser output into stable campaign/breakdown rows for audit use."""
    rows: list[dict[str, Any]] = []
    campaign_items = sorted(
        parsed.get("campaign_metrics", {}).items(),
        key=lambda pair: str(pair[0]),
    )
    for key, item in campaign_items:
        rows.append({
            "row_type": "campaign",
            "campaign_name": item.get("campaign_name", ""),
            "metric_level": "",
            "metric_name": item.get("universal_name", ""),
            "metric_value": item.get("value"),
            "source_key": str(key),
        })
    level_rows = list(parsed.get("level_data", []))
    level_rows.sort(key=lambda row: (
        str(row.get("_campaign", "")), str(row.get("metric_level", "")),
        str(row.get("metric_name", "")), repr(row.get("metric_value")),
    ))
    for row in level_rows:
        rows.append({
            "row_type": "breakdown",
            "campaign_name": row.get("_campaign", ""),
            "metric_level": row.get("metric_level", ""),
            "metric_name": row.get("metric_name", ""),
            "metric_value": row.get("metric_value"),
            "source_key": (
                f"{row.get('_campaign', '')}|{row.get('metric_level', '')}|"
                f"{row.get('metric_name', '')}"
            ),
        })
    return rows


def _numeric(value: Any) -> float | int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if math.isfinite(float(value)):
            return value
        return None
    if isinstance(value, str):
        cleaned = value.replace("$", "").replace(",", "").replace("%", "").strip()
        if not cleaned:
            return None
        try:
            number = float(cleaned)
        except ValueError:
            return None
        if math.isfinite(number):
            return int(number) if number.is_integer() else number
    return None


def _profile_sheet_for_table(profile: Mapping[str, Any] | None,
                             table: Mapping[str, Any], table_index: int,
                             table_count: int) -> Mapping[str, Any] | None:
    if not profile:
        return None
    sheets = profile.get("sheets", [])
    name = normalize_structure_label(table.get("sheet_name", ""))
    for sheet in sheets:
        if normalize_structure_label(sheet.get("sheet_name", "")) == name:
            return sheet
    if len(sheets) == 1 and table_count == 1:
        return sheets[0]
    return None


def reconcile_source(parsed: Mapping[str, Any], structure: Mapping[str, Any],
                     profile: Mapping[str, Any] | None = None,
                     unified_rows: list[Mapping[str, Any]] | None = None
                     ) -> dict[str, Any]:
    """Return deterministic source-cell versus resolved-row reconciliation data."""
    rows = unified_rows if unified_rows is not None else flatten_unified_rows(parsed)
    parser_tables = list(parsed.get("detected_tables", []))
    if not parser_tables and parsed.get("metrics"):
        parser_tables = [{
            "sheet_name": "HTML Report",
            "headers": list(parsed["metrics"].keys()),
            "rows": [dict(parsed["metrics"])],
            "row_count": 1,
            "found_metrics": [
                (key, match_metric_alias(key) or key)
                for key in parsed["metrics"]
            ],
        }]

    source_counts: Counter[str] = Counter()
    source_totals: defaultdict[str, float] = defaultdict(float)
    source_rows = 0
    for index, table in enumerate(parser_tables):
        table_rows = list(table.get("rows", []))
        source_rows += len(table_rows)
        profile_sheet = _profile_sheet_for_table(
            profile, table, index, len(parser_tables))
        pairs = _metric_columns(profile_sheet, table)
        for row in table_rows:
            for column, metric in pairs:
                value = row.get(column, "")
                if value == "" or value is None:
                    continue
                source_counts[metric] += 1
                numeric = _numeric(value)
                if numeric is not None:
                    source_totals[metric] += float(numeric)

    resolved_counts: Counter[str] = Counter(
        str(row.get("metric_name", "")) for row in rows
        if row.get("metric_name", "")
    )
    resolved_totals: defaultdict[str, float] = defaultdict(float)
    for row in rows:
        metric = str(row.get("metric_name", ""))
        numeric = _numeric(row.get("metric_value"))
        if metric and numeric is not None:
            resolved_totals[metric] += float(numeric)

    metrics = {}
    review_notes = []
    for metric in sorted(set(source_counts) | set(resolved_counts)):
        source_total = source_totals.get(metric, 0.0)
        resolved_total = resolved_totals.get(metric, 0.0)
        count_match = source_counts[metric] == resolved_counts[metric]
        metrics[metric] = {
            "source_cells": source_counts[metric],
            "resolved_rows": resolved_counts[metric],
            "source_numeric_total": source_total,
            "resolved_numeric_total": resolved_total,
            "numeric_delta": resolved_total - source_total,
        }
        if not count_match:
            review_notes.append(
                f"{metric}: {source_counts[metric]} source value(s) became "
                f"{resolved_counts[metric]} resolved row(s)."
            )

    return {
        "schema_version": PROFILE_SCHEMA_VERSION,
        "status": "review" if review_notes else "ok",
        "source_table_count": len(structure.get("sheets", [])),
        "source_row_count": source_rows,
        "resolved_row_count": len(rows),
        "campaign_metric_row_count": len(parsed.get("campaign_metrics", {})),
        "breakdown_row_count": len(parsed.get("level_data", [])),
        "metrics": metrics,
        "notes": review_notes,
    }


def profile_path(fingerprint: str) -> str:
    """Return the configured path for a version-1 profile."""
    return import_profile_path(fingerprint)
