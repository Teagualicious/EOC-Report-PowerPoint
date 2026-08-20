"""Stage 1 ingestion, profile, reconciliation, and fixture contracts."""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from config import settings
from engine import campaign_dictionary, workflow
from engine.errors import ConfigError, ParserError, UserInputError
from engine.synthetic_fixtures import write_synthetic_export


def _use_temp_profiles(monkeypatch, tmp_path):
    monkeypatch.setattr(settings, "MAPPINGS_DIR", str(tmp_path / "mappings"))


def test_unknown_structure_requires_profile_and_column_reorder_keeps_fingerprint(
        tmp_path, monkeypatch):
    _use_temp_profiles(monkeypatch, tmp_path)
    first_path = tmp_path / "first.csv"
    reordered_path = tmp_path / "renamed.csv"
    first_path.write_text(
        "Campaign,Impressions,Clicks,Cost\nFAKE-001,1000,25,12.50\n",
        encoding="utf-8",
    )
    reordered_path.write_text(
        "Cost,Clicks,Campaign,Impressions\n12.50,25,FAKE-001,1000\n",
        encoding="utf-8",
    )

    first = workflow.parse_dump(str(first_path))
    reordered = workflow.parse_dump(str(reordered_path))

    assert first["status"] == "profile_required"
    assert first["profile_required"] is True
    assert "campaign_metrics" not in first
    assert first["fingerprint"] == reordered["fingerprint"]
    assert first["structure"]["sheets"][0]["row_count"] == 1

    parsed = workflow.parse_dump(str(first_path), profile=first["profile"])
    assert parsed["status"] == "parsed"
    assert parsed["profile_required"] is False
    assert parsed["campaign_metrics"]["FAKE-001|Impressions"]["value"] == 1000
    assert parsed["campaign_dictionary"]["version"] == "v0"
    assert parsed["reconciliation"]["status"] == "ok"
    assert Path(parsed["profile_path"]).is_file()

    replay = workflow.parse_dump(str(first_path))
    assert replay == parsed


def test_structure_change_requires_a_new_profile(tmp_path, monkeypatch):
    _use_temp_profiles(monkeypatch, tmp_path)
    original = tmp_path / "original.csv"
    changed = tmp_path / "changed.csv"
    original.write_text(
        "Campaign,Impressions\nFAKE-001,1000\n", encoding="utf-8")
    changed.write_text(
        "Campaign,Impressions,Clicks\nFAKE-001,1000,25\n", encoding="utf-8")

    first = workflow.parse_dump(str(original))
    second = workflow.parse_dump(str(changed))

    assert first["fingerprint"] != second["fingerprint"]
    assert second["status"] == "profile_required"
    assert second["profile_required"] is True
    with pytest.raises(ConfigError) as mismatch:
        workflow.parse_dump(str(changed), profile=first["profile"])
    assert mismatch.value.code == "PROFILE_FINGERPRINT_MISMATCH"


@pytest.mark.parametrize("suffix", [".csv", ".xlsx", ".xlsm", ".html"])
def test_supported_formats_share_the_stage1_contract(tmp_path, monkeypatch, suffix):
    _use_temp_profiles(monkeypatch, tmp_path)
    path = tmp_path / f"synthetic{suffix}"
    write_synthetic_export(path, count=2)

    required = workflow.parse_dump(str(path))
    parsed = workflow.parse_dump(str(path), profile=required["profile"])

    assert required["status"] == "profile_required"
    assert parsed["status"] == "parsed"
    assert parsed["source"]["type"] in {"csv", "excel", "html"}
    assert parsed["unified_rows"]
    assert parsed["reconciliation"]["source_row_count"] == 2


def test_campaign_dictionary_v0_is_identity_with_notes():
    rows = [{"campaign_name": "FAKE-001", "metric_name": "Impressions",
             "metric_value": 1000}]
    transformed, notes = campaign_dictionary.apply(rows)

    assert transformed == rows
    assert transformed is not rows
    assert any("identity passthrough" in note for note in notes)
    assert any("FAKE" not in note for note in notes)


def test_malformed_inputs_have_actionable_errors(tmp_path):
    duplicate = tmp_path / "duplicate.csv"
    duplicate.write_text(
        "Campaign,Impressions,impressions\nFAKE-001,1,2\n", encoding="utf-8")
    with pytest.raises(ParserError) as duplicate_error:
        workflow.parse_dump(str(duplicate))
    assert duplicate_error.value.code == "DUPLICATE_HEADERS"
    assert "duplicate" in duplicate_error.value.user_message.lower()

    invalid_excel = tmp_path / "broken.xlsx"
    invalid_excel.write_text("not an xlsx", encoding="utf-8")
    with pytest.raises(ParserError) as excel_error:
        workflow.parse_dump(str(invalid_excel))
    assert "valid .xlsx" in excel_error.value.user_message

    unsupported = tmp_path / "notes.txt"
    unsupported.write_text("not a campaign export", encoding="utf-8")
    with pytest.raises(UserInputError) as unsupported_error:
        workflow.parse_dump(str(unsupported))
    assert unsupported_error.value.code == "UNSUPPORTED_EXTENSION"


def test_fifty_thousand_row_synthetic_export_is_within_budget(tmp_path, monkeypatch):
    _use_temp_profiles(monkeypatch, tmp_path)
    path = tmp_path / "large.csv"
    write_synthetic_export(path, count=50000)

    started = time.perf_counter()
    required = workflow.parse_dump(str(path))
    parsed = workflow.parse_dump(str(path), profile=required["profile"])
    elapsed = time.perf_counter() - started

    assert parsed["reconciliation"]["source_row_count"] == 50000
    assert parsed["reconciliation"]["resolved_row_count"] == 150000
    assert elapsed < 15, f"Stage 1 synthetic replay took {elapsed:.2f}s"
