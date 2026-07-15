"""Tests for parsers/dictionary.py — alias matching and column classification."""

from parsers.dictionary import (
    build_campaign_rows,
    build_level_rows,
    classify_columns,
    extract_row_context,
    is_skip_column,
    match_context,
    match_level,
    match_metric_alias,
)

TEST_DICT = {
    "metric_aliases": {
        "Impressions": ["Impressions", "Imps", "Ad Views"],
        "100% Completions": ["100% Complete", "100% Completions"],
        "Cost": ["Cost", "Spend"],
    },
    "level_definitions": {
        "device": {"prefix": "device", "id_column_aliases": ["Device", "Device Type"]},
        "zip": {"prefix": "zip", "id_column_aliases": ["Zip Code", "Zipcode"]},
    },
    "context_columns": {"campaign_name": ["Campaign Name", "Campaign", "Order"]},
    "skip_columns": ["Flight Name", "Order ID"],
}


class TestMatchMetricAlias:
    def test_exact_match(self):
        assert match_metric_alias("Impressions", TEST_DICT) == "Impressions"

    def test_alias_match(self):
        assert match_metric_alias("Imps", TEST_DICT) == "Impressions"
        assert match_metric_alias("Spend", TEST_DICT) == "Cost"

    def test_case_insensitive(self):
        assert match_metric_alias("IMPRESSIONS", TEST_DICT) == "Impressions"
        assert match_metric_alias("  ad views ", TEST_DICT) == "Impressions"

    def test_no_match_returns_none(self):
        assert match_metric_alias("Bananas", TEST_DICT) is None

    def test_real_dictionary_loads(self):
        # Against the shipped metric_dictionary.json
        assert match_metric_alias("Total Impressions") == "Impressions"
        assert match_metric_alias("Video Completions") == "Contributions"


class TestMatchLevelAndContext:
    def test_level_match(self):
        assert match_level("Device", TEST_DICT)["prefix"] == "device"
        assert match_level("zip code", TEST_DICT)["prefix"] == "zip"

    def test_level_no_match(self):
        assert match_level("Campaign", TEST_DICT) is None

    def test_context_match(self):
        assert match_context("Order", TEST_DICT) == "campaign_name"
        assert match_context("campaign", TEST_DICT) == "campaign_name"

    def test_skip_column(self):
        assert is_skip_column("Flight Name", TEST_DICT)
        assert not is_skip_column("Impressions", TEST_DICT)


class TestClassifyColumns:
    def test_full_classification(self):
        result = classify_columns(
            ["Order", "Device", "Impressions", "Flight Name", "Unknown Col"],
            TEST_DICT)
        assert result["context"] == {"campaign_name": "Order"}
        assert result["levels"][0][0] == "Device"
        assert ("Impressions", "Impressions") in result["metrics"]
        assert "Flight Name" in result["skip"]
        # Unknown columns fall through as potential metrics under original name
        assert ("Unknown Col", "Unknown Col") in result["metrics"]

    def test_context_wins_over_metric(self):
        result = classify_columns(["Campaign"], TEST_DICT)
        assert result["context"] == {"campaign_name": "Campaign"}
        assert result["metrics"] == []


class TestRowBuilders:
    def test_extract_row_context(self):
        ctx = extract_row_context(
            {"Order": " Acme Motors - Used VLA ", "Impressions": 45230},
            {"campaign_name": "Order"})
        assert ctx == {"campaign_name": "Acme Motors - Used VLA"}

    def test_extract_row_context_missing(self):
        ctx = extract_row_context({}, {"campaign_name": "Order"})
        assert ctx == {"campaign_name": ""}

    def test_build_level_rows(self):
        levels = [("Device", {"prefix": "device"})]
        metrics = [("Impressions", "Impressions"), ("Spend", "Cost")]
        rows = build_level_rows(
            {"Device": "Roku", "Impressions": 20000, "Spend": 150.5},
            levels, metrics, {"campaign_name": "Campaign A"})
        assert {"metric_level": "device:Roku", "metric_name": "Impressions",
                "metric_value": 20000, "_campaign": "Campaign A"} in rows
        assert {"metric_level": "device:Roku", "metric_name": "Cost",
                "metric_value": 150.5, "_campaign": "Campaign A"} in rows

    def test_build_level_rows_skips_empty(self):
        levels = [("Device", {"prefix": "device"})]
        metrics = [("Impressions", "Impressions")]
        assert build_level_rows({"Device": "", "Impressions": 100},
                                levels, metrics, {}) == []
        assert build_level_rows({"Device": "Roku", "Impressions": ""},
                                levels, metrics, {}) == []

    def test_build_campaign_rows(self):
        metrics = [("Imps", "Impressions")]
        result = build_campaign_rows({"Imps": 500}, metrics,
                                     {"campaign_name": "Camp X"})
        assert result == {"Camp X|Impressions": {
            "value": 500, "universal_name": "Impressions",
            "campaign_name": "Camp X"}}

    def test_build_campaign_rows_no_campaign_key(self):
        metrics = [("Imps", "Impressions")]
        result = build_campaign_rows({"Imps": 500}, metrics, {})
        assert list(result.keys()) == ["Impressions"]
