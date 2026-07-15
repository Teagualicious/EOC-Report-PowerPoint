"""Tests for KPI calculation (engine/kpi.py — extracted during the refactor).

Skipped while the logic still lives inline in the review UI.
"""

import pytest

from conftest import import_kpi

kpi = import_kpi()

pytestmark = pytest.mark.skipif(
    kpi is None, reason="engine.kpi not extracted yet (pre-refactor)")


class TestComputeKpis:
    def test_max_across_sources(self, client_data):
        totals, details, flags = kpi.compute_kpis(client_data)
        # Campaign A Impressions: device sums 50,000, zip sums 60,000 → max 60,000
        assert details["Campaign A"]["Impressions"] == 60000
        # Campaign B comes from campaign-level metrics
        assert details["Campaign B"]["Impressions"] == 88450
        # Grand total = sum of per-campaign bests
        assert totals["Impressions"] == 148450

    def test_alias_folding(self, client_data):
        totals, details, _ = kpi.compute_kpis(client_data)
        # "100% Completions" folds into display metric "Completions"
        assert "Completions" in totals
        assert totals["Completions"] == 27000

    def test_zero_value_flagged(self):
        data = [{
            "campaign_metrics": {
                "X|Clicks": {"value": 0, "universal_name": "Clicks",
                             "campaign_name": "X"}},
            "level_data": [],
        }]
        _, _, flags = kpi.compute_kpis(data)
        assert ("X", "Clicks", "Zero value") in flags

    def test_empty_data(self):
        totals, details, flags = kpi.compute_kpis([])
        assert totals == {}
        assert details == {}
        assert flags == []

    def test_constants_exported(self):
        assert kpi.KPI_METRICS == ["Impressions", "Clicks", "Completions", "Cost"]
        assert kpi.KPI_ALIASES["100% Completions"] == "Completions"


class TestCompletionRate:
    @staticmethod
    def _data(with_starts):
        cm = {
            "X|Impressions": {"value": 1000, "universal_name": "Impressions",
                              "campaign_name": "X"},
            "X|Completions": {"value": 855, "universal_name": "Completions",
                              "campaign_name": "X"},
        }
        if with_starts:
            cm["X|Video Starts"] = {"value": 900,
                                    "universal_name": "Video Starts",
                                    "campaign_name": "X"}
        return [{"campaign_metrics": cm, "level_data": []}]

    def test_rate_uses_video_starts_when_present(self):
        """Approved 2026-07-13: vendor dashboards compute VCR against video
        starts — a several-point VCR mismatch on a real order traced to the
        impressions denominator."""
        totals, details, _ = kpi.compute_kpis(self._data(with_starts=True))
        assert totals["Completion Rate"] == 95.0     # 855 / 900
        assert details["X"]["Completion Rate"] == 95.0

    def test_rate_falls_back_to_impressions_without_starts(self):
        totals, details, _ = kpi.compute_kpis(self._data(with_starts=False))
        assert totals["Completion Rate"] == 85.5     # 855 / 1000
        assert details["X"]["Completion Rate"] == 85.5


class TestNonDedupOmission:
    """Cross-campaign Reach/Frequency cannot be deduplicated from campaign
    aggregates (a real order's campaign sum ran roughly 3x the vendor's deduplicated total).
    Approved 2026-07-14 (supersedes the 2026-07-13 relabel-and-flag rule):
    OMIT them from review KPIs entirely — totals AND per-campaign detail —
    until household-level data/formulas exist. The raw values still land in
    the exported workbook rows."""

    @staticmethod
    def _two_campaign_reach():
        return [{
            "campaign_metrics": {
                "A|Reach": {"value": 600, "universal_name": "Reach",
                            "campaign_name": "A"},
                "B|Reach": {"value": 400, "universal_name": "Reach",
                            "campaign_name": "B"},
                "A|Frequency": {"value": 1.5, "universal_name": "Frequency",
                                "campaign_name": "A"},
                "B|Frequency": {"value": 2.5, "universal_name": "Frequency",
                                "campaign_name": "B"},
                "A|Impressions": {"value": 1000,
                                  "universal_name": "Impressions",
                                  "campaign_name": "A"},
            },
            "level_data": [],
        }]

    def test_reach_and_frequency_omitted_everywhere(self):
        totals, details, _ = kpi.compute_kpis(self._two_campaign_reach())

        assert not any("Reach" in k for k in totals)
        assert not any("Frequency" in k for k in totals)
        assert "Reach" not in details["A"]
        assert "Reach" not in details["B"]
        assert "Frequency" not in details["A"]
        assert "Frequency" not in details["B"]

    def test_other_metrics_unaffected(self):
        totals, details, _ = kpi.compute_kpis(self._two_campaign_reach())

        assert totals["Impressions"] == 1000
        assert details["A"]["Impressions"] == 1000


class TestFormatKpiValue:
    """One formatter for every review display (cards, campaign summaries,
    detail rows) — the cards previously mixed trailing-.00 counts, whole
    counts, and bare unmarked rates on the same screen."""

    def test_counts_render_whole_with_commas(self):
        assert kpi.format_kpi_value("Impressions", 2345678) == "2,345,678"
        assert kpi.format_kpi_value("Completions", 8765432.00) == "8,765,432"

    def test_float_noise_from_sums_is_not_decimals(self):
        assert kpi.format_kpi_value("Completions",
                                    8765432.0000001) == "8,765,432"

    def test_rates_get_percent_sign(self):
        assert kpi.format_kpi_value("Completion Rate", 87.65) == "87.65%"
        assert kpi.format_kpi_value("CTR", 0.42) == "0.42%"
        assert kpi.format_kpi_value("Impression Share", 55.5) == "55.50%"

    def test_money_gets_dollar_sign(self):
        assert kpi.format_kpi_value("Cost", 1234.5) == "$1,234.50"
        assert kpi.format_kpi_value("CPM", 12.5) == "$12.50"
        assert kpi.format_kpi_value("Budget Spend", 99) == "$99.00"

    def test_genuine_ratios_keep_two_decimals(self):
        assert kpi.format_kpi_value("Avg Session Depth", 2.47) == "2.47"

    def test_numpy_and_text_values_are_safe(self):
        import numpy as np
        assert kpi.format_kpi_value("Impressions",
                                    np.int64(1204556)) == "1,204,556"
        assert kpi.format_kpi_value("Impressions",
                                    np.float64(1204556.0)) == "1,204,556"
        assert kpi.format_kpi_value("Notes", "n/a") == "n/a"
