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
        starts — a 91.98 vs 98.36 mismatch on a real order traced to the
        impressions denominator."""
        totals, details, _ = kpi.compute_kpis(self._data(with_starts=True))
        assert totals["Completion Rate"] == 95.0     # 855 / 900
        assert details["X"]["Completion Rate"] == 95.0

    def test_rate_falls_back_to_impressions_without_starts(self):
        totals, details, _ = kpi.compute_kpis(self._data(with_starts=False))
        assert totals["Completion Rate"] == 85.5     # 855 / 1000
        assert details["X"]["Completion Rate"] == 85.5


class TestNonDedupTotals:
    """Cross-campaign Reach/Frequency cannot be deduplicated from campaign
    aggregates (a real order read 1,039,763 vs the vendor's 325,644).
    Approved 2026-07-13: totals stay visible but honestly labeled + flagged;
    per-campaign values are vendor-computed and keep their plain names."""

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
            },
            "level_data": [],
        }]

    def test_totals_are_relabeled_and_flagged(self):
        totals, details, flags = kpi.compute_kpis(self._two_campaign_reach())

        assert "Reach" not in totals
        assert "Frequency" not in totals
        assert totals["Combined Reach (not deduplicated)"] == 1000
        assert totals["Avg Campaign Frequency"] == 2.0
        assert any("deduplication" in str(f) for f in flags)
        assert any("per-campaign" in str(f) for f in flags)

    def test_per_campaign_values_keep_plain_names(self):
        _, details, _ = kpi.compute_kpis(self._two_campaign_reach())

        assert details["A"]["Reach"] == 600
        assert details["B"]["Frequency"] == 2.5
