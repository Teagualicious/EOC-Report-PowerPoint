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
