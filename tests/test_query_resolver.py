"""Tests for the pandas query engine."""

from conftest import import_query_resolver

qr = import_query_resolver()


class TestStringKeys:
    def test_client_name(self):
        assert qr.resolve_query("__client_name__", [], "Acme Appliance Co") == "Acme Appliance Co"

    def test_date_range(self):
        assert qr.resolve_query("__date_range__", [], "",
                                "2026-05-01", "2026-05-31") == "2026-05-01 - 2026-05-31"

    def test_date_range_partial(self):
        assert qr.resolve_query("__date_range__", [], "", "2026-05-01", "") == "2026-05-01"

    def test_months_and_year(self):
        assert qr.resolve_query("__start_month__", [], "", "2026-05-01", "") == "May"
        assert qr.resolve_query("__end_month__", [], "", "", "2026-06-30") == "June"
        assert qr.resolve_query("__year__", [], "", "2026-05-01", "") == "2026"

    def test_bad_date_falls_back(self):
        assert qr.resolve_query("__start_month__", [], "", "not-a-date", "") == "not-a-date"
        assert qr.resolve_query("__year__", [], "", "garbage", "") == ""

    def test_unknown_key(self):
        assert qr.resolve_query("__bogus__", []) == ""

    def test_total_key(self, client_data):
        # Campaign A contributes its best source (zip: 60,000) and
        # Campaign B contributes its summary source (88,450).
        val = qr.resolve_query("__total_Impressions__", client_data)
        assert val == 148450

    def test_total_key_with_source(self, client_data):
        assert qr.resolve_query("__total_device_Impressions__", client_data) == 50000


class TestDictQueries:
    def test_sum_specific_breakdown(self, client_data):
        q = {"metric": "Impressions", "breakdown": "zip", "filter": "all", "agg": "sum"}
        assert qr.resolve_query(q, client_data) == 60000

    def test_best_source_for_all(self, client_data):
        q = {"metric": "Impressions", "breakdown": "all", "filter": "all", "agg": "sum"}
        assert qr.resolve_query(q, client_data) == 148450

    def test_filter_by_level_value(self, client_data):
        q = {"metric": "Impressions", "breakdown": "device", "filter": "Roku",
             "agg": "sum"}
        assert qr.resolve_query(q, client_data) == 30000

    def test_filter_case_insensitive(self, client_data):
        q = {"metric": "Impressions", "breakdown": "device", "filter": "roku",
             "agg": "sum"}
        assert qr.resolve_query(q, client_data) == 30000

    def test_aggregations(self, client_data):
        base = {"metric": "Impressions", "breakdown": "device", "filter": "all"}
        assert qr.resolve_query({**base, "agg": "max"}, client_data) == 30000
        assert qr.resolve_query({**base, "agg": "min"}, client_data) == 20000
        assert qr.resolve_query({**base, "agg": "avg"}, client_data) == 25000
        assert qr.resolve_query({**base, "agg": "count"}, client_data) == 2

    def test_missing_metric_returns_zero(self, client_data):
        q = {"metric": "Nonexistent", "breakdown": "all", "filter": "all", "agg": "sum"}
        assert qr.resolve_query(q, client_data) == 0

    def test_empty_data_returns_zero(self):
        q = {"metric": "Impressions", "breakdown": "all", "filter": "all", "agg": "sum"}
        assert qr.resolve_query(q, []) == 0

    def test_non_dict_non_string(self):
        assert qr.resolve_query(42, []) == ""


class TestHelpers:
    def test_get_available_breakdowns(self, client_data):
        bd = qr.get_available_breakdowns(client_data)
        assert bd["device"] == ["Mobile", "Roku"]
        assert bd["zip"] == ["33607", "33609"]


    def test_build_simple_options(self):
        structured = {"totals": {"Impressions": 100, "100% Completions": 50},
                      "breakdowns": {"device": []}}
        options = qr.build_simple_options(structured)
        keys = [o["key"] for o in options]
        assert "__client_name__" in keys
        assert "__total_Impressions__" in keys
        assert "__total_device_Impressions__" in keys
        # Display alias applied to label but key keeps the raw metric
        completions = [o for o in options if o["key"] == "__total_100% Completions__"]
        assert completions[0]["label"] == "📊 Total Completions"


# ── Builder-query re-resolution (resolver fidelity, 2026-07-15) ──────────────
# Queries saved by the Advanced Query Builder carry the full selection
# (campaigns/sources/values/top_n). resolve_query used to ignore those keys,
# so a refill in a later session could silently diverge from the pivot value
# the user saw at apply time. Builder queries now re-resolve through the
# same engine.pivot.build_pivot the builder previewed.

def _builder_client_data():
    return [{
        "campaign_metrics": {
            "A|Impressions": {"value": 1000, "universal_name": "Impressions",
                              "campaign_name": "A"},
            "B|Impressions": {"value": 500, "universal_name": "Impressions",
                              "campaign_name": "B"},
        },
        "level_data": [
            {"metric_name": "Impressions", "metric_value": 600,
             "metric_level": "zip:11111", "_campaign": "A"},
            {"metric_name": "Impressions", "metric_value": 400,
             "metric_level": "zip:22222", "_campaign": "A"},
            {"metric_name": "Impressions", "metric_value": 300,
             "metric_level": "zip:11111", "_campaign": "B"},
        ],
    }]


def test_builder_query_matches_apply_time_pivot_total():
    """The resolved value must equal pivot_total(build_pivot(...)) — the
    exact number the builder displayed and cached when the user applied."""
    import pandas as pd
    from engine.pivot import build_pivot, pivot_total
    from engine.query_resolver import resolve_query

    query = {"metric": "Impressions", "agg": "sum", "top_n": "1",
             "campaigns": [], "sources": ["zip"], "values": []}
    resolved = resolve_query(query, _builder_client_data())

    rows = pd.DataFrame([
        {"campaign": "A", "metric": "Impressions", "value": 1000,
         "source": "campaign", "level_value": ""},
        {"campaign": "B", "metric": "Impressions", "value": 500,
         "source": "campaign", "level_value": ""},
        {"campaign": "A", "metric": "Impressions", "value": 600,
         "source": "zip", "level_value": "11111"},
        {"campaign": "A", "metric": "Impressions", "value": 400,
         "source": "zip", "level_value": "22222"},
        {"campaign": "B", "metric": "Impressions", "value": 300,
         "source": "zip", "level_value": "11111"},
    ])
    pivot, _ = build_pivot(rows, "Impressions", "sum", "1", [], ["zip"], [])
    assert resolved == pivot_total(pivot)
    assert resolved == 900  # top-1 zip row: 11111 = 600 + 300


def test_builder_query_honors_campaign_and_value_selections():
    from engine.query_resolver import resolve_query
    data = _builder_client_data()

    only_a = resolve_query({"metric": "Impressions", "agg": "sum",
                            "top_n": "all", "campaigns": ["A"],
                            "sources": ["zip"], "values": []}, data)
    assert only_a == 1000  # A's zip rows only: 600 + 400

    one_zip = resolve_query({"metric": "Impressions", "agg": "sum",
                             "top_n": "all", "campaigns": [],
                             "sources": ["zip"], "values": ["22222"]}, data)
    assert one_zip == 400


def test_builder_query_empty_sources_means_campaign_totals():
    from engine.query_resolver import resolve_query
    total = resolve_query({"metric": "Impressions", "agg": "sum",
                           "top_n": "all", "campaigns": [], "sources": [],
                           "values": []}, _builder_client_data())
    assert total == 1500  # campaign-total rows only, never pooled levels


def test_plain_queries_keep_historical_resolution():
    """Queries without builder keys must go through the unchanged
    metric/breakdown/filter/agg path (best-source-per-campaign etc.)."""
    from engine.query_resolver import resolve_query
    v = resolve_query({"metric": "Impressions", "breakdown": "all",
                       "filter": "all", "agg": "sum"}, _builder_client_data())
    assert v == 1500  # best source per campaign: A=1000 summary, B=500
