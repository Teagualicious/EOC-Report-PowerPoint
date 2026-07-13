"""Tests for the pandas query engine."""

from conftest import import_query_resolver

qr = import_query_resolver()


class TestStringKeys:
    def test_client_name(self):
        assert qr.resolve_query("__client_name__", [], "Famous Tate") == "Famous Tate"

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
