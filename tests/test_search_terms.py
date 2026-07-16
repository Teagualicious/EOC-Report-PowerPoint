"""Tests for the query search term grammar (engine.search_terms) — the
Excel Search-sheet vocabulary parsed against client data, plus the
display ordering helpers. tk-free (the query builder UI is a thin shell
over these)."""

import pandas as pd
import pytest

from engine.search_terms import (apply_sort, default_metric,
                                 parse_search_terms)

METRICS = ["Impressions", "Clicks", "100% Completions"]
BREAKDOWNS = {"zip": ["33607", "33609"], "network": ["HGTV", "Food Network"],
              "device": ["Roku", "Mobile"]}
CAMPAIGNS = ["Campaign A", "Campaign B"]


def parse(raw):
    return parse_search_terms(raw, METRICS, BREAKDOWNS, CAMPAIGNS)


def test_blank_search_is_campaign_totals():
    parsed, ignored = parse("")
    assert parsed["metric"] is None          # UI substitutes the default
    assert parsed["sources"] == [] and parsed["values"] == []
    assert parsed["campaigns"] == [] and ignored == []


def test_terms_map_to_their_vocabularies():
    parsed, ignored = parse("impressions, zip, Campaign A")
    assert parsed["metric"] == "Impressions"
    assert parsed["sources"] == ["zip"]
    assert parsed["campaigns"] == ["Campaign A"]
    assert ignored == []


def test_level_value_implies_its_breakdown_type():
    parsed, _ = parse("33607, impressions")
    assert parsed["values"] == ["33607"]
    assert parsed["sources"] == ["zip"]      # like typing "33607" in Excel


def test_campaign_keyword_means_campaign_totals():
    parsed, ignored = parse("campaign, impressions")
    assert parsed["sources"] == []           # no breakdown = campaign rows
    assert parsed["metric"] == "Impressions"
    assert ignored == []


def test_top_and_bottom_terms_set_order():
    parsed, _ = parse("top 5, zip, impressions")
    assert parsed["top_n"] == "5" and parsed["sort"] == "largest"
    parsed, _ = parse("bottom 3, zip")
    assert parsed["top_n"] == "3" and parsed["sort"] == "smallest"


def test_alphabetical_terms_set_order():
    assert parse("a-z, network")[0]["sort"] == "az"
    assert parse("z-a, network")[0]["sort"] == "za"


def test_unique_prefix_matches():
    parsed, ignored = parse("netw, impr")
    assert parsed["sources"] == ["network"]
    assert parsed["metric"] == "Impressions"
    assert ignored == []


def test_ambiguous_prefix_is_ignored_loudly():
    # "c" prefixes Clicks, 100%... no; Clicks + Campaign A/B — ambiguous
    parsed, ignored = parse("c")
    assert parsed["metric"] is None and parsed["campaigns"] == []
    assert len(ignored) == 1 and ignored[0][0] == "c"


def test_second_metric_is_ignored_with_reason():
    parsed, ignored = parse("impressions, clicks")
    assert parsed["metric"] == "Impressions"
    assert ignored and "one metric" in ignored[0][1]


def test_unknown_term_reported_never_silent():
    _, ignored = parse("impressions, banana")
    assert ("banana", "not a metric, breakdown type, value, or campaign") \
        in ignored


def test_matching_is_case_insensitive():
    parsed, _ = parse("IMPRESSIONS, ROKU, campaign a")
    assert parsed["metric"] == "Impressions"
    assert parsed["values"] == ["Roku"]
    assert parsed["campaigns"] == ["Campaign A"]


def test_default_metric_prefers_impressions():
    assert default_metric(["Clicks", "Impressions"]) == "Impressions"
    assert default_metric(["Spend", "Clicks"]) == "Clicks"
    assert default_metric([]) == ""


# ── apply_sort ────────────────────────────────────────────────────────────────

@pytest.fixture
def pivot():
    df = pd.DataFrame({"Campaign A": [30, 10, 20], "Total": [30, 10, 20]},
                      index=pd.Index(["b-val", "c-val", "a-val"],
                                     name="level_value"))
    return df.sort_values("Total", ascending=False)  # build_pivot's order


def test_apply_sort_largest_keeps_native_order(pivot):
    assert list(apply_sort(pivot, "largest").index) == \
        ["b-val", "a-val", "c-val"]


def test_apply_sort_smallest_inverts(pivot):
    assert list(apply_sort(pivot, "smallest").index) == \
        ["c-val", "a-val", "b-val"]


def test_apply_sort_alphabetical(pivot):
    assert list(apply_sort(pivot, "az").index) == ["a-val", "b-val", "c-val"]
    assert list(apply_sort(pivot, "za").index) == ["c-val", "b-val", "a-val"]


def test_bottom_n_is_smallest_n(pivot):
    sliced = apply_sort(pivot, "smallest", top_n="2")
    assert list(sliced["Total"]) == [10, 20]


def test_bad_top_n_is_tolerated(pivot):
    assert len(apply_sort(pivot, "largest", top_n="junk")) == 3
    assert apply_sort(None, "largest") is None
