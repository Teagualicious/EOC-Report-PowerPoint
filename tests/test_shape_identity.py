"""Unit tests for engine.shape_identity — the shared resolver both the
python-pptx fill engine and the COM live preview use to find a mapping
entry's target shape (mapper roadmap Phase 4).

Resolution contract: uid match → unique-name match → positional index
(legacy entries only) → None (skip, never guess).
"""

from engine.shape_identity import resolve_shape_index

PAIRS = [(10, "Title 1"), (42, "TextBox 3"), (7, "Picture 2"),
         (13, "TextBox 3")]  # duplicate name on purpose


def test_uid_match_wins_regardless_of_stored_position():
    assert resolve_shape_index(PAIRS, 0, shape_uid=7,
                               shape_name="Picture 2") == 2
    assert resolve_shape_index(PAIRS, 3, shape_uid=10) == 0


def test_uid_miss_falls_back_to_unique_name():
    assert resolve_shape_index(PAIRS, 0, shape_uid=999,
                               shape_name="Picture 2") == 2


def test_duplicate_names_never_match():
    """PowerPoint shape names are not unique — matching an ambiguous name
    would just be a different flavor of the wrong-shape bug."""
    assert resolve_shape_index(PAIRS, 1, shape_uid=999,
                               shape_name="TextBox 3") is None


def test_legacy_entry_resolves_positionally():
    assert resolve_shape_index(PAIRS, 2) == 2
    assert resolve_shape_index(PAIRS, 0, shape_uid=None, shape_name="") == 0


def test_legacy_entry_out_of_range_returns_none():
    assert resolve_shape_index(PAIRS, 9) is None
    assert resolve_shape_index(PAIRS, -1) is None


def test_stored_identity_unmatched_never_falls_back_to_position():
    """A mapped shape that was deleted must NOT resolve to the positional
    slot — writing there is exactly the wrong-shape bug this fixes."""
    assert resolve_shape_index(PAIRS, 1, shape_uid=999,
                               shape_name="Deleted Box") is None


def test_empty_collection_resolves_nothing():
    assert resolve_shape_index([], 0) is None
    assert resolve_shape_index([], 0, shape_uid=1, shape_name="X") is None
