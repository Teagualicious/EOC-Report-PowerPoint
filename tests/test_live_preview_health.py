"""Tests for the live-preview health monitor (engine.pptx_live).

Real COM never runs here — pywin32 is absent on Linux, so __init__ leaves
the preview inactive. The tests then activate the object against small
stub "presentations" to drive the _tracked wrapper: consecutive-failure
counting, reset-on-success, self-disable, and the one-shot on_disabled
callback. This is exactly the logic that must hold on Windows regardless
of what PowerPoint throws.
"""

import pytest

from engine.pptx_live import PPTXLivePreview


class _BrokenPresentation:
    """Every COM touch explodes — simulates PowerPoint dying mid-session."""

    class _Slides:
        def __call__(self, n):
            raise OSError("RPC server is unavailable")

        @property
        def Count(self):
            raise OSError("RPC server is unavailable")

    def __init__(self):
        self.Slides = self._Slides()

    def SaveAs(self, path):
        raise OSError("RPC server is unavailable")


class _WorkingPresentation:
    """Just enough COM surface for export_slide_image to succeed."""

    class _Slide:
        def Export(self, path, fmt, width):
            open(path, "wb").write(b"PNG")

    def Slides(self, n):
        return self._Slide()


@pytest.fixture
def preview(tmp_path):
    tpl = tmp_path / "t.pptx"
    tpl.write_bytes(b"x")
    p = PPTXLivePreview(str(tpl))
    assert p.active is False  # no COM on this platform
    return p


def _activate(p, presentation):
    """Simulate a live Windows session: app handle present, deck open."""
    p.active = True
    p.pptx_app = object()
    p.presentation = presentation


def test_inactive_preview_records_nothing(preview):
    preview.update_shape_text(1, 0, "x")
    preview.save_as("out.pptx")
    assert preview.get_health()["calls"] == 0


def test_consecutive_failures_disable_preview_loudly(preview, tmp_path):
    _activate(preview, _BrokenPresentation())
    notices = []
    preview.on_disabled = notices.append

    preview.update_shape_text(1, 0, "a")
    preview.update_shape_text(1, 1, "b")
    assert preview.active is True  # two strikes: still trying

    preview.go_to_slide(2)  # strike three

    assert preview.active is False
    health = preview.get_health()
    assert health["failures"] == 3
    assert health["consecutive_failures"] == 3
    assert health["last_error"] == "go_to_slide"
    assert "3 consecutive" in health["disabled_reason"]
    assert notices == [health["disabled_reason"]]


def test_success_resets_consecutive_counter(preview, tmp_path):
    broken, working = _BrokenPresentation(), _WorkingPresentation()
    _activate(preview, broken)

    preview.update_shape_text(1, 0, "a")
    preview.update_shape_text(1, 1, "b")
    preview.presentation = working
    out = str(tmp_path / "s.png")
    assert preview.export_slide_image(1, out) == out  # success

    health = preview.get_health()
    assert health["consecutive_failures"] == 0  # reset
    assert health["failures"] == 2              # totals preserved
    assert preview.active is True

    # Sporadic flakiness after a success starts the count fresh
    preview.presentation = broken
    preview.update_shape_text(1, 0, "c")
    preview.update_shape_text(1, 1, "d")
    assert preview.active is True
    preview.update_shape_text(1, 2, "e")
    assert preview.active is False


def test_on_disabled_fires_exactly_once_and_callback_errors_are_safe(preview):
    _activate(preview, _BrokenPresentation())
    calls = []

    def bad_callback(reason):
        calls.append(reason)
        raise RuntimeError("UI is gone")

    preview.on_disabled = bad_callback
    for i in range(5):  # keep hammering past the threshold
        preview.update_shape_text(1, i, "x")
        if not preview.active:
            preview.active = True  # simulate stubborn caller re-enabling

    assert len(calls) == 1  # fired once, exception swallowed


def test_save_as_failure_counts_and_returns_false(preview):
    _activate(preview, _BrokenPresentation())

    assert preview.save_as("out.pptx") is False
    health = preview.get_health()
    assert health["failures"] == 1
    assert health["last_error"] == "save_as"


def test_mapper_fallback_condition_after_disable(preview):
    """mapper_window falls back to the python-pptx fill when
    live_preview.is_active() is False — self-disable must flip exactly
    that switch."""
    _activate(preview, _BrokenPresentation())
    for i in range(3):
        preview.update_shape_text(1, i, "x")

    assert preview.is_active() is False


# ── _resolve_shape identity resolution (mapper roadmap Phase 4) ───────────────
# The COM-side drift logic CI can pin: a fake 1-indexed Shapes collection
# where the positional slot no longer holds the mapped shape.

def _stub_slide(*id_names):
    shapes = [type("_Shape", (), {"Id": i, "Name": n})()
              for i, n in id_names]

    class _Shapes:
        def __call__(self, n):  # COM collections are 1-indexed
            return shapes[n - 1]

        @property
        def Count(self):
            return len(shapes)

    class _Slide:
        Shapes = _Shapes()

    return _Slide()


def test_resolve_shape_fast_path_when_index_still_matches(preview):
    slide = _stub_slide((10, "A"), (42, "Target"))
    shape = preview._resolve_shape(slide, 1, 1, shape_uid=42,
                                   expected_name="Target")
    assert shape.Id == 42


def test_resolve_shape_retargets_by_uid_after_drift(preview):
    """The stored index points at the wrong shape — the uid wins."""
    slide = _stub_slide((10, "A"), (7, "B"), (42, "Target"))
    shape = preview._resolve_shape(slide, 1, 0, shape_uid=42,
                                   expected_name="Target")
    assert shape.Id == 42 and shape.Name == "Target"


def test_resolve_shape_legacy_positional_without_uid(preview):
    """Old mappings carry no uid — positional lookup is unchanged."""
    slide = _stub_slide((10, "A"), (42, "Target"))
    shape = preview._resolve_shape(slide, 1, 1)
    assert shape.Id == 42


def test_resolve_shape_unmatched_identity_returns_none(preview):
    """Mapped shape gone from the deck: never write to the positional
    slot's new occupant."""
    slide = _stub_slide((10, "A"), (42, "B"))
    assert preview._resolve_shape(slide, 1, 0, shape_uid=999,
                                  expected_name="Gone") is None
