"""Regression tests for Windows DPI sizing and template preview fallbacks."""

from pathlib import Path


def test_fit_window_scales_geometry_for_windows_dpi(monkeypatch):
    from ui import utils

    class FakeTk:
        def call(self, *args):
            assert args == ("tk", "scaling")
            return 2.0  # 150% relative to Tk's 96-DPI baseline of 4/3

    class FakeWindow:
        tk = FakeTk()

        def __init__(self):
            self.value = None

        def geometry(self, value):
            self.value = value

        def winfo_screenwidth(self):
            return 1920

        def winfo_screenheight(self):
            return 1080

    win = FakeWindow()
    monkeypatch.setattr(utils.sys, "platform", "win32")
    monkeypatch.setattr(utils, "_work_area", lambda _win: (0, 0, 1920, 1040))

    utils.fit_window(win, 560, 520)

    assert win.value.startswith("840x780+")
    _, position = win.value.split("+", 1)
    x, y = (int(v) for v in position.split("+"))
    assert x >= 0
    assert y >= 0


def test_fit_window_clamps_scaled_dialog_to_work_area(monkeypatch):
    from ui import utils

    class FakeTk:
        def call(self, *args):
            return 2.4

    class FakeWindow:
        tk = FakeTk()

        def __init__(self):
            self.value = None

        def geometry(self, value):
            self.value = value

        def winfo_screenwidth(self):
            return 1280

        def winfo_screenheight(self):
            return 720

    win = FakeWindow()
    monkeypatch.setattr(utils.sys, "platform", "win32")
    monkeypatch.setattr(utils, "_work_area", lambda _win: (0, 0, 1280, 680))

    utils.fit_window(win, 940, 700)

    size = win.value.split("+", 1)[0]
    width, height = (int(v) for v in size.split("x"))
    assert width < 1280
    assert height < 680


def test_template_preview_uses_text_when_thumbnail_unavailable(tmp_path, monkeypatch):
    from pptx import Presentation
    from engine import pptx_thumbs

    path = tmp_path / "sample.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    box = slide.shapes.add_textbox(0, 0, 2_000_000, 500_000)
    box.text_frame.text = "Monthly Performance Summary"
    prs.save(path)

    monkeypatch.setattr(pptx_thumbs, "get_template_thumbnail",
                        lambda *_args, **_kwargs: None)

    preview = pptx_thumbs.get_template_preview(str(path))

    assert preview["kind"] == "text"
    assert "Monthly Performance Summary" in preview["value"]


def test_template_preview_returns_cached_image_path(tmp_path, monkeypatch):
    from engine import pptx_thumbs

    template = tmp_path / "sample.pptx"
    template.write_bytes(b"not needed for mocked thumbnail")
    png = Path(tmp_path / "sample.png")
    png.write_bytes(b"png")
    monkeypatch.setattr(pptx_thumbs, "get_template_thumbnail",
                        lambda *_args, **_kwargs: str(png))

    preview = pptx_thumbs.get_template_preview(str(template))

    assert preview == {"kind": "image", "value": str(png)}
