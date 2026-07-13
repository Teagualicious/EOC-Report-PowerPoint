"""Round-trip test for template bundles: export -> wipe -> import."""
import json
import os
import zipfile

import pytest

from config import paths
from engine import template_bundle
from engine.pptx_mapper import load_template_mapping, save_template_mapping


@pytest.fixture
def store(tmp_path, monkeypatch):
    tdir = tmp_path / "templates"
    idir = tdir / "images"
    mdir = tmp_path / "mappings"
    for d in (tdir, idir, mdir):
        d.mkdir(parents=True)
    monkeypatch.setattr(paths, "TEMPLATES_DIR", str(tdir))
    monkeypatch.setattr(paths, "IMAGES_DIR", str(idir))
    monkeypatch.setattr(paths, "MAPPINGS_DIR", str(mdir))
    import engine.pptx_mapper as pm
    monkeypatch.setattr(pm, "MAPPINGS_DIR", str(mdir))
    monkeypatch.setattr(pm, "TEMPLATES_DIR", str(tdir))
    return tmp_path


def test_bundle_roundtrip(store):
    tpl = os.path.join(paths.TEMPLATES_DIR, "Team Report.pptx")
    open(tpl, "wb").write(b"PPTX-BYTES")
    img = store / "logo.png"
    img.write_bytes(b"PNG-BYTES")
    save_template_mapping("Team Report.pptx", {
        "slides": {"1": {"shape_mappings": {
            "7": {"type": "image", "value": str(img)},
            "8": {"type": "metric", "value": "Total Impressions"}}}}})

    dest = str(store / "bundle.zip")
    n = template_bundle.export_template_bundle("Team Report.pptx", dest)
    assert n == 1
    with zipfile.ZipFile(dest) as z:
        names = z.namelist()
        assert "manifest.json" in names
        assert "template/Team Report.pptx" in names
        assert "images/logo.png" in names
        bundled = json.loads(z.read("mapping.json"))
        assert bundled["slides"]["1"]["shape_mappings"]["7"]["value"] == \
            "{{BUNDLE}}/images/logo.png"

    # Wipe local store (simulate the teammate's machine)
    os.remove(tpl)
    os.remove(os.path.join(paths.MAPPINGS_DIR,
                           os.listdir(paths.MAPPINGS_DIR)[0]))
    img.unlink()

    name, imgs = template_bundle.import_template_bundle(dest)
    assert name == "Team Report.pptx" and imgs == 1
    assert os.path.exists(tpl)
    mapping = load_template_mapping("Team Report.pptx")
    new_img = mapping["slides"]["1"]["shape_mappings"]["7"]["value"]
    assert os.path.isfile(new_img), new_img
    assert mapping["slides"]["1"]["shape_mappings"]["8"]["value"] == \
        "Total Impressions"

    # Second import without overwrite must refuse; with overwrite succeeds
    with pytest.raises(FileExistsError):
        template_bundle.import_template_bundle(dest)
    template_bundle.import_template_bundle(dest, overwrite=True)



def test_bundle_export_resolves_relative_image_path(store, monkeypatch):
    tpl = os.path.join(paths.TEMPLATES_DIR, "Relative Report.pptx")
    open(tpl, "wb").write(b"PPTX-BYTES")
    img = store / "templates" / "images" / "relative.png"
    img.write_bytes(b"PNG-BYTES")
    monkeypatch.setattr(paths, "APP_ROOT", str(store))
    monkeypatch.setattr(paths, "WORKSPACE_DIR", str(store))
    save_template_mapping("Relative Report.pptx", {
        "slides": {"1": {"shape_mappings": {
            "2": {"image_path": "templates/images/relative.png"}}}}})

    dest = str(store / "relative_bundle.zip")
    assert template_bundle.export_template_bundle("Relative Report.pptx", dest) == 1
    with zipfile.ZipFile(dest) as z:
        assert "images/relative.png" in z.namelist()
        bundled = json.loads(z.read("mapping.json"))
        assert bundled["slides"]["1"]["shape_mappings"]["2"]["image_path"] == \
            "{{BUNDLE}}/images/relative.png"

def test_import_rejects_foreign_zip(store):
    bad = str(store / "notabundle.zip")
    with zipfile.ZipFile(bad, "w") as z:
        z.writestr("random.txt", "hi")
    with pytest.raises(ValueError):
        template_bundle.import_template_bundle(bad)


def test_import_rejects_unsafe_template_path(store):
    bad = str(store / "unsafe_bundle.zip")
    manifest = {
        "format": template_bundle.FORMAT_NAME,
        "version": 1,
        "template": "../escape.pptx",
    }
    with zipfile.ZipFile(bad, "w") as z:
        z.writestr("manifest.json", json.dumps(manifest))
        z.writestr("mapping.json", "{}")
        z.writestr("template/../escape.pptx", b"PPTX")
    with pytest.raises(ValueError):
        template_bundle.import_template_bundle(bad)


def test_import_rejects_archive_traversal(store):
    bad = str(store / "traversal_bundle.zip")
    manifest = {
        "format": template_bundle.FORMAT_NAME,
        "version": 1,
        "template": "Safe.pptx",
    }
    with zipfile.ZipFile(bad, "w") as z:
        z.writestr("manifest.json", json.dumps(manifest))
        z.writestr("mapping.json", "{}")
        z.writestr("template/Safe.pptx", b"PPTX")
        z.writestr("images/../escape.png", b"PNG")
    with pytest.raises(ValueError):
        template_bundle.import_template_bundle(bad)


def test_import_rejects_duplicate_image_basenames(store):
    bad = str(store / "duplicate_images.zip")
    manifest = {
        "format": template_bundle.FORMAT_NAME,
        "version": 1,
        "template": "Safe.pptx",
    }
    with zipfile.ZipFile(bad, "w") as z:
        z.writestr("manifest.json", json.dumps(manifest))
        z.writestr("mapping.json", "{}")
        z.writestr("template/Safe.pptx", b"PPTX")
        z.writestr("images/logo.png", b"PNG1")
        z.writestr("images/LOGO.PNG", b"PNG2")
    with pytest.raises(ValueError):
        template_bundle.import_template_bundle(bad)
