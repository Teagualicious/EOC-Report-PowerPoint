"""Import/export mapped PowerPoint templates as portable ZIP bundles."""

from __future__ import annotations

import json
import logging
import os
import shutil
import tempfile
import zipfile
from datetime import datetime

from config.naming import safe_component

logger = logging.getLogger(__name__)

_IMG_EXTS = (".png", ".jpg", ".jpeg", ".gif", ".bmp")
_TOKEN = "{{BUNDLE}}/images/"
FORMAT_NAME = "sr-template-bundle"
_MAX_FILES = 500
_MAX_UNCOMPRESSED = 500 * 1024 * 1024


def _walk_rewrite(obj, fn):
    """Recursively rewrite string values via ``fn(value)``."""
    if isinstance(obj, dict):
        return {k: _walk_rewrite(v, fn) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_walk_rewrite(v, fn) for v in obj]
    if isinstance(obj, str):
        new = fn(obj)
        return obj if new is None else new
    return obj


def _validate_template_name(name):
    if not isinstance(name, str) or not name.lower().endswith(".pptx"):
        raise ValueError("Bundle template must be a .pptx file.")
    if name != os.path.basename(name) or safe_component(name) != name:
        raise ValueError("Bundle contains an unsafe template filename.")
    return name


def _validate_archive(z):
    infos = z.infolist()
    if len(infos) > _MAX_FILES:
        raise ValueError("Template bundle contains too many files.")
    total = sum(i.file_size for i in infos)
    if total > _MAX_UNCOMPRESSED:
        raise ValueError("Template bundle is too large when extracted.")
    for info in infos:
        normalized = info.filename.replace("\\", "/")
        parts = [p for p in normalized.split("/") if p]
        if normalized.startswith("/") or any(p in (".", "..") for p in parts):
            raise ValueError("Template bundle contains an unsafe archive path.")


def export_template_bundle(template_filename, dest_zip):
    """Package a template, mapping, and referenced images into ``dest_zip``."""
    from config import paths
    from engine.pptx_mapper import load_template_mapping

    template_filename = _validate_template_name(template_filename)
    src_pptx = os.path.join(paths.TEMPLATES_DIR, template_filename)
    if not os.path.isfile(src_pptx):
        raise ValueError(f"Template not found: {template_filename}")
    mapping = load_template_mapping(template_filename)
    if mapping is None:
        raise ValueError(f"No saved mapping for {template_filename} — map it before exporting.")

    images = {}

    def collect(value):
        if not value.lower().endswith(_IMG_EXTS):
            return None

        source = value
        if not os.path.isabs(source):
            source = os.path.join(paths.APP_ROOT, source)
        if not os.path.isfile(source) and not os.path.isabs(value):
            # Compatibility with older mappings whose path began at the
            # former writable root, e.g. templates/images/logo.png.
            source = os.path.join(paths.WORKSPACE_DIR, value)
        if not os.path.isfile(source):
            return None

        source = os.path.abspath(source)
        base = safe_component(os.path.basename(source), fallback="image.png")
        stem, ext = os.path.splitext(base)
        candidate = base
        i = 2
        existing = {name.casefold() for name in images.values()}
        while candidate.casefold() in existing and images.get(source) != candidate:
            candidate = f"{stem}_{i}{ext}"
            i += 1
        images[source] = candidate
        return _TOKEN + candidate

    bundled_mapping = _walk_rewrite(mapping, collect)
    manifest = {
        "format": FORMAT_NAME,
        "version": 1,
        "template": template_filename,
        "exported": datetime.now().isoformat(timespec="seconds"),
        "images": len(images),
    }
    with zipfile.ZipFile(dest_zip, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("manifest.json", json.dumps(manifest, indent=2))
        z.write(src_pptx, f"template/{template_filename}")
        z.writestr("mapping.json", json.dumps(bundled_mapping, indent=2))
        for src, base in images.items():
            z.write(src, f"images/{base}")
    logger.info("Exported template bundle: %s (%d images)", dest_zip, len(images))
    return len(images)


def import_template_bundle(zip_path, overwrite=False):
    """Install a validated bundle and return ``(template_name, image_count)``."""
    from config import paths
    from engine.pptx_mapper import save_template_mapping

    with zipfile.ZipFile(zip_path) as z:
        _validate_archive(z)
        names = set(z.namelist())
        if "manifest.json" not in names or "mapping.json" not in names:
            raise ValueError("Not a template bundle (manifest/mapping missing).")

        try:
            manifest = json.loads(z.read("manifest.json"))
            mapping = json.loads(z.read("mapping.json"))
        except (json.JSONDecodeError, UnicodeDecodeError, KeyError) as exc:
            raise ValueError("Template bundle metadata is invalid.") from exc

        if manifest.get("format") != FORMAT_NAME or manifest.get("version") != 1:
            raise ValueError("Unsupported Spectrum Reach template bundle.")
        tpl_name = _validate_template_name(manifest.get("template"))
        tpl_arc = f"template/{tpl_name}"
        if tpl_arc not in names:
            raise ValueError("Bundle is missing its template file.")

        dest_pptx = os.path.join(paths.TEMPLATES_DIR, tpl_name)
        if os.path.exists(dest_pptx) and not overwrite:
            raise FileExistsError(tpl_name)

        os.makedirs(paths.TEMPLATES_DIR, exist_ok=True)
        stem = safe_component(os.path.splitext(tpl_name)[0]).replace(" ", "_")
        img_dir = os.path.join(paths.IMAGES_DIR, stem)

        extracted_images = {}
        extracted_names = set()
        os.makedirs(img_dir, exist_ok=True)
        for name in sorted(names):
            if not name.startswith("images/") or name.endswith("/"):
                continue
            base = safe_component(os.path.basename(name), fallback="image")
            if not base.lower().endswith(_IMG_EXTS):
                continue
            folded = base.casefold()
            if folded in extracted_names:
                raise ValueError(f"Bundle contains duplicate image names: {base}")
            extracted_names.add(folded)
            target = os.path.join(img_dir, base)
            with z.open(name) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)
            extracted_images[base] = target

        def restore(value):
            if value.startswith(_TOKEN):
                base = safe_component(os.path.basename(value[len(_TOKEN):]), fallback="image")
                target = extracted_images.get(base)
                if target is None:
                    raise ValueError(f"Bundle mapping references a missing image: {base}")
                return target
            return None

        mapping = _walk_rewrite(mapping, restore)

        fd, temp_pptx = tempfile.mkstemp(prefix=".template_", suffix=".pptx",
                                         dir=paths.TEMPLATES_DIR)
        os.close(fd)
        try:
            with z.open(tpl_arc) as src, open(temp_pptx, "wb") as out:
                shutil.copyfileobj(src, out)
            os.replace(temp_pptx, dest_pptx)
        except Exception:
            try:
                os.remove(temp_pptx)
            except OSError:
                pass
            raise

        save_template_mapping(tpl_name, mapping)

    logger.info("Imported template bundle: %s (%d images)", tpl_name, len(extracted_images))
    return tpl_name, len(extracted_images)
