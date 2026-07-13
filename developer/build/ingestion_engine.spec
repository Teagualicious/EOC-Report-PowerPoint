# PyInstaller one-folder build for the cleaned project layout.
#
# Run from the project root:
#   python -m pip install pyinstaller
#   python -m PyInstaller developer/build/ingestion_engine.spec --noconfirm
#
# Ship the entire dist/IngestionEngine folder. On first launch the app creates
# workspace/ beside the executable. Copy the source workspace/templates and
# workspace/mappings folders into the distribution when the stand-in template
# should be included.

import os

PROJECT_ROOT = os.path.abspath(os.path.join(SPECPATH, "..", ".."))
APP_DIR = os.path.join(PROJECT_ROOT, "app")
RESOURCE_DIR = os.path.join(APP_DIR, "resources")


a = Analysis(
    [os.path.join(APP_DIR, "main.py")],
    pathex=[APP_DIR],
    binaries=[],
    datas=[
        (os.path.join(RESOURCE_DIR, "metric_dictionary.json"), "resources"),
        (os.path.join(RESOURCE_DIR, "assets"), "resources/assets"),
    ],
    hiddenimports=[
        "win32com", "win32com.client", "pythoncom", "win32timezone",
        "tkcalendar", "babel.numbers", "PIL._tkinter_finder",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["pytest"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="IngestionEngine",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="IngestionEngine",
)
