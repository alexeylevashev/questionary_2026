# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for transport survey CLI tool.
Backend: geopandas + pyogrio (no fiona).
"""

import os
from pathlib import Path

SITE = Path("C:/Users/user/AppData/Local/Programs/Python/Python311/Lib/site-packages")

block_cipher = None

# ---------------------------------------------------------------------------
# Data files to bundle
# ---------------------------------------------------------------------------
datas = [
    # Project config and data folder
    ("config.yaml", "."),
    ("data", "data"),

    # pyogrio bundled GDAL + PROJ data
    (str(SITE / "pyogrio" / "gdal_data"), "pyogrio/gdal_data"),
    (str(SITE / "pyogrio" / "proj_data"), "pyogrio/proj_data"),

    # pyproj PROJ database
    (str(SITE / "pyproj" / "proj_dir" / "share" / "proj"), "pyproj/proj_dir/share/proj"),

    # geopandas datasets (small, needed for some internals)
    (str(SITE / "geopandas" / "datasets"), "geopandas/datasets"),
]

# ---------------------------------------------------------------------------
# Hidden imports (not auto-detected by PyInstaller)
# ---------------------------------------------------------------------------
hiddenimports = [
    # pyogrio
    "pyogrio",
    "pyogrio._ogr",
    "pyogrio._geometry",
    "pyogrio.raw",

    # pyproj
    "pyproj",
    "pyproj.transformer",
    "pyproj.crs",

    # shapely
    "shapely",
    "shapely.geometry",
    "shapely.geometry.point",
    "shapely.geometry.linestring",
    "shapely.geometry.polygon",
    "shapely.ops",
    "shapely.prepared",
    "shapely.strtree",

    # geopandas
    "geopandas",
    "geopandas.io.file",
    "geopandas.io.arrow",
    "geopandas._compat",

    # pandas / numpy
    "pandas",
    "pandas._libs.tslibs.timedeltas",
    "pandas._libs.tslibs.np_datetime",
    "pandas._libs.tslibs.nattype",
    "numpy",

    # scipy
    "scipy",
    "scipy.optimize",
    "scipy.optimize._minpack_py",
    "scipy.optimize._zeros_py",
    "scipy.special",
    "scipy.special._ufuncs",
    "scipy.linalg",
    "scipy.linalg._decomp",
    "scipy._lib.messagestream",

    # openpyxl
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.chart",
    "openpyxl.chart.label",
    "openpyxl.utils",

    # yaml
    "yaml",

    # src package
    "src",
    "src.cli",
    "src.config",
    "src.io_utils",
    "src.coords",
    "src.filters",
    "src.status",
    "src.od_matrix",
    "src.eva",
    "src.excel_report",
    "src.export_gis",
    "src.qgis_project",
]

# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------
a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # GUI toolkits not needed
        "tkinter", "matplotlib", "PIL", "Pillow",
        "IPython", "jupyter", "notebook",
        "PyQt5", "PyQt6", "PySide2", "PySide6",
        "wx", "gi",
        # Heavy ML packages — not used in this project
        "torch", "torchvision", "torchaudio",
        "numba", "llvmlite",
        "tensorflow", "keras",
        "onnxruntime", "onnx",
        "sklearn", "scikit_learn",
        "cv2", "imageio",
        "sympy",
        "numexpr",
        "bottleneck",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="questionary",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,          # console app — window stays open
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="questionary",
)
