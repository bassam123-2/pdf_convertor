# -*- mode: python ; coding: utf-8 -*-
"""
PDF Converter for Windows: one folder that runs on a PC with no Python and
no Tesseract installed.

    pyinstaller PDFTool.spec --noconfirm

Builds dist/PDF Converter/PDF Converter.exe. packaging/PDFTool.iss then
turns that folder into one installer. packaging/build.bat runs both, and
the "Windows app" workflow on GitHub does the same on every push.

Tesseract is copied from its usual install folder on the build PC, with the
English model from there and the Arabic model from tessdata/ in this repo.
"""

import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules, copy_metadata

ROOT = Path(SPECPATH)
VENDOR = ROOT / "vendor" / "tesseract"
SYSTEM_TESS = Path(os.environ.get("TESSERACT_HOME", r"C:\Program Files\Tesseract-OCR"))


def _stage_tesseract() -> None:
    """Copy the build PC's Tesseract program and the models into vendor/."""
    import shutil

    if not (SYSTEM_TESS / "tesseract.exe").is_file():
        raise SystemExit(f"Install Tesseract once on this build PC ({SYSTEM_TESS}) so it can be copied into the app.")
    tessdata = VENDOR / "tessdata"
    tessdata.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SYSTEM_TESS / "tesseract.exe", VENDOR / "tesseract.exe")
    for dll in SYSTEM_TESS.glob("*.dll"):
        shutil.copy2(dll, VENDOR / dll.name)
    for name in ("eng.traineddata", "osd.traineddata"):
        source = SYSTEM_TESS / "tessdata" / name
        if source.is_file():
            shutil.copy2(source, tessdata / name)
    # "tsv" output is a config file in tessdata/configs; without it the word
    # boxes come back empty.
    for folder in ("configs", "tessconfigs"):
        source = SYSTEM_TESS / "tessdata" / folder
        if source.is_dir():
            shutil.copytree(source, tessdata / folder, dirs_exist_ok=True)
    if not (tessdata / "configs" / "tsv").is_file():
        shutil.copytree(ROOT / "tessdata" / "configs", tessdata / "configs", dirs_exist_ok=True)
    if not (tessdata / "eng.traineddata").is_file():
        raise SystemExit("eng.traineddata was not found in the Tesseract install.")
    ara = ROOT / "tessdata" / "ara.traineddata"
    if not ara.is_file():
        raise SystemExit("tessdata/ara.traineddata is missing.")
    shutil.copy2(ara, tessdata / "ara.traineddata")


_stage_tesseract()

datas = [
    ("app.py", "."),
    ("excel_review.py", "."),
    ("page_images.py", "."),
    ("desktop_review.py", "."),
    ("arabic_pdf_to_docx.py", "."),
    ("pdf_to_excel.py", "."),
    ("core", "core"),
    ("utils", "utils"),
    ("engine", "engine"),
    (".streamlit", ".streamlit"),
    ("tessdata", "tessdata"),
    ("packaging/icon.png", "packaging"),
    ("packaging/icon.ico", "packaging"),
    ("packaging/UPDATE.txt", "."),
    (str(VENDOR / "tessdata"), "tesseract/tessdata"),
]
binaries = [(str(path), "tesseract") for path in VENDOR.glob("*.dll")]
binaries.append((str(VENDOR / "tesseract.exe"), "tesseract"))

# Modules the page imports only when a certain kind of PDF comes in (a
# scan, a broken font, a table in a picture): named here, or the frozen app
# would miss them on that page alone.
hiddenimports = ["excel_review", "page_images", "desktop_review", "multiprocessing", "tkinter", "tkinter.ttk", "tkinter.messagebox"]
for package in ("core", "engine", "utils"):
    hiddenimports += collect_submodules(package)
for package in ("fontTools", "img2table", "pdfplumber", "pdfminer", "docx", "openpyxl", "thefuzz", "rapidfuzz", "winrt"):
    try:
        hiddenimports += [name for name in collect_submodules(package) if ".tests" not in name]
    except Exception:
        pass
hiddenimports += [
    name for name in collect_submodules("streamlit")
    if ".tests" not in name and ".testing" not in name
]

# Data files libraries read at run time: Word's blank template, the PDF
# character maps, Streamlit's page, the fonts' metric tables.
for package in ("streamlit", "altair", "docx", "pdfminer", "fontTools", "img2table", "openpyxl", "pypdfium2_raw"):
    try:
        datas += collect_data_files(package)
    except Exception:
        pass
# python-docx finds its header and footer templates as "parts/../templates";
# the "parts" folder must exist on disk for that path to resolve.
datas += collect_data_files("docx", include_py_files=True)
binaries += collect_dynamic_libs("streamlit")
for package in ("cv2", "pypdfium2", "pypdfium2_raw"):
    try:
        binaries += collect_dynamic_libs(package)
    except Exception:
        pass

for distribution in (
    "streamlit", "altair", "pandas", "pyarrow", "pillow", "pymupdf", "pdfplumber", "pdfminer.six",
    "python-docx", "openpyxl", "fonttools", "img2table", "opencv-contrib-python", "pypdfium2",
    "numpy", "thefuzz", "rapidfuzz",
):
    try:
        datas += copy_metadata(distribution)
    except Exception:
        pass

a = Analysis(
    ["desktop_app.pyw"],
    pathex=[str(ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=["packaging/runtime_hook.py"],
    excludes=["pytest", "IPython", "matplotlib.tests", "numpy.tests"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PDF Converter",
    icon="packaging/icon.ico",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version="packaging/file_version_info.txt",
    contents_directory="_internal",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="PDF Converter",
)
