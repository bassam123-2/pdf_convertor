# -*- mode: python ; coding: utf-8 -*-
"""One-folder package. Another Windows PC does not need Python or a Tesseract install."""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules, copy_metadata

ROOT = Path(SPECPATH)
VENDOR = ROOT / "vendor" / "tesseract"
SYSTEM_TESS = Path(r"C:\Program Files\Tesseract-OCR")


def _stage_tesseract() -> None:
    """Copy the local Tesseract program and the Arabic model into vendor/."""
    import shutil

    if not (SYSTEM_TESS / "tesseract.exe").is_file():
        raise SystemExit("Install Tesseract once on this build PC so it can be copied into the app.")
    tessdata = VENDOR / "tessdata"
    tessdata.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SYSTEM_TESS / "tesseract.exe", VENDOR / "tesseract.exe")
    for dll in SYSTEM_TESS.glob("*.dll"):
        shutil.copy2(dll, VENDOR / dll.name)
    for name in ("eng.traineddata", "osd.traineddata"):
        source = SYSTEM_TESS / "tessdata" / name
        if source.is_file():
            shutil.copy2(source, tessdata / name)
    # The plate re-read ends with the config name "tsv". Tesseract loads that
    # from tessdata/configs. Traineddata alone makes the re-read return nothing.
    for folder in ("configs", "tessconfigs"):
        source = SYSTEM_TESS / "tessdata" / folder
        if source.is_dir():
            shutil.copytree(source, tessdata / folder, dirs_exist_ok=True)
    if not (tessdata / "configs" / "tsv").is_file():
        raise SystemExit("Tesseract tessdata/configs/tsv was not copied.")
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
    (str(VENDOR / "tessdata"), "tesseract/tessdata"),
    ("packaging/UPDATE.txt", "."),
]
binaries = [(str(path), "tesseract") for path in VENDOR.glob("*.dll")]
binaries.append((str(VENDOR / "tesseract.exe"), "tesseract"))
hiddenimports = [
    "engine",
    "engine.host",
    "engine.paths",
    "engine.resources",
    "engine.tesseract_bin",
    "engine.ocr_status",
    "engine.benchmark",
    "engine.checkpoint",
    "engine.offline",
    "engine.version",
    "excel_review",
    "page_images",
    "desktop_review",
    "core",
    "core.pdf_loader",
    "core.arabic_fixer",
    "core.docx_converter",
    "core.layout_check",
    "core.html_layout",
    "core.excel_converter",
    "utils",
    "utils.file_manager",
    "fitz",
    "pdfplumber",
    "pdfminer",
    "pdfminer.high_level",
    "docx",
    "openpyxl",
    "pandas",
    "PIL",
    "altair",
    "pyarrow",
    "streamlit",
    "multiprocessing",
    "winrt",
    "winrt.windows.media.ocr",
    "winrt.windows.graphics.imaging",
    "winrt.windows.storage.streams",
    "winrt.windows.globalization",
    "winrt.windows.foundation",
    "winrt.windows.foundation.collections",
]

datas += collect_data_files("streamlit")
datas += collect_data_files("altair")
binaries += collect_dynamic_libs("streamlit")
hiddenimports += [
    name
    for name in collect_submodules("streamlit")
    if ".tests" not in name and ".testing" not in name
]
hiddenimports += [
    name
    for name in collect_submodules("winrt")
    if ".tests" not in name
]

for distribution in (
    "streamlit",
    "altair",
    "pandas",
    "pyarrow",
    "pillow",
    "pymupdf",
    "pdfplumber",
    "python-docx",
    "openpyxl",
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
    excludes=["pytest", "IPython"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PDF Tool",
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
    name="PDF Tool",
)
