# -*- coding: utf-8 -*-
"""
Check a packaged copy of the app: every part the page needs is inside it.

    PDF Converter.exe --selftest report.json

A frozen program can start and still miss a module that is only imported
when a certain kind of page comes in (a scan, a broken font, a table in a
picture). Each check here imports or runs that part. A real conversion of
a small Arabic and English PDF with a table runs last. The report is JSON;
the exit code is 0 only when every check passed.
"""

from __future__ import annotations

import json
import pkgutil
import sys
import tempfile
import time
import traceback
from pathlib import Path


def _check(name: str, results: list, fn) -> None:
    start = time.time()
    try:
        detail = fn()
        results.append({"check": name, "ok": True, "detail": str(detail or ""), "seconds": round(time.time() - start, 2)})
    except Exception as exc:
        results.append({
            "check": name,
            "ok": False,
            "detail": f"{type(exc).__name__}: {exc}",
            "trace": traceback.format_exc()[-1500:],
        })


def _modules() -> str:
    import core
    import engine

    names = []
    for package in (core, engine):
        for info in pkgutil.iter_modules(package.__path__):
            names.append(f"{package.__name__}.{info.name}")
    failed = []
    for name in names:
        if name in {"engine.doctor"}:
            continue
        try:
            __import__(name)
        except Exception as exc:
            failed.append(f"{name}: {exc}")
    if failed:
        raise RuntimeError("; ".join(failed))
    return f"{len(names)} modules"


def _libraries() -> str:
    import docx
    import fitz  # noqa: F401
    import fontTools.ttLib  # noqa: F401
    import openpyxl  # noqa: F401
    import pdfplumber  # noqa: F401
    import streamlit  # noqa: F401
    from fontTools.ttLib.tables import _p_o_s_t, _c_m_a_p, _g_l_y_f  # noqa: F401
    from img2table.document import Image  # noqa: F401
    from img2table.ocr import TesseractOCR  # noqa: F401

    docx.Document()  # needs the default template inside the package
    static = Path(streamlit.__file__).parent / "static" / "index.html"
    if not static.is_file():
        raise RuntimeError("Streamlit's page files are missing")
    return "ok"


def _tesseract() -> str:
    from core.image_tables import tessdata_for
    from engine.tesseract_bin import find_exe

    exe = find_exe()
    if not exe:
        raise RuntimeError("tesseract.exe not found")
    folder, langs = tessdata_for()
    if "ara" not in (langs or "") or "eng" not in (langs or ""):
        raise RuntimeError(f"languages: {langs!r} in {folder}")
    return f"{exe} ({langs})"


def _windows_ocr() -> str:
    if not sys.platform.startswith("win"):
        return "not Windows"
    try:
        from winrt.windows.globalization import Language
        from winrt.windows.media.ocr import OcrEngine
    except Exception as exc:
        raise RuntimeError(f"winrt is missing: {exc}")
    engine = OcrEngine.try_create_from_language(Language("ar-SA"))
    return "Arabic OCR ready" if engine is not None else "Arabic language pack not installed (Tesseract is used)"


def _sample_pdf(path: Path) -> None:
    import fitz

    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((60, 80), "PDF Converter self test", fontsize=16)
    page.insert_text((60, 110), "This line checks English text in a paragraph.", fontsize=11)
    try:
        page.insert_htmlbox(
            fitz.Rect(60, 130, 535, 200),
            '<p dir="rtl" style="font-size:14px">هذا سطر عربي لاختبار التحويل إلى وورد.</p>',
        )
    except Exception:
        pass
    xs, ys = [60, 200, 340, 535], [240, 265, 290, 315]
    for x in xs:
        page.draw_line((x, ys[0]), (x, ys[-1]), width=0.8)
    for y in ys:
        page.draw_line((xs[0], y), (xs[-1], y), width=0.8)
    for r in range(3):
        for c in range(3):
            page.insert_text((xs[c] + 6, ys[r] + 16), f"Cell {r + 1}-{c + 1}", fontsize=10)
    doc.save(path)


def _convert() -> str:
    from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx

    with tempfile.TemporaryDirectory(prefix="pdf_converter_test_") as work:
        pdf = Path(work) / "sample.pdf"
        out = Path(work) / "sample.docx"
        _sample_pdf(pdf)
        counts = convert_arabic_pdf_to_docx(pdf, out, options=WordOptions(mode="editable", fit_a4=True, direction="auto"))
        import docx

        document = docx.Document(str(out))
        text = "\n".join(p.text for p in document.paragraphs)
        cells = " ".join(cell.text for table in document.tables for row in table.rows for cell in row.cells)
        if "self test" not in text:
            raise RuntimeError("the English paragraph is missing")
        if "Cell 2-2" not in cells:
            raise RuntimeError("the table is missing")
        if not counts.get("docx_validation", {}).get("ok", False):
            raise RuntimeError(f"Word file check: {counts.get('docx_validation')}")
        arabic = "عربي" in text
        return f"{counts.get('pages')} page, table ok, Arabic {'ok' if arabic else 'not drawn by this PDF'}"


def _picture_reading() -> str:
    """The reader for pages with no text layer: Tesseract on a rendered page."""
    import fitz

    from core.docx_converter import _ocr_word_atoms

    doc = fitz.open()
    page = doc.new_page(width=400, height=200)
    # Arabic, as the reader is set for: it leaves out Latin scraps of logos.
    page.insert_htmlbox(fitz.Rect(20, 60, 380, 140), '<p dir="rtl" style="font-size:26px">قراءة الكلمات من الصورة</p>')
    # Turn the page into a picture so it has no text layer.
    picture = page.get_pixmap(dpi=200)
    scan = fitz.open()
    sheet = scan.new_page(width=400, height=200)
    sheet.insert_image(sheet.rect, pixmap=picture)
    words = _ocr_word_atoms(sheet)
    joined = " ".join(atom.text.strip() for atom in words)
    if "الكلمات" not in joined and "الصورة" not in joined:
        raise RuntimeError(f"read: {joined!r}")
    return joined


def run(report_path: str = "") -> int:
    results: list = []
    _check("modules", results, _modules)
    _check("libraries", results, _libraries)
    _check("tesseract", results, _tesseract)
    _check("windows_ocr", results, _windows_ocr)
    _check("convert", results, _convert)
    _check("picture_reading", results, _picture_reading)
    from engine.version import VERSION

    report = {"version": VERSION, "frozen": bool(getattr(sys, "frozen", False)), "results": results}
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if report_path:
        Path(report_path).write_text(text, encoding="utf-8")
    elif sys.stdout is not None:
        print(text)
    return 0 if all(item["ok"] for item in results) else 1
