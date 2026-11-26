# -*- coding: utf-8 -*-
"""
Check what this machine gives the converter, in Arabic and English.

    python -m engine.doctor

Each line says whether a part is there and what is lost without it.
"""

from __future__ import annotations

import sys


def _line(ok: bool, name: str, lost: str) -> str:
    mark = "OK " if ok else "--"
    tail = "" if ok else f"   ← بدونه: {lost}"
    return f"[{mark}] {name}{tail}"


def report() -> list[str]:
    lines = []
    try:
        import fontTools  # noqa: F401

        fonttools = True
    except Exception:
        fonttools = False
    lines.append(_line(fonttools, "fonttools (قراءة الحروف من شكلها)", "الخطوط المعطوبة تُقرأ بالـ OCR فقط"))

    try:
        import img2table  # noqa: F401

        tables = True
    except Exception:
        tables = False
    lines.append(_line(tables, "img2table (جداول الصور)", "الجداول التي هي صور تبقى صوراً"))

    try:
        from engine.tesseract_bin import find_exe

        exe = find_exe()
    except Exception:
        exe = None
    lines.append(_line(bool(exe), f"Tesseract ({exe or 'غير موجود'})", "لا تُقرأ خلايا جداول الصور"))
    if exe:
        try:
            from core.image_tables import tessdata_for

            folder, langs = tessdata_for()
        except Exception:
            folder, langs = None, ""
        have = set(langs.split("+")) if langs else set()
        lines.append(_line("ara" in have, "Tesseract: العربية (ara)", "جداول الصور العربية لا تُقرأ"))
        lines.append(_line("eng" in have, "Tesseract: الإنجليزية (eng)", "جداول الصور الإنجليزية لا تُقرأ"))

    windows_ocr = False
    if sys.platform.startswith("win"):
        try:
            from winrt.windows.media.ocr import OcrEngine  # noqa: F401
            from winrt.windows.globalization import Language

            windows_ocr = OcrEngine.try_create_from_language(Language("ar-SA")) is not None
        except Exception:
            windows_ocr = False
    lines.append(_line(windows_ocr, "Windows OCR العربي", "الصفحات الممسوحة والخطوط المرسومة لا تُقرأ نصاً"))

    try:
        from core.font_repair import _candidates

        for family in ("SakkalMajalla", "TimesNewRomanPSMT", "ArialMT", "Calibri"):
            lines.append(_line(bool(_candidates(family)), f"خط مثبت: {family}", "الحروف المعطوبة بهذا الخط لا تُصلح من شكلها"))
    except Exception:
        pass
    return lines


def explain(pdf: str, pages: str) -> list[str]:
    """How the editable reader takes each page, and why."""
    from core.docx_converter import WordOptions, read_word_build

    build = read_word_build(pdf, options=WordOptions(mode="editable", fit_a4=True, direction="auto", page_spec=pages))
    lines = []
    for page in build.pages:
        kinds: dict[str, int] = {}
        for block in page.blocks:
            name = type(block).__name__
            kinds[name] = kinds.get(name, 0) + 1
        how = "faithful" if page.faithful is not None else ("positioned" if page.frames else "flow")
        lines.append(f"page {page.number}: {how} ({page.kind}) {kinds}")
    for note in build.counts.get("retries", []):
        lines.append(f"  note: {note}")
    for note in build.counts.get("review", []):
        lines.append(f"  review: {note}")
    return lines


def main() -> int:
    for line in report():
        print(line)
    # python -m engine.doctor file.pdf 4   (or 3-5)
    if len(sys.argv) >= 3:
        print()
        for line in explain(sys.argv[1], sys.argv[2]):
            print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
