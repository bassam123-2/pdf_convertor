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


def main() -> int:
    for line in report():
        print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
