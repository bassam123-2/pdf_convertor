# -*- coding: utf-8 -*-
"""Which OCR engines can run on this PC. Nothing is sent off the machine."""

from __future__ import annotations

import subprocess

from engine.tesseract_bin import ara_tessdata_dir, find_exe

_CACHE: tuple[str, str] | None = None


def tesseract_status() -> tuple[bool, str]:
    exe = find_exe()
    data = ara_tessdata_dir()
    if not exe:
        return False, "Tesseract غير موجود."
    if data is None:
        return False, "نموذج العربية ara غير موجود بجانب التطبيق."
    try:
        result = subprocess.run(
            [exe, "--tessdata-dir", str(data), "--list-langs"],
            capture_output=True,
            timeout=40,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return False, f"تعذّر تشغيل Tesseract ({exc})."
    text = (result.stdout or b"").decode("utf-8", "replace")
    text += "\n" + (result.stderr or b"").decode("utf-8", "replace")
    langs = {line.strip() for line in text.splitlines()}
    if "ara" not in langs:
        return False, "Tesseract يعمل، لكن نموذج ara غير محمّل."
    return True, "Tesseract مع نموذج العربية ara جاهز، ومن غير اتصال بالإنترنت."


def windows_ocr_status() -> tuple[bool, str]:
    try:
        from winrt.windows.globalization import Language
        from winrt.windows.media.ocr import OcrEngine

        engine = OcrEngine.try_create_from_language(Language("ar-SA"))
    except Exception:
        return False, "قراءة ويندوز غير متاحة على هذا الجهاز."
    if engine is None:
        return (
            False,
            "حزمة العربية لويندوز غير مثبتة. لا يحتاج التطبيق إليها: Tesseract ara يقرأ الصورة.",
        )
    return True, "قراءة ويندوز للعربية متاحة كمحرك إضافي عندما تكون مثبتة."


def describe(refresh: bool = False) -> tuple[str, str]:
    """ok, warn, or error, plus one sentence for the window."""
    global _CACHE
    if _CACHE is not None and not refresh:
        return _CACHE
    tess_ok, tess_msg = tesseract_status()
    win_ok, win_msg = windows_ocr_status()
    if tess_ok and win_ok:
        level = "ok"
        text = tess_msg + " " + win_msg
    elif tess_ok:
        level = "warn"
        text = tess_msg + " " + win_msg
    elif win_ok:
        level = "warn"
        text = tess_msg + " القراءة تستخدم ويندوز ما دام يعمل."
    else:
        level = "error"
        text = "لا محرك لقراءة الصفحات المصورة. " + tess_msg + " " + win_msg
    _CACHE = (level, text)
    return _CACHE
