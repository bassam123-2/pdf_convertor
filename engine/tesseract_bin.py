# -*- coding: utf-8 -*-
"""Locate Tesseract and the bundled Arabic model. This does not read a page."""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path


def find_exe() -> str | None:
    """
    The installed search the converter already used.

    A packaged app prefers the Tesseract copied into the bundle, so a PC
    without a separate install still has the same program.
    """
    override = os.environ.get("PDF_TOOL_TESSERACT", "").strip()
    if override and Path(override).is_file():
        return override
    if getattr(sys, "frozen", False):
        bundled = _beside_app("tesseract.exe")
        if bundled is not None:
            return str(bundled)
    for candidate in (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ):
        if Path(candidate).is_file():
            return candidate
    bundled = _beside_app("tesseract.exe")
    if bundled is not None:
        return str(bundled)
    return shutil.which("tesseract")


def legacy_tessdata_dir(module_file: str) -> Path:
    """The directory the converter already checks: next to that module."""
    return Path(module_file).resolve().parent / "tessdata"


def ara_tessdata_dir() -> Path | None:
    """A tessdata folder that contains ara.traineddata, for the offline fallback."""
    candidates: list[Path] = []
    override = os.environ.get("PDF_TOOL_TESSDATA", "").strip()
    if override:
        candidates.append(Path(override))
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        root = Path(meipass)
        candidates.append(root / "tesseract" / "tessdata")
        candidates.append(root / "tessdata")
    if getattr(sys, "frozen", False):
        internal = Path(sys.executable).resolve().parent / "_internal"
        candidates.append(internal / "tesseract" / "tessdata")
        candidates.append(internal / "tessdata")
    candidates.append(_vendor_tessdata())
    candidates.append(Path(__file__).resolve().parent.parent / "tessdata")
    for folder in candidates:
        if folder is not None and (folder / "ara.traineddata").is_file():
            return folder
    return None


def _vendor_tessdata() -> Path:
    return Path(__file__).resolve().parent.parent / "vendor" / "tesseract" / "tessdata"


def _beside_app(name: str) -> Path | None:
    roots: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        roots.append(Path(meipass) / "tesseract")
    if getattr(sys, "frozen", False):
        roots.append(Path(sys.executable).resolve().parent / "_internal" / "tesseract")
    roots.append(Path(__file__).resolve().parent.parent / "vendor" / "tesseract")
    for root in roots:
        candidate = root / name
        if candidate.is_file():
            return candidate
    return None
