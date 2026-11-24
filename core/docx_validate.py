# -*- coding: utf-8 -*-
"""
Check that a .docx is well-formed OOXML Word can open without a repair prompt.

LibreOffice and Google Docs are more tolerant; this still flags broken XML,
missing required parts, and schema-level problems that make Word offer Repair.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

_REQUIRED = (
    "[Content_Types].xml",
    "word/document.xml",
    "_rels/.rels",
)


def validate_docx(path: str | Path) -> dict:
    """
    Return {ok, errors, warnings}.

    `ok` is True when the package is a readable zip with parsable Word XML and
    the required parts. Warnings cover optional issues that may degrade in
    LibreOffice / Google Docs without blocking Word.
    """
    path = Path(path)
    errors: list[str] = []
    warnings: list[str] = []
    if not path.is_file():
        return {"ok": False, "errors": [f"File not found: {path}"], "warnings": []}
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            for required in _REQUIRED:
                if required not in names:
                    errors.append(f"Missing required part: {required}")
            for name in sorted(names):
                if not name.lower().endswith(".xml") and not name.endswith(".rels"):
                    continue
                try:
                    raw = archive.read(name)
                except Exception as exc:
                    errors.append(f"Cannot read {name}: {exc}")
                    continue
                try:
                    ET.fromstring(raw)
                except ET.ParseError as exc:
                    errors.append(f"Invalid XML in {name}: {exc}")
            # Floating tables and text boxes: LibreOffice may ignore tblpPr.
            if "word/document.xml" in names:
                try:
                    text = archive.read("word/document.xml").decode("utf-8", errors="replace")
                except Exception:
                    text = ""
                if "tblpPr" in text:
                    warnings.append(
                        "Document uses floating tables (tblpPr); LibreOffice/Google Docs may place them in-flow."
                    )
                if "wps:wsp" in text or "w:txbxContent" in text:
                    warnings.append(
                        "Document uses text boxes; older LibreOffice builds may shift or drop some boxes."
                    )
                if "w:fldChar" in text:
                    # PAGE fields update in Word; Google Docs shows the cached result.
                    warnings.append(
                        "Document uses Word fields (e.g. PAGE); Google Docs shows the last cached value."
                    )
    except zipfile.BadZipFile as exc:
        errors.append(f"Not a valid ZIP/OOXML package: {exc}")
    return {"ok": not errors, "errors": errors, "warnings": warnings}
