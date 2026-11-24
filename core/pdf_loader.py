# -*- coding: utf-8 -*-
"""Open a PDF, tell a text page from a scan, and hand pages out in small groups.

Word and Excel both use this module. Neither keeps the whole file's pages
alive at once. A group is released, and its memory is collected, before the
next group is read.
"""

from __future__ import annotations

import gc
from pathlib import Path

import fitz

# Five pages is enough to recognise a repeated header, and small enough that
# a long file does not keep every page image in memory.
PAGE_BATCH = 5


def open_document(path: str | Path, password: str = "") -> fitz.Document:
    """
    Open one PDF.

    FileNotFoundError means the path is missing. PermissionError means the
    file asked for a password and the one given did not open it. RuntimeError
    means the file could not be read.
    """
    source = Path(path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(source)
    try:
        document = fitz.open(source)
    except Exception as exc:
        raise RuntimeError(str(exc)) from exc
    if document.needs_pass and not document.authenticate(password or ""):
        document.close()
        raise PermissionError("password")
    return document


def close_document(document: fitz.Document | None) -> None:
    """Close the PDF and return its pages to the system."""
    if document is None:
        return
    try:
        document.close()
    except Exception:
        pass
    gc.collect()


def text_layer_count(page: fitz.Page) -> int:
    """How many letters the page stores. A scan of paper returns zero."""
    try:
        text = page.get_text("text") or ""
    except Exception:
        return 0
    count = 0
    for char in text:
        code = ord(char)
        if char.isascii() and char.isalpha():
            count += 1
        elif (
            0x0600 <= code <= 0x06FF
            or 0x0750 <= code <= 0x077F
            or 0x08A0 <= code <= 0x08FF
            or 0xFB50 <= code <= 0xFDFF
            or 0xFE70 <= code <= 0xFEFF
        ):
            count += 1
    return count


def has_text_layer(page: fitz.Page, minimum: int = 40) -> bool:
    """True when the page has a real text layer. False means it should be read as a picture."""
    return text_layer_count(page) >= minimum


def iter_page_batches(document: fitz.Document, size: int = PAGE_BATCH):
    """
    Yield (first_index, pages) for each group.

    The pages of a group are dropped, and garbage-collected, as soon as the
    caller finishes that group. The next group is not loaded before then.
    """
    step = max(1, int(size))
    total = document.page_count
    for start in range(0, total, step):
        end = min(start + step, total)
        pages: list[fitz.Page] = []
        try:
            for index in range(start, end):
                pages.append(document.load_page(index))
            yield start, pages
        finally:
            pages.clear()
            gc.collect()
