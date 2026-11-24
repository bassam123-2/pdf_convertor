# -*- coding: utf-8 -*-
"""
Offline PDF to Excel converter.

The PDF stays on this computer. Nothing is uploaded. A digital page is read
with pdfplumber. Vertical and horizontal snap tolerances are tight, so a
plate, a colour, and a date stay in their own columns. PyMuPDF keeps the
table when its rows are already complete. A scan is only a picture, so
Windows OCR reads it on this PC. Photos, charts, and diagrams are covered
before that read and stored on a Figures sheet. When Tesseract is installed,
it re-reads the Arabic words on the same picture.

Each page is checked on its own. With strict reading, a blank cell stays
blank: a plate is not invented from another column, and a make, colour, or
year is not copied from the row above. A colour that is actually printed is
still rewritten from a fixed Arabic list. Arabic is stored as plain Unicode. It is not reshaped and it is not
reordered for display. After the workbook is saved, the sheet view is set
right to left, and Excel draws the letters. Every row records the page it
came from. Columns are whatever the page printed, whether the file is an
invoice or a registry list. A header is renamed only when it clearly matches
one known Arabic or English label. An English sheet stays left to right.
Cells are not centered. Stamps, signatures, and lines that repeat on every
page are left out of the table.

    pip install -r requirements.txt
    python pdf_to_excel.py
    python pdf_to_excel.py invoice.pdf invoice.xlsx
"""

from __future__ import annotations

import contextvars
from datetime import date
import hashlib
import json
import os
import pickle
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
from collections.abc import Callable
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path


# ---------------------------------------------------------------------------
# Dependency check
# ---------------------------------------------------------------------------

def _missing_libraries() -> list[str]:
    missing = []
    for name in ("fitz", "pdfplumber", "pandas", "openpyxl"):
        try:
            __import__(name)
        except ImportError:
            missing.append("pymupdf" if name == "fitz" else name)
    return missing


_MISSING = _missing_libraries()
if _MISSING:
    message = (
        "Install the missing libraries, then run the program again:\n\n"
        "pip install " + " ".join(_MISSING)
    )
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("PDF to Excel", message)
        root.destroy()
    except Exception:
        print(message, file=sys.stderr)
    raise SystemExit(1)

import fitz
import pandas as pd

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.pdf_loader import close_document, iter_page_batches, open_document

if hasattr(fitz, "no_recommend_layout"):
    fitz.no_recommend_layout()
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


class ConversionError(Exception):
    """A problem the user can act on: missing file, locked workbook, no table."""


class ExcelWriteCount(int):
    """Row count that also remembers how many figures were stored."""

    def __new__(cls, rows: int, figures: int = 0):
        value = int.__new__(cls, rows)
        value.figures = figures
        return value


# A lone page number, in English or Arabic. It is not a table row.
_PAGE_RE = re.compile(
    r"^(?:page|pg|الصفحه|صفحه|صفحة)\s*\d{1,4}(?:\s*(?:of|من|/|-)\s*\d{1,4})?$",
    re.IGNORECASE,
)


@dataclass
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    confidence: float = 0.0

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2.0


@dataclass
class Column:
    name: str
    x0: float
    x1: float


def _is_arabic_letter(char: str) -> bool:
    """Arabic letters, including the presentation forms some PDFs store."""
    code = ord(char)
    return (
        0x0600 <= code <= 0x06FF
        or 0x0750 <= code <= 0x077F
        or 0x08A0 <= code <= 0x08FF
        or 0xFB50 <= code <= 0xFDFF
        or 0xFE70 <= code <= 0xFEFF
    )


def _normalize(text: str) -> str:
    """Make Arabic header text comparable: drop tatweel, diacritics, and alef variants."""
    if not text:
        return ""
    value = unicodedata.normalize("NFKC", text)
    value = value.replace("\u0640", "")
    value = re.sub(r"[\u064b-\u0652\u0670]", "", value)
    for source, target in (("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ى", "ي"), ("ة", "ه")):
        value = value.replace(source, target)
    return re.sub(r"\s+", " ", value).strip()


def _looks_like_seal(rect: fitz.Rect, page: fitz.Rect) -> bool:
    """A stamp is compact and sits at the side. A table rule runs across the page."""
    if rect.width < 40 or rect.height < 40:
        return False
    if rect.width > page.width * 0.4 or rect.height > page.height * 0.35:
        return False
    ratio = rect.width / max(rect.height, 1.0)
    if not 0.5 <= ratio <= 1.9:
        return False
    center = (rect.x0 + rect.x1) / 2.0
    return center < page.x0 + page.width * 0.3 or center > page.x0 + page.width * 0.7


def _merge_boxes(boxes: list[fitz.Rect], padding: float = 6.0) -> list[fitz.Rect]:
    pending = [fitz.Rect(box) for box in boxes]
    changed = True
    passes = 0
    while changed and passes < 8:
        changed = False
        passes += 1
        merged: list[fitz.Rect] = []
        used = [False] * len(pending)
        for index, rect in enumerate(pending):
            if used[index]:
                continue
            acc = fitz.Rect(rect)
            used[index] = True
            for other_index, other in enumerate(pending):
                if used[other_index]:
                    continue
                inflated = fitz.Rect(acc.x0 - padding, acc.y0 - padding, acc.x1 + padding, acc.y1 + padding)
                if not inflated.intersects(other):
                    continue
                acc |= other
                used[other_index] = True
                changed = True
            merged.append(acc)
        pending = merged
    return pending


def _seal_boxes(page: fitz.Page) -> list[fitz.Rect]:
    """
    Find stamp and signature graphics so their text never becomes a table row.

    Image seals and tight clusters of vector curves are both treated as stamps.
    Long table rules are ignored because they span the data, not a side mark.
    """
    found: list[fitz.Rect] = []
    try:
        for info in page.get_image_info() or []:
            raw = info.get("bbox")
            if not raw:
                continue
            rect = fitz.Rect(raw)
            if _looks_like_seal(rect, page.rect):
                found.append(rect)
    except Exception:
        pass
    pieces: list[fitz.Rect] = []
    try:
        for drawing in page.get_drawings() or []:
            raw = drawing.get("rect")
            if raw is None:
                continue
            rect = fitz.Rect(raw)
            if rect.width > page.rect.width * 0.45:
                continue
            if max(rect.width, rect.height) < 1.5:
                continue
            pieces.append(rect)
    except Exception:
        pieces = []
    for cluster in _merge_boxes(pieces):
        if _looks_like_seal(cluster, page.rect):
            found.append(cluster)
    expanded = []
    for rect in found:
        expanded.append(fitz.Rect(rect.x0 - 8, rect.y0 - 8, rect.x1 + 8, rect.y1 + 8))
    return expanded


def _inside_any(point: tuple[float, float], boxes: list[fitz.Rect]) -> bool:
    for box in boxes:
        if box.x0 <= point[0] <= box.x1 and box.y0 <= point[1] <= box.y1:
            return True
    return False


def _token_text(chars: list[tuple[float, float, float, float, str, int]]) -> str:
    """
    Rebuild one word in reading order.

    Isolated Arabic forms are painted left to right, so left-to-right positions
    are the real word. Joining Arabic is already stored in reading order.
    """
    isolated = 0
    joining = 0
    base = 0
    for *_, char, _index in chars:
        name = unicodedata.name(char, "")
        if "ISOLATED FORM" in name:
            isolated += 1
        elif name.endswith((" INITIAL FORM", " MEDIAL FORM", " FINAL FORM")):
            joining += 1
        elif _is_arabic_letter(char):
            base += 1
    if isolated and not joining:
        ordered = sorted(chars, key=lambda item: item[0])
        return unicodedata.normalize("NFKC", "".join(item[4] for item in ordered))
    # One line often holds several words. Judge each word on its own.
    # If the first letter sits to the right of the last, the word is already
    # in reading order (رقم, تويوتا). Reversing it would store اتويوت.
    # Reverse only a word whose letters were stored left to right.
    stream = sorted(chars, key=lambda item: item[5])
    pieces: list[str] = []
    run: list[tuple[float, str]] = []

    def flush() -> None:
        if not run:
            return
        letters = "".join(char for _, char in run)
        if len(run) >= 2 and run[-1][0] > run[0][0] + 0.4:
            letters = letters[::-1]
        pieces.append(unicodedata.normalize("NFKC", letters))
        run.clear()

    for x0, _y0, _x1, _y1, char, _index in stream:
        if _is_arabic_letter(char):
            run.append((x0, char))
        else:
            flush()
            pieces.append(unicodedata.normalize("NFKC", char))
    flush()
    return "".join(pieces)


def _page_glyphs(page: fitz.Page) -> list[tuple[float, float, float, float, str, int]]:
    """Horizontal characters in the order the PDF stored them."""
    glyphs: list[tuple[float, float, float, float, str, int]] = []
    try:
        data = page.get_text("rawdict") or {}
    except Exception:
        return []
    for block in data.get("blocks") or []:
        if block.get("type") != 0:
            continue
        for line in block.get("lines") or []:
            direction = line.get("dir") or (1, 0)
            try:
                horizontal = abs(float(direction[0]))
            except (TypeError, ValueError, IndexError):
                horizontal = 1.0
            if horizontal < 0.85:
                continue
            for span in line.get("spans") or []:
                for glyph in span.get("chars") or []:
                    char = str(glyph.get("c") or "")
                    if not char or char in "\u200b\ufeff":
                        continue
                    if char == "\u00ad":
                        char = "-"
                    box = glyph.get("bbox")
                    if not box:
                        continue
                    x0, y0, x1, y1 = (float(box[0]), float(box[1]), float(box[2]), float(box[3]))
                    glyphs.append((x0, y0, x1, y1, char, len(glyphs)))
    return glyphs


def _words_from_glyphs(glyphs: list[tuple[float, float, float, float, str, int]]) -> list[Word]:
    """Group characters into words. Arabic stays in reading order."""
    if not glyphs:
        return []
    glyphs = sorted(glyphs, key=lambda item: ((item[1] + item[3]) / 2.0, item[0]))
    lines: list[list[tuple[float, float, float, float, str, int]]] = [[glyphs[0]]]
    for glyph in glyphs[1:]:
        current = lines[-1]
        y_now = sum((item[1] + item[3]) / 2.0 for item in current) / len(current)
        if abs((glyph[1] + glyph[3]) / 2.0 - y_now) <= 4:
            current.append(glyph)
        else:
            lines.append([glyph])

    words: list[Word] = []
    for line in lines:
        line.sort(key=lambda item: item[0])
        token: list[tuple[float, float, float, float, str, int]] = []

        def flush_token() -> None:
            if not token:
                return
            text = _token_text(token).strip(" |")
            if text:
                words.append(
                    Word(
                        token[0][0],
                        min(item[1] for item in token),
                        token[-1][2],
                        max(item[3] for item in token),
                        text,
                    )
                )
            token.clear()

        for glyph in line:
            # A pipe is a column boundary in a text table, not part of the word.
            if glyph[4] == "|":
                flush_token()
                continue
            if token:
                gap = glyph[0] - token[-1][2]
                if gap > 8:
                    flush_token()
            token.append(glyph)
        flush_token()
    return words


def _page_words(page: fitz.Page) -> list[Word]:
    """
    Read horizontal characters and rebuild words from their positions.

    A stamp's rotated lettering is skipped. Characters that sit far apart
    become different cells, so a plate number is not glued to the colour.
    """
    return _words_from_glyphs(_page_glyphs(page))


def _logical_text_in_box(
    glyphs: list[tuple[float, float, float, float, str, int]],
    bbox,
) -> str:
    """
    Letters that belong to one table cell, in reading order.

    A grid extractor often returns those letters backwards (ةكرش). The
    characters inside the cell box are still in reading order (شركة).
    Excel stores that reading order and draws the sheet from the right.
    """
    if not bbox or len(bbox) < 4:
        return ""
    try:
        x0, y0, x1, y1 = (float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3]))
    except (TypeError, ValueError):
        return ""
    if x1 < x0:
        x0, x1 = x1, x0
    if y1 < y0:
        y0, y1 = y1, y0
    chosen = []
    for glyph in glyphs:
        center_x = (glyph[0] + glyph[2]) / 2.0
        center_y = (glyph[1] + glyph[3]) / 2.0
        if x0 - 1 <= center_x <= x1 + 1 and y0 - 1 <= center_y <= y1 + 1:
            chosen.append(glyph)
    words = _words_from_glyphs(chosen)
    if not words:
        return ""
    words.sort(key=lambda word: (word.center_y, word.x0))
    lines: list[list[Word]] = [[words[0]]]
    for word in words[1:]:
        if abs(word.center_y - lines[-1][-1].center_y) <= 4:
            lines[-1].append(word)
        else:
            lines.append([word])
    parts: list[str] = []
    for line in lines:
        sample = " ".join(word.text for word in line)
        arabic, latin = _letter_counts(sample)
        line.sort(key=lambda word: word.x0, reverse=arabic > latin)
        parts.append(" ".join(word.text.strip() for word in line if word.text.strip()))
    return _clean_cell(" ".join(parts))


@dataclass
class ExtractedTable:
    """One table. Column names come from the page, and anchors keep the x position."""

    columns: list[str]
    rows: list[list[str]]
    anchors: list[float] | None = None
    sources: list[int] | None = field(default=None, repr=False)
    # Every token on the printed line, including a plate that missed its column.
    lines: list[str] | None = field(default=None, repr=False)
    kind: str = ""
    # Plates seen on the page, and subtotal lines that are not vehicle rows.
    detected: int = 0
    statements: list[str] = field(default_factory=list)
    detected_pages: dict = field(default_factory=dict)
    # page|plate|field -> agreed, one-engine, normalized, suggested, or manual.
    provenance: dict[str, str] = field(default_factory=dict, repr=False)
    # Ranked candidates. Written beside the raw reading, not into the Table cell.
    suggestions: dict[str, str] = field(default_factory=dict, repr=False)
    # page|plate|field -> the OCR text before the Table cell was normalized.
    raw_readings: dict[str, str] = field(default_factory=dict, repr=False)
    # auto, rtl, or ltr. Empty follows the letters on the sheet.
    direction: str = ""
    # page|plate -> baseline of that plate, in PDF points.
    baseline_at: dict[str, float] = field(default_factory=dict, repr=False)
    # page -> field -> (x0, x1) of that column on the page.
    column_spans: dict[str, dict[str, tuple[float, float]]] = field(default_factory=dict, repr=False)
    # Shown when the text layer is boxes, or the vehicle columns were not named.
    column_warning: str = ""
    # The rendered page was already read, so a bad cache is not opened again.
    ocr_attempted: bool = False
    # Review edits: who, before, after. Written to the سجل sheet.
    audit: list = field(default_factory=list, repr=False)


def _source_key(page: int, plate: str, field_name: str) -> str:
    return f"{int(page)}|{plate}|{field_name}"


def _remember_source(
    table: ExtractedTable,
    page: int,
    plate: str,
    field_name: str,
    source: str,
    suggestion: str = "",
) -> None:
    """Record how a cell was accepted. A suggestion is not written into the cell."""
    if not plate or not field_name:
        return
    key = _source_key(page, plate, field_name)
    table.provenance[key] = source
    if suggestion:
        table.suggestions[key] = suggestion


def _remember_raw(table: ExtractedTable, page: int, plate: str, field_name: str, raw: str) -> None:
    """Keep the OCR text for the provenance sheet. The Table cell can change later."""
    if not plate or not field_name:
        return
    text = str(raw or "").strip()
    if text:
        table.raw_readings[_source_key(page, plate, field_name)] = text


def _clean_cell(value) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).replace("\n", " ")
    text = text.replace("|", " ")
    return re.sub(r"\s+", " ", text).strip()


def _unique_headers(cells: list[str]) -> list[str]:
    """Keep every column the PDF printed, and make duplicate titles unique."""
    names: list[str] = []
    seen: dict[str, int] = {}
    for index, cell in enumerate(cells, start=1):
        name = cell.strip() or f"Column {index}"
        count = seen.get(name, 0) + 1
        seen[name] = count
        names.append(name if count == 1 else f"{name} {count}")
    return names


@dataclass(frozen=True)
class _HeaderConcept:
    """One field, with the label to print in Arabic and in English."""

    key: str
    arabic: str
    english: str
    aliases: tuple[str, ...]


# Known labels only. A title that matches none of these is kept as printed,
# so an invoice with extra columns is not forced into a fixed layout.
_HEADER_CONCEPTS: tuple[_HeaderConcept, ...] = (
    _HeaderConcept("purpose", "غرض الترخيص", "Licence purpose", ("غرض الترخيص", "license purpose", "licence purpose")),
    _HeaderConcept("plate", "رقم اللوحة", "Plate number", ("رقم اللوحة", "plate number", "plate no", "plate no.", "plate #")),
    _HeaderConcept("admin", "رقم إداري", "Admin number", ("رقم اداري", "الرقم الاداري", "admin number", "admin no", "administrative number")),
    _HeaderConcept("make", "الصنع", "Make", ("الصنع", "الماركة", "make", "manufacturer")),
    _HeaderConcept("year", "السنة", "Year", ("السنة", "سنة الصنع", "model year", "year")),
    _HeaderConcept("color", "اللون", "Color", ("اللون", "اللون السائد", "color", "colour")),
    _HeaderConcept("insurance", "انتهاء التأمين", "Insurance expiry", ("انتهاء التأمين", "insurance expiry", "insurance end")),
    _HeaderConcept("item", "الصنف", "Item", ("الصنف", "البند", "البيان", "item", "description", "product")),
    _HeaderConcept("quantity", "الكمية", "Quantity", ("الكمية", "qty", "quantity")),
    _HeaderConcept("price", "السعر", "Price", ("السعر", "سعر الوحدة", "unit price", "price")),
    _HeaderConcept("total", "الإجمالي", "Total", ("الإجمالي", "المجموع", "line total", "total")),
    _HeaderConcept("date", "التاريخ", "Date", ("التاريخ", "invoice date", "date")),
    _HeaderConcept("page", "رقم الصفحة", "Page", ("رقم الصفحة", "page", "page no", "page number")),
)
_CONCEPT_BY_KEY = {concept.key: concept for concept in _HEADER_CONCEPTS}


def _header_norm(text: str) -> str:
    """Lower-case, comparable form of a header in either language."""
    value = _normalize(text).lower()
    value = value.replace(".", " ").replace("#", " ").replace(":", " ")
    return re.sub(r"\s+", " ", value).strip()


def _concept_for(text: str) -> str | None:
    """
    The field a header names, when Arabic and English agree on one field.

    "رقم اللوحة" and "Plate No" are the same field. A garbled title that
    mentions both the year and the colour is left alone, so those columns
    stay separate.
    """
    header = _header_norm(text)
    if not header or re.fullmatch(r"column \d+(?: \d+)?", header):
        return None
    compact = header.replace(" ", "")
    hits: list[tuple[int, str]] = []
    for concept in _HEADER_CONCEPTS:
        score = 0
        for alias in concept.aliases:
            alias_norm = _header_norm(alias)
            if not alias_norm:
                continue
            alias_compact = alias_norm.replace(" ", "")
            if header == alias_norm or compact == alias_compact:
                score = max(score, 100)
                continue
            phrase = " " in alias_norm or len(alias_compact) >= 8
            if phrase and f" {alias_norm} " in f" {header} " and len(alias_compact) >= len(compact) * 0.6:
                score = max(score, 70)
        if score:
            hits.append((score, concept.key))
    compact_ar = header.replace(" ", "")
    # "السنة اللون السائد" is one garbled title over two fields. It must not
    # become the colour column, or the year 2025 is replaced by صدفي.
    if "لون" in compact_ar and ("السنه" in compact_ar or "سنه" in compact_ar):
        return None
    if not hits:
        # "انتهاء التأمي ي" is the insurance column with a letter the PDF dropped.
        if "انتهاء" in compact_ar and "تام" in compact_ar:
            return "insurance"
        return None
    hits.sort(reverse=True)
    if len(hits) > 1 and hits[1][0] == hits[0][0]:
        return None
    return hits[0][1]


def _is_printed_page_label(text: str) -> bool:
    """A page number printed in a footer, such as 'صفحة 13' or '- 13 -'. Not the loop counter."""
    value = _normalize(text)
    if not value:
        return False
    compact = re.sub(r"\s+", "", value)
    if _PAGE_RE.fullmatch(value) or _PAGE_RE.fullmatch(compact):
        return True
    return bool(re.fullmatch(r"[-–—]+\d{1,4}[-–—]+", compact))


def _header_is_printed_page(name: str) -> bool:
    """True when the PDF itself titled a column as a page number."""
    if _concept_for(name) == "page":
        return True
    return _header_norm(name) in {"صفحه", "page", "pg"}


_NOTES_HEADER = "ملاحظات"
_NONDATA_HEADER = "ترويسة"
_NONDATA_BREAK = "فاصل"
_NONDATA_MARKS = {_NONDATA_HEADER, _NONDATA_BREAK}
# A blank cell is left blank. These words are never written into a header or a gap.
_FABRICATED_MARKERS = {"مش موجود", "غير موجود", "مشموجود", "غيرموجود"}
_STRICT: contextvars.ContextVar[bool] = contextvars.ContextVar("pdf_excel_strict", default=False)
# baseline-6 stores each plate on the digits of that row. The cell text is unchanged.
_CACHE_VERSION = "baseline-6"
_LAST_TIMINGS: list[dict[str, float | str]] = []
_CHECKPOINT: contextvars.ContextVar[str] = contextvars.ContextVar("pdf_checkpoint", default="")


class ExtractionTimings:
    """Seconds for each stage of one extract. Parallel pages overlap, so the sum is not the wait."""

    def __init__(self) -> None:
        self.events: list[dict[str, float | str]] = []

    def add(self, name: str, seconds: float) -> None:
        self.events.append({"name": name, "seconds": round(max(0.0, float(seconds)), 3)})

    def rows(self) -> list[dict[str, float | str]]:
        return [dict(item) for item in self.events]


_TIMINGS: contextvars.ContextVar[ExtractionTimings | None] = contextvars.ContextVar("pdf_excel_timings", default=None)


class _Stage:
    def __init__(self, name: str) -> None:
        self.name = name
        self.started = 0.0

    def __enter__(self) -> "_Stage":
        self.started = time.perf_counter()
        return self

    def __exit__(self, *_exc: object) -> bool:
        _remember_time(self.name, time.perf_counter() - self.started)
        return False


def _stage(name: str) -> _Stage:
    return _Stage(name)


def _remember_time(name: str, seconds: float) -> None:
    clock = _TIMINGS.get()
    if clock is not None:
        clock.add(name, seconds)


def last_timings() -> list[dict[str, float | str]]:
    """Stages from the most recent conversion, in the order they finished."""
    return [dict(item) for item in _LAST_TIMINGS]


def _cache_dir() -> Path:
    """Hash cache. A packaged app writes it under %LOCALAPPDATA%, not the install folder."""
    try:
        from engine.paths import cache_dir

        return cache_dir()
    except Exception:
        folder = Path(__file__).resolve().parent.parent / ".extract_cache"
        folder.mkdir(exist_ok=True)
        return folder


def _cache_stamp(pdf_bytes: bytes, password: str, strict: bool, direction: str = "") -> str:
    digest = hashlib.sha256(pdf_bytes).hexdigest()
    extra = hashlib.sha256(
        f"{_CACHE_VERSION}|{int(bool(strict))}|{password or ''}|{direction or ''}".encode("utf-8")
    ).hexdigest()[:16]
    return f"{digest}-{extra}"


def workbook_cache_path(pdf_bytes: bytes, password: str, strict: bool, direction: str = "") -> Path:
    """Extraction cache when direction is empty. The workbook cache includes the sheet direction."""
    kind = "book" if direction else "extract"
    return _cache_dir() / f"{kind}-{_cache_stamp(pdf_bytes, password, strict, direction)}.pkl"


def _load_pickle(path: Path) -> dict | None:
    try:
        with path.open("rb") as handle:
            state = pickle.load(handle)
    except Exception:
        return None
    return state if isinstance(state, dict) else None


def _save_pickle(path: Path, state: dict) -> None:
    try:
        path.parent.mkdir(exist_ok=True)
        temporary = path.with_suffix(".tmp")
        with temporary.open("wb") as handle:
            pickle.dump(state, handle, protocol=pickle.HIGHEST_PROTOCOL)
        temporary.replace(path)
    except Exception:
        try:
            path.with_suffix(".tmp").unlink(missing_ok=True)
        except Exception:
            return


def _load_cached_tables(source: Path, password: str, strict: bool) -> list[ExtractedTable] | None:
    try:
        pdf_bytes = source.read_bytes()
    except OSError:
        return None
    path = workbook_cache_path(pdf_bytes, password, strict)
    if not path.is_file():
        return None
    started = time.perf_counter()
    state = _load_pickle(path)
    tables = state.get("tables") if state else None
    if not tables:
        return None
    if any(_extract_needs_reread(table) for table in tables):
        return None
    _remember_time("cache", time.perf_counter() - started)
    return list(tables)


def _store_cached_tables(source: Path, password: str, strict: bool, tables: list[ExtractedTable]) -> None:
    try:
        pdf_bytes = source.read_bytes()
    except OSError:
        return
    _save_pickle(workbook_cache_path(pdf_bytes, password, strict), {"tables": tables})


def load_cached_workbook(pdf_bytes: bytes, password: str, strict: bool, direction: str) -> dict | None:
    """The last workbook for this file and these settings, including the Excel bytes."""
    path = workbook_cache_path(pdf_bytes, password, strict, direction or "auto")
    if not path.is_file():
        return None
    started = time.perf_counter()
    state = _load_pickle(path)
    if not state or not state.get("rows") or not state.get("bytes"):
        return None
    probe = ExtractedTable(
        list(state.get("columns") or []),
        [list(row) for row in state.get("rows") or []],
        kind=str(state.get("kind") or ""),
    )
    probe.ocr_attempted = bool(state.get("ocr_attempted"))
    if _extract_needs_reread(probe):
        return None
    state["timings"] = [{"name": "cache", "seconds": round(time.perf_counter() - started, 3)}]
    return state


def save_cached_workbook(pdf_bytes: bytes, password: str, strict: bool, direction: str, state: dict) -> None:
    if not state.get("rows") or not state.get("bytes"):
        return
    stored = {
        key: value
        for key, value in state.items()
        if key not in {"pdf", "undo", "review_cache", "timings"}
    }
    _save_pickle(workbook_cache_path(pdf_bytes, password, strict, direction or "auto"), stored)


def strict_values() -> bool:
    """True when blank cells must stay blank and nothing is copied from another row."""
    return bool(_STRICT.get())


class _StrictScope:
    """Apply the no-guess rule for one extract, re-scan, or export."""

    def __init__(self, enabled: bool) -> None:
        self.enabled = bool(enabled)
        self.token: contextvars.Token[bool] | None = None

    def __enter__(self) -> "_StrictScope":
        self.token = _STRICT.set(self.enabled)
        return self

    def __exit__(self, *exc: object) -> None:
        if self.token is not None:
            _STRICT.reset(self.token)


def _is_notes_header(name: str) -> bool:
    """The review column. It is not a vehicle field and must not move the others."""
    return _header_norm(name) in {"ملاحظات", "ملاحظه", "notes", "note"}


def _add_page_column(table: ExtractedTable) -> None:
    """
    Write the sequential page counter into its own column.

    page 1 of the file is Page 1, page 2 is Page 2, and so on. A number printed
    in the footer is never copied into this column.
    """
    keep = [index for index, name in enumerate(table.columns) if not _header_is_printed_page(name)]
    if len(keep) != len(table.columns):
        table.columns = [table.columns[index] for index in keep]
        table.rows = [[row[index] if index < len(row) else "" for index in keep] for row in table.rows]
        if table.anchors is not None:
            table.anchors = [table.anchors[index] for index in keep if index < len(table.anchors)]
    arabic = _sheet_is_rtl(table)
    header = "رقم الصفحة" if arabic else "Page"
    word = "صفحة" if arabic else "Page"
    # sources holds the 1-based loop counter assigned when that page was read.
    if table.sources and len(table.sources) == len(table.rows):
        numbers = [int(page_num) for page_num in table.sources]
    else:
        numbers = [1] * len(table.rows)
    labels = [f"{word} \u2066{page_num}\u2069" for page_num in numbers]
    vehicle = [name for _key, name in _VEHICLE_COLUMNS]
    note_indexes = [index for index, name in enumerate(table.columns) if _is_notes_header(name)]
    data_indexes = [index for index in range(len(table.columns)) if index not in note_indexes]
    data_names = [table.columns[index] for index in data_indexes]
    note_names = [table.columns[index] for index in note_indexes]

    def _cells(row: list[str], indexes: list[int]) -> list[str]:
        return [row[index] if index < len(row) else "" for index in indexes]

    data_rows = [_cells(row, data_indexes) for row in table.rows]
    note_rows = [_cells(row, note_indexes) for row in table.rows]
    # The seven register columns stay in printed order. The page counter
    # follows them, then the review note. A notes column must not push the
    # page counter in front of the plate.
    if data_names == vehicle:
        table.columns = data_names + [header] + note_names
        table.rows = [
            data + [label] + notes
            for data, notes, label in zip(data_rows, note_rows, labels)
        ]
        if table.anchors is not None:
            table.anchors = [table.anchors[index] for index in data_indexes if index < len(table.anchors)]
            table.anchors = list(table.anchors) + [0.0] * (1 + len(note_names))
        return
    table.columns = [header] + data_names + note_names
    # The digits are isolated so Excel does not split 13 into "3 ... 1".
    table.rows = [
        [label] + data + notes
        for label, data, notes in zip(labels, data_rows, note_rows)
    ]
    if table.anchors is not None:
        kept = [table.anchors[index] for index in data_indexes if index < len(table.anchors)]
        table.anchors = [0.0] + kept + [0.0] * len(note_names)


def _apply_bilingual_headers(table: ExtractedTable) -> None:
    """Use the Arabic label on an Arabic sheet, and the English label on an English sheet."""
    try:
        arabic_sheet = _sheet_is_rtl(table)
        renamed: list[str] = []
        for name in table.columns:
            if name in _VEHICLE_HEADER_TEXT:
                renamed.append(name)
                continue
            concept = _concept_for(name)
            if concept is None:
                renamed.append(name)
                continue
            spec = _CONCEPT_BY_KEY[concept]
            renamed.append(spec.arabic if arabic_sheet else spec.english)
        table.columns = _unique_headers(renamed)
    except Exception:
        return


def _is_noise_row(cells: list[str]) -> bool:
    """Page numbers, blank lines, and stamp captions. A total row on an invoice is kept."""
    filled = [cell for cell in cells if cell.strip()]
    if not filled:
        return True
    text = _normalize(" ".join(filled))
    compact = re.sub(r"\s+", "", text)
    if _PAGE_RE.fullmatch(text) or _PAGE_RE.fullmatch(compact):
        return True
    if re.fullmatch(r"[-–—]*\d{1,3}[-–—]*", compact):
        return True
    if len(filled) == 1 and re.fullmatch(r"\d{1,3}", compact):
        return True
    folded = text.replace(" ", "")
    if "هذاالكشف" in folded or "هذا الكشف" in text:
        return True
    if "صالح" in text and ("لمده" in folded or "لمدة" in text or "الكشف" in text):
        return True
    return False


def _header_index(rows: list[list[str]]) -> int:
    """The header is the first row that is mostly labels rather than amounts."""
    for index, row in enumerate(rows[:4]):
        filled = [cell for cell in row if cell.strip()]
        if len(filled) < 2:
            continue
        labels = 0
        for cell in filled:
            if not re.fullmatch(r"[-+]?\d[\d.,]*", cell):
                labels += 1
        if labels >= max(1, len(filled) // 2):
            return index
    return 0


def _table_from_matrix(matrix: list[list[str]], page_index: int = 0) -> ExtractedTable | None:
    """Turn a grid into columns plus data rows. The header is whatever the PDF used."""
    width = max((len(row) for row in matrix), default=0)
    if width < 2:
        return None
    rows = [row + [""] * (width - len(row)) for row in matrix]
    rows = [[_clean_cell(cell) for cell in row][:width] for row in rows]
    if not rows:
        return None
    header_at = _header_index(rows)
    columns = _unique_headers(rows[header_at])
    data: list[list[str]] = []
    header_key = tuple(_normalize(name) for name in columns)
    for row in rows[header_at + 1 :]:
        if _is_noise_row(row):
            continue
        if tuple(_normalize(cell) for cell in row) == header_key:
            continue
        if not any(cell.strip() for cell in row):
            continue
        data.append(row)
    if not data:
        return None
    filled = sum(1 for row in data for cell in row if cell.strip())
    if filled < 2:
        return None
    return ExtractedTable(columns, data, None, [page_index] * len(data))


def _line_tables(page: fitz.Page) -> list:
    """Prefer ruled grids. Invoices and statements usually draw those lines."""
    try:
        found = page.find_tables(vertical_strategy="lines", horizontal_strategy="lines")
        tables = list(getattr(found, "tables", []) or [])
        if tables:
            return tables
    except TypeError:
        pass
    except Exception:
        return []
    try:
        found = page.find_tables()
        return list(getattr(found, "tables", []) or [])
    except Exception:
        return []


def _matrix_with_logical_text(raw_rows, row_boxes, glyphs) -> list[list[str]]:
    """Keep the grid, and replace each cell with the letters actually inside it."""
    matrix: list[list[str]] = []
    for index, raw in enumerate(raw_rows or []):
        if not raw:
            continue
        boxes = row_boxes[index] if index < len(row_boxes) else []
        built: list[str] = []
        for cell_index, cell in enumerate(raw):
            fallback = "" if cell is None else str(cell)
            bbox = boxes[cell_index] if cell_index < len(boxes) else None
            logical = _logical_text_in_box(glyphs, bbox) if bbox else ""
            built.append(logical if logical else fallback)
        matrix.append(built)
    return matrix


def _tables_from_rules(page: fitz.Page, page_index: int = 0) -> list[ExtractedTable]:
    extracted: list[ExtractedTable] = []
    glyphs = _page_glyphs(page)
    for table in _line_tables(page):
        try:
            raw_rows = table.extract()
        except Exception:
            continue
        if not raw_rows:
            continue
        row_boxes = []
        for row in getattr(table, "rows", []) or []:
            row_boxes.append(list(getattr(row, "cells", []) or []))
        matrix = _matrix_with_logical_text(raw_rows, row_boxes, glyphs)
        parsed = _table_from_matrix(matrix, page_index)
        if parsed is not None:
            extracted.append(parsed)
    return extracted


def _cluster_words(words: list[Word]) -> list[list[Word]]:
    """Words that sit in the same horizontal slot belong to one column."""
    ordered = sorted(words, key=lambda word: word.x0)
    clusters: list[list[Word]] = [[ordered[0]]]
    for word in ordered[1:]:
        previous = max(item.x1 for item in clusters[-1])
        if word.x0 <= previous + 14:
            clusters[-1].append(word)
        else:
            clusters.append([word])
    return clusters


def _cluster_name(cluster: list[Word]) -> str:
    arabic = sum(1 for word in cluster for char in word.text if _is_arabic_letter(char))
    latin = sum(1 for word in cluster for char in word.text if char.isascii() and char.isalpha())
    ordered = sorted(cluster, key=lambda word: word.x0, reverse=arabic > latin)
    return _clean_cell(" ".join(word.text for word in ordered))


def _columns_from_header(header: list[Word]) -> list[Column]:
    clusters = _cluster_words(header)
    named = []
    for cluster in clusters:
        name = _cluster_name(cluster)
        if name:
            named.append((name, min(word.x0 for word in cluster), max(word.x1 for word in cluster)))
    if len(named) < 2:
        return []
    names = _unique_headers([name for name, _, _ in named])
    columns: list[Column] = []
    for index, (name, x0, x1) in enumerate(zip(names, [item[1] for item in named], [item[2] for item in named])):
        if index == 0:
            left = x0 - 16
        else:
            left = (named[index - 1][2] + x0) / 2.0
        if index == len(named) - 1:
            right = x1 + 24
        else:
            right = (x1 + named[index + 1][1]) / 2.0
        columns.append(Column(name, left, max(right, left + 8)))
    return columns


def _assign_cells(row: list[Word], columns: list[Column]) -> list[str]:
    buckets: list[list[Word]] = [[] for _ in columns]
    for word in row:
        target = None
        for index, column in enumerate(columns):
            if column.x0 <= word.center_x <= column.x1:
                target = index
                break
        if target is None:
            target = min(
                range(len(columns)),
                key=lambda index: abs(((columns[index].x0 + columns[index].x1) / 2.0) - word.center_x),
            )
            center = (columns[target].x0 + columns[target].x1) / 2.0
            if abs(word.center_x - center) > 40:
                continue
        buckets[target].append(word)
    values = []
    for index, bucket in enumerate(buckets):
        arabic = sum(1 for word in bucket for char in word.text if _is_arabic_letter(char))
        latin = sum(1 for word in bucket for char in word.text if char.isascii() and char.isalpha())
        bucket.sort(key=lambda word: word.x0, reverse=arabic > latin)
        values.append(_clean_cell(" ".join(word.text for word in bucket)))
    return values


def _tables_from_words(page: fitz.Page, page_index: int = 0) -> list[ExtractedTable]:
    """
    Build a table from aligned text when the PDF has no ruled grid.

    The header is the widest row of short labels. Later rows use those columns,
    however many the invoice happens to have.
    """
    words = _page_words(page)
    visual_rows = _group_rows(words)
    if len(visual_rows) < 2:
        return []
    chosen: tuple[int, list[Column]] | None = None
    for index, row in enumerate(visual_rows):
        if len(row) < 2 or row[0].center_y > page.rect.height * 0.8:
            continue
        span = max(word.x1 for word in row) - min(word.x0 for word in row)
        if span < page.rect.width * 0.35:
            continue
        columns = _columns_from_header(row)
        if len(columns) < 2:
            continue
        if chosen is None or len(columns) > len(chosen[1]):
            chosen = (index, columns)
    if chosen is None:
        return []
    header_index, columns = chosen
    header_bottom = max(word.y1 for word in visual_rows[header_index])
    left = min(column.x0 for column in columns)
    right = max(column.x1 for column in columns)
    seals = _seal_boxes(page)
    body_rows = []
    for row in visual_rows[header_index + 1 :]:
        kept = []
        for word in row:
            outside = word.center_x < left or word.center_x > right
            if outside and _inside_any((word.center_x, word.center_y), seals):
                continue
            if word.y0 <= header_bottom:
                continue
            kept.append(word)
        if kept:
            body_rows.append(kept)
    matrix = [[column.name for column in columns]]
    header_key = tuple(_normalize(column.name) for column in columns)
    for row in body_rows:
        values = _assign_cells(row, columns)
        if _is_noise_row(values):
            continue
        if tuple(_normalize(cell) for cell in values) == header_key:
            continue
        matrix.append(values)
    parsed = _table_from_matrix(matrix, page_index)
    return [parsed] if parsed is not None else []


# A date piece such as "04" or "21-" is still part of the insurance date.
_DATE_TOKEN = re.compile(r"^[\d./\-]+$")
_LOOKALIKE = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "B": "8", "Z": "2"})
_OCR_ENGINES: dict[str, object] = {}


def _letter_counts(text: str) -> tuple[int, int]:
    arabic = sum(1 for char in text if _is_arabic_letter(char))
    latin = sum(1 for char in text if char.isascii() and char.isalpha())
    return arabic, latin


def _text_char_count(page: fitz.Page) -> int:
    """Letters already stored in the PDF. A scan of paper returns zero."""
    try:
        text = page.get_text("text") or ""
    except Exception:
        return 0
    return sum(1 for char in text if char.isalpha() or _is_arabic_letter(char))


def _glyph_is_unmapped(char: str) -> bool:
    """A box, a private-use code, or a dingbat. These are not letters."""
    if not char:
        return False
    if char in "\ufffd■□▪▫●◼◻▢▣":
        return True
    code = ord(char)
    if 0xE000 <= code <= 0xF8FF or 0xF000 <= code <= 0xF8FF:
        return True
    if 0x25A0 <= code <= 0x25FF or 0x2700 <= code <= 0x27BF:
        return True
    return False


def _token_is_garbage(text: str) -> bool:
    """A run of one Latin letter, or a run of boxes. A real word is left alone."""
    raw = "".join(char for char in str(text) if not char.isspace())
    if len(raw) < 3:
        return False
    if sum(1 for char in raw if _glyph_is_unmapped(char)) / len(raw) >= 0.5:
        return True
    letters = [char for char in raw if char.isalpha() and not _is_arabic_letter(char)]
    if len(letters) >= 3 and len({char.lower() for char in letters}) == 1 and len(letters) / len(raw) >= 0.8:
        return True
    return False


def _headers_look_garbage(columns: list[str]) -> bool:
    filled = [str(name) for name in columns if str(name).strip()]
    if len(filled) < 3:
        return False
    bad = sum(1 for name in filled if _token_is_garbage(name))
    return bad / len(filled) >= 0.35


def text_layer_is_garbage(page: fitz.Page) -> bool:
    """
    True when the stored letters are symbol-font boxes or unmapped codes.

    Digits in a normal font do not count. A page of real Arabic or English
    stays on its text layer.
    """
    try:
        data = page.get_text("rawdict") or {}
    except Exception:
        return False
    letters = 0
    garbage = 0
    symbol_marks = ("zapfdingbat", "wingding", "webding", "dingbat")
    for block in data.get("blocks") or []:
        if block.get("type") != 0:
            continue
        for line in block.get("lines") or []:
            for span in line.get("spans") or []:
                font = str(span.get("font") or "").lower()
                symbol_font = any(mark in font for mark in symbol_marks)
                for glyph in span.get("chars") or []:
                    char = str(glyph.get("c") or "")
                    if not char or char.isspace():
                        continue
                    if not (char.isalpha() or _is_arabic_letter(char) or _glyph_is_unmapped(char)):
                        continue
                    letters += 1
                    if symbol_font or _glyph_is_unmapped(char):
                        garbage += 1
    if letters < 12:
        return False
    return garbage / letters >= 0.35


def _median(values: list[float]) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    mid = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[mid]
    return (ordered[mid - 1] + ordered[mid]) / 2.0


def _assemble_date(tokens: list[str]) -> str | None:
    """Join split OCR pieces such as 2027, 04, 21 into 2027-04-21."""
    groups: list[str] = []
    for token in tokens:
        cleaned = token.translate(_LOOKALIKE).strip(" _|")
        if not cleaned or not _DATE_TOKEN.fullmatch(cleaned):
            return None
        groups.extend(re.findall(r"\d+", cleaned))
    if len(groups) < 3:
        return None
    year = ""
    rest: list[str] = []
    for piece in groups:
        if not year and len(piece) == 4 and piece.startswith(("19", "20")):
            year = piece
        elif not year and len(piece) == 3 and piece[0] == "0":
            year = "2" + piece
        else:
            rest.append(piece)
    if not year or len(rest) < 2 or len(rest[0]) > 2 or len(rest[1]) > 2:
        return None
    month, day = int(rest[0]), int(rest[1])
    if not (1 <= month <= 12 and 1 <= day <= 31):
        if 1 <= day <= 12 and 1 <= month <= 31:
            month, day = day, month
        else:
            return None
    if not 1990 <= int(year) <= 2099:
        return None
    return f"{year}-{month:02d}-{day:02d}"


def _cell_text(tokens: list[Word]) -> str:
    """One cell. Arabic is read right to left. A split date becomes one value."""
    if not tokens:
        return ""
    ordered = sorted(tokens, key=lambda word: word.x0)
    assembled = _assemble_date([word.text for word in ordered])
    if assembled:
        return assembled
    sample = "".join(word.text for word in ordered)
    arabic, latin = _letter_counts(sample)
    if arabic > latin:
        ordered = list(reversed(ordered))
    parts = [ordered[0].text.strip()]
    for previous, word in zip(ordered, ordered[1:]):
        gap = abs(word.x0 - previous.x1)
        same_line = abs(word.center_y - previous.center_y) <= 4
        if same_line and gap < 3.5 and arabic > latin:
            parts.append(word.text.strip())
        else:
            parts.append(" ")
            parts.append(word.text.strip())
    return _clean_cell("".join(parts))


def _arabic_only_text(text: str) -> bool:
    """A make or a colour, not a plate, a year, or a date."""
    if not text or re.search(r"\d", text):
        return False
    return any(_is_arabic_letter(char) for char in text)


def _tidy_arabic(text: str) -> str:
    """Join a word the reader split, and turn alef maqsura into yeh."""
    value = unicodedata.normalize("NFKC", text or "")
    value = re.sub(r"[\u064b-\u0652\u0670]", "", value)
    for mark in ("\u200e", "\u200f", "\u200c", "\u200d", "\u0640", '"', "'", "«", "»", "“", "”"):
        value = value.replace(mark, "")
    value = value.replace("ى", "ي")
    parts = [part for part in value.split() if part]
    if not parts:
        return ""
    merged = [parts[0]]
    for part in parts[1:]:
        previous = merged[-1]
        shorter = min(len(previous), len(part))
        longer = max(len(previous), len(part))
        # "تو يو تا" and "صد في" are one word. "باص خاص" stays two words.
        glue = shorter <= 2 and not (longer >= 5 and shorter == 2)
        if (
            glue
            and _is_arabic_letter(previous[-1])
            and _is_arabic_letter(part[0])
        ):
            merged[-1] = previous + part
        else:
            merged.append(part)
    return " ".join(merged)


def _collapse_spaced_arabic(text: str) -> str:
    """
    Join Arabic letters that were read with a space between each one.

    "ت و ي و ت ا" becomes تويوتا, "ص د ف ى" becomes صدفي, and "ا ب ي ض"
    becomes ابيض. A real two-word value such as "باص خاص" stays two words.
    Alef maqsura is folded to yeh here. Hamza is folded only while matching
    a dictionary entry, so an unmatched word keeps the hamza it was read with.
    """
    value = _tidy_arabic(text)
    parts = [part for part in value.split() if part]
    if not parts:
        return ""
    merged: list[str] = []
    singles: list[str] = []

    def flush() -> None:
        if singles:
            merged.append("".join(singles))
            singles.clear()

    for part in parts:
        if len(part) == 1 and _is_arabic_letter(part):
            singles.append(part)
        else:
            flush()
            merged.append(part)
    flush()
    return " ".join(merged)


# Words this registry prints. A reversed OCR or PDF token such as "اتويوت"
# is put back only when the reversed spelling is one of these words.
# Excel stores logical Arabic and draws it right to left itself. Shaping the
# letters into presentation forms, or running a bidi reorder, would make
# Excel show the word backwards a second time.
_LOGICAL_WORDS = (
    "تويوتا",
    "شيفروليه",
    "شيفرولية",
    "رقم",
    "اللوحة",
    "لوحة",
    "إداري",
    "الصنع",
    "صنع",
    "السنة",
    "اللون",
    "انتهاء",
    "التأمين",
    "الصفحة",
    "صفحة",
    "كشف",
    "تجريبي",
    "لترخيص",
    "وسيارات",
    "الأفضل",
    "غرض",
    "الترخيص",
    "إيسوزو",
    "ايسوزو",
    "رينو",
    "باص",
    "خاص",
    "خصوصي",
)


def _logical_key(text: str) -> str:
    """Folded Arabic, so أبيض and ابيض are the same word."""
    value = _tidy_arabic(text)
    for source, target in (("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ؤ", "و"), ("ئ", "ي"), ("ة", "ه")):
        value = value.replace(source, target)
    return value.replace(" ", "")


_LOGICAL_CACHE: set[str] | None = None


def _logical_words() -> set[str]:
    global _LOGICAL_CACHE
    if _LOGICAL_CACHE is None:
        found = {_logical_key(word) for word in _LOGICAL_WORDS}
        found.update(_logical_key(color) for color in _COLOR_CANON)
        _LOGICAL_CACHE = found
    return _LOGICAL_CACHE


def _logical_arabic(text: str) -> str:
    """
    Restore a word that was stored backwards, and leave a correct word alone.

    "اتويوت" becomes تويوتا. "تويوتا" stays تويوتا. The letters are not
    reshaped and the word is not reordered for display: Excel does that
    when the sheet is right to left.
    """
    if not text or not any(_is_arabic_letter(char) for char in text):
        return text
    known = _logical_words()
    pieces: list[str] = []
    for word in text.split():
        if not any(_is_arabic_letter(char) for char in word):
            pieces.append(word)
            continue
        # A plate is digits, sometimes with letters. Reversing that token
        # turns 18018 into 81081. Only a pure Arabic word may be restored.
        if re.search(r"\d", word):
            pieces.append(word.translate(_INDIC_DIGITS))
            continue
        current = _logical_key(word)
        flipped = word[::-1]
        flipped_key = _logical_key(flipped)
        # Only an exact known word is restored. A fuzzy colour match would
        # turn a header fragment into فضي or رمادي.
        if current not in known and flipped_key in known:
            pieces.append(flipped)
        else:
            pieces.append(word)
    return " ".join(pieces)


# A whole value, after the cell has already been cleaned.
_DATE_VALUE = re.compile(r"^(?:19|20)\d{2}-\d{2}-\d{2}$")
_YEAR_VALUE = re.compile(r"^(?:19|20)\d{2}$")
_PLATE_VALUE = re.compile(r"^\d{4,5}$")
# These headers already name a field. A column of item codes must not become a plate column.
_ROLE_LOCKED = {"purpose", "admin", "make", "color", "item", "quantity", "price", "total", "page"}


def _blank_text(value) -> str:
    """Empty stays empty. The literal NaN that pandas would write is removed."""
    if value is None:
        return ""
    if isinstance(value, float) and value != value:
        return ""
    text = _clean_cell(value)
    if text.lower() in {"nan", "none", "null", "<na>", "nat"}:
        return ""
    return text


def _value_kind(text: str) -> str | None:
    """Plate, year, or date. Anything else is left where it was read."""
    if _DATE_VALUE.fullmatch(text):
        return "date"
    if _YEAR_VALUE.fullmatch(text):
        return "year"
    if _PLATE_VALUE.fullmatch(text):
        return "plate"
    return None


def _pattern_ratios(values: list[str]) -> dict[str, float]:
    if not values:
        return {"date": 0.0, "year": 0.0, "plate": 0.0}
    counts = {"date": 0, "year": 0, "plate": 0}
    for value in values:
        kind = _value_kind(value)
        if kind:
            counts[kind] += 1
    total = len(values)
    return {kind: counts[kind] / total for kind in counts}


def _column_role(name: str, values: list[str]) -> tuple[str | None, float]:
    """
    Decide what a column holds from its header, then from the values themselves.

    A header such as Color is never treated as a plate column. When the header
    was not read, a column that is mostly 4 or 5 digit numbers is the plate
    column, one that is mostly years starting with 19 or 20 is the year, and
    one that is mostly YYYY-MM-DD is the date.
    """
    concept = _concept_for(name)
    if concept in _ROLE_LOCKED:
        return None, 0.0
    if concept == "insurance":
        concept = "date"
    filled = [value for value in values if value]
    ratios = _pattern_ratios(filled)
    if concept in {"plate", "year", "date"} and (not filled or ratios[concept] >= 0.25):
        return concept, max(ratios[concept], 0.25)
    if len(filled) < 4:
        return None, 0.0
    best = max(("date", "year", "plate"), key=lambda kind: ratios[kind])
    if ratios[best] >= 0.55:
        return best, ratios[best]
    return None, 0.0


def _repair_shifted_values(table: ExtractedTable) -> None:
    """
    Move a value that landed in the wrong column back, and only when the right cell is empty.

    A date sitting in the year column is moved into the date column. Two values
    that each match the other's column are swapped. A filled cell is never overwritten.
    """
    scored: list[tuple[str | None, float]] = []
    width = len(table.columns)
    for index, name in enumerate(table.columns):
        column = [row[index] if index < len(row) else "" for row in table.rows]
        scored.append(_column_role(name, column))
    role_at: dict[str, int] = {}
    for index, (role, score) in enumerate(scored):
        if role is None:
            continue
        current = role_at.get(role)
        if current is None or score > scored[current][1]:
            role_at[role] = index
    roles = [None] * width
    for role, index in role_at.items():
        roles[index] = role
    if not any(roles):
        return
    # A vehicle sheet is already pinned to its seven headers. Moving a plate
    # or a year between those columns is what scrambles the register.
    # The page counter and the review note sit beside those seven columns.
    pinned = [
        name
        for name in table.columns
        if not _header_is_printed_page(name) and not _is_notes_header(name)
    ]
    if pinned == [name for _key, name in _VEHICLE_COLUMNS]:
        return
    for row in table.rows:
        while len(row) < width:
            row.append("")
        for index in range(width):
            kind = _value_kind(row[index])
            if kind is None or roles[index] == kind or kind not in role_at:
                continue
            target = role_at[kind]
            other = row[target]
            # A plate, year, or date in the wrong column moves home when that cell is empty.
            # Two swapped values trade places. A cell that already has a value is left alone.
            if not other:
                row[target] = row[index]
                row[index] = ""
            elif roles[index] is not None and _value_kind(other) == roles[index]:
                row[index], row[target] = other, row[index]


# Canonical makes, plus the OCR and visual-order slips these reports produce.
# The digit order of a plate is never passed through this list.
_BRAND_ALIASES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("تويوتا", ("تويوتا", "اتويوت", "تويوته", "تويووتا", "تويوتاء", "تويو تا", "تويوطة", "تويوت")),
    (
        "شيفروليه",
        (
            "شيفروليه",
            "شيفرولية",
            "شيفروليت",
            "شفروليه",
            "شيقروليه",
            "شيفقروليه",
            "شيفرولىه",
            "شيفروليةه",
            "شيفروليهه",
            "شيفرولي",
            "شيفرولى",
            "شيفورليه",
            "هيلورفيش",
        ),
    ),
    ("إيسوزو", ("إيسوزو", "ايسوزو", "ايسوزوا", "ايسوزوو", "وزوسيا", "وزوسيإ", "إيسوزو")),
    ("رينو", ("رينو", "رينوو", "ونير", "رينوا")),
    (
        "جي إم سي",
        (
            "جي إم سي",
            "جي ام سي",
            "جي  ام  سي",
            "ج ي ام سي",
            "ج ي إم سي",
            "جي إم سي",
            "جي أم سي",
            "ج إم سي",
            "جيامسي",
            "جي إمسي",
            "جيمسي",
            "جيإمسي",
            "GMC",
            "gmc",
        ),
    ),
)


# Official makes in the Saudi market. The الصنع column is rewritten to one of these.
_MANUFACTURERS: tuple[str, ...] = (
    "تويوتا",
    "شيفروليه",
    "إيسوزو",
    "رينو",
    "لكزس",
    "جي إم سي",
    "فورد",
    "نيسان",
    "هوندا",
    "هيونداي",
)
_MANUFACTURER_SCORE = 75


def _manufacturer_catalog() -> dict[str, str]:
    """Folded spelling to the official name. Hamza and spaces do not make a new make."""
    return {_logical_key(name): name for name in _MANUFACTURERS}


def _alias_brand(text: str) -> str | None:
    """A known reversed or OCR spelling, such as اتويوت or هيلورفيش."""
    folded = _logical_key(text)
    if not folded:
        return None
    catalog = _manufacturer_catalog()
    if folded in catalog:
        return catalog[folded]
    for canon, aliases in _BRAND_ALIASES:
        keys = {_logical_key(alias) for alias in aliases}
        keys.add(_logical_key(canon))
        if folded in keys:
            return canon if canon in _MANUFACTURERS else catalog.get(_logical_key(canon), canon)
    return None


def _manufacturer_scorer(query: str, choice: str) -> int:
    """
    Full-string similarity, plus a fragment score for a shortened make.

    تويو and توبو are pieces of تويوتا. A longer slip such as هونداي is not
    given to the shorter name هوندا just because those letters sit inside it.
    """
    from thefuzz import fuzz

    ratio = int(fuzz.ratio(query, choice))
    if len(query) < len(choice) and len(choice) - len(query) <= 2:
        return max(ratio, int(fuzz.partial_ratio(query, choice)))
    return ratio


def _extract_manufacturer(folded: str) -> tuple[str, int] | None:
    """Best official make and its score. Uses thefuzz extractOne."""
    choices = list(_manufacturer_catalog())
    if not folded or not choices:
        return None
    try:
        from thefuzz import process

        match = process.extractOne(folded, choices, scorer=_manufacturer_scorer)
    except Exception:
        match = None
        best = 0
        for choice in choices:
            limit = max(len(folded), len(choice))
            distance = _edit_distance(folded, choice, limit=limit)
            score = int(round(100 * (1 - distance / limit))) if limit else 0
            if score > best:
                best = score
                match = (choice, score)
    if not match:
        return None
    choice, score = match[0], int(match[1])
    official = _manufacturer_catalog().get(choice)
    if official is None:
        return None
    return official, score


def _make_skeleton(text: str) -> str:
    """Fold the letters OCR swaps inside a manufacturer name."""
    value = _logical_key(text).casefold()
    for source, target in (("ب", "ي"), ("و", "ي"), ("ق", "ف"), ("ى", "ي")):
        value = value.replace(source, target)
    return value


def _prefix_manufacturer(folded: str) -> str:
    """
    One official make when the OCR text is a damaged prefix of that name.

    توو stays with تويوتا. Two makes that share the same start are left alone.
    """
    skeleton = _make_skeleton(folded)
    if len(skeleton) < 3:
        return ""
    hits: list[str] = []
    for name in _MANUFACTURERS:
        key = _make_skeleton(name)
        if not key.startswith(skeleton[:2]):
            continue
        from thefuzz import fuzz

        if int(fuzz.ratio(skeleton, key)) >= 64 or key.startswith(skeleton):
            hits.append(name)
    unique = list(dict.fromkeys(hits))
    if len(unique) == 1:
        return unique[0]
    return ""


def _match_manufacturer(text: str) -> str:
    """
    Replace a corrupted الصنع value when it is close to an official make.

    The cell is rewritten in place. Plates, dates, years, and admin numbers
    contain digits and are returned unchanged, so this step cannot move them.
    """
    cleaned = _tidy_arabic(text)
    if not cleaned:
        return text
    if re.search(r"\d", cleaned):
        return cleaned
    known = _alias_brand(cleaned)
    if known is not None:
        return known
    folded = _logical_key(cleaned)
    if len(folded) < 3:
        return cleaned
    found = _extract_manufacturer(folded)
    if found is not None and found[1] >= _MANUFACTURER_SCORE:
        return found[0]
    prefixed = _prefix_manufacturer(folded)
    if prefixed:
        return prefixed
    return cleaned


def _gmc_reading_ok(text: str) -> bool:
    """جي إم سي only from a multi-token reading whose first word is جي.

    Spacing is read before short pieces are glued, so جي ام سي counts and
    the single token جيامسي does not.
    """
    value = unicodedata.normalize("NFKC", text or "")
    for mark in ("\u064b", "\u064c", "\u064d", "\u064e", "\u064f", "\u0650", "\u0651", "\u0652", "\u0670"):
        value = value.replace(mark, "")
    tokens = [part for part in value.split() if any(_is_arabic_letter(char) for char in part)]
    if len(tokens) < 2:
        return False
    return _logical_key(tokens[0]) == "جي"


def _explicit_gmc(text: str) -> bool:
    """Reversed, spaced, or Latin OCR of جي إم سي. A fuzzy short slip is not this."""
    raw = unicodedata.normalize("NFKC", text or "").strip()
    if not raw:
        return False
    if not any(_is_arabic_letter(char) for char in raw):
        return re.sub(r"[^A-Za-z]", "", raw).casefold() == "gmc"
    key = _logical_key(raw)
    known = {
        _logical_key("ام جي"),
        _logical_key("ام جى"),
        _logical_key("يو لو سى ام جى"),
        _logical_key("يولوسي امجي"),
    }
    if key in known:
        return True
    tokens = [_logical_key(part) for part in raw.split() if _logical_key(part)]
    return tokens in (["ام", "جي"], ["جي", "ام"])


def _arabic_letter_count(text: str) -> int:
    collapsed = _collapse_spaced_arabic(text)
    return sum(1 for char in collapsed if _is_arabic_letter(char))


def _short_make_fragment(text: str) -> bool:
    """A 2–3 letter slip such as توو. It is not normalized on its own."""
    if not str(text or "").strip() or _explicit_gmc(text) or _gmc_reading_ok(text):
        return False
    count = _arabic_letter_count(text)
    return 0 < count < 4


_MAKE_FLOOR = 80
_MAKE_MARGIN = 10


def _rank_makes(text: str) -> list[tuple[str, int]]:
    """Official makes scored against one OCR reading. Best first."""
    cleaned = re.sub(r"[^\u0600-\u06FF\s]", "", _collapse_spaced_arabic(text))
    cleaned = _tidy_arabic(cleaned)
    if not cleaned or re.search(r"\d", cleaned):
        return []
    folded = _make_skeleton(cleaned)
    if len(folded) < 2:
        return []
    from thefuzz import fuzz

    allow_gmc = _gmc_reading_ok(text)
    ranked: list[tuple[str, int]] = []
    for name in _MANUFACTURERS:
        if name == "جي إم سي" and not allow_gmc:
            continue
        key = _make_skeleton(name)
        ratio = int(fuzz.ratio(folded, key))
        partial = int(fuzz.partial_ratio(folded, key)) if 3 <= len(folded) <= len(key) else 0
        ranked.append((name, max(ratio, partial)))
    alias = _alias_brand(cleaned)
    if alias == "جي إم سي" and not allow_gmc:
        alias = None
    if alias:
        ranked = [(name, 100 if name == alias else score) for name, score in ranked]
        if alias not in {name for name, _score in ranked}:
            ranked.append((alias, 100))
    ranked.sort(key=lambda item: (-item[1], item[0]))
    return ranked


def _accept_ranked(ranked: list[tuple[str, int]], floor: int, margin: int) -> str:
    """The best name when it clears the floor and beats the runner-up."""
    if not ranked:
        return ""
    best_name, best = ranked[0]
    second = ranked[1][1] if len(ranked) > 1 else 0
    if best >= floor and best - second >= margin:
        return best_name
    return ""


def _format_ranked(ranked: list[tuple[str, int]], limit: int = 3) -> str:
    """تويوتا 100، رينو 67. A zero score is not a candidate."""
    return "، ".join(f"{name} {score}" for name, score in ranked[:limit] if score > 0)


def _accept_make(text: str) -> str:
    """The official make when this reading wins by a margin. Otherwise empty.

    جي إم سي is recognized from its reversed and Latin OCR forms. A fragment
    of two or three letters, such as توو, is never normalized.
    """
    if _explicit_gmc(text) or _gmc_reading_ok(text):
        return "جي إم سي"
    if _short_make_fragment(text):
        return ""
    return _accept_ranked(_rank_makes(text), _MAKE_FLOOR, _MAKE_MARGIN)


def _raw_is_slip(field: str, raw: str, official: str) -> bool:
    """True when this reading is close to one official name and is not that name already.

    توبو and ت و ي و ت ا may become تويوتا. The official spelling itself is a
    reading, not a slip, and the catalog ranking of that name is not a reason
    to write it into an empty cell.
    """
    if not official or not str(raw or "").strip():
        return False
    if str(raw).strip() == official:
        return False
    accepted = _accept_make(raw) if field == "make" else _accept_color(raw)
    kept = _keep_arabic_value(field, raw)
    return accepted == official or kept == official


def _add_reasoned_suggestion(
    table: ExtractedTable,
    page: int,
    plate: str,
    field_name: str,
    value: str,
    reason: str,
) -> None:
    """Append one per-row suggestion. The cell is not changed."""
    text = str(value or "").strip()
    if not plate or not field_name or not text:
        return
    key = _source_key(page, plate, field_name)
    piece = f"{text} — {reason}" if reason else text
    current = str(table.suggestions.get(key) or "")
    have = {
        part.split(" — ", 1)[0].strip()
        for part in re.split(r"[؛;]", current)
        if part.strip()
    }
    if text in have:
        return
    table.suggestions[key] = f"{current}؛ {piece}".strip("؛ ")


def _exact_manufacturer(text: str) -> str:
    """The official make when the text is that name. A resemblance is not accepted."""
    cleaned = _collapse_spaced_arabic(text)
    if not cleaned or re.search(r"\d", cleaned):
        return ""
    known = _alias_brand(cleaned)
    if known == "جي إم سي" and not _gmc_reading_ok(text):
        return ""
    return known or ""


def _normalize_brand(text: str) -> str:
    """The الصنع value. Only an exact dictionary spelling is rewritten."""
    exact = _exact_manufacturer(text)
    return exact if exact else _tidy_arabic(text)


def _fix_make(text: str) -> str:
    """An exact known slip. Fuzzy matching is reserved for the الصنع column."""
    cleaned = _tidy_arabic(text)
    known = _alias_brand(cleaned)
    return known if known is not None else (cleaned or text)


def _manufacturer_column(table: ExtractedTable) -> int | None:
    """The الصنع column, by its header. Other columns are not candidates."""
    for index, name in enumerate(table.columns):
        if _concept_for(name) == "make" or _logical_key(name) in {"الصنع", "صنع"}:
            return index
    return None


def _apply_manufacturer_dictionary(table: ExtractedTable) -> None:
    """
    Clean only the manufacturer column.

    An exact dictionary spelling is written in its official form. Other Arabic
    stays in the cell. A fuzzy resemblance is only a suggestion, so one
    reading cannot become فورد or إيسوزو.
    """
    index = _manufacturer_column(table)
    if index is None:
        return
    notes_at = next((i for i, name in enumerate(table.columns) if _is_notes_header(name)), None)
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    sources = list(table.sources or [])
    for row_index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        while len(row) <= index:
            row.append("")
        original = str(row[index]).strip()
        if not original:
            continue
        page = int(sources[row_index]) if row_index < len(sources) else 0
        plate = _plate_token([row[plate_at]]) if plate_at is not None and plate_at < len(row) else ""
        kept = _keep_arabic_value("make", original)
        if not kept:
            continue
        exact = kept if kept in _MANUFACTURERS else ""
        if exact:
            if exact != original and _raw_is_slip("make", original, exact):
                _append_change(row, notes_at, original, exact)
                row[index] = exact
                _remember_raw(table, page, plate, "make", original)
                _remember_source(table, page, plate, "make", "normalized", f"{exact} — تطبيع: {original}")
            continue
        if kept != original:
            row[index] = kept
        guess = _match_manufacturer(original)
        if guess in _MANUFACTURERS and guess != kept and not _short_make_fragment(original):
            _add_reasoned_suggestion(table, page, plate, "make", guess, "قراءة المقطع")
            if table.provenance.get(_source_key(page, plate, "make")) not in {"agreed", "manual", "normalized"}:
                _remember_source(table, page, plate, "make", "suggested", table.suggestions.get(_source_key(page, plate, "make"), ""))


_SHORT_DIGIT = str.maketrans(
    {
        "O": "0",
        "o": "0",
        "Ο": "0",
        "ο": "0",
        "Q": "0",
        "I": "1",
        "l": "1",
        "|": "1",
        "و": "9",
        "٩": "9",
        "۹": "9",
        "٠": "0",
        "۰": "0",
        "١": "1",
        "۱": "1",
    }
)


def _short_number(text: str) -> str:
    """A 1 to 3 digit admin number read as letters, such as 'و O' for 90."""
    token = "".join(char for char in text.translate(_SHORT_DIGIT) if not char.isspace())
    if token.isdigit() and 1 <= len(token) <= 3:
        return token
    return ""


def _repair_short_numbers(table: ExtractedTable) -> None:
    """In a column of small integers, turn a letter-shaped reading back into digits."""
    width = len(table.columns)
    for index in range(width):
        values = [row[index].strip() if index < len(row) else "" for row in table.rows]
        filled = [value for value in values if value]
        if len(filled) < 4:
            continue
        digits = [value for value in filled if re.fullmatch(r"\d{1,3}", value)]
        if len(digits) < 3 or len(digits) / len(filled) < 0.55:
            continue
        for row in table.rows:
            while len(row) <= index:
                row.append("")
            current = row[index].strip()
            if re.fullmatch(r"\d{1,3}", current):
                continue
            fixed = _short_number(current)
            if fixed:
                row[index] = fixed


def _clean_table(table: ExtractedTable) -> None:
    """Drop empty markers, join a split Arabic word, then correct a shifted plate, year, or date."""
    table.columns = [
        _logical_arabic(_tidy_arabic(name)) if any(_is_arabic_letter(char) for char in name) else name
        for name in table.columns
    ]
    for row in table.rows:
        if _nondata_kind(row, table.columns):
            continue
        for index, value in enumerate(row):
            text = _blank_text(value)
            if _is_printed_page_label(text):
                text = ""
            elif any(_is_arabic_letter(char) for char in text):
                text = _fix_make(_logical_arabic(_tidy_arabic(text)))
            row[index] = text
    _repair_short_numbers(table)
    _repair_shifted_values(table)


def _read_cell(tokens: list[Word], arabic_words: list[Word] | None) -> str:
    """
    Keep the number from the first reader.

    An Arabic cell is replaced when the second reader saw at least as many letters.
    """
    base = _cell_text(tokens)
    if not _arabic_only_text(base):
        return base
    chosen = base
    if arabic_words and tokens:
        x0 = min(word.x0 for word in tokens) - 3
        y0 = min(word.y0 for word in tokens) - 3
        x1 = max(word.x1 for word in tokens) + 3
        y1 = max(word.y1 for word in tokens) + 3
        hits = []
        for word in arabic_words:
            if not _arabic_only_text(word.text):
                continue
            overlap_x = min(x1, word.x1) - max(x0, word.x0)
            overlap_y = min(y1, word.y1) - max(y0, word.y0)
            if overlap_x >= 4 and overlap_y >= 1:
                hits.append(word)
        if hits:
            kept: list[Word] = []
            for word in sorted(hits, key=lambda item: (item.y0, -len(item.text))):
                duplicate = False
                for index, other in enumerate(kept):
                    overlap_x = min(word.x1, other.x1) - max(word.x0, other.x0)
                    overlap_y = min(word.y1, other.y1) - max(word.y0, other.y0)
                    if overlap_x < 4 or overlap_y < 1:
                        continue
                    if len(word.text) > len(other.text):
                        kept[index] = word
                    duplicate = True
                    break
                if not duplicate:
                    kept.append(word)
            base_tidy = _tidy_arabic(base)
            if " " not in base_tidy:
                best = max(
                    kept,
                    key=lambda word: (
                        min(x1, word.x1) - max(x0, word.x0),
                        len(_tidy_arabic(word.text)),
                    ),
                )
                candidate = _tidy_arabic(best.text)
            else:
                kept.sort(key=lambda word: word.x0, reverse=True)
                pieces = [_tidy_arabic(kept[0].text)]
                for word in kept[1:]:
                    piece = _tidy_arabic(word.text)
                    if not piece or any(piece in item or item in piece for item in pieces):
                        continue
                    pieces.append(piece)
                candidate = " ".join(part for part in pieces if part)
            if _arabic_only_text(candidate):
                # A reading that is a real colour wins over a scrambled one.
                # A much longer word is a make, and it is not replaced by a colour.
                base_hit = _match_color(base)
                cand_hit = _match_color(candidate)
                base_len = len(_tidy_arabic(base).replace(" ", ""))
                cand_len = len(_tidy_arabic(candidate).replace(" ", ""))
                if cand_hit and not base_hit and base_len < cand_len + 3:
                    chosen = cand_hit
                elif base_hit and not cand_hit and cand_len < base_len + 3:
                    chosen = base_hit
                elif base_hit and cand_hit:
                    chosen = base_hit
                else:
                    base_letters = sum(1 for char in base if _is_arabic_letter(char))
                    new_letters = sum(1 for char in candidate if _is_arabic_letter(char))
                    if new_letters >= base_letters:
                        chosen = candidate
    return _tidy_arabic(chosen) or base


def _digits_only(text: str) -> bool:
    return bool(re.fullmatch(r"[\d./\-]+", text.strip()))


def _row_cells(row: list[Word], gap: float) -> list[list[Word]]:
    """
    Split a line where a wide empty gap means a new column.

    A year and the colour beside it are different columns even when the gap is
    small. Digits stay with digits, so a split date is still one value.
    """
    ordered = sorted(row, key=lambda word: word.x0)
    cells: list[list[Word]] = [[ordered[0]]]
    for word in ordered[1:]:
        previous = cells[-1][-1]
        space = word.x0 - max(item.x1 for item in cells[-1])
        digit_then_word = _digits_only(previous.text) and any(_is_arabic_letter(char) for char in word.text)
        word_then_digit = any(_is_arabic_letter(char) for char in previous.text) and _digits_only(word.text)
        if space > gap or (space > 5 and (digit_then_word or word_then_digit)):
            cells.append([word])
        else:
            cells[-1].append(word)
    return cells


def _peak_centers(boxes: list[tuple[float, float, float]], page_width: float) -> list[float]:
    """
    Column positions are the peaks where cell centers pile up.

    A sparse column, such as a note that is filled on only some invoice lines,
    still counts. A one-off smudge does not.
    """
    if not boxes:
        return []
    bin_size = 4.0
    count = int(page_width / bin_size) + 2
    hist = [0] * count
    for center, _x0, _x1 in boxes:
        index = int(center / bin_size)
        if 0 <= index < count:
            hist[index] += 1
    strongest = max(hist) or 1
    raw: list[tuple[float, float]] = []
    for index in range(1, count - 1):
        if hist[index] < 1:
            continue
        if hist[index] < hist[index - 1] or hist[index] < hist[index + 1]:
            continue
        raw.append((index * bin_size + bin_size / 2.0, float(hist[index])))
    merged: list[tuple[float, float]] = []
    for center, weight in raw:
        if not merged:
            merged.append((center, weight))
            continue
        previous_center, previous_weight = merged[-1]
        # Pieces of one word sit a few points apart. Real columns are much further apart.
        close = center - previous_center <= 22
        weak = min(weight, previous_weight) < max(3.0, strongest * 0.08)
        if close or (center - previous_center <= 30 and weak):
            total = previous_weight + weight
            merged[-1] = ((previous_center * previous_weight + center * weight) / total, total)
            continue
        merged.append((center, weight))
    return [center for center, _weight in merged]


def _ranges_from_centers(centers: list[float]) -> list[tuple[float, float]]:
    ranges: list[tuple[float, float]] = []
    for index, center in enumerate(centers):
        if index == 0:
            left = center - (centers[1] - center) / 2.0 if len(centers) > 1 else center - 40
        else:
            left = (centers[index - 1] + center) / 2.0
        if index == len(centers) - 1:
            right = center + (center - centers[index - 1]) / 2.0 if len(centers) > 1 else center + 40
        else:
            right = (center + centers[index + 1]) / 2.0
        ranges.append((left, max(right, left + 8)))
    return ranges


def _fold_letter_slivers(
    names: list[str], data: list[list[str]], centers: list[float]
) -> tuple[list[str], list[list[str]], list[float]]:
    """A one-letter column is a word the reader split, so it joins the neighbouring word."""
    slivers: set[int] = set()
    for index in range(len(centers)):
        values = [row[index].strip() for row in data if row[index].strip()]
        if len(values) < 3:
            continue
        if any(any(char.isdigit() for char in value) for value in values):
            continue
        nearest = 999.0
        if index:
            nearest = min(nearest, abs(centers[index] - centers[index - 1]))
        if index + 1 < len(centers):
            nearest = min(nearest, abs(centers[index + 1] - centers[index]))
        if nearest > 32:
            continue
        if sum(len(value) for value in values) / len(values) > 2.2:
            continue
        slivers.add(index)
    if not slivers:
        return names, data, centers
    for index in sorted(slivers):
        neighbours = [other for other in (index - 1, index + 1) if 0 <= other < len(centers) and other not in slivers]
        if not neighbours:
            continue
        target = max(
            neighbours,
            key=lambda other: sum(
                1
                for row in data[:20]
                for char in row[other]
                if char.isalpha() or _is_arabic_letter(char)
            ),
        )
        for row in data:
            piece = row[index].strip()
            if not piece:
                continue
            host = row[target].strip()
            if host and (_is_arabic_letter(piece[0]) or _is_arabic_letter(host[0])):
                row[target] = piece + host if centers[index] > centers[target] else host + piece
            else:
                row[target] = f"{piece} {host}".strip() if centers[index] < centers[target] else f"{host} {piece}".strip()
            row[index] = ""
    keep = [index for index in range(len(centers)) if index not in slivers]
    return (
        [names[index] for index in keep],
        [[row[index] for index in keep] for row in data],
        [centers[index] for index in keep],
    )


def _table_from_positioned_words(
    words: list[Word],
    page_width: float,
    page_index: int,
    arabic_words: list[Word] | None = None,
    band_plates: list[Word] | None = None,
    band_colors: list[Word] | None = None,
) -> ExtractedTable | None:
    """
    Build a table from word positions only.

    A column is a vertical band that the lines share. The header is the label
    line just above the first band of values, including a header that wraps
    onto a second line. Empty cells stay empty, so a missing plate purpose
    does not shift the colour into the wrong column.
    """
    usable = []
    for word in words:
        text = word.text.strip()
        if not text:
            continue
        if not re.search(r"[\dA-Za-z\u0600-\u06FF-]", text):
            continue
        usable.append(word)
    if len(usable) < 8:
        return None
    heights = [max(1.0, word.y1 - word.y0) for word in usable]
    # Tight enough that the next printed line stays separate. Wide enough for one line.
    # Letters on one printed line can sit a few points above the digits. Keep them together.
    tolerance = max(4.2, min(5.2, _median(heights) * 0.6))
    rows = _split_double_plate_lines(_group_rows(usable, tolerance))
    if len(rows) < 3:
        return None
    gap = 16.0

    def is_candidate(row: list[Word]) -> bool:
        # A line that already shows a plate is a vehicle even when the other
        # cells were read as few words. The footer gap must not drop it.
        if _plate_token([word.text for word in row]):
            return True
        if len(row) < 4:
            return False
        span = max(word.x1 for word in row) - min(word.x0 for word in row)
        if span < page_width * 0.42:
            return False
        numbers = sum(1 for word in row if re.search(r"\d", word.text))
        return numbers >= 2

    candidates = [row for row in rows if is_candidate(row)]
    if len(candidates) < 3:
        return None
    pitches = [
        candidates[index][0].center_y - candidates[index - 1][0].center_y
        for index in range(1, len(candidates))
        if 2 < candidates[index][0].center_y - candidates[index - 1][0].center_y < 40
    ]
    pitch = _median(pitches) or 12.0
    body = []
    for index, row in enumerate(candidates):
        previous_y = candidates[index - 1][0].center_y if index else None
        next_y = candidates[index + 1][0].center_y if index + 1 < len(candidates) else None
        nearest = 999.0
        if previous_y is not None:
            nearest = min(nearest, row[0].center_y - previous_y)
        if next_y is not None:
            nearest = min(nearest, next_y - row[0].center_y)
        limit = pitch * 4.2 if _plate_token([word.text for word in row]) else pitch * 2.6
        if nearest <= limit:
            body.append(row)
    if len(body) < 3:
        body = candidates
    first_y = min(word.center_y for word in body[0])
    label_rows: list[list[Word]] = []
    for row in rows:
        row_y = sum(word.center_y for word in row) / len(row)
        if row_y >= first_y - 1:
            break
        if first_y - row_y > pitch * 3.4:
            continue
        letters = sum(1 for word in row for char in word.text if char.isalpha() or _is_arabic_letter(char))
        numbers = sum(1 for word in row if re.search(r"\d", word.text))
        if letters >= 2 and numbers < 2:
            label_rows.append(row)
    label_rows = label_rows[-3:]

    boxes: list[tuple[float, float, float]] = []
    for row in body + label_rows:
        for cell in _row_cells(row, gap):
            x0 = min(word.x0 for word in cell)
            x1 = max(word.x1 for word in cell)
            boxes.append(((x0 + x1) / 2.0, x0, x1))
    if len(boxes) < 6:
        return None
    centers = _peak_centers(boxes, page_width)
    if len(centers) < 2:
        return None
    ranges = _ranges_from_centers(centers)

    def column_of(word: Word) -> int | None:
        for index, (left, right) in enumerate(ranges):
            if left <= word.center_x <= right:
                return index
        nearest = min(range(len(ranges)), key=lambda index: abs(word.center_x - centers[index]))
        if abs(word.center_x - centers[nearest]) <= 18:
            return nearest
        return None

    first_y = min(word.center_y for word in body[0])
    header_rows = []
    for row in rows:
        row_y = sum(word.center_y for word in row) / len(row)
        if row_y >= first_y - 1:
            break
        if first_y - row_y > pitch * 3.4:
            continue
        letters = sum(1 for word in row for char in word.text if char.isalpha() or _is_arabic_letter(char))
        numbers = sum(1 for word in row if re.search(r"\d", word.text))
        if letters >= 2 and numbers < 2:
            header_rows.append(row)
    header_rows = header_rows[-3:]

    def column_of_span(x0: float, x1: float) -> int | None:
        center = (x0 + x1) / 2.0
        for index, (left, right) in enumerate(ranges):
            if left <= center <= right:
                return index
        nearest = min(range(len(ranges)), key=lambda index: abs(center - centers[index]))
        if abs(center - centers[nearest]) <= 22:
            return nearest
        return None

    names_tokens: list[list[Word]] = [[] for _ in ranges]
    for row in header_rows:
        for cell in _row_cells(row, gap):
            index = column_of_span(min(word.x0 for word in cell), max(word.x1 for word in cell))
            if index is not None:
                names_tokens[index].extend(cell)
    names = _unique_headers([_read_cell(tokens, arabic_words) for tokens in names_tokens])

    data: list[list[str]] = []
    loose_lines: list[list[str]] = []
    row_ys: list[float] = []
    for row in body:
        buckets: list[list[Word]] = [[] for _ in ranges]
        loose: list[str] = []
        for cell in _row_cells(row, gap):
            index = column_of_span(min(word.x0 for word in cell), max(word.x1 for word in cell))
            if index is not None:
                buckets[index].extend(cell)
                continue
            # Layout shift: this token sits outside every column band. Keep it
            # so the plate regex can still read a 4 or 5 digit number later.
            orphan = _read_cell(cell, arabic_words).strip()
            if orphan:
                loose.append(orphan)
        values = [_read_cell(bucket, arabic_words) for bucket in buckets]
        if _is_noise_row(values):
            continue
        filled = [cell for cell in values if cell]
        has_plate = bool(_plate_token(values + loose))
        if not has_plate and len(filled) < 3:
            continue
        if not has_plate and sum(1 for cell in filled if re.search(r"\d", cell)) < 2:
            continue
        if tuple(_normalize(cell) for cell in values) == tuple(_normalize(name) for name in names):
            continue
        data.append(values)
        loose_lines.append(loose)
        row_ys.append(sum(word.center_y for word in row) / len(row))
    if not data:
        return None
    # A plate the first reader skipped is filled from a digit-only re-read of
    # the same column, matched to this row by its height on the page.
    if band_plates:
        _apply_band_plates(data, row_ys, loose_lines, band_plates, words)
        _splice_missed_plates(data, row_ys, loose_lines, band_plates, words)
    if band_colors:
        _apply_band_colors(data, row_ys, loose_lines, band_colors)

    for index, column_rows in enumerate(zip(*data)):
        filled = [cell for cell in column_rows if cell.strip()]
        if len(filled) < 3:
            continue
        numeric = 0
        for cell in filled:
            mapped = cell.translate(_LOOKALIKE)
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", mapped) or re.fullmatch(r"[-+]?\d[\d.,]*", mapped):
                numeric += 1
        if numeric < max(2, int(len(filled) * 0.65)):
            continue
        mostly_dates = sum(1 for cell in filled if re.fullmatch(r"\d{4}-\d{2}-\d{2}", cell)) >= len(filled) * 0.5
        for row in data:
            mapped = row[index].translate(_LOOKALIKE)
            if mostly_dates:
                salvaged = _assemble_date(re.findall(r"\d+", mapped))
                if salvaged:
                    row[index] = salvaged
                    continue
            if mapped and re.fullmatch(r"[0-9.\-]+", mapped):
                row[index] = mapped

    names, data, centers = _fold_letter_slivers(names, data, centers)
    # Full line text, taken before a sparse column is removed. A plate that
    # shifted into that column is still available to the regex fallback.
    lines = [
        " ".join([cell.strip() for cell in row if cell and cell.strip()] + loose)
        for row, loose in zip(data, loose_lines)
    ]
    if data and centers:
        kept = []
        minimum = max(2, int(len(data) * 0.04))
        for index in range(len(centers)):
            filled = sum(1 for row in data if index < len(row) and row[index].strip())
            placeholder = bool(re.fullmatch(r"Column \d+(?: \d+)?", names[index].strip()))
            # One real plate is enough to keep the column. A later page can
            # shift that column so only a few rows land in it.
            holds_plate = any(_plate_token([row[index] if index < len(row) else ""]) for row in data)
            if placeholder and filled < max(4, int(len(data) * 0.2)) and not holds_plate:
                continue
            if filled >= minimum or holds_plate:
                kept.append(index)
        if len(kept) >= 2:
            names = [names[index] for index in kept]
            data = [[row[index] if index < len(row) else "" for index in kept] for row in data]
            centers = [centers[index] for index in kept]
    sample = " ".join(names) + " " + " ".join(" ".join(row) for row in data[:8])
    arabic, latin = _letter_counts(sample)
    anchors = list(centers)
    if arabic > latin:
        names = list(reversed(names))
        data = [list(reversed(row)) for row in data]
        anchors = list(reversed(anchors))
    parsed = ExtractedTable(names, data, anchors, [page_index] * len(data))
    parsed.lines = lines
    return parsed


def _ocr_engine(language: str):
    cached = _OCR_ENGINES.get(language)
    if cached is not None:
        return cached
    try:
        from winrt.windows.globalization import Language
        from winrt.windows.media.ocr import OcrEngine
    except Exception as exc:
        raise ConversionError(
            "This PDF is a scanned picture. Install the PDF tools again so Windows OCR can read it."
        ) from exc
    engine = OcrEngine.try_create_from_language(Language(language))
    if engine is None:
        return None
    _OCR_ENGINES[language] = engine
    return engine


def _prepare_ocr_apartment() -> None:
    """
    OCR has to run off the window thread.

    The window owns the single-threaded apartment, and a blocking read there
    never returns. A worker thread can use the multithreaded apartment.
    """
    try:
        from winrt.runtime import ApartmentType, init_apartment

        init_apartment(ApartmentType.MTA)
    except Exception:
        pass


_MASK_MARKUP = True


def _without_markup(image_bytes: bytes) -> bytes:
    """Whiten pink highlighter and blue stamp ink before OCR. Black text stays."""
    if not _MASK_MARKUP:
        return image_bytes
    import io

    import numpy as np
    from PIL import Image

    try:
        image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    except Exception:
        return image_bytes
    try:
        hsv = np.asarray(image.convert("HSV"))
        tone = hsv[:, :, 0]
        saturation = hsv[:, :, 1]
        value = hsv[:, :, 2]
        pink = (saturation > 80) & (value > 110) & ((tone < 18) | (tone > 225))
        blue = (saturation > 80) & (value > 70) & (tone >= 135) & (tone <= 175)
        mask = pink | blue
        if not bool(mask.any()):
            return image_bytes
        rgb = np.asarray(image).copy()
        rgb[mask] = (255, 255, 255)
        cleaned = Image.fromarray(rgb, "RGB")
        output = io.BytesIO()
        try:
            cleaned.save(output, format="PNG")
            return output.getvalue()
        finally:
            cleaned.close()
    finally:
        image.close()


def _recognize_image(image_bytes: bytes, language: str):
    image_bytes = _without_markup(image_bytes)
    from winrt.windows.graphics.imaging import BitmapAlphaMode, BitmapDecoder, BitmapPixelFormat
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

    _prepare_ocr_apartment()
    engine = _ocr_engine(language)
    if engine is None:
        return None
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(image_bytes)
    writer.store_async().get()
    writer.detach_stream()
    stream.seek(0)
    decoder = BitmapDecoder.create_async(stream).get()
    bitmap = decoder.get_software_bitmap_async(BitmapPixelFormat.BGRA8, BitmapAlphaMode.PREMULTIPLIED).get()
    return engine.recognize_async(bitmap).get()


def _scan_image(page: fitz.Page) -> tuple[bytes, fitz.Rect, float, float]:
    """
    Render the page for OCR.

    The embedded scan is used as the picture, but a fresh render keeps each
    printed line on one baseline. The long side stays under 4500 pixels.
    """
    scale = 3.0
    long_side = max(page.rect.width, page.rect.height) * scale
    if long_side > 4500:
        scale = 4500 / max(page.rect.width, page.rect.height)
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    png = pixmap.tobytes("png")
    width = float(pixmap.width)
    height = float(pixmap.height)
    del pixmap
    return png, fitz.Rect(page.rect), width, height


def _words_from_result(result, rect: fitz.Rect, pixel_width: float, pixel_height: float) -> list[Word]:
    words: list[Word] = []
    if result is None:
        return words
    for line in result.lines:
        for item in line.words:
            text = _clean_cell(getattr(item, "text", ""))
            if not text:
                continue
            box = item.bounding_rect
            x0 = rect.x0 + (float(box.x) / pixel_width) * rect.width
            y0 = rect.y0 + (float(box.y) / pixel_height) * rect.height
            x1 = rect.x0 + ((float(box.x) + float(box.width)) / pixel_width) * rect.width
            y1 = rect.y0 + ((float(box.y) + float(box.height)) / pixel_height) * rect.height
            raw_confidence = getattr(item, "confidence", 0) or 0
            try:
                confidence = float(raw_confidence)
            except (TypeError, ValueError):
                confidence = 0.0
            if 0 < confidence <= 1:
                confidence *= 100.0
            words.append(Word(x0, y0, x1, y1, text, confidence))
    return words


def _score_words(words: list[Word]) -> tuple[int, int]:
    sample = " ".join(word.text for word in words)
    return _letter_counts(sample)


def _recognize_rendered(
    image: bytes,
    rect: fitz.Rect,
    pixel_width: float,
    pixel_height: float,
    language_box: dict[str, str],
) -> list[Word]:
    """Read one rendered page. The language is chosen once, from the first page."""
    language = language_box.get("language", "")
    if not language:
        arabic_words = _ocr_words(image, "ar-SA", rect, pixel_width, pixel_height)
        arabic, latin = _score_words(arabic_words)
        if arabic >= latin or arabic >= 8:
            language_box["language"] = "ar-SA"
            return arabic_words
        english_words = _ocr_words(image, "en-US", rect, pixel_width, pixel_height)
        language_box["language"] = "en-US" if english_words else "ar-SA"
        return english_words or arabic_words
    return _ocr_words(image, language, rect, pixel_width, pixel_height)


def _tesseract_exe() -> str | None:
    try:
        from engine.tesseract_bin import find_exe

        return find_exe()
    except Exception:
        for candidate in (
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ):
            if Path(candidate).is_file():
                return candidate
        return shutil.which("tesseract")


def _tessdata_dir() -> Path:
    """Where the existing second read looks. A fallback may pass its own folder."""
    try:
        from engine.tesseract_bin import legacy_tessdata_dir

        return legacy_tessdata_dir(__file__)
    except Exception:
        return Path(__file__).resolve().parent / "tessdata"


def _tesseract_words(
    image: bytes,
    rect: fitz.Rect,
    pixel_width: float,
    pixel_height: float,
    language: str = "ara",
    data_dir: Path | None = None,
) -> list[Word]:
    """Second reading. ara+eng is the alternate engine on a retry."""
    image = _without_markup(image)
    exe = _tesseract_exe()
    folder = data_dir or _tessdata_dir()
    if exe is None or not (folder / "ara.traineddata").is_file():
        return []
    picture = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    try:
        picture.write(image)
        picture.close()
        result = subprocess.run(
            [
                exe,
                picture.name,
                "stdout",
                "-l",
                language,
                "--tessdata-dir",
                str(folder),
                "--psm",
                "6",
                "tsv",
            ],
            capture_output=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return []
    finally:
        Path(picture.name).unlink(missing_ok=True)
    text = result.stdout.decode("utf-8", errors="replace")
    words: list[Word] = []
    scale_x = rect.width / pixel_width
    scale_y = rect.height / pixel_height
    for line in text.splitlines()[1:]:
        parts = line.split("\t")
        if len(parts) < 12:
            continue
        if parts[0] != "5":
            continue
        raw = parts[11].strip()
        if not raw:
            continue
        try:
            confidence = float(parts[10])
            left, top, width, height = (float(parts[6]), float(parts[7]), float(parts[8]), float(parts[9]))
        except ValueError:
            continue
        if confidence < 20 or width < 1 or height < 1:
            continue
        cleaned = _clean_cell(raw)
        if not cleaned:
            continue
        words.append(
            Word(
                rect.x0 + left * scale_x,
                rect.y0 + top * scale_y,
                rect.x0 + (left + width) * scale_x,
                rect.y0 + (top + height) * scale_y,
                cleaned,
                confidence,
            )
        )
    return words


def _windows_language_missing(language: str) -> bool:
    """True only when Windows OCR cannot be created. An empty page is not missing."""
    try:
        return _ocr_engine(language) is None
    except Exception:
        return True


def _ocr_words(
    image: bytes,
    language: str,
    rect: fitz.Rect,
    pixel_width: float,
    pixel_height: float,
) -> list[Word]:
    """
    Windows OCR when that language is installed.

    The word list is unchanged in that case. Tesseract ara is used only when
    Windows has no engine for the language, so a PC without the language pack
    can still read the scan.
    """
    try:
        words = _words_from_result(_recognize_image(image, language), rect, pixel_width, pixel_height)
    except ConversionError:
        words = []
    if words or not _windows_language_missing(language):
        return words
    try:
        from engine.tesseract_bin import ara_tessdata_dir

        folder = ara_tessdata_dir()
    except Exception:
        folder = None
    if folder is None:
        return words
    tess_lang = "ara+eng" if language == "ar-SA" else "eng"
    return _tesseract_words(image, rect, pixel_width, pixel_height, tess_lang, folder)


def _image_kind(data: bytes) -> str | None:
    """Container Excel can store without recompressing the pixels."""
    if len(data) < 12:
        return None
    if data.startswith(b"\xff\xd8\xff"):
        return "jpeg"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith((b"GIF87a", b"GIF89a")):
        return "gif"
    return None


def _covers_page(rect: fitz.Rect, page: fitz.Page) -> bool:
    return rect.width >= page.rect.width * 0.9 and rect.height >= page.rect.height * 0.9


def _blank_embedded_figures(page: fitz.Page) -> None:
    """
    Cover photos and charts before a scan is read.

    OCR is only for words. A picture that is not the page itself is painted
    white here so its pixels cannot become a column. The original bytes are
    saved separately and are not changed by this paint.
    """
    try:
        infos = page.get_image_info() or []
    except Exception:
        return
    for info in infos:
        raw = info.get("bbox")
        if not raw:
            continue
        rect = fitz.Rect(raw)
        if rect.width < 24 or rect.height < 24 or _covers_page(rect, page):
            continue
        try:
            page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), width=0, overlay=True)
        except Exception:
            continue


def _dominant_plate_length(words: list[Word]) -> int:
    """The plate width this page actually uses. A year is not a plate."""
    counts: dict[int, int] = {}
    for word in words:
        token = re.sub(r"\D", "", word.text.translate(_INDIC_DIGITS))
        if (
            token.isdigit()
            and 4 <= len(token) <= 6
            and not token.startswith("0")
            and not _YEAR_VALUE.fullmatch(token)
        ):
            counts[len(token)] = counts.get(len(token), 0) + 1
    if not counts:
        return 5
    length = max(counts, key=lambda item: counts[item])
    if counts[length] >= max(3, int(0.55 * sum(counts.values()))):
        return length
    return 5


def _plate_band_words(
    image: bytes,
    rect: fitz.Rect,
    pixel_width: float,
    pixel_height: float,
    words: list[Word],
) -> list[Word]:
    """
    Read the plate column again, as digits only.

    The first reader misses some plates when that column sits a little apart
    from the rest of the line. The column is not a fixed coordinate: it is the
    vertical band where this page's own 5-digit plates were already found.
    A second read of that band, restricted to the digits 0-9, fills the gaps.
    """
    plate_length = _dominant_plate_length(words)
    known = [word for word in words if len(_plate_token([word.text])) == plate_length]
    if len(known) < 3:
        return []
    exe = _tesseract_exe()
    if exe is None:
        return []
    left_pt = _median([word.x0 for word in known]) - 14
    right_pt = _median([word.x1 for word in known]) + 14
    scale_x = pixel_width / rect.width
    left = max(0, int(left_pt * scale_x))
    right = min(int(pixel_width), int(right_pt * scale_x))
    if right - left < 8:
        return []
    try:
        from PIL import Image
    except ImportError:
        return []
    try:
        picture = Image.open(BytesIO(image))
        crop = picture.crop((left, 0, right, picture.height))
        handle = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        crop.save(handle.name)
        handle.close()
        result = subprocess.run(
            [
                exe,
                handle.name,
                "stdout",
                "-l",
                "eng",
                "--psm",
                "6",
                "-c",
                "tessedit_char_whitelist=0123456789",
                "tsv",
            ],
            capture_output=True,
            timeout=60,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return []
    finally:
        if "crop" in locals():
            crop.close()
        if "picture" in locals():
            picture.close()
        if "handle" in locals():
            Path(handle.name).unlink(missing_ok=True)
    scale_y = rect.height / pixel_height
    found: list[Word] = []
    for line in result.stdout.decode("utf-8", errors="replace").splitlines()[1:]:
        parts = line.split("\t")
        if len(parts) < 12 or parts[0] != "5" or not parts[11].strip():
            continue
        try:
            confidence = float(parts[10])
            top, width, height = float(parts[7]), float(parts[8]), float(parts[9])
        except ValueError:
            continue
        token = re.sub(r"\D", "", parts[11])
        # Low-confidence scraps and years are not plates. The length is the page's own.
        if confidence < 80 or len(token) != plate_length or token.startswith("0") or _YEAR_VALUE.fullmatch(token):
            continue
        y0 = rect.y0 + top * scale_y
        y1 = rect.y0 + (top + height) * scale_y
        x0 = rect.x0 + (left + float(parts[6])) * (rect.width / pixel_width)
        x1 = x0 + width * (rect.width / pixel_width)
        found.append(Word(x0, y0, x1, y1, token))
    return found


def _splice_missed_plates(
    data: list[list[str]],
    row_ys: list[float],
    loose_lines: list[list[str]],
    band_plates: list[Word],
    words: list[Word],
) -> None:
    """
    Add a row for a plate that was read in the plate column but never got a line.

    The row keeps that plate and the words on the same baseline. A plate that
    already sits on a row is not added again.
    """
    if not data or not row_ys or not band_plates:
        return
    present = {token for row in data if (token := _plate_token(row))}
    ordered = sorted(row_ys)
    pitches = [b - a for a, b in zip(ordered, ordered[1:]) if 2 < b - a < 40]
    pitch = _median(pitches) or 12.0
    width = max(len(row) for row in data)
    counts = [
        sum(1 for row in data if index < len(row) and _plate_token([row[index]]))
        for index in range(width)
    ]
    plate_col = counts.index(max(counts)) if counts and max(counts) else 0
    additions: list[tuple[float, list[str], list[str]]] = []
    for band in band_plates:
        if band.text in present:
            continue
        nearest = min(range(len(row_ys)), key=lambda index: abs(row_ys[index] - band.center_y))
        gap = abs(row_ys[nearest] - band.center_y)
        if gap <= 12 and not _plate_token(data[nearest]):
            continue
        if gap > max(pitch * 2.4, 18):
            continue
        values = [""] * width
        values[plate_col] = band.text
        loose = [
            word.text
            for word in words
            if abs(word.center_y - band.center_y) <= max(8.0, pitch * 0.85) and word.text != band.text
        ]
        additions.append((band.center_y, values, loose))
        present.add(band.text)
    for y_value, values, loose in sorted(additions):
        insert_at = 0
        while insert_at < len(row_ys) and row_ys[insert_at] < y_value:
            insert_at += 1
        data.insert(insert_at, values)
        row_ys.insert(insert_at, y_value)
        loose_lines.insert(insert_at, loose)


def _split_double_plate_lines(rows: list[list[Word]]) -> list[list[Word]]:
    """A line that holds two different plates is two vehicles, one under the other."""
    output: list[list[Word]] = []
    for row in rows:
        grouped: dict[str, list[Word]] = {}
        for word in row:
            token = _plate_token([word.text])
            if token:
                grouped.setdefault(token, []).append(word)
        if len(grouped) < 2:
            output.append(row)
            continue
        centers = {
            token: sum(word.center_y for word in items) / len(items) for token, items in grouped.items()
        }
        buckets: dict[str, list[Word]] = {token: list(items) for token, items in grouped.items()}
        for word in row:
            if _plate_token([word.text]):
                continue
            nearest = min(centers, key=lambda token: abs(word.center_y - centers[token]))
            buckets[nearest].append(word)
        for token in sorted(buckets, key=lambda item: centers[item]):
            output.append(buckets[token])
    return output


def _scan_page_facts(words: list[Word], band_plates: list[Word]) -> tuple[int, list[str]]:
    """How many plates this page showed, and the subtotal lines that are not vehicles."""
    band_left = min((word.x0 for word in band_plates), default=None)
    band_right = max((word.x1 for word in band_plates), default=None)
    seen: list[tuple[float, str]] = []
    for word in list(words) + list(band_plates):
        token = _plate_token([word.text])
        if not token:
            continue
        if band_left is not None and not (band_left - 10 <= word.center_x <= band_right + 10):
            continue
        if any(abs(word.center_y - y_value) <= 9 and token == old for y_value, old in seen):
            continue
        seen.append((word.center_y, token))
    if seen:
        prefixes: dict[str, int] = {}
        for _y_value, token in seen:
            prefixes[token[:2]] = prefixes.get(token[:2], 0) + 1
        dominant = max(prefixes, key=lambda item: prefixes[item])
        seen = [
            item
            for item in seen
            if item[1].startswith(dominant) or prefixes.get(item[1][:2], 0) >= 4
        ]
    statements: list[str] = []
    for line in _group_rows(words, 5):
        text = " ".join(word.text for word in line)
        folded = _normalize(text).replace(" ", "")
        plated = bool(_plate_token([word.text for word in line]))
        if plated:
            continue
        mentions_total = any(token in folded for token in ("اجمالي", "جمالي", "المجموع")) or (
            "عدد" in folded and ("مركبات" in folded or "سيارات" in folded)
        )
        if mentions_total or (_purpose_token([text]) and re.search(r"\d", text) and len(line) <= 6):
            statements.append(text)
    return len(seen), statements


def _apply_band_plates(
    data: list[list[str]],
    row_ys: list[float],
    loose_lines: list[list[str]],
    band_plates: list[Word],
    words: list[Word],
) -> None:
    """Put a re-read plate into the empty plate cell of the row at the same height."""
    width = max((len(row) for row in data), default=0)
    counts = [
        sum(1 for row in data if index < len(row) and _plate_token([row[index]]))
        for index in range(width)
    ]
    plate_col = counts.index(max(counts)) if counts and max(counts) >= 1 else None
    windows: dict[str, float] = {}
    for word in words:
        token = _plate_token([word.text])
        if token and token not in windows:
            windows[token] = word.center_y
    deltas = [band.center_y - windows[band.text] for band in band_plates if band.text in windows]
    offset = _median(deltas) if deltas else 0.0
    present = {token for row in data if (token := _plate_token(row))}
    pairs: list[tuple[float, int, int]] = []
    for row_index, row in enumerate(data):
        if _plate_token(row):
            continue
        for band_index, band in enumerate(band_plates):
            if band.text in present:
                continue
            gap = abs((band.center_y - offset) - row_ys[row_index])
            if gap <= 12:
                pairs.append((gap, row_index, band_index))
    pairs.sort()
    used_rows: set[int] = set()
    used_bands: set[int] = set()
    for _gap, row_index, band_index in pairs:
        if row_index in used_rows or band_index in used_bands:
            continue
        text = band_plates[band_index].text
        if text in present:
            continue
        used_rows.add(row_index)
        used_bands.add(band_index)
        present.add(text)
        row = data[row_index]
        if plate_col is not None:
            while len(row) <= plate_col:
                row.append("")
            if not _plate_token([row[plate_col]]):
                row[plate_col] = text
        loose_lines[row_index].append(text)


def _color_band_words(
    image: bytes,
    rect: fitz.Rect,
    pixel_width: float,
    pixel_height: float,
    words: list[Word],
) -> list[Word]:
    """
    Read the colour column again.

    The first reader misses some colours when that word sits a little off the
    row. The band is the x range where this page already found colours.
    """
    known = [word for word in words if _match_color(word.text)]
    if len(known) < 3:
        return []
    exe = _tesseract_exe()
    data_dir = Path(__file__).resolve().parent / "tessdata"
    if exe is None or not (data_dir / "ara.traineddata").is_file():
        return []
    left_pt = _median([word.x0 for word in known]) - 16
    right_pt = _median([word.x1 for word in known]) + 16
    scale_x = pixel_width / rect.width
    left = max(0, int(left_pt * scale_x))
    right = min(int(pixel_width), int(right_pt * scale_x))
    if right - left < 8:
        return []
    try:
        from PIL import Image
    except ImportError:
        return []
    try:
        picture = Image.open(BytesIO(image))
        crop = picture.crop((left, 0, right, picture.height))
        handle = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        crop.save(handle.name)
        handle.close()
        result = subprocess.run(
            [
                exe,
                handle.name,
                "stdout",
                "-l",
                "ara",
                "--tessdata-dir",
                str(data_dir),
                "--psm",
                "6",
                "tsv",
            ],
            capture_output=True,
            timeout=60,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return []
    finally:
        if "crop" in locals():
            crop.close()
        if "picture" in locals():
            picture.close()
        if "handle" in locals():
            Path(handle.name).unlink(missing_ok=True)
    scale_y = rect.height / pixel_height
    scale_back = rect.width / pixel_width
    found: list[Word] = []
    for line in result.stdout.decode("utf-8", errors="replace").splitlines()[1:]:
        parts = line.split("\t")
        if len(parts) < 12 or parts[0] != "5" or not parts[11].strip():
            continue
        try:
            confidence = float(parts[10])
            box_left, top, width, height = (float(parts[6]), float(parts[7]), float(parts[8]), float(parts[9]))
        except ValueError:
            continue
        if confidence < 30 or width < 1 or height < 1:
            continue
        color = _match_color(_clean_cell(parts[11]))
        if not color:
            continue
        y0 = rect.y0 + top * scale_y
        y1 = rect.y0 + (top + height) * scale_y
        x0 = rect.x0 + (left + box_left) * scale_back
        x1 = x0 + width * scale_back
        found.append(Word(x0, y0, x1, y1, color))
    return found


def _apply_band_colors(
    data: list[list[str]],
    row_ys: list[float],
    loose_lines: list[list[str]],
    band_colors: list[Word],
) -> None:
    """Put a re-read colour into the empty colour cell of the row at the same height."""
    width = max((len(row) for row in data), default=0)
    counts = [
        sum(1 for row in data if index < len(row) and _match_color(row[index]))
        for index in range(width)
    ]
    if not counts or max(counts) < 1:
        return
    color_col = counts.index(max(counts))
    deltas = []
    for row_index, row in enumerate(data):
        current = _match_color(row[color_col]) if color_col < len(row) else ""
        if not current:
            continue
        nearest = min(band_colors, key=lambda word: abs(word.center_y - row_ys[row_index]))
        if nearest.text == current and abs(nearest.center_y - row_ys[row_index]) < 18:
            deltas.append(nearest.center_y - row_ys[row_index])
    offset = _median(deltas) if deltas else 0.0
    pairs: list[tuple[float, int, int]] = []
    for row_index, row in enumerate(data):
        while len(row) <= color_col:
            row.append("")
        if _match_color(row[color_col]):
            continue
        for band_index, band in enumerate(band_colors):
            gap = abs((band.center_y - offset) - row_ys[row_index])
            if gap <= 7:
                pairs.append((gap, row_index, band_index))
    pairs.sort()
    used_rows: set[int] = set()
    used_bands: set[int] = set()
    for _gap, row_index, band_index in pairs:
        if row_index in used_rows or band_index in used_bands:
            continue
        used_rows.add(row_index)
        used_bands.add(band_index)
        data[row_index][color_col] = band_colors[band_index].text
        loose_lines[row_index].append(band_colors[band_index].text)


def _tables_from_scan(page: fitz.Page, page_index: int, language_box: dict[str, str]) -> list[ExtractedTable]:
    _blank_embedded_figures(page)
    image, rect, pixel_width, pixel_height = _scan_image(page)
    with _stage(f"page_{page_index + 1}_ocr"):
        words = _recognize_rendered(image, rect, pixel_width, pixel_height, language_box)
    # The second engine is not run on the whole page. A weak cell asks for it later.
    arabic_words: list[Word] = []
    with _stage(f"page_{page_index + 1}_plate_band"):
        band_plates = _plate_band_words(image, rect, pixel_width, pixel_height, words)
    with _stage(f"page_{page_index + 1}_color_band"):
        band_colors = _color_band_words(image, rect, pixel_width, pixel_height, words)
    detected, statements = _scan_page_facts(words, band_plates)
    language_box.setdefault("detected_pages", {})[int(page_index) + 1] = detected
    language_box.setdefault("statements", []).extend(statements)
    table = _table_from_positioned_words(
        words, page.rect.width, page_index, arabic_words, band_plates, band_colors
    )
    return [table] if table is not None else []


def _table_score(table: ExtractedTable) -> tuple[int, int, int]:
    """More stable rows win. A one-column dump or a letter-by-letter split does not."""
    columns = len(table.columns)
    if columns < 2 or columns > 14 or not table.rows:
        return (0, 0, 0)
    filled = [sum(1 for cell in row if str(cell).strip()) for row in table.rows]
    median = sorted(filled)[len(filled) // 2]
    stable = sum(1 for count in filled if abs(count - median) <= 1)
    return (stable, len(table.rows), -columns)


def _choose_digital_tables(
    positioned: list[ExtractedTable], plumber: list[ExtractedTable]
) -> list[ExtractedTable]:
    """
    Keep PyMuPDF when it already has the rows.

    pdfplumber replaces it only when the borderless read keeps at least two
    more complete rows. A scan never reaches this choice.
    """
    usable = [table for table in plumber if _table_score(table)[0] >= 1]
    if not usable:
        return positioned
    if not positioned:
        return usable
    plumber_score = _table_score(max(usable, key=_table_score))
    positioned_score = _table_score(max(positioned, key=_table_score))
    if plumber_score[0] >= positioned_score[0] + 2:
        return usable
    return positioned


# Snap is separate for the two axes. A wide snap joins the plate column to
# the colour column and the whole row shifts. One point is enough to close a
# broken rule. Text columns use two points, still less than the gap between
# neighbouring fields.
_PDFPLUMBER_SETTINGS: tuple[dict[str, object], ...] = (
    {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_x_tolerance": 1,
        "snap_y_tolerance": 1,
        "join_x_tolerance": 1,
        "join_y_tolerance": 1,
        "intersection_x_tolerance": 1,
        "intersection_y_tolerance": 1,
        "edge_min_length": 8,
    },
    {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_x_tolerance": 2,
        "snap_y_tolerance": 2,
        "join_x_tolerance": 1,
        "join_y_tolerance": 2,
        "text_x_tolerance": 2,
        "text_y_tolerance": 3,
        "intersection_x_tolerance": 2,
        "intersection_y_tolerance": 3,
        "min_words_vertical": 3,
        "min_words_horizontal": 2,
    },
)


def _reading_order_cell(value, words: list[Word]) -> str:
    """
    Plain Arabic in reading order.

    pdfplumber sometimes returns the visual order (ةكرش). When that cell is
    exactly one word already read in reading order, the reading-order word
    is kept. The letters are not reshaped.
    """
    text = "" if value is None else _clean_cell(value)
    if not text or not words or not any(_is_arabic_letter(char) for char in text):
        return text
    flipped = text[::-1]
    for word in words:
        if word.text == text or word.text == flipped:
            return word.text
    return text


def _tables_from_pdfplumber(
    document,
    page_index: int,
    source: fitz.Page | None = None,
) -> list[ExtractedTable]:
    """Ruled and borderless grids. Column edges snap by one or two points."""
    if document is None:
        return []
    try:
        page = document.pages[page_index]
    except Exception:
        return []
    words = _page_words(source) if source is not None else []
    best: list[ExtractedTable] = []
    best_score = (0, 0, 0)
    for settings in _PDFPLUMBER_SETTINGS:
        batch: list[ExtractedTable] = []
        try:
            raw_tables = page.extract_tables(settings) or []
        except Exception:
            continue
        for raw in raw_tables:
            if not raw:
                continue
            matrix = [[_reading_order_cell(cell, words) for cell in row] for row in raw if row]
            parsed = _table_from_matrix(matrix, page_index)
            if parsed is not None:
                batch.append(parsed)
        if not batch:
            continue
        score = _table_score(max(batch, key=_table_score))
        if score > best_score:
            best = batch
            best_score = score
    return best


def _open_pdfplumber(pdf_path: Path, password: str):
    """Open the same file for pdfplumber. A failure leaves PyMuPDF in charge."""
    try:
        import pdfplumber
    except ImportError:
        return None
    try:
        options = {"password": password} if password else {}
        return pdfplumber.open(str(pdf_path), **options)
    except Exception:
        return None


def _tables_from_aligned_words(page: fitz.Page, page_index: int = 0) -> list[ExtractedTable]:
    """
    A text table whose columns are separated by a pipe or a wide gap.

    Each word is already one cell. Clustering those words by x would glue
    "رقم اللوحة" to "إداري" because the gap is only a few points.
    """
    matrix: list[list[str]] = []
    for row in _group_rows(_page_words(page)):
        if len(row) < 2:
            continue
        ordered = sorted(row, key=lambda word: word.x0)
        span = ordered[-1].x1 - ordered[0].x0
        if span < page.rect.width * 0.25:
            continue
        matrix.append([word.text.strip() for word in ordered])
    if len(matrix) < 2:
        return []
    width = max(len(row) for row in matrix)
    usable = [row for row in matrix if len(row) >= max(2, width - 1)]
    if len(usable) < 2:
        return []
    usable = [row + [""] * (width - len(row)) for row in usable]

    def label_score(row: list[str]) -> tuple[int, int]:
        letters = sum(1 for cell in row for char in cell if _is_arabic_letter(char))
        numbers = sum(1 for cell in row if re.search(r"\d", cell))
        return (numbers == 0, letters)

    header_at = max(range(len(usable)), key=lambda index: label_score(usable[index]))
    ordered = [usable[header_at]] + [row for index, row in enumerate(usable) if index != header_at]
    parsed = _table_from_matrix(ordered, page_index)
    return [parsed] if parsed is not None else []


def _visual_lines(page: fitz.Page) -> int:
    """Text lines stored on this page. Used to spot a grid that invented rows."""
    try:
        data = page.get_text("dict") or {}
    except Exception:
        return 0
    count = 0
    for block in data.get("blocks") or []:
        for line in block.get("lines") or []:
            spans = line.get("spans") or []
            if any((span.get("text") or "").strip() for span in spans):
                count += 1
    return count


# Printed left-to-right on the vehicle report. Column 0 is انتهاء التأمين.
# The sheet stays right to left; these names are not rearranged by the slot guesser.
_VEHICLE_COLUMNS: tuple[tuple[str, str], ...] = (
    ("date", "انتهاء التأمين"),
    ("color", "اللون السائد"),
    ("admin", "رقم اداري"),
    ("make", "الصنع"),
    ("year", "السنة"),
    ("plate", "رقم اللوحة"),
    ("purpose", "غرض الترخيص"),
)
_VEHICLE_HEADER_TEXT = {name for _key, name in _VEHICLE_COLUMNS}
_VEHICLE_LINE_SETTINGS: dict[str, object] = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_x_tolerance": 2,
    "snap_y_tolerance": 2,
    "join_x_tolerance": 2,
    "join_y_tolerance": 2,
    "intersection_x_tolerance": 3,
    "intersection_y_tolerance": 3,
    "edge_min_length": 10,
}
_VEHICLE_TEXT_SETTINGS: dict[str, object] = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "snap_x_tolerance": 2,
    "snap_y_tolerance": 2,
    "join_x_tolerance": 1,
    "join_y_tolerance": 2,
    "text_x_tolerance": 2,
    "text_y_tolerance": 2,
    "intersection_x_tolerance": 3,
    "intersection_y_tolerance": 3,
    "min_words_vertical": 2,
    "min_words_horizontal": 1,
}


def _cluster_positions(values: list[float], tolerance: float) -> list[float]:
    """Collapse line positions that are the same rule seen twice."""
    if not values:
        return []
    ordered = sorted(values)
    groups: list[list[float]] = [[ordered[0]]]
    for value in ordered[1:]:
        if value - groups[-1][-1] <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [sum(group) / len(group) for group in groups]


def _vehicle_field(text: str) -> str | None:
    """The vehicle column a header cell names, or nothing for a data cell."""
    concept = _concept_for(text)
    if concept == "insurance":
        return "date"
    if concept in {key for key, _name in _VEHICLE_COLUMNS}:
        return concept
    folded = _normalize(text).replace(" ", "")
    if not folded:
        return None
    if "انتهاء" in folded or ("تامين" in folded and "انته" in folded):
        return "date"
    if "لون" in folded:
        return "color"
    if "اداري" in folded:
        return "admin"
    if folded in {"الصنع", "صنع", "الماركه", "ماركه"}:
        return "make"
    if folded in {"السنه", "سنه"}:
        return "year"
    if "لوح" in folded:
        return "plate"
    if "ترخيص" in folded or folded == "غرض":
        return "purpose"
    return None


def _vehicle_header_map(row: list[str]) -> dict[str, int]:
    found: dict[str, int] = {}
    for index, cell in enumerate(row):
        field_name = _vehicle_field(cell)
        if field_name and field_name not in found:
            found[field_name] = index
    return found


def _is_vehicle_note(cells: list[str]) -> bool:
    """A footer sentence such as 'هذا الكشف صالح ...' is not a vehicle."""
    text = _normalize(" ".join(cell for cell in cells if cell))
    if not text:
        return True
    folded = text.replace(" ", "")
    if "هذاالكشف" in folded or "هذا الكشف" in text:
        return True
    if "صالح" in text and ("لمده" in folded or "الكشف" in text or "يوم" in text):
        return True
    if _is_printed_page_label(text) or _is_noise_row(cells):
        return True
    return False


_BIDI_DATE_MARKS = (
    "\u200e",
    "\u200f",
    "\u202a",
    "\u202b",
    "\u202c",
    "\u202d",
    "\u202e",
    "\u2066",
    "\u2067",
    "\u2068",
    "\u2069",
)


def _strip_bidi_date_noise(text: str) -> str:
    """Drop bidi marks, quotes, and the dashes that only wrap a date token."""
    value = _clean_cell(text).translate(_INDIC_DIGITS)
    for mark in _BIDI_DATE_MARKS:
        value = value.replace(mark, "")
    for quote in ("'", "’", "`", "´", "«", "»", "“", "”", '"'):
        value = value.replace(quote, " ")
    value = value.replace("–", "-").replace("—", "-").replace("/", "-").replace(".", "-")
    value = re.sub(r"(?<=\d)\s+(?=\d)", "-", value)
    value = re.sub(r"\s*-\s*", "-", value)
    value = re.sub(r"-{2,}", "-", value)
    return value.strip(" -")


def _iso_from_parts(year: str, month: str, day: str) -> str | None:
    """YYYY-MM-DD when the three parts are a real calendar date."""
    if not (year.isdigit() and month.isdigit() and day.isdigit()):
        return None
    year_n, month_n, day_n = int(year), int(month), int(day)
    if not (1900 <= year_n <= 2099 and 1 <= month_n <= 12 and 1 <= day_n <= 31):
        return None
    iso = f"{year_n:04d}-{month_n:02d}-{day_n:02d}"
    if _DATE_VALUE.fullmatch(iso):
        return iso
    return None


def _iso_date_from_text(text: str) -> str | None:
    """A year-first date, or a day-first date, from one cleaned string."""
    match = re.search(r"(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})", text)
    if match:
        found = _iso_from_parts(*match.groups())
        if found:
            return found
    match = re.search(r"(\d{1,2})\s*-\s*(\d{1,2})\s*-\s*(\d{4})", text)
    if match:
        day, month, year = match.groups()
        return _iso_from_parts(year, month, day)
    groups = re.findall(r"\d+", text)
    for start in range(0, max(len(groups) - 2, 0)):
        window = groups[start : start + 3]
        if len(window[0]) == 4:
            found = _iso_from_parts(window[0], window[1], window[2])
        elif len(window[2]) == 4:
            found = _iso_from_parts(window[2], window[1], window[0])
        else:
            found = None
        if found:
            return found
    return None


def _date_fragment(text: str) -> str:
    """
    Two short number groups and no 4-digit year.

    '- 04 '09' stays 04-09. A year is not invented from the السنة column
    or from the row above.
    """
    groups = re.findall(r"\d+", text)
    if len(groups) != 2 or any(len(group) > 2 for group in groups):
        return ""
    month, day = int(groups[0]), int(groups[1])
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return ""
    return f"{month:02d}-{day:02d}"


def _rtl_month_day(text: str) -> tuple[str, str, str] | None:
    """'- 04 '09' is month then day. A year counts only when it is in this text."""
    value = _clean_cell(text).translate(_INDIC_DIGITS)
    for mark in _BIDI_DATE_MARKS:
        value = value.replace(mark, "")
    match = re.search(
        r"-\s*['’`]?\s*(\d{1,2})\s*['’`]\s*(\d{1,2})(?:\D+((?:19|20)\d{2}))?",
        value,
    )
    if not match:
        return None
    return match.group(1), match.group(2), match.group(3) or ""


def _normalize_date_cell(text: str) -> str:
    """
    The انتهاء التأمين cell.

    A full date, written YYYY-MM-DD or DD-MM-YYYY with '-', '/', or '.',
    becomes YYYY-MM-DD. A fully reversed token such as 90-40-6202 is put
    back the same way. '- 04 '09' stays 04-09 unless that same text also
    contains the year.
    """
    mirrored = _rtl_month_day(text)
    if mirrored is not None:
        month, day, year = mirrored
        if year:
            found = _iso_from_parts(year, month, day)
            if found:
                return found
        fragment = _date_fragment(f"{month}-{day}")
        if fragment:
            return fragment
    prepared = _strip_bidi_date_noise(text)
    if not prepared:
        return ""
    found = _iso_date_from_text(prepared)
    if found:
        return found
    # The whole token was stored backwards, and the year digits are in it.
    if any(len(group) == 4 for group in re.findall(r"\d+", prepared)):
        found = _iso_date_from_text(_strip_bidi_date_noise(prepared[::-1]))
        if found:
            return found
    fragment = _date_fragment(prepared)
    if fragment:
        return fragment
    return prepared


def _normalize_year_cell(text: str) -> str:
    value = _clean_cell(text).translate(_INDIC_DIGITS)
    match = re.search(r"(?<!\d)(?:19|20)\d{2}(?!\d)", value)
    if match:
        return match.group()
    return value


def _normalize_plate_cell(text: str) -> str:
    """
    Digits stay in extracted order.

    Spaces that split 18018 into 180 18 are closed. The string is not reversed,
    so a plate cannot become 81081.
    """
    value = _clean_cell(text).translate(_INDIC_DIGITS)
    value = value.replace("\u200e", "").replace("\u200f", "")
    value = re.sub(r"(?<=\d)\s+(?=\d)", "", value)
    digits = re.sub(r"\D", "", value)
    letters = re.sub(r"[\d\s./\-]+", "", value)
    if digits and not letters and 4 <= len(digits) <= 5 and not digits.startswith("0"):
        if not _YEAR_VALUE.fullmatch(digits):
            return digits
    return re.sub(r"\s+", " ", value).strip()


def _normalize_admin_cell(text: str) -> str:
    value = _clean_cell(text).translate(_INDIC_DIGITS)
    if re.fullmatch(r"[\d\s./\-]+", value or ""):
        digits = re.sub(r"\D", "", value)
        if 1 <= len(digits) <= 6:
            return digits
    return value


def _normalize_vehicle_cells(cells: list[str]) -> list[str]:
    """One vehicle in the fixed column order, with brands and numbers normalized."""
    date, color, admin, make, year, plate, purpose = cells
    return [
        _normalize_date_cell(date),
        _correct_color(color) or _tidy_arabic(color),
        _normalize_admin_cell(admin),
        _normalize_brand(make),
        _normalize_year_cell(year),
        _normalize_plate_cell(plate),
        _purpose_token([purpose]) or _tidy_arabic(purpose),
    ]


def _infer_vehicle_map(rows: list[list[str]]) -> dict[str, int]:
    """
    Assign the seven fields from the values when the header line is unreadable.

    A column is used once. An invoice without a plate column, a year column,
    and a date column is not given this layout.
    """
    if len(rows) < 3:
        return {}
    width = max(len(row) for row in rows)
    scores: list[tuple[float, str, int]] = []
    for index in range(width):
        values = [row[index].strip() if index < len(row) else "" for row in rows]
        filled = [value for value in values if value]
        if len(filled) < 3:
            continue
        total = len(filled)

        def ratio(predicate) -> float:
            return sum(1 for value in filled if predicate(value)) / total

        scores.append((ratio(lambda value: bool(_normalize_date_cell(value) and _DATE_VALUE.fullmatch(_normalize_date_cell(value)))), "date", index))
        scores.append((ratio(lambda value: bool(_YEAR_VALUE.fullmatch(_normalize_year_cell(value)))), "year", index))
        scores.append((ratio(lambda value: bool(_PLATE_VALUE.fullmatch(_normalize_plate_cell(value)))), "plate", index))
        scores.append((ratio(lambda value: bool(re.fullmatch(r"\d{1,3}", _normalize_admin_cell(value)))), "admin", index))
        scores.append((ratio(lambda value: _normalize_brand(value) in {item[0] for item in _BRAND_ALIASES}), "make", index))
        scores.append((ratio(lambda value: bool(_match_color(value))), "color", index))
        scores.append((ratio(lambda value: bool(_purpose_token([value]))), "purpose", index))
    found: dict[str, int] = {}
    used: set[int] = set()
    for score, field_name, index in sorted(scores, reverse=True):
        if score < 0.45 or field_name in found or index in used:
            continue
        found[field_name] = index
        used.add(index)
    if not {"plate", "year", "date"} <= set(found):
        return {}
    return found


def _vehicle_value_field(text: str) -> str | None:
    """Which vehicle column a data value belongs in, judged from the value itself."""
    cleaned = _clean_cell(text)
    if not cleaned or _vehicle_field(cleaned):
        return None
    if _DATE_VALUE.fullmatch(_normalize_date_cell(cleaned) or ""):
        return "date"
    year = _normalize_year_cell(cleaned)
    if _YEAR_VALUE.fullmatch(year or ""):
        return "year"
    plate = _normalize_plate_cell(cleaned)
    if _PLATE_VALUE.fullmatch(plate or "") and not _YEAR_VALUE.fullmatch(plate or ""):
        return "plate"
    admin = _normalize_admin_cell(cleaned)
    if re.fullmatch(r"\d{1,3}", admin or ""):
        return "admin"
    if _match_official_color(cleaned) or _match_color(cleaned):
        return "color"
    brand = _normalize_brand(cleaned)
    if brand in _MANUFACTURERS or brand in {item[0] for item in _BRAND_ALIASES}:
        return "make"
    if _purpose_token([cleaned]):
        return "purpose"
    return None


def _lock_vehicle_cells(raw: list[str], mapping: dict[str, int], order: list[str]) -> list[str]:
    """
    Read one printed row into the fixed columns.

    A full row keeps each cell at the header's index, so a blank purpose or
    plate stays blank and the cells beside it do not slide. A short row, whose
    empty cells were dropped, is placed by what each value is, not by packing
    the survivors to the left.
    """
    width = max(mapping.values(), default=-1) + 1
    if width > 0 and len(raw) >= width:
        return [raw[mapping[key]] if key in mapping and mapping[key] < len(raw) else "" for key in order]
    slots = {key: "" for key in order}
    for cell in raw:
        field_name = _vehicle_value_field(cell)
        if field_name and field_name in slots and not slots[field_name]:
            slots[field_name] = cell
    return [slots[key] for key in order]


def _separator_vehicle_row(columns: list[str], row: list[str]) -> bool:
    """A header echo or a subtotal, read against the current column names."""
    order = [key for key, _name in _VEHICLE_COLUMNS]
    keys = {key: "" for key in order}
    for index, name in enumerate(columns):
        field_name = _vehicle_field(name)
        if field_name in keys and index < len(row):
            keys[field_name] = str(row[index])
    if not any(keys.values()) and not any(str(cell).strip() for cell in row):
        return False
    if any(keys.values()):
        return _is_separator_vehicle(keys)
    blob = _normalize(" ".join(str(cell) for cell in row)).replace(" ", "")
    return any(token in blob for token in ("المجموع", "الاجمالي", "عددالمركبات", "المجموعالفرعي"))


def _nondata_kind(row: list[str], columns: list[str] | None = None) -> str:
    """ترويسة or فاصل when this row was isolated from the vehicle records."""
    if columns:
        for index, name in enumerate(columns):
            if _is_notes_header(name) and index < len(row):
                text = str(row[index]).strip()
                return text if text in _NONDATA_MARKS else ""
        return ""
    text = str(row[-1]).strip() if row else ""
    if text in _NONDATA_MARKS and len(row) > len(_VEHICLE_COLUMNS):
        return text
    return ""


def _marked_nondata(kind: str) -> list[str]:
    """A header or a break, with every vehicle cell left blank."""
    return [""] * len(_VEHICLE_COLUMNS) + [kind]


def _is_separator_vehicle(keys: dict[str, str]) -> bool:
    """A repeated header, a subtotal, or a row whose plate cell says اللوحة."""
    plate = _normalize_plate_cell(keys.get("plate", ""))
    if _PLATE_VALUE.fullmatch(plate or "") and not _YEAR_VALUE.fullmatch(plate or ""):
        return False
    plate_name = _normalize(keys.get("plate", "")).replace(" ", "")
    if plate_name in {"اللوحه", "رقماللوحه", "لوحه"}:
        return True
    blob = _normalize(" ".join(keys.values())).replace(" ", "")
    if any(token in blob for token in ("المجموع", "الاجمالي", "عددالمركبات", "المجموعالفرعي")):
        return True
    named = sum(1 for value in keys.values() if value and _vehicle_field(value))
    return named >= 3


def _project_vehicle_rows(matrix: list[list[str]]) -> list[list[str]] | None:
    """
    Header row, then one list per vehicle.

    A wrapped fragment that holds only the rest of the make, colour, or purpose
    is joined onto the vehicle above it. A blank gap, a repeated header, or a
    subtotal ends that vehicle. A real ruled row, which still has a plate, a
    year, a date, or an admin number, is never dropped.
    """
    if not matrix:
        return None
    header_at = -1
    mapping: dict[str, int] = {}
    for index, row in enumerate(matrix[:6]):
        found = _vehicle_header_map(row)
        if len(found) > len(mapping):
            mapping = found
            header_at = index
    body = matrix[header_at + 1 :] if header_at >= 0 else matrix
    if len(mapping) < 4:
        inferred = _infer_vehicle_map(body)
        if len(inferred) > len(mapping):
            mapping = inferred
    if len(mapping) < 4 or "plate" not in mapping:
        return None
    rows: list[list[str]] = []
    order = [key for key, _name in _VEHICLE_COLUMNS]
    anchor: list[str] | None = None
    for raw in body:
        if raw and raw[0] == "__GAP__":
            # The gap splits this page. The section below starts with no memory of the section above.
            anchor = None
            continue
        if _vehicle_header_map(raw) and len(_vehicle_header_map(raw)) >= 4:
            anchor = None
            rows.append(_marked_nondata(_NONDATA_HEADER))
            continue
        projected = _lock_vehicle_cells(raw, mapping, order)
        if not any(cell.strip() for cell in projected):
            anchor = None
            continue
        # A header or a subtotal is named before any brand or colour rewrite.
        if _is_separator_vehicle(dict(zip(order, projected))):
            anchor = None
            raw_keys = dict(zip(order, projected))
            label = _NONDATA_HEADER if _normalize(raw_keys.get("plate", "")).replace(" ", "") in {
                "اللوحه",
                "رقماللوحه",
                "لوحه",
            } or sum(1 for value in raw_keys.values() if value and _vehicle_field(value)) >= 3 else _NONDATA_BREAK
            rows.append(_marked_nondata(label))
            continue
        normalized = _normalize_vehicle_cells(projected)
        keys = dict(zip(order, normalized))
        has_vehicle = bool(
            _DATE_VALUE.fullmatch(keys["date"])
            or _YEAR_VALUE.fullmatch(keys["year"])
            or (_PLATE_VALUE.fullmatch(keys["plate"]) and not _YEAR_VALUE.fullmatch(keys["plate"]))
        )
        if not has_vehicle and _is_vehicle_note(raw):
            continue
        # A wrapped colour, make, or purpose belongs to the vehicle above it.
        # A date, a year, a plate, or an admin number is never copied onto that vehicle.
        if anchor is not None and not has_vehicle and not any(_looks_like_date(value) for value in keys.values()):
            for offset, name in enumerate(order):
                if name in {"date", "year", "plate", "admin"}:
                    continue
                extra = keys[name]
                if extra and extra not in anchor[offset]:
                    anchor[offset] = f"{anchor[offset]} {extra}".strip()
            anchor[:] = _normalize_vehicle_cells(anchor)
            continue
        if anchor is not None and not has_vehicle:
            anchor = None
        if not has_vehicle and not any(keys[name] for name in ("make", "admin", "color", "purpose")):
            continue
        rows.append(normalized)
        anchor = normalized
    if len(rows) < 1:
        return None
    return rows


def _plumber_matrix(plumber_page, settings: dict) -> list[list[str]]:
    try:
        raw_tables = plumber_page.extract_tables(settings) or []
    except Exception:
        return []
    best: list[list[str]] = []
    for raw in raw_tables:
        rows = [[_clean_cell(cell) for cell in (row or [])] for row in raw if row]
        if len(rows) > len(best):
            best = rows
    return best


def _explicit_vehicle_settings(plumber_page) -> dict | None:
    """Column and row rules taken from the drawn grid, not from a text guess."""
    vertical: list[float] = []
    horizontal: list[float] = []
    try:
        edges = list(plumber_page.edges or [])
    except Exception:
        return None
    for edge in edges:
        try:
            x0, x1 = float(edge["x0"]), float(edge["x1"])
            top, bottom = float(edge["top"]), float(edge["bottom"])
        except (KeyError, TypeError, ValueError):
            continue
        if abs(x1 - x0) <= 1.8 and abs(bottom - top) >= 12:
            vertical.append((x0 + x1) / 2.0)
        elif abs(bottom - top) <= 1.8 and abs(x1 - x0) >= 20:
            horizontal.append((top + bottom) / 2.0)
    xs = _cluster_positions(vertical, 2.5)
    ys = _cluster_positions(horizontal, 2.0)
    if len(xs) < 6 or len(ys) < 3:
        return None
    return {
        "vertical_strategy": "explicit",
        "horizontal_strategy": "explicit",
        "explicit_vertical_lines": xs,
        "explicit_horizontal_lines": ys,
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "intersection_tolerance": 4,
    }


def _matrix_from_fitz(page: fitz.Page) -> list[list[str]]:
    best: list[list[str]] = []
    for table in _line_tables(page):
        try:
            raw = table.extract() or []
        except Exception:
            continue
        rows = [[_clean_cell(cell) for cell in (row or [])] for row in raw if row]
        if len(rows) > len(best):
            best = rows
    return best


def _vehicle_data_count(table: ExtractedTable) -> int:
    """Vehicle rows on this sheet. A marked header or gap is not one of them."""
    return sum(1 for row in table.rows if not _nondata_kind(row, table.columns))


def _vehicle_table_from_matrix(matrix: list[list[str]], page_index: int) -> ExtractedTable | None:
    rows = _project_vehicle_rows(matrix)
    if not rows:
        return None
    columns = [name for _key, name in _VEHICLE_COLUMNS]
    if any(len(row) > len(columns) for row in rows):
        columns = columns + [_NOTES_HEADER]
        rows = [list(row) + [""] * (len(columns) - len(row)) for row in rows]
    if not any(not _nondata_kind(row, columns) for row in rows):
        return None
    return ExtractedTable(columns, rows, None, [page_index] * len(rows))


def _page_rules(page: fitz.Page) -> tuple[list[float], list[float]]:
    """Vertical and horizontal rules drawn on the page, in points."""
    vertical: list[float] = []
    horizontal: list[float] = []
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is not None:
            width = float(rect.x1) - float(rect.x0)
            height = float(rect.y1) - float(rect.y0)
            if width <= 1.8 and height >= 12:
                vertical.append((float(rect.x0) + float(rect.x1)) / 2.0)
            elif height <= 1.8 and width >= 18:
                horizontal.append((float(rect.y0) + float(rect.y1)) / 2.0)
        for item in drawing.get("items") or []:
            kind = item[0] if item else ""
            if kind == "l" and len(item) >= 3:
                start, end = item[1], item[2]
                x0 = float(getattr(start, "x", start[0]))
                y0 = float(getattr(start, "y", start[1]))
                x1 = float(getattr(end, "x", end[0]))
                y1 = float(getattr(end, "y", end[1]))
                if abs(x1 - x0) <= 1.8 and abs(y1 - y0) >= 12:
                    vertical.append((x0 + x1) / 2.0)
                elif abs(y1 - y0) <= 1.8 and abs(x1 - x0) >= 18:
                    horizontal.append((y0 + y1) / 2.0)
            elif kind == "re" and len(item) >= 2:
                box = item[1]
                vertical.append(float(box.x0))
                vertical.append(float(box.x1))
                horizontal.append(float(box.y0))
                horizontal.append(float(box.y1))
    return _cluster_positions(vertical, 2.0), _cluster_positions(horizontal, 2.0)


def _vehicle_header_spans(words: list[Word]) -> list[tuple[str, Word]] | None:
    """The seven titles, left to right, from the line that names the most fields."""
    labeled = [(field, word) for word in words if (field := _vehicle_field(word.text))]
    if len(labeled) < 4:
        return None
    lines = _group_rows([word for _field, word in labeled], 4.0)
    best = max(lines, key=len)
    if len(best) < 4:
        return None
    center = sum(word.center_y for word in best) / len(best)
    spans: dict[str, Word] = {}
    for field, word in labeled:
        if abs(word.center_y - center) > 5:
            continue
        current = spans.get(field)
        if current is None or (word.x1 - word.x0) > (current.x1 - current.x0):
            spans[field] = word
    if "plate" not in spans or len(spans) < 4:
        return None
    return sorted(spans.items(), key=lambda item: item[1].x0)


def _vehicle_column_edges(
    spans: list[tuple[str, Word]],
    vertical: list[float],
    page_width: float,
) -> list[float]:
    """One boundary between each header, taken from a rule when one is drawn there."""
    centers = [word.center_x for _field, word in spans]
    edges: list[float] = []
    left_lines = [line for line in vertical if line <= centers[0] + 4]
    edges.append(max(left_lines) if left_lines else max(0.0, spans[0][1].x0 - 6))
    for index in range(len(centers) - 1):
        low, high = centers[index], centers[index + 1]
        between = [line for line in vertical if low + 4 < line < high - 4]
        if between:
            middle = (low + high) / 2.0
            edges.append(min(between, key=lambda line: abs(line - middle)))
        else:
            edges.append((low + high) / 2.0)
    right_lines = [line for line in vertical if line >= centers[-1] - 4]
    edges.append(min(right_lines) if right_lines else min(page_width, spans[-1][1].x1 + 6))
    return edges


def _vehicle_cell_text(
    glyphs: list[tuple[float, float, float, float, str, int]],
    x0: float,
    y0: float,
    x1: float,
    y1: float,
) -> str:
    """
    Text inside one ruled cell.

    Digits stay in left-to-right order, so a plate is not turned around.
    Arabic is restored only when the stored order is a known reversed word.
    """
    chosen = []
    for glyph in glyphs:
        center_x = (glyph[0] + glyph[2]) / 2.0
        center_y = (glyph[1] + glyph[3]) / 2.0
        if x0 - 1.0 <= center_x <= x1 + 1.0 and y0 - 1.0 <= center_y <= y1 + 1.0:
            chosen.append(glyph)
    if not chosen:
        return ""
    chosen.sort(key=lambda glyph: glyph[0])
    parts: list[str] = []
    run: list[tuple[float, float, float, float, str, int]] = []

    def flush() -> None:
        if not run:
            return
        parts.append(_token_text(run))
        run.clear()

    for glyph in chosen:
        if _is_arabic_letter(glyph[4]):
            if run and glyph[0] - run[-1][2] > 8:
                flush()
                parts.append(" ")
            run.append(glyph)
            continue
        flush()
        parts.append(glyph[4])
    flush()
    text = _clean_cell("".join(parts))
    text = re.sub(r"(?<=\d)\s+(?=\d)", "", text)
    return _logical_arabic(text)


def _line_is_standalone_break(text: str) -> bool:
    """A header echo or a subtotal line, with no vehicle plate on it."""
    folded = _normalize(text).replace(" ", "")
    if not folded:
        return False
    has_plate = bool(_plate_token([text]))
    if has_plate:
        return False
    if any(token in folded for token in ("المجموع", "الاجمالي", "عددالمركبات", "المجموعالفرعي")):
        return True
    hits = sum(1 for word in text.split() if _vehicle_field(word))
    if hits >= 3:
        return True
    return folded in {"اللوحه", "رقماللوحه", "لوحه"}


def _looks_like_date(text: str) -> bool:
    """A finished date, or a fragment that still has digits and a separator."""
    raw = _clean_cell(text)
    if not raw:
        return False
    if _DATE_VALUE.fullmatch(_normalize_date_cell(raw) or ""):
        return True
    return bool(re.search(r"\d", raw) and re.search(r"[-/'’]", raw))


def _vehicle_bands(page: fitz.Page) -> list[tuple[str, float, float]]:
    """
    Upper and lower blocks on one page.

    A tall blank, a horizontal or dashed rule sitting in that blank, or a
    subtotal/header line ends the block above it. The next vehicles are a
    new block.
    """
    words = [word for word in _page_words(page) if word.center_y < float(page.rect.height) * 0.94]
    lines = _group_rows(words, 3.0)
    if not lines:
        return [("rows", float(page.rect.y0), float(page.rect.y1))]
    spans: list[tuple[float, float, str]] = []
    for line in lines:
        top = min(word.y0 for word in line)
        bottom = max(word.y1 for word in line)
        text = " ".join(word.text for word in sorted(line, key=lambda item: item.x0))
        spans.append((top, bottom, text))
    heights = [bottom - top for top, bottom, _text in spans]
    typical = sorted(heights)[len(heights) // 2] if heights else 12.0
    gap_limit = max(16.0, typical * 2.2)
    _vertical, horizontal = _page_rules(page)
    segments: list[tuple[str, float, float]] = []
    band_top = spans[0][0] - 1.0
    previous_bottom: float | None = None
    for top, bottom, text in spans:
        if _line_is_standalone_break(text):
            if previous_bottom is not None and band_top < previous_bottom:
                segments.append(("rows", band_top, previous_bottom + 1.0))
            kind = "header" if sum(1 for word in text.split() if _vehicle_field(word)) >= 3 or _normalize(text).replace(" ", "") in {
                "اللوحه",
                "رقماللوحه",
                "لوحه",
            } else "break"
            segments.append((kind, top, bottom))
            band_top = bottom + 1.0
            previous_bottom = None
            continue
        split = False
        if previous_bottom is not None and top - previous_bottom > gap_limit:
            split = True
        if previous_bottom is not None and top - previous_bottom > max(8.0, typical * 0.9):
            if any(previous_bottom + 1.5 < rule < top - 1.5 for rule in horizontal):
                split = True
        if split and previous_bottom is not None:
            segments.append(("rows", band_top, previous_bottom + 1.0))
            segments.append(("break", previous_bottom, top))
            band_top = top - 1.0
        previous_bottom = bottom
    if previous_bottom is not None and band_top < previous_bottom:
        segments.append(("rows", band_top, previous_bottom + 2.0))
    return segments or [("rows", float(page.rect.y0), float(page.rect.y1))]


def _vehicle_matrix_from_grid(
    page: fitz.Page,
    row_tolerance: float = 3.0,
    y0: float | None = None,
    y1: float | None = None,
) -> list[list[str]]:
    """
    One matrix row per printed baseline, cut on the header rules.

    A missing horizontal line does not merge two vehicles, and a wrapped
    line stays in the same columns instead of being dropped. y0 and y1
    limit the read to one band, so a gap is not crossed.
    """
    glyphs = _page_glyphs(page)
    words = _words_from_glyphs(glyphs)
    spans = _vehicle_header_spans(words)
    if not spans:
        return []
    vertical, _horizontal = _page_rules(page)
    edges = _vehicle_column_edges(spans, vertical, float(page.rect.width))
    if len(edges) != len(spans) + 1:
        return []
    header_bottom = max(word.y1 for _field, word in spans)
    body = [
        word
        for word in words
        if word.y0 >= header_bottom - 1 and word.center_y < float(page.rect.height) * 0.94
    ]
    if y0 is not None:
        body = [word for word in body if word.center_y >= y0]
    if y1 is not None:
        body = [word for word in body if word.center_y <= y1]
    # Detected titles, left to right. The projector maps each title to its field
    # and writes the sheet in the fixed order.
    names = {key: name for key, name in _VEHICLE_COLUMNS}
    matrix = [[names.get(field, field) for field, _word in spans]]
    lines = _group_rows(body, row_tolerance)
    heights = [max(word.y1 for word in line) - min(word.y0 for word in line) for line in lines]
    typical = sorted(heights)[len(heights) // 2] if heights else 12.0
    gap_limit = max(16.0, typical * 2.2)
    previous_bottom: float | None = None
    for line in lines:
        top = min(word.y0 for word in line)
        bottom = max(word.y1 for word in line)
        if previous_bottom is not None and top - previous_bottom > gap_limit:
            matrix.append(["__GAP__"] + [""] * (len(spans) - 1))
        cells = [
            _vehicle_cell_text(glyphs, edges[index], top, edges[index + 1], bottom)
            for index in range(len(spans))
        ]
        if any(cell.strip() for cell in cells):
            matrix.append(cells)
            previous_bottom = bottom
    return matrix if len(matrix) > 1 else []


def _break_table(page_index: int, kind: str = "break") -> ExtractedTable:
    """One blank header or separator row. It carries no date, name, or colour."""
    columns = [name for _key, name in _VEHICLE_COLUMNS] + [_NOTES_HEADER]
    label = _NONDATA_HEADER if kind == "header" else _NONDATA_BREAK
    return ExtractedTable(columns, [_marked_nondata(label)], None, [page_index])


def _concat_vehicle_parts(parts: list[ExtractedTable], page_index: int) -> ExtractedTable | None:
    """Stack band tables in reading order. A band does not lend cells to the next one."""
    parts = [part for part in parts if part.rows]
    if not parts:
        return None
    columns = [name for _key, name in _VEHICLE_COLUMNS] + [_NOTES_HEADER]
    rows: list[list[str]] = []
    for part in parts:
        fields = _vehicle_field_at(part.columns)
        for row in part.rows:
            built = []
            for key, _name in _VEHICLE_COLUMNS:
                index = fields.get(key)
                built.append(row[index] if index is not None and index < len(row) else "")
            built.append(_nondata_kind(row, part.columns) or _nondata_kind(row))
            rows.append(built)
    if not any(not _nondata_kind(row, columns) for row in rows):
        return None
    return ExtractedTable(columns, rows, None, [page_index] * len(rows))


def _page_has_mid_band(segments: list[tuple[str, float, float]]) -> bool:
    """True when a gap or a subtotal sits after the first block of vehicles."""
    seen_rows = False
    for kind, _top, _bottom in segments:
        if kind == "rows":
            if seen_rows:
                return True
            seen_rows = True
        elif kind in {"break", "header"} and seen_rows:
            return True
    return False


def _vehicle_tables_banded(page: fitz.Page, page_index: int, plumber_page) -> list[ExtractedTable]:
    """Read each band on its own, then stack them with a separator between bands."""
    parts: list[ExtractedTable] = []
    for kind, top, bottom in _vehicle_bands(page):
        if kind in {"break", "header"}:
            parts.append(_break_table(page_index, kind))
            continue
        parsed = _vehicle_table_from_matrix(_vehicle_matrix_from_grid(page, y0=top, y1=bottom), page_index)
        if parsed is None:
            continue
        _pass_word_coordinates(page, parsed, plumber_page, y_floor=top, y_limit=bottom)
        parts.append(parsed)
    stacked = _concat_vehicle_parts(parts, page_index)
    return [stacked] if stacked is not None else []


def _vehicle_tables(page: fitz.Page, page_index: int, plumber) -> list[ExtractedTable]:
    """
    A vehicle register is read from its ruled grid.

    Explicit line positions keep a wrapped cell inside its own row and stop
    neighbouring columns from trading a plate for a year. The seven headers
    stay in the printed order. A mid-page gap is read as a separate band.
    """
    plumber_page = None
    if plumber is not None:
        try:
            plumber_page = plumber.pages[page_index]
        except Exception:
            plumber_page = None
    segments = _vehicle_bands(page)
    if _page_has_mid_band(segments):
        return _vehicle_tables_banded(page, page_index, plumber_page)
    candidates: list[tuple[int, int, ExtractedTable]] = []
    matrices: list[tuple[int, list[list[str]]]] = [(-1, _vehicle_matrix_from_grid(page))]
    if plumber_page is not None:
        explicit = _explicit_vehicle_settings(plumber_page)
        if explicit is not None:
            matrices.append((0, _plumber_matrix(plumber_page, explicit)))
        matrices.append((1, _plumber_matrix(plumber_page, _VEHICLE_LINE_SETTINGS)))
        matrices.append((2, _plumber_matrix(plumber_page, _VEHICLE_TEXT_SETTINGS)))
    matrices.append((3, _matrix_from_fitz(page)))
    for priority, matrix in matrices:
        parsed = _vehicle_table_from_matrix(matrix, page_index)
        if parsed is not None:
            candidates.append((_vehicle_data_count(parsed), -priority, parsed))
    if not candidates:
        return []
    # Pass 1 keeps the explicit grid. Pass 2 then reads words that fell outside a cell.
    _count, _rank, chosen = max(candidates, key=lambda item: (item[0], item[1]))
    _pass_word_coordinates(page, chosen, plumber_page)
    return [chosen]


def _recognised_vehicle_fields(columns: list[str]) -> set[str]:
    """Header names that map to a vehicle field. A box title maps to nothing."""
    found: set[str] = set()
    for name in columns:
        field_name = _vehicle_field(str(name))
        if field_name:
            found.add(field_name)
    return found


def _vehicle_columns_missing(columns: list[str]) -> bool:
    """The plate is unnamed, or both the make and the colour are unnamed."""
    found = _recognised_vehicle_fields(columns)
    if "plate" not in found:
        return True
    return "make" not in found and "color" not in found


def build_column_warning(columns: list[str], *, garbage: bool, kind: str, forced: bool = False) -> str:
    """
    A visible reason when the text layer is junk or the columns have no names.

    An empty review list is not described as clean while this warning is set.
    """
    missing = _vehicle_columns_missing(columns)
    relevant = garbage or forced or kind == "register"
    if not relevant:
        return ""
    if not garbage and not missing:
        return ""
    parts: list[str] = []
    if garbage:
        parts.append("طبقة النص مربعات أو رموز غير مقروءة، وقُرئت الصفحة المصورة بالتعرف الضوئي.")
    if missing:
        parts.append("تعذر التعرف على أعمدة المركبات.")
        parts.append("قائمة فارغة لا تعني أن الملف سليم.")
    return " ".join(parts)


def _extract_needs_reread(table: ExtractedTable) -> bool:
    """A cached read of boxes, before the rendered page was tried, is stale."""
    if getattr(table, "ocr_attempted", False):
        return False
    if _headers_look_garbage(list(table.columns)):
        return True
    return table.kind == "register" and "plate" not in _recognised_vehicle_fields(list(table.columns))


def _mark_ocr_attempted(tables: list[ExtractedTable]) -> list[ExtractedTable]:
    for table in tables:
        table.ocr_attempted = True
    return tables


def _tables_on_page(
    page: fitz.Page,
    page_index: int = 0,
    language_box: dict[str, str] | None = None,
    plumber=None,
    notes: list[str] | None = None,
    kind: str = "register",
) -> list[ExtractedTable]:
    """
    Read one page with the strategy chosen for this document.

    A register tries the vehicle grid first. An invoice or a financial
    statement stays on the ruled grid and the word boxes. A free-text page
    uses words, then OCR. An empty or failed grid never stops the batch.
    """
    label = str(page_index + 1)
    box = language_box if language_box is not None else {}

    def note(message: str) -> None:
        if notes is not None:
            notes.append(message)

    def attempt(reader, *args) -> list[ExtractedTable]:
        try:
            found = reader(*args) or []
        except Exception as exc:
            note(f"Page {label}: {getattr(reader, '__name__', 'reader')} failed ({exc}).")
            return []
        return [table for table in found if table.rows and any(any(str(cell).strip() for cell in row) for row in table.rows)]

    try:
        garbage = text_layer_is_garbage(page)
        if garbage:
            box["text_garbage"] = "1"
            note(f"Page {label}: the text layer is boxes or unmapped glyphs. Reading the rendered page.")
            scanned = _mark_ocr_attempted(attempt(_tables_from_scan, page, page_index, box))
            if scanned:
                return scanned
        digital = _text_char_count(page) >= 20
        if not digital:
            return attempt(_tables_from_scan, page, page_index, box)
        if kind == "register":
            vehicle = attempt(_vehicle_tables, page, page_index, plumber)
            if vehicle and _vehicle_columns_missing(vehicle[0].columns):
                box["columns_unrecognised"] = "1"
                note(f"Page {label}: the columns were not identified. Reading the rendered page.")
                scanned = _mark_ocr_attempted(attempt(_tables_from_scan, page, page_index, box))
                if scanned and len(_recognised_vehicle_fields(scanned[0].columns)) > len(_recognised_vehicle_fields(vehicle[0].columns)):
                    return scanned
                _mark_ocr_attempted(vehicle)
            if vehicle:
                return vehicle
        lines = _visual_lines(page)
        ruled = attempt(_tables_from_rules, page, page_index) if kind != "freetext" else []
        ruled_rows = sum(len(table.rows) for table in ruled)
        if lines >= 4 and ruled_rows > max(int(lines * 1.6), lines + 6):
            note(
                f"Page {label}: the ruled grid produced {ruled_rows} rows from {lines} lines. "
                "The page was read from its words instead."
            )
            ruled = []
        positioned = list(ruled)
        if not positioned and kind != "financial":
            positioned = attempt(_tables_from_words, page, page_index)
        if not positioned:
            positioned = attempt(_tables_from_aligned_words, page, page_index)
        plumber_tables = attempt(_tables_from_pdfplumber, plumber, page_index, page)
        chosen = _choose_digital_tables(positioned, plumber_tables) if positioned or plumber_tables else []
        chosen = [table for table in chosen if table.rows]
        if chosen:
            return chosen
        note(f"Page {label}: the grid was empty. Falling back to word positions.")
        fallback = attempt(_tables_from_words, page, page_index) or attempt(_tables_from_aligned_words, page, page_index)
        if fallback:
            return fallback
        note(f"Page {label}: no text grid. Falling back to OCR.")
        return attempt(_tables_from_scan, page, page_index, box)
    except Exception as exc:
        note(f"Page {label}: {exc}. Trying the other reader.")
        fallback = attempt(_tables_from_aligned_words, page, page_index)
        if fallback:
            return fallback
        return attempt(_tables_from_scan, page, page_index, box)


def _header_key(table: ExtractedTable) -> tuple[str, ...]:
    """Arabic and English names of the same field stack into one sheet."""
    parts: list[str] = []
    for name in table.columns:
        parts.append(_concept_for(name) or _normalize(name))
    return tuple(parts)


_SLOT_ORDER = ("purpose", "plate", "admin", "make", "year", "color", "date")
_SLOT_HEADER = {
    "purpose": "غرض الترخيص",
    "plate": "رقم اللوحة",
    "admin": "رقم إداري",
    "make": "الصنع",
    "year": "السنة",
    "color": "اللون",
    "date": "انتهاء التأمين",
}


def _slot_from_values(name: str, values: list[str]) -> str:
    """
    What this column holds on this page alone.

    The decision uses the values on this page. A wrong header does not move
    the column onto another page's grid.
    """
    filled = [value for value in values if value]
    if len(filled) >= 3:
        ratios = _pattern_ratios(filled)
        best = max(("date", "year", "plate"), key=lambda kind: ratios[kind])
        if ratios[best] >= 0.45:
            return best
        short = sum(1 for value in filled if re.fullmatch(r"\d{1,3}", value)) / len(filled)
        if short >= 0.5:
            return "admin"
        arabic = sum(1 for value in filled if any(_is_arabic_letter(char) for char in value)) / len(filled)
        if arabic >= 0.45:
            return "words"
    concept = _concept_for(name)
    if concept == "insurance":
        concept = "date"
    if concept in _SLOT_ORDER:
        return concept
    return ""


def _slots_for_page(table: ExtractedTable) -> dict[str, int]:
    """One slot per field, chosen from this page's own columns."""
    described: list[tuple[int, str, float]] = []
    for index, name in enumerate(table.columns):
        values = [row[index].strip() if index < len(row) and row[index] else "" for row in table.rows]
        kind = _slot_from_values(name, values)
        filled = sum(1 for value in values if value)
        described.append((index, kind, filled / max(len(values), 1)))
    assigned: dict[str, int] = {}
    used: set[int] = set()

    def take(slot: str, index: int) -> None:
        if slot not in assigned and index not in used:
            assigned[slot] = index
            used.add(index)

    for slot in ("date", "year", "plate", "admin"):
        choices = [item for item in described if item[1] == slot]
        if choices:
            take(slot, max(choices, key=lambda item: item[2])[0])
    plate_at = assigned.get("plate")
    year_at = assigned.get("year")
    date_at = assigned.get("date")
    if "purpose" not in assigned and plate_at is not None:
        for index, name in enumerate(table.columns):
            if index >= plate_at or index in used:
                continue
            values = [row[index].strip() if index < len(row) and row[index] else "" for row in table.rows]
            arabic_count = sum(1 for value in values if any(_is_arabic_letter(char) for char in value))
            if arabic_count >= 2:
                take("purpose", index)
                break
    words = [item[0] for item in described if item[1] == "words" and item[0] not in used]
    # In the stored reading order the purpose sits before the plate, the make
    # sits before the year, and the colour sits between the year and the date.
    for index in words:
        if plate_at is not None and index < plate_at:
            take("purpose", index)
    between = [index for index in words if index not in used and (date_at is None or index < date_at)]
    if len(between) >= 2:
        take("color", between[-1])
        take("make", between[-2])
    elif len(between) == 1:
        if year_at is not None and between[0] > year_at:
            take("color", between[0])
        else:
            take("make", between[0])
    return assigned


# Canonical colour spellings. Comparison ignores alef hamza and alef maqsura,
# so "ابيض" and "أبيض" are the same colour. A one-letter OCR slip such as
# "صدفح" is listed here and also accepted by the one-letter fuzzy match.
_COLOR_CANON = (
    "صدفي",
    "أبيض",
    "فضي",
    "أسود",
    "رمادي",
    "كحلي",
    "أحمر",
    "أزرق",
    "بني",
    "أخضر",
    "ذهبي",
    "بيج",
    "برتقالي",
)
# OCR slips that are not a single-letter edit after hamza folding, or that
# we want to pin to one colour even when another colour is equally close.
_COLOR_EXTRA = {
    "صدفح": "صدفي",
    "صدقي": "صدفي",
    "صدفو": "صدفي",
    "صدفه": "صدفي",
    "صدفى": "صدفي",
    "ابيظ": "أبيض",
    "اييض": "أبيض",
    "ابيضه": "أبيض",
    "اسوذ": "أسود",
    "فضى": "فضي",
    "رمادى": "رمادي",
    "كحلى": "كحلي",
    "كحله": "كحلي",
    "احمر": "أحمر",
    "ازرق": "أزرق",
    "اخضر": "أخضر",
    "ذهبى": "ذهبي",
    "ذهبي": "ذهبي",
    # These slips contain بني, but the registry reads them as أبيض.
    "ابنيفر": "أبيض",
    "ابنير": "أبيض",
    "ابير": "أبيض",
    "ايبدي": "أبيض",
    "ايبضي": "أبيض",
}
# Arabic-Indic and Persian digits. A plate read as ١٨٠١٨ is still 18018.
_INDIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


# Letters this scan confuses inside a colour. ح is read for ي, and ف for ق, س, or ض.
_COLOR_CONFUSION = (
    frozenset("يحخى"),
    frozenset("فقسضش"),
)


def _fold_color(text: str) -> str:
    """Comparable form of a colour: no spaces, no hamza, yeh instead of alef maqsura."""
    value = _tidy_arabic(text)
    for source, target in (("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ؤ", "و"), ("ئ", "ي"), ("ة", "ه")):
        value = value.replace(source, target)
    return value.replace(" ", "")


def _collapse_repeats(text: str) -> str:
    """
    صددفي is صدفي, and صدقفقفي is صدفي.

    An exact double letter is removed. A run of three or more lookalike
    letters (ف ق ف ق) is also one letter. A real pair such as فض in فضي stays.
    """
    chars: list[str] = []
    for char in text:
        if not chars or chars[-1] != char:
            chars.append(char)
    output: list[str] = []
    index = 0
    while index < len(chars):
        end = index + 1
        while end < len(chars) and any(
            chars[end] in group and chars[index] in group for group in _COLOR_CONFUSION
        ):
            end += 1
        if end - index >= 3:
            output.append(chars[index])
        else:
            output.extend(chars[index:end])
        index = end
    return "".join(output)


def _color_distance(left: str, right: str, limit: int = 2, groups: tuple[frozenset[str], ...] | None = None) -> int:
    """Edit distance where a known OCR confusion, such as ح for ي, costs nothing."""
    if abs(len(left) - len(right)) > limit:
        return limit + 1
    confused = groups if groups is not None else _COLOR_CONFUSION
    previous = list(range(len(right) + 1))
    for index, char in enumerate(left, 1):
        current = [index]
        lowest = current[0]
        for other_index, other in enumerate(right, 1):
            same = char == other or any(char in group and other in group for group in confused)
            cost = 0 if same else 1
            current.append(
                min(previous[other_index] + 1, current[other_index - 1] + 1, previous[other_index - 1] + cost)
            )
            lowest = min(lowest, current[-1])
        if lowest > limit:
            return limit + 1
        previous = current
    return previous[-1]


def _color_lookup() -> dict[str, str]:
    """Folded spelling to the canonical colour, including the explicit OCR list."""
    found: dict[str, str] = {}
    for color in _COLOR_CANON:
        found[_fold_color(color)] = color
    for source, target in _COLOR_EXTRA.items():
        found[_fold_color(source)] = target
    return found


_COLOR_FOLDED = _color_lookup()


def _match_color_key(text: str) -> str:
    """
    Canonical colour for one token, or empty when it is not a colour.

    The token is folded, a leading "ال" is ignored, and the dictionary is
    tried first. A unique spelling within one letter is accepted. A make such
    as تويوتا is too far from every colour, so it is left unchanged. Two
    colours at the same distance are not guessed.
    """
    key = _collapse_repeats(_fold_color(text))
    if not key or len(key) < 3 or len(key) > 12 or re.search(r"\d", key):
        return ""
    bare = key[2:] if key.startswith("ال") and len(key) >= 5 else key
    if len(bare) < 3:
        return ""
    for candidate in (key, bare):
        pinned = _COLOR_FOLDED.get(candidate)
        if pinned:
            return pinned
    # Nearest colour within two edits. A confused letter such as ح for ي is free.
    # The nearest colour has to be the only one at that distance, so تويوتا is
    # not rewritten and two equally close colours are not guessed.
    best = 99
    matches: list[str] = []
    for color in _COLOR_CANON:
        distance = _color_distance(bare, _collapse_repeats(_fold_color(color)), 2)
        if distance < best:
            best = distance
            matches = [color]
        elif distance == best and color not in matches:
            matches.append(color)
    # Distance 2 is only for a word long enough to be a real colour.
    # A one-letter scrap must not become فضي.
    if len(matches) == 1 and (best <= 1 or (best <= 2 and len(bare) >= 4)):
        return matches[0]
    return ""


def _match_color(text: str) -> str:
    """The colour in a cell. Digits beside it are ignored. Two different colours are not guessed."""
    tidied = _tidy_arabic(text)
    if not tidied:
        return ""
    letters = re.sub(r"[^\u0600-\u06FF]", "", tidied)
    direct = _match_color_key(letters or tidied)
    if direct:
        return direct
    found: list[str] = []
    for word in re.findall(r"[\u0600-\u06FF]+", tidied):
        hit = _match_color_key(word)
        if hit and hit not in found:
            found.append(hit)
    if len(found) == 1:
        return found[0]
    return ""


def _correct_color(text: str) -> str:
    """Write the canonical colour when the cell is a known colour. Otherwise keep the tidied text."""
    match = _match_color(text)
    if match:
        return match
    return _tidy_arabic(text) or _blank_text(text)


# Official colours on a traffic register. The اللون column is rewritten to one of these.
_OFFICIAL_COLORS: tuple[str, ...] = (
    "أبيض",
    "أسود",
    "رمادي",
    "فضي",
    "أحمر",
    "أخضر",
    "صدفي",
    "كحلي",
    "بني",
    "رمادي غامق",
)
_COLOR_SCORE = 75


def _official_color_catalog() -> dict[str, str]:
    """Folded spelling to the official colour, including the pinned OCR slips."""
    found = {_fold_color(name): name for name in _OFFICIAL_COLORS}
    for source, target in _COLOR_EXTRA.items():
        if target in _OFFICIAL_COLORS:
            found[_fold_color(source)] = target
    return found


def _color_scorer(query: str, choice: str) -> int:
    """Full-string score. A short colour sitting inside a longer slip does not win."""
    from thefuzz import fuzz

    if abs(len(query) - len(choice)) > 3:
        return 0
    return int(fuzz.ratio(query, choice))


def _match_official_color(text: str) -> str:
    """
    Replace a corrupted اللون value when it is close to an official colour.

    ابنيفر and ابنير are pinned to أبيض. A plate or a make is left unchanged.
    """
    cleaned = _tidy_arabic(text)
    if not cleaned or re.search(r"\d", cleaned):
        return cleaned or text
    catalog = _official_color_catalog()
    folded = _fold_color(cleaned)
    if "رمادي" in folded and "غ" in folded:
        return "رمادي غامق"
    if folded in catalog:
        return catalog[folded]
    known = _match_color(cleaned)
    if known in _OFFICIAL_COLORS:
        return known
    if len(folded) < 3:
        return cleaned
    choices = list(catalog)
    match = None
    try:
        from thefuzz import process

        match = process.extractOne(folded, choices, scorer=_color_scorer)
    except Exception:
        best = 0
        for choice in choices:
            limit = max(len(folded), len(choice), 1)
            distance = _edit_distance(folded, choice, limit=limit)
            score = int(round(100 * (1 - distance / limit)))
            if abs(len(folded) - len(choice)) <= 2 and score > best:
                best = score
                match = (choice, score)
    if not match:
        return cleaned
    official = catalog.get(match[0])
    if official and int(match[1]) >= _COLOR_SCORE:
        return official
    return cleaned


def _color_column(table: ExtractedTable) -> int | None:
    """The اللون column. Other columns are not candidates."""
    for index, name in enumerate(table.columns):
        if _concept_for(name) == "color" or _logical_key(name) in {"اللون", "اللون السايد"}:
            return index
    return None


def _exact_color(text: str) -> str:
    """A colour that is in the dictionary. A fuzzy resemblance is not accepted."""
    cleaned = _collapse_spaced_arabic(text)
    if not cleaned or re.search(r"\d", cleaned):
        return ""
    folded = _collapse_repeats(_fold_color(cleaned))
    if "رمادي" in folded and "غ" in folded:
        return "رمادي غامق"
    bare = folded[2:] if folded.startswith("ال") and len(folded) >= 5 else folded
    for candidate in (folded, bare):
        pinned = _COLOR_FOLDED.get(candidate)
        if pinned:
            return pinned
    return ""


# ص/س and ج/ح are the extra swaps in سدفسح and صدفسج. They stay out of the
# global colour fold so an unrelated word is not collapsed.
_COLOR_RANK_GROUPS = _COLOR_CONFUSION + (frozenset("صس"), frozenset("جح"))


def _rank_colors(text: str) -> list[tuple[str, int]]:
    """Official colours scored from edit distance. Best first."""
    cleaned = re.sub(r"[^\u0600-\u06FF\s]", "", _collapse_spaced_arabic(text))
    folded = _collapse_repeats(_fold_color(cleaned))
    if not folded or re.search(r"\d", cleaned):
        return []
    exact = _exact_color(cleaned)
    ranked: list[tuple[str, int]] = []
    for name in _OFFICIAL_COLORS:
        target = _collapse_repeats(_fold_color(name))
        distance = _color_distance(folded, target, limit=4, groups=_COLOR_RANK_GROUPS)
        score = max(0, 100 - 25 * distance)
        if exact and name == exact:
            score = 100
        ranked.append((name, score))
    ranked.sort(key=lambda item: (-item[1], item[0]))
    return ranked


def _accept_color(text: str) -> str:
    """The official colour when the nearest name is a full step closer than the next."""
    ranked = _rank_colors(text)
    if not ranked:
        return ""
    best_name, best = ranked[0]
    second = ranked[1][1] if len(ranked) > 1 else 0
    if best >= 75 and best - second >= 25:
        return best_name
    return ""


def _apply_color_dictionary(table: ExtractedTable) -> None:
    """Clean only the colour column. A resemblance is a suggestion, not a value."""
    index = _color_column(table)
    if index is None:
        return
    notes_at = next((i for i, name in enumerate(table.columns) if _is_notes_header(name)), None)
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    sources = list(table.sources or [])
    for row_index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        while len(row) <= index:
            row.append("")
        original = str(row[index]).strip()
        if not original:
            continue
        page = int(sources[row_index]) if row_index < len(sources) else 0
        plate = _plate_token([row[plate_at]]) if plate_at is not None and plate_at < len(row) else ""
        kept = _keep_arabic_value("color", original)
        if not kept:
            continue
        exact = kept if kept in _OFFICIAL_COLORS else ""
        if exact:
            if exact != original and _raw_is_slip("color", original, exact):
                _append_change(row, notes_at, original, exact)
                row[index] = exact
                _remember_raw(table, page, plate, "color", original)
                _remember_source(table, page, plate, "color", "normalized", f"{exact} — تطبيع: {original}")
            continue
        if kept != original:
            row[index] = kept
        guess = _match_official_color(original)
        if guess in _OFFICIAL_COLORS and guess != kept:
            _add_reasoned_suggestion(table, page, plate, "color", guess, "قراءة المقطع")
            if table.provenance.get(_source_key(page, plate, "color")) not in {"agreed", "manual", "normalized"}:
                _remember_source(table, page, plate, "color", "suggested", table.suggestions.get(_source_key(page, plate, "color"), ""))


def _column_title_keys() -> set[str]:
    """Folded header labels. A crop that only saw الصنع is not a make."""
    labels = [name for _key, name in _VEHICLE_COLUMNS]
    labels.extend(("الصنع", "الصنج", "اللون", "اللون السائد", "السنة", "رقم اللوحة", "غرض الترخيص", "الصفحة", "صفحة"))
    return {_logical_key(label) for label in labels if label}


def _keep_arabic_value(field: str, text: str) -> str:
    """
    Collapse spaced Arabic and keep the cleaned word.

    An exact dictionary make or colour is rewritten to the official spelling.
    Anything else that is still Arabic stays in the cell. The cell is not
    cleared just because the dictionary did not name it.
    """
    collapsed = _collapse_spaced_arabic(text)
    words = [part for part in collapsed.split() if part and not re.fullmatch(r"[\d./\-]+", part)]
    collapsed = " ".join(words).strip()
    if field == "make" and _explicit_gmc(text):
        return "جي إم سي"
    if not collapsed or not any(_is_arabic_letter(char) for char in collapsed):
        return ""
    if _logical_key(collapsed) in _column_title_keys():
        return ""
    letters = sum(1 for char in collapsed if _is_arabic_letter(char))
    if field == "make":
        if _gmc_reading_ok(text) or _explicit_gmc(text):
            return "جي إم سي"
        if _short_make_fragment(text):
            return ""
        exact = _exact_manufacturer(collapsed) or _exact_make_in_text(collapsed)
        if exact:
            return exact
        if letters < 2 or _purpose_token([collapsed]) or _exact_color(collapsed):
            return ""
        return collapsed
    if field == "color":
        exact = _exact_color(collapsed) or _exact_color_in_text(collapsed)
        if exact:
            return exact
        if letters < 2 or _purpose_token([collapsed]) or _exact_manufacturer(collapsed):
            return ""
        return collapsed
    return collapsed if letters >= 2 else ""


def _arabic_cell_settled(field: str, text: str) -> bool:
    """An official make or colour is already good enough to leave alone."""
    value = str(text or "").strip()
    if field == "make":
        return value in _MANUFACTURERS
    if field == "color":
        return value in _OFFICIAL_COLORS
    return False


_REPEAT_FIELDS = ("make", "color", "year")


def _repeat_columns(table: ExtractedTable) -> list[tuple[int, str]]:
    """Manufacturer, colour, and year. Plate, date, and admin are not included."""
    found: list[tuple[int, str]] = []
    seen: set[str] = set()
    for index, name in enumerate(table.columns):
        field = _vehicle_field(name) or _concept_for(name)
        if field not in _REPEAT_FIELDS or field in seen:
            continue
        seen.add(field)
        found.append((index, field))
    return found


def _repeatable_value(field: str, text: str) -> str:
    """An official make, an official colour, or a model year. Anything else is refused."""
    if field == "year":
        year = _normalize_year_cell(text)
        return year if _YEAR_VALUE.fullmatch(year) else ""
    if field == "color":
        color = _match_official_color(text)
        return color if color in _OFFICIAL_COLORS else ""
    if field == "make":
        cleaned = _tidy_arabic(text)
        if not cleaned or re.search(r"\d", cleaned):
            return ""
        known = _alias_brand(cleaned)
        if known:
            return known
        found = _extract_manufacturer(_logical_key(cleaned))
        if found and found[1] >= _MANUFACTURER_SCORE:
            return found[0]
    return ""


def _plumber_words(plumber_page) -> list[Word]:
    """pdfplumber word boxes. An empty or failed read leaves the glyph words in charge."""
    if plumber_page is None:
        return []
    try:
        raw = plumber_page.extract_words(keep_blank_chars=False, use_text_flow=False) or []
    except Exception:
        return []
    words: list[Word] = []
    for item in raw:
        try:
            text = _logical_arabic(_clean_cell(item.get("text") or ""))
            if not text:
                continue
            words.append(Word(float(item["x0"]), float(item["top"]), float(item["x1"]), float(item["bottom"]), text))
        except (KeyError, TypeError, ValueError):
            continue
    return words


def _merge_coordinate_words(primary: list[Word], extra: list[Word]) -> list[Word]:
    """Glyph words first. A plumber word is added only when that spot is still empty."""
    if not extra:
        return list(primary)
    merged = list(primary)
    for word in extra:
        if any(abs(word.center_x - old.center_x) < 2.5 and abs(word.center_y - old.center_y) < 2.5 for old in primary):
            continue
        merged.append(word)
    return merged


def _attribute_column_boxes(page: fitz.Page, words: list[Word]) -> dict[str, tuple[float, float]]:
    """The x span of الصنع, اللون, and السنة, taken from the ruled header."""
    spans = _vehicle_header_spans(words)
    if not spans:
        return {}
    vertical, _horizontal = _page_rules(page)
    edges = _vehicle_column_edges(spans, vertical, float(page.rect.width))
    if len(edges) != len(spans) + 1:
        return {}
    boxes: dict[str, tuple[float, float]] = {}
    for index, (field, _word) in enumerate(spans):
        if field in _REPEAT_FIELDS:
            boxes[field] = (edges[index], edges[index + 1])
    return boxes


def _plate_anchor_hits(words: list[Word], plate: str) -> list[Word]:
    """Words that are this row's plate. The plate's y is the row's band."""
    digits = re.sub(r"\D", "", (plate or "").translate(_INDIC_DIGITS))
    if len(digits) < 4:
        return []
    return [word for word in words if re.sub(r"\D", "", word.text.translate(_INDIC_DIGITS)) == digits]


def _fill_blank_attributes(
    table: ExtractedTable,
    words: list[Word],
    boxes: dict[str, tuple[float, float]],
    y_floor: float | None = None,
    y_limit: float | None = None,
) -> None:
    """
    Pass 2. Fill an empty manufacturer, colour, or year from words in that row's band.

    The band runs from this vehicle's plate down to the next vehicle's plate,
    and it stops at the edge of this page-band. A word in another column, on
    the next vehicle, or below a gap is not copied. A date is never filled here.
    """
    columns = _repeat_columns(table)
    if not words or not boxes or not columns:
        return
    plate_at = _column_index(table, "plate", "رقماللوحه")
    if plate_at is None:
        return
    located: list[tuple[float, float] | None] = []
    for row in table.rows:
        if _nondata_kind(row, table.columns):
            located.append(None)
            continue
        plate = row[plate_at].strip() if plate_at < len(row) else ""
        hits = _plate_anchor_hits(words, plate)
        if not hits:
            located.append(None)
            continue
        mid = hits[len(hits) // 2].center_y
        near = [word for word in hits if abs(word.center_y - mid) <= 14]
        located.append((min(word.y0 for word in near) - 2.0, max(word.y1 for word in near)))
    centres = [((band[0] + band[1]) / 2.0) if band is not None else None for band in located]
    known = [center for center in centres if center is not None]
    shift = _text_line_offset(words, known)
    for row_index, focus in enumerate(centres):
        if focus is None:
            continue
        row = table.rows[row_index]
        for index, field in columns:
            while len(row) <= index:
                row.append("")
            if row[index].strip():
                continue
            box = boxes.get(field)
            if not box:
                continue
            chosen = [
                word
                for word in words
                if box[0] <= word.center_x <= box[1]
                and _glyph_owns_row(word.center_y, focus, known, shift)
                and (y_floor is None or word.center_y >= y_floor)
                and (y_limit is None or word.center_y <= y_limit)
            ]
            if not chosen:
                continue
            chosen.sort(key=lambda word: (word.center_y, word.x0))
            text = _logical_arabic(_clean_cell(" ".join(word.text for word in chosen)))
            value = _repeatable_value(field, text)
            if value:
                row[index] = value


def _date_tokens_in(text: str) -> tuple[list[str], list[str]]:
    """Years and short day/month groups in one date string, in reading order."""
    prepared = _strip_bidi_date_noise(text)
    years: list[str] = []
    shorts: list[str] = []
    for group in re.findall(r"\d+", prepared):
        if _YEAR_VALUE.fullmatch(group):
            years.append(group)
        elif len(group) <= 2:
            shorts.append(group)
    return years, shorts


def _assemble_date_text(parts: list[str]) -> str:
    """
    Join the date tokens that belong to one row.

    '2026 - 04 - 09' becomes 2026-04-09. A leftover '04-09' becomes that
    date only when a 19xx or 20xx token is among the same parts. A year
    that is not in those parts is not invented.
    """
    pieces = [part for part in parts if _clean_cell(part)]
    if not pieces:
        return ""
    for piece in pieces:
        one = _normalize_date_cell(piece)
        if _DATE_VALUE.fullmatch(one or ""):
            return one
    blob = " ".join(pieces)
    years, _shorts = _date_tokens_in(blob)
    years = list(dict.fromkeys(years))
    groups = re.findall(r"\d+", _strip_bidi_date_noise(blob))
    if len(years) == 1 and len(groups) >= 3 and groups[0] == years[0]:
        iso = _iso_from_parts(years[0], groups[1], groups[2])
        if iso:
            return iso
    fragment = ""
    for piece in pieces:
        fragment = _date_fragment(_normalize_date_cell(piece)) or _date_fragment(piece)
        if fragment:
            break
    if len(years) == 1 and fragment:
        month, day = fragment.split("-")
        iso = _iso_from_parts(years[0], month, day)
        if iso:
            return iso
    if fragment and len(years) != 1:
        return fragment
    return _normalize_date_cell(blob)


def _assemble_dates_from_words(
    table: ExtractedTable,
    words: list[Word],
    date_box: tuple[float, float],
    y_floor: float | None = None,
    y_limit: float | None = None,
) -> None:
    """
    Complete انتهاء التأمين from the date tokens in that row's Y band.

    The search stays inside the date column. اللون is copied back at the end,
    so a colour that was already read is not cleared. A full YYYY-MM-DD already
    in the cell is left as it is.
    """
    fields = _vehicle_field_at(table.columns)
    date_at = fields.get("date")
    plate_at = fields.get("plate")
    color_at = fields.get("color")
    if date_at is None or not words:
        return
    saved_colors = [
        row[color_at] if color_at is not None and color_at < len(row) else ""
        for row in table.rows
    ]
    x0, x1 = date_box
    date_words = [
        word
        for word in words
        if x0 - 1.0 <= word.center_x <= x1 + 1.0
        and (y_floor is None or word.center_y >= y_floor)
        and (y_limit is None or word.center_y <= y_limit)
        and (_date_tokens_in(word.text)[0] or _date_tokens_in(word.text)[1] or re.search(r"[-/'’.]", word.text))
    ]
    centres: list[float | None] = []
    for row in table.rows:
        if _nondata_kind(row, table.columns) or plate_at is None or plate_at >= len(row):
            centres.append(None)
            continue
        hits = _plate_anchor_hits(words, row[plate_at])
        if y_floor is not None:
            hits = [word for word in hits if word.center_y >= y_floor]
        if y_limit is not None:
            hits = [word for word in hits if word.center_y <= y_limit]
        centres.append(hits[len(hits) // 2].center_y if hits else None)
    known = [center for center in centres if center is not None]
    owned_text: dict[int, str] = {}
    for index, center in enumerate(centres):
        if center is None:
            owned_text[index] = ""
            continue
        # Each date token goes to the nearest plate. A token in the gap is not used.
        owned = [
            word
            for word in date_words
            if _glyph_owns_row(word.center_y, center, known)
            and (y_floor is None or word.center_y >= y_floor)
            and (y_limit is None or word.center_y <= y_limit)
        ]
        owned.sort(key=lambda item: (item.center_y, item.x0))
        owned_text[index] = _assemble_date_text([word.text for word in owned])
    for index, center in enumerate(centres):
        row = table.rows[index]
        while len(row) <= date_at:
            row.append("")
        current = _normalize_date_cell(row[date_at])
        assembled = owned_text.get(index) or ""
        if _DATE_VALUE.fullmatch(current or ""):
            if not known or not date_words or assembled == current:
                row[date_at] = current
                continue
            if _DATE_VALUE.fullmatch(assembled or ""):
                row[date_at] = assembled
                continue
            # The stored date is printed on another vehicle's line.
            if any(other == current for other_index, other in owned_text.items() if other_index != index and other):
                row[date_at] = ""
                continue
            row[date_at] = current
            continue
        parts = [current] if current else []
        if assembled:
            parts.append(assembled)
        rebuilt = _assemble_date_text(parts)
        if rebuilt:
            row[date_at] = rebuilt

    # A date is never copied onto the neighbouring vehicle, even when the
    # ink sits between two rows. An empty cell stays empty and is flagged.

    if color_at is None:
        return
    for row, color in zip(table.rows, saved_colors):
        while len(row) <= color_at:
            row.append("")
        row[color_at] = color


def _assemble_row_dates(
    page: fitz.Page,
    table: ExtractedTable,
    y_floor: float | None = None,
    y_limit: float | None = None,
) -> None:
    """Read the date column on this page and complete split dates in the table."""
    try:
        boxes = _preview_column_boxes(page, table.columns)
    except Exception:
        return
    fields = _vehicle_field_at(table.columns)
    date_at = fields.get("date")
    if date_at is None or date_at >= len(boxes) or boxes[date_at] is None:
        return
    _assemble_dates_from_words(
        table,
        _page_words(page),
        boxes[date_at],
        y_floor=y_floor,
        y_limit=y_limit,
    )


def _pass_word_coordinates(
    page: fitz.Page,
    table: ExtractedTable,
    plumber_page=None,
    y_floor: float | None = None,
    y_limit: float | None = None,
) -> None:
    """Pass 2 on one band. A failure here keeps the grid cells already read."""
    try:
        glyph_words = _page_words(page)
        boxes = _attribute_column_boxes(page, glyph_words)
        extra = _plumber_words(plumber_page)
        if not boxes:
            boxes = _attribute_column_boxes(page, extra)
        _fill_blank_attributes(
            table,
            _merge_coordinate_words(glyph_words, extra),
            boxes,
            y_floor=y_floor,
            y_limit=y_limit,
        )
        _assemble_row_dates(page, table, y_floor=y_floor, y_limit=y_limit)
    except Exception:
        return


def _rescan_blank_attributes(document: fitz.Document, table: ExtractedTable) -> None:
    """Run the word-band fill again for rows that are still missing a repeating attribute."""
    columns = _repeat_columns(table)
    if not columns or not table.sources or len(table.sources) != len(table.rows):
        return
    pending = [
        index
        for index, row in enumerate(table.rows)
        if any(field_at >= len(row) or not row[field_at].strip() for field_at, _field in columns)
    ]
    if not pending:
        return
    by_page: dict[int, list[int]] = {}
    for index in pending:
        by_page.setdefault(int(table.sources[index]), []).append(index)
    for page_num, indexes in by_page.items():
        if page_num <= 0:
            continue
        try:
            page = document.load_page(page_num - 1)
        except Exception:
            continue
        subset = ExtractedTable(list(table.columns), [list(table.rows[index]) for index in indexes])
        _pass_word_coordinates(page, subset)
        for local, index in enumerate(indexes):
            table.rows[index] = subset.rows[local]


def _inheritable_attribute(field: str, text: str) -> str:
    """A make, colour, or year worth copying into the plated row below."""
    official = _repeatable_value(field, text)
    if official:
        return official
    if field == "year":
        return ""
    cleaned = _tidy_arabic(text)
    if cleaned and not re.search(r"\d", cleaned) and any(_is_arabic_letter(char) for char in cleaned):
        return cleaned
    return ""


def _plated_row(row: list[str], plate_at: int | None) -> bool:
    """True when رقم اللوحة holds a 4 or 5 digit plate."""
    if plate_at is None or plate_at >= len(row):
        return False
    token = _normalize_plate_cell(row[plate_at])
    return bool(_PLATE_VALUE.fullmatch(token) and not _YEAR_VALUE.fullmatch(token))


def _forward_fill_attributes(table: ExtractedTable) -> None:
    """
    Fill a blank year on a row that already has a plate.

    The value comes from the preceding filled cell on the same page.
    A plate, a date, a manufacturer, and a colour are never copied from the
    row above. Strict reading also leaves the year blank. A header or a gap
    ends whatever year is still being carried inside this page.
    """
    if strict_values():
        return
    columns = [(index, field) for index, field in _repeat_columns(table) if field == "year"]
    plate_at = _column_index(table, "plate", "رقماللوحه")
    if len(table.rows) < 2 or not columns or plate_at is None:
        return
    sources = table.sources if table.sources and len(table.sources) == len(table.rows) else None
    # Rows recovered from a later pass are appended, so group by page first.
    # A value never crosses from one page onto the next.
    groups: dict[int, list[int]] = {}
    page_order: list[int] = []
    for row_index in range(len(table.rows)):
        page = sources[row_index] if sources is not None else 0
        if page not in groups:
            groups[page] = []
            page_order.append(page)
        groups[page].append(row_index)
    for index, field in columns:
        for page in page_order:
            carried = ""
            for row_index in groups[page]:
                row = table.rows[row_index]
                if _nondata_kind(row, table.columns):
                    carried = ""
                    continue
                while len(row) <= index:
                    row.append("")
                value = _inheritable_attribute(field, row[index])
                if value:
                    carried = value
                    row[index] = value
                    continue
                if carried and _plated_row(row, plate_at) and not row[index].strip():
                    row[index] = carried


def _blank_repeat_cells(table: ExtractedTable) -> int:
    """Blank manufacturer, year, or colour cells on rows that already have a plate."""
    columns = _repeat_columns(table)
    plate_at = _column_index(table, "plate", "رقماللوحه")
    if not columns or plate_at is None:
        return 0
    missing = 0
    for row in table.rows:
        if plate_at >= len(row) or not row[plate_at].strip():
            continue
        for index, _field in columns:
            if index >= len(row) or not row[index].strip():
                missing += 1
    return missing


def _color_in_text(parts: list[str]) -> str:
    """The only colour named anywhere in the row. Mixed colours are left alone."""
    found: list[str] = []
    for word in re.findall(r"[\u0600-\u06FF]+", " ".join(parts)):
        hit = _match_color_key(word)
        if hit and hit not in found:
            found.append(hit)
    if len(found) == 1:
        return found[0]
    return ""


def _digit_blob(parts: list[str]) -> str:
    """
    One stream of the row's numbers.

    Arabic-Indic digits become Western digits. A token that is only digits
    and lookalike letters (O, l, I) is converted. A word such as تويوتا is not.
    Dates and years are masked so their digits cannot become a plate.
    """
    pieces: list[str] = []
    for part in parts:
        if not part:
            continue
        for token in str(part).translate(_INDIC_DIGITS).split():
            if re.fullmatch(r"[0-9OoIl|SBZ.\-/_]+", token):
                token = token.translate(_LOOKALIKE)
            pieces.append(token)
    blob = " ".join(pieces)
    blob = re.sub(r"(?:19|20)\d{2}\s*[-./]\s*\d{1,2}\s*[-./]\s*\d{1,2}", " ", blob)
    blob = re.sub(r"(?:19|20)\d{2}(?:\s+\d{1,2}){2}", " ", blob)
    blob = re.sub(r"(?<!\d)(?:19|20)\d{2}(?!\d)", " ", blob)
    return blob


def _plate_token(row: list[str], extra: str = "") -> str:
    """
    A 4 or 5 digit plate anywhere in the row's text, not in a fixed column.

    A whole number is preferred. "895 54" and "89.554" are joined when the
    reader split the plate. A year such as 2025 is not a plate, and the digits
    inside a date such as 2028-02-22 are not a plate. A number that starts
    with 0 is a date fragment, not a plate.
    """
    blob = _digit_blob(list(row) + ([extra] if extra else []))
    found = [
        token
        for token in re.findall(r"(?<!\d)\d{4,5}(?!\d)", blob)
        if not token.startswith("0") and not _YEAR_VALUE.fullmatch(token)
    ]
    if not found:
        # Split plate: short digit groups separated only by space, dot, or hyphen.
        numbers = list(re.finditer(r"\d+", blob))
        for index, match in enumerate(numbers):
            if len(match.group()) >= 4:
                continue
            groups = [match.group()]
            end = match.end()
            for follow in numbers[index + 1 :]:
                gap = blob[end : follow.start()]
                if not re.fullmatch(r"[\s./\-]+", gap) or len(follow.group()) >= 4:
                    break
                groups.append(follow.group())
                end = follow.end()
                joined = "".join(groups)
                if len(joined) > 5:
                    break
                # "22" + "23" is two admin numbers, not a plate. A 4-digit join
                # has to include a 3-digit piece, such as "895" + "5".
                if (
                    len(joined) >= 4
                    and not joined.startswith("0")
                    and not _YEAR_VALUE.fullmatch(joined)
                    and (len(joined) == 5 or any(len(group) == 3 for group in groups))
                ):
                    found.append(joined)
                    break
    if not found:
        return ""
    fives = [token for token in found if len(token) == 5]
    return fives[0] if fives else found[0]


def _recover_plate(cells: list[str], plate_index: int, sources: list[str]) -> None:
    """Fill an empty or invalid plate cell from any 4 or 5 digit token in the row."""
    current = cells[plate_index].strip()
    if (
        _PLATE_VALUE.fullmatch(current)
        and not _YEAR_VALUE.fullmatch(current)
        and not current.startswith("0")
    ):
        return
    found = _plate_token(sources)
    if not found:
        return
    cells[plate_index] = found
    for other, value in enumerate(cells):
        if other == plate_index or not value:
            continue
        if _DATE_VALUE.fullmatch(value) or _YEAR_VALUE.fullmatch(value):
            continue
        cells[other] = re.sub(rf"(?<!\d){re.escape(found)}(?!\d)", " ", value).strip()


# Licence purpose, plus the OCR slips this registry actually produces.
# "تقل خاص" is "نقل خاص" with the noon read as a teh.
_PURPOSE_CANON = (
    ("نقل خاص", "نقل خاص"),
    ("نفل خاص", "نقل خاص"),
    ("تقل خاص", "نقل خاص"),
    ("باص خاص", "باص خاص"),
    ("خصوصي", "خصوصي"),
    ("عمومي", "عمومي"),
    ("تأجير", "تأجير"),
    ("تجاري", "تجاري"),
)


def _purpose_token(parts: list[str]) -> str:
    """The licence purpose printed on this row, or empty when the row has none."""
    text = _tidy_arabic(" ".join(part for part in parts if part))
    if not text:
        return ""
    compact = text.replace(" ", "")
    for raw, canon in sorted(_PURPOSE_CANON, key=lambda item: len(item[0]), reverse=True):
        if raw in text or raw.replace(" ", "") in compact:
            return canon
    return ""


def _recover_purpose(cells: list[str], purpose_index: int, sources: list[str]) -> None:
    """Fill an empty purpose cell from a phrase that missed its column."""
    current = _purpose_token([cells[purpose_index]])
    if current:
        cells[purpose_index] = current
        return
    found = _purpose_token(sources)
    if found:
        cells[purpose_index] = found


def _recover_color(cells: list[str], color_index: int, sources: list[str]) -> None:
    """Normalize the colour cell, or copy the only colour token found elsewhere on the row."""
    current = cells[color_index].strip()
    if _YEAR_VALUE.fullmatch(current) or _DATE_VALUE.fullmatch(current) or _PLATE_VALUE.fullmatch(current):
        return
    match = _match_color(cells[color_index])
    taken_elsewhere = False
    if not match:
        match = _color_in_text(sources)
        taken_elsewhere = bool(match)
    if not match:
        return
    cells[color_index] = match
    if not taken_elsewhere:
        return
    for other, value in enumerate(cells):
        if other == color_index or not value:
            continue
        kept = [word for word in value.split() if _match_color(word) != match]
        if len(kept) != len(value.split()):
            cells[other] = " ".join(kept)


def _pattern_columns(table: ExtractedTable) -> tuple[int | None, int | None]:
    """Plate and colour columns, from the header or from the values on this table."""
    plate_index: int | None = None
    color_index: int | None = None
    for index, name in enumerate(table.columns):
        concept = _concept_for(name)
        if concept == "plate" and plate_index is None:
            plate_index = index
        elif concept == "color" and color_index is None:
            color_index = index
    # No plate header: only a registry page (its own year column and date column)
    # may be recognized from the values. An invoice item code stays where it is.
    if plate_index is None and table.rows and _has_year_and_date(table):
        best_index = -1
        best_ratio = 0.0
        for index in range(len(table.columns)):
            values = [row[index].strip() for row in table.rows if index < len(row) and row[index].strip()]
            if len(values) < 3:
                continue
            hits = sum(1 for value in values if _plate_token([value]))
            ratio = hits / len(values)
            if ratio > best_ratio:
                best_ratio = ratio
                best_index = index
        if best_ratio >= 0.45:
            plate_index = best_index
    return plate_index, color_index


def _has_year_and_date(table: ExtractedTable) -> bool:
    """True when this table already has a year column and a date column."""
    seen: set[str] = set()
    for index in range(len(table.columns)):
        values = [row[index].strip() for row in table.rows if index < len(row) and row[index].strip()]
        if len(values) < 3:
            continue
        ratios = _pattern_ratios(values)
        for kind in ("year", "date"):
            if ratios[kind] >= 0.45:
                seen.add(kind)
    return seen == {"year", "date"}


def _nondata_label(columns: list[str], row: list[str]) -> str:
    """ترويسة for a repeated header, فاصل for a subtotal or other non-record line."""
    fields = _vehicle_field_at(columns)
    plate = ""
    if "plate" in fields and fields["plate"] < len(row):
        plate = str(row[fields["plate"]])
    folded = _normalize(plate).replace(" ", "")
    if folded in {"اللوحه", "رقماللوحه", "لوحه"}:
        return _NONDATA_HEADER
    named = sum(1 for cell in row if _vehicle_field(str(cell)))
    if named >= 3:
        return _NONDATA_HEADER
    return _NONDATA_BREAK


def _isolate_nondata_rows(table: ExtractedTable) -> None:
    """
    Blank a header, a repeated title, or a subtotal and name it in الملاحظات.

    The vehicle cells stay empty. No plate is invented and no filler word is written.
    """
    if "plate" not in _vehicle_field_at(table.columns) and not any(
        _nondata_kind(row) for row in table.rows
    ):
        return
    if "plate" not in _vehicle_field_at(table.columns) and _NOTES_HEADER not in table.columns:
        return
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    if notes_at is None:
        table.columns = list(table.columns) + [_NOTES_HEADER]
        notes_at = len(table.columns) - 1
    fields = list(_vehicle_field_at(table.columns).values())
    named = _vehicle_field_at(table.columns)
    sources = table.sources if table.sources and len(table.sources) == len(table.rows) else [0] * len(table.rows)
    plate_seen: set[int] = set()
    for index, row in enumerate(table.rows):
        while len(row) <= notes_at:
            row.append("")
        page = int(sources[index])
        marked = _nondata_kind(row, table.columns) or _nondata_kind(row)
        if not marked and _separator_vehicle_row(table.columns, row):
            marked = _nondata_label(table.columns, row)
        if not marked and page not in plate_seen and not _plate_token(row):
            make_at = named.get("make")
            make = row[make_at] if make_at is not None and make_at < len(row) else ""
            recognized = bool(_alias_brand(make)) or _match_manufacturer(make) in _MANUFACTURERS
            if not recognized:
                marked = _NONDATA_HEADER
        if _plate_token(row):
            plate_seen.add(page)
        if not marked:
            for index, cell in enumerate(row):
                if index == notes_at:
                    continue
                if str(cell).strip() in _FABRICATED_MARKERS or _normalize(str(cell)).replace(" ", "") in _FABRICATED_MARKERS:
                    row[index] = ""
            continue
        for index in fields:
            if index < len(row) and index != notes_at:
                row[index] = ""
        row[notes_at] = marked


def _normalize_present_cells(table: ExtractedTable) -> None:
    """Rewrite a colour or a purpose that is already in its own cell. Blanks stay blank."""
    _plate_index, color_index = _pattern_columns(table)
    purpose_index = next(
        (index for index, name in enumerate(table.columns) if _concept_for(name) == "purpose"),
        None,
    )
    for row in table.rows:
        if _nondata_kind(row, table.columns):
            continue
        if color_index is not None and color_index < len(row):
            match = _match_color(row[color_index])
            if match:
                row[color_index] = match
        if purpose_index is not None and purpose_index < len(row):
            current = _purpose_token([row[purpose_index]])
            if current:
                row[purpose_index] = current


def _validate_patterns(table: ExtractedTable) -> None:
    """
    Check one page's table before it is joined to any other page.

    A header or a gap is blanked and labeled first. Strict reading then only
    normalizes a colour or a purpose that is already in its cell. Otherwise an
    empty plate may be filled from another number on the same row.
    """
    _isolate_nondata_rows(table)
    if strict_values():
        _normalize_present_cells(table)
        return
    plate_index, color_index = _pattern_columns(table)
    purpose_index = next(
        (index for index, name in enumerate(table.columns) if _concept_for(name) == "purpose"),
        None,
    )
    if plate_index is None and color_index is None and purpose_index is None:
        return
    width = len(table.columns)
    for index, row in enumerate(table.rows):
        while len(row) < width:
            row.append("")
        if _nondata_kind(row, table.columns):
            continue
        line = table.lines[index] if table.lines and index < len(table.lines) else ""
        sources = list(row) + ([line] if line else [])
        if plate_index is not None:
            _recover_plate(row, plate_index, sources)
        if color_index is not None:
            _recover_color(row, color_index, sources)
        if purpose_index is not None:
            _recover_purpose(row, purpose_index, sources)


def _stack_scan_pages(tables: list[ExtractedTable]) -> ExtractedTable | None:
    """
    Put every scanned page on one sheet, in page order.

    Each page keeps the columns it read. Those columns are matched by what
    they contain (plate, year, date, and so on), not by a position voted
    across the file. A page with fewer columns still contributes its rows.
    """
    if not tables:
        return None
    page_slots = [_slots_for_page(table) for table in tables]
    present = [slot for slot in _SLOT_ORDER if any(slot in mapping for mapping in page_slots)]
    # Keep a plate column whenever any row still contains a 4 or 5 digit plate,
    # even if that page's layout never put the plate in its own column.
    def row_plate(table: ExtractedTable, row: list[str], index: int) -> str:
        line = table.lines[index] if table.lines and index < len(table.lines) else ""
        return _plate_token(row, line)

    if "plate" not in present and any(
        row_plate(table, row, index) for table in tables for index, row in enumerate(table.rows)
    ):
        present = [slot for slot in _SLOT_ORDER if slot in present or slot == "plate"]

    def row_purpose(table: ExtractedTable, row: list[str], index: int) -> str:
        line = table.lines[index] if table.lines and index < len(table.lines) else ""
        return _purpose_token(list(row) + ([line] if line else []))

    if "purpose" not in present and any(
        row_purpose(table, row, index) for table in tables for index, row in enumerate(table.rows)
    ):
        present = [slot for slot in _SLOT_ORDER if slot in present or slot == "purpose"]
    if len(present) < 2:
        return None
    plate_index = present.index("plate") if "plate" in present else None
    color_index = present.index("color") if "color" in present else None
    purpose_index = present.index("purpose") if "purpose" in present else None
    rows: list[list[str]] = []
    sources: list[int] = []
    lines: list[str] = []
    for table, mapping in zip(tables, page_slots):
        page_numbers = table.sources if table.sources and len(table.sources) == len(table.rows) else [1] * len(table.rows)
        for index, (row, page_num) in enumerate(zip(table.rows, page_numbers)):
            line = table.lines[index] if table.lines and index < len(table.lines) else " ".join(row)
            built = []
            for slot in present:
                column = mapping.get(slot)
                built.append(row[column] if column is not None and column < len(row) else "")
            # First pass on this row: plate and colour come from the text, not the column x position.
            pool = list(row) + [line]
            # The plate printed on this line is not a guess. Colour and purpose
            # stay in their own cells when strict reading is on.
            if plate_index is not None:
                _recover_plate(built, plate_index, pool)
            if not strict_values():
                if color_index is not None:
                    _recover_color(built, color_index, pool)
                if purpose_index is not None:
                    _recover_purpose(built, purpose_index, pool)
            if _plate_token(built, line) or sum(bool(cell) for cell in built) >= 2:
                rows.append(built)
                sources.append(int(page_num))
                lines.append(line)
    if not rows:
        return None
    names = [_SLOT_HEADER[slot] for slot in present]
    stacked = ExtractedTable(_unique_headers(names), rows, None, sources)
    stacked.lines = lines
    return stacked


def _same_layout(left: ExtractedTable, right: ExtractedTable) -> bool:
    """Pages of one report share column positions even when OCR reads the title differently."""
    if not left.anchors or not right.anchors or len(left.anchors) != len(right.anchors):
        return False
    return all(abs(a - b) <= 22 for a, b in zip(left.anchors, right.anchors))


def _header_richness(columns: list[str]) -> int:
    return sum(1 for name in columns for char in name if char.isalpha() or _is_arabic_letter(char))


def _drop_repeated_rows(table: ExtractedTable) -> None:
    """A line printed on three or more pages is a footer or a page number, not a record."""
    if not table.sources or len(table.sources) != len(table.rows):
        return
    pages_for_row: dict[tuple[str, ...], set[int]] = {}
    for page_index, row in zip(table.sources, table.rows):
        pages_for_row.setdefault(tuple(row), set()).add(page_index)
    chrome = {signature for signature, pages in pages_for_row.items() if len(pages) >= 3}
    if not chrome:
        return
    kept_rows: list[list[str]] = []
    kept_pages: list[int] = []
    kept_lines: list[str] = []
    lines = table.lines if table.lines and len(table.lines) == len(table.rows) else None
    for row_index, (page_index, row) in enumerate(zip(table.sources, table.rows)):
        # A plate, an amount, or a date makes a repeated line a record. A footer has none of that variety.
        if (
            tuple(row) in chrome
            and len(re.findall(r"\d{4,}", " ".join(row))) < 2
            and not _plate_token(row)
        ):
            continue
        kept_rows.append(row)
        kept_pages.append(page_index)
        if lines is not None:
            kept_lines.append(lines[row_index])
    table.rows = kept_rows
    table.sources = kept_pages
    if lines is not None:
        table.lines = kept_lines


def _edit_distance(left: str, right: str, limit: int = 1) -> int:
    if abs(len(left) - len(right)) > limit:
        return limit + 1
    previous = list(range(len(right) + 1))
    for index, char in enumerate(left, 1):
        current = [index]
        lowest = current[0]
        for other_index, other in enumerate(right, 1):
            cost = 0 if char == other else 1
            current.append(
                min(previous[other_index] + 1, current[other_index - 1] + 1, previous[other_index - 1] + cost)
            )
            lowest = min(lowest, current[-1])
        if lowest > limit:
            return limit + 1
        previous = current
    return previous[-1]


def _same_spelling(left: str, right: str) -> bool:
    """One letter off, or a word that was cut short by one or two letters."""
    if left == right:
        return True
    if _edit_distance(left, right, 1) <= 1:
        return True
    short, long = (left, right) if len(left) <= len(right) else (right, left)
    return len(short) >= 3 and long.startswith(short) and len(long) - len(short) <= 2


def _smooth_spellings(table: ExtractedTable) -> None:
    """
    In one column, a rare misspelling follows the common longer spelling.

    تويو follows تويوتا. A different make stays as it was read.
    Strict reading keeps the spelling that was printed.
    """
    if strict_values() or not table.rows:
        return
    width = max(len(row) for row in table.rows)
    for index in range(width):
        keys: list[str | None] = []
        counts: dict[str, int] = {}
        for row in table.rows:
            if _nondata_kind(row, table.columns):
                keys.append(None)
                continue
            value = row[index] if index < len(row) else ""
            if not _arabic_only_text(value) or " " in value:
                keys.append(None)
                continue
            key = _tidy_arabic(value)
            keys.append(key or None)
            if key:
                counts[key] = counts.get(key, 0) + 1
        forms = list(counts)
        parent = {form: form for form in forms}

        def find(form: str) -> str:
            while parent[form] != form:
                parent[form] = parent[parent[form]]
                form = parent[form]
            return form

        for left_index, left in enumerate(forms):
            for right in forms[left_index + 1 :]:
                if _same_spelling(left, right):
                    parent[find(right)] = find(left)
        chosen: dict[str, str] = {}
        clusters: dict[str, list[str]] = {}
        for form in forms:
            clusters.setdefault(find(form), []).append(form)
        for members in clusters.values():
            best = max(members, key=lambda form: (counts[form], len(form)))
            # A word cut short, such as تويو, follows the full word. An extra letter in the middle does not.
            longer = [
                form
                for form in members
                if form.startswith(best)
                and len(form) > len(best)
                and counts[form] >= max(8, int(counts[best] * 0.2))
            ]
            if longer:
                best = max(longer, key=lambda form: (len(form), counts[form]))
            for form in members:
                chosen[form] = best
        for row, key in zip(table.rows, keys):
            if _nondata_kind(row, table.columns):
                continue
            if key and index < len(row):
                row[index] = chosen.get(key, key)


_DOC_LABELS = {
    "register": "Structured Grid Register",
    "invoice": "Line-item Invoice",
    "financial": "Financial Report",
    "freetext": "Free-text Document",
}
_REGISTER_MARKS = ("رقم اللوحة", "غرض الترخيص", "انتهاء التامين", "الصنع", "اللون السائد")
_INVOICE_MARKS = ("فاتوره", "invoice", "الكميه", "سعر الوحده", "الاجمالي", "unit price", "qty")
_FINANCIAL_MARKS = (
    "قائمة الدخل",
    "قائمه الدخل",
    "الميزانيه",
    "بيان الدخل",
    "balance sheet",
    "income statement",
    "الاصول",
    "التدفقات النقديه",
)
_KEY_CONCEPTS = {"plate", "date", "insurance", "year", "item", "total", "quantity"}
_SHEET_ROW_LIMIT = 5000


def _classify_samples(samples: list[tuple[str, bool, bool]]) -> str:
    """
    register, invoice, financial, or freetext.

    Each sample is the page text, whether it has real letters, and whether
    it has a drawn grid.
    """
    register_hits = 0
    invoice_hits = 0
    financial_hits = 0
    text_pages = 0
    ruled_pages = 0
    for raw, has_text, ruled in samples:
        folded = _normalize(raw).lower()
        compact = folded.replace(" ", "")
        if has_text:
            text_pages += 1
        register_hits += sum(1 for mark in _REGISTER_MARKS if _normalize(mark).replace(" ", "") in compact)
        invoice_hits += sum(1 for mark in _INVOICE_MARKS if _normalize(mark) in folded or _normalize(mark).replace(" ", "") in compact)
        financial_hits += sum(1 for mark in _FINANCIAL_MARKS if _normalize(mark) in folded or _normalize(mark).replace(" ", "") in compact)
        if ruled:
            ruled_pages += 1
    if register_hits >= 3:
        return "register"
    if invoice_hits and invoice_hits >= financial_hits:
        return "invoice"
    if financial_hits:
        return "financial"
    if text_pages == 0:
        return "freetext"
    if ruled_pages and text_pages:
        return "register"
    return "freetext"


def detect_document_type(document: fitz.Document) -> str:
    """
    Classify the file from the first pages, before the full extract.

    A vehicle register, an invoice, a financial statement, and a prose or
    scanned file each take a different reader.
    """
    sample = min(3, document.page_count)
    samples: list[tuple[str, bool, bool]] = []
    for index in range(sample):
        try:
            page = document.load_page(index)
        except Exception:
            continue
        try:
            raw = page.get_text("text") or ""
        except Exception:
            raw = ""
        vertical, horizontal = _page_rules(page)
        samples.append((raw, _text_char_count(page) >= 20, len(vertical) >= 4 and len(horizontal) >= 4))
    return _classify_samples(samples)


def _unify_cell(value: str) -> str:
    """One space between words, and one ISO date when the cell is a date."""
    text = _clean_cell(value)
    if not text:
        return ""
    if re.fullmatch(r"[-–—.•]+", text):
        return ""
    dated = _normalize_date_cell(text)
    if _DATE_VALUE.fullmatch(dated):
        return dated
    return text


def _identifier_indexes(columns: list[str]) -> list[int]:
    """Columns that identify a record: plate, date, year, item, or amount."""
    found: list[int] = []
    for index, name in enumerate(columns):
        concept = _concept_for(name)
        if concept in _KEY_CONCEPTS:
            found.append(index)
    return found


def _date_column_indexes(table: ExtractedTable) -> list[int]:
    """Insurance expiry and any other date column. A plate column is not a date."""
    found: list[int] = []
    for index, name in enumerate(table.columns):
        concept = _concept_for(name)
        if concept == "insurance":
            concept = "date"
        if concept == "date" or _vehicle_field(name) == "date":
            found.append(index)
    return found


def _normalize_table_dates(table: ExtractedTable) -> None:
    """Rewrite RTL-mirrored dates to YYYY-MM-DD before they reach the DataFrame."""
    for index in _date_column_indexes(table):
        for row in table.rows:
            while len(row) <= index:
                row.append("")
            row[index] = _normalize_date_cell(row[index])


def _merge_wrapped_fragment_rows(table: ExtractedTable) -> None:
    """
    Join a wrapped make, colour, or purpose onto the vehicle above it.

    A second visual line with no plate is not a new vehicle. A row that
    still has رقم اللوحة stays its own record. A date is never copied onto
    the vehicle above, and a header or a gap is not folded into it.
    """
    plate_at = _column_index(table, "plate", "رقماللوحه")
    date_at = _column_index(table, "date")
    if plate_at is None or len(table.rows) < 2:
        return
    fold: list[tuple[int, str]] = []
    for index, name in enumerate(table.columns):
        field = _vehicle_field(name) or _concept_for(name)
        if field in {"color", "make", "year", "purpose"}:
            fold.append((index, field))
    if not fold:
        return
    sources = table.sources if table.sources and len(table.sources) == len(table.rows) else None
    lines = table.lines if table.lines and len(table.lines) == len(table.rows) else None
    kept_rows: list[list[str]] = []
    kept_sources: list[int] = []
    kept_lines: list[str] = []
    for index, row in enumerate(table.rows):
        page = sources[index] if sources is not None else 0
        cells = list(row)
        has_date = any(_looks_like_date(cell) for cell in cells)
        if date_at is not None and date_at < len(cells):
            has_date = has_date or bool(_DATE_VALUE.fullmatch(_normalize_date_cell(cells[date_at])))
        if _nondata_kind(cells, table.columns) or _plated_row(cells, plate_at) or has_date or not kept_rows:
            kept_rows.append(cells)
            kept_sources.append(page)
            if lines is not None:
                kept_lines.append(lines[index])
            continue
        if _is_noise_row(cells):
            continue
        previous = kept_rows[-1]
        same_page = kept_sources[-1] == page
        if (
            not same_page
            or not _plated_row(previous, plate_at)
            or _nondata_kind(previous, table.columns)
        ):
            kept_rows.append(cells)
            kept_sources.append(page)
            if lines is not None:
                kept_lines.append(lines[index])
            continue
        for column, field in fold:
            extra = cells[column].strip() if column < len(cells) else ""
            if not extra:
                continue
            while len(previous) <= column:
                previous.append("")
            current = previous[column].strip()
            if current:
                continue
            if field in {"date", "make", "color"} and _looks_like_date(extra):
                continue
            previous[column] = extra
        if lines is not None and kept_lines:
            extra_line = lines[index].strip()
            if extra_line and extra_line not in kept_lines[-1]:
                kept_lines[-1] = f"{kept_lines[-1]} {extra_line}".strip()
    table.rows = kept_rows
    if sources is not None:
        table.sources = kept_sources
    if lines is not None:
        table.lines = kept_lines


def _sanitize_table(table: ExtractedTable) -> None:
    """
    Drop repeated headers, empty records, and tidy every remaining cell.

    A row with a plate, date, year, item, or amount column is kept when one
    of those cells is filled. Other cells are not moved.
    """
    _normalize_table_dates(table)
    _merge_wrapped_fragment_rows(table)
    width = len(table.columns)
    if width == 0:
        table.rows = []
        return
    header_key = tuple(_normalize(name) for name in table.columns)
    keys = _identifier_indexes(table.columns)
    sources = table.sources if table.sources and len(table.sources) == len(table.rows) else None
    lines = table.lines if table.lines and len(table.lines) == len(table.rows) else None
    kept_rows: list[list[str]] = []
    kept_sources: list[int] = []
    kept_lines: list[str] = []
    for index, row in enumerate(table.rows):
        cells = [_unify_cell(row[column] if column < len(row) else "") for column in range(width)]
        if _nondata_kind(list(row), table.columns) or _nondata_kind(cells, table.columns):
            kept_rows.append(list(row) + [""] * max(0, width - len(row)))
            if sources is not None:
                kept_sources.append(sources[index])
            if lines is not None:
                kept_lines.append(lines[index])
            continue
        if not any(cells):
            continue
        if tuple(_normalize(cell) for cell in cells) == header_key:
            continue
        if _is_noise_row(cells):
            continue
        if keys and not any(cells[column].strip() for column in keys if column < len(cells)):
            continue
        # A plate keeps the vehicle even when الصنع or اللون is still empty.
        kept_rows.append(cells)
        if sources is not None:
            kept_sources.append(sources[index])
        if lines is not None:
            kept_lines.append(lines[index])
    table.rows = kept_rows
    if sources is not None:
        table.sources = kept_sources
    if lines is not None:
        table.lines = kept_lines
    _drop_repeated_rows(table)


def _append_change(row: list[str], notes_at: int | None, original: str, corrected: str) -> None:
    """Record one correction. An empty source is written as فراغ so the change is visible."""
    if notes_at is None or original == corrected:
        return
    if not str(original).strip() and not str(corrected).strip():
        return
    note = f"{str(original).strip() or 'فراغ'} -> {corrected}"
    while len(row) <= notes_at:
        row.append("")
    current = str(row[notes_at]).strip()
    if note in current or current in _NONDATA_MARKS:
        return
    row[notes_at] = f"{current}؛ {note}".strip("؛ ")


def _dominant_length_from_rows(rows: list[list[str]], plate_at: int) -> int | None:
    counts: dict[int, int] = {}
    for row in rows:
        token = _plate_token([row[plate_at] if plate_at < len(row) else ""])
        if token.isdigit():
            counts[len(token)] = counts.get(len(token), 0) + 1
    if not counts:
        return None
    length = max(counts, key=lambda item: counts[item])
    if counts[length] < max(3, int(0.6 * sum(counts.values()))):
        return None
    return length


def _repair_plate_text(text: str, length: int | None, votes: dict[int, dict[str, int]]) -> tuple[str, str]:
    """
    Map a look-alike letter inside a plate. An ambiguous letter is left in place.

    O and o become 0. C becomes 0 or 8 only when the same position on the
    other plates of this file votes for one digit.
    """
    raw = _clean_cell(text).translate(_INDIC_DIGITS)
    if not raw or raw.isdigit():
        return raw, ""
    sure = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "B": "8", "Z": "2"})
    chars = list(raw.translate(sure))
    ambiguous = False
    for index, char in enumerate(chars):
        if char.isdigit():
            continue
        if char in {"C", "c"}:
            tally = votes.get(index, {})
            zero = tally.get("0", 0)
            eight = tally.get("8", 0)
            if zero > eight + 1:
                chars[index] = "0"
            elif eight > zero + 1:
                chars[index] = "8"
            else:
                ambiguous = True
            continue
        ambiguous = True
    repaired = "".join(chars)
    if ambiguous or not repaired.isdigit():
        return raw, ""
    if length is not None and len(repaired) != length:
        return raw, ""
    if _YEAR_VALUE.fullmatch(repaired):
        return raw, ""
    return repaired, raw


def _retire_plate_outliers(table: ExtractedTable) -> None:
    """
    A 5-digit token from a date, such as 22898 beside plates 898xx, is not a vehicle.

    It is kept and labeled ترويسة. A page whose plates do not share one prefix
    is left unchanged, so a mixed fleet is not stripped.
    """
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    if plate_at is None or not table.sources or len(table.sources) != len(table.rows):
        return
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    if notes_at is None:
        table.columns = list(table.columns) + [_NOTES_HEADER]
        notes_at = len(table.columns) - 1
    date_at = fields.get("date")
    by_page: dict[int, list[int]] = {}
    for index, page in enumerate(table.sources):
        by_page.setdefault(int(page), []).append(index)
    vehicle_fields = list(fields.values())
    for indexes in by_page.values():
        days: dict[str, int] = {}
        prefixes: dict[str, int] = {}
        for index in indexes:
            row = table.rows[index]
            if date_at is not None and date_at < len(row):
                iso = _normalize_date_cell(row[date_at])
                if _DATE_VALUE.fullmatch(iso or ""):
                    days[iso[8:10]] = days.get(iso[8:10], 0) + 1
            if _nondata_kind(row, table.columns) or plate_at >= len(row):
                continue
            token = _normalize_plate_cell(row[plate_at])
            if token.isdigit() and len(token) >= 4:
                prefixes[token[:2]] = prefixes.get(token[:2], 0) + 1
        if not days or not prefixes:
            continue
        dominant = max(prefixes, key=lambda item: prefixes[item])
        common_days = {value for value, count in days.items() if count >= 3}
        for index in indexes:
            row = table.rows[index]
            if _nondata_kind(row, table.columns) or plate_at >= len(row):
                continue
            token = _normalize_plate_cell(row[plate_at])
            if not token.isdigit() or token.startswith(dominant):
                continue
            if token[:2] not in common_days or prefixes.get(token[:2], 0) >= 4:
                continue
            for field_index in vehicle_fields:
                if field_index < len(row) and field_index != notes_at:
                    row[field_index] = ""
            row[notes_at] = _NONDATA_HEADER


def _repair_plates(table: ExtractedTable) -> None:
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    if plate_at is None:
        return
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    length = _dominant_length_from_rows(table.rows, plate_at)
    votes: dict[int, dict[str, int]] = {}
    for row in table.rows:
        if _nondata_kind(row, table.columns) or plate_at >= len(row):
            continue
        token = str(row[plate_at]).translate(_INDIC_DIGITS)
        if token.isdigit():
            for index, char in enumerate(token):
                votes.setdefault(index, {})
                votes[index][char] = votes[index].get(char, 0) + 1
    for row in table.rows:
        if _nondata_kind(row, table.columns) or plate_at >= len(row):
            continue
        repaired, original = _repair_plate_text(str(row[plate_at]), length, votes)
        if original and repaired != str(row[plate_at]):
            _append_change(row, notes_at, original, repaired)
            row[plate_at] = repaired


def _fill_purpose_down(table: ExtractedTable) -> None:
    """Copy غرض الترخيص forward until the next group, including across pages."""
    fields = _vehicle_field_at(table.columns)
    purpose_at = fields.get("purpose")
    if purpose_at is None:
        return
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    carried = ""
    for row in table.rows:
        while len(row) <= purpose_at:
            row.append("")
        if _nondata_kind(row, table.columns):
            continue
        current = _purpose_token([row[purpose_at]])
        if current:
            if current != str(row[purpose_at]).strip():
                _append_change(row, notes_at, str(row[purpose_at]).strip(), current)
            row[purpose_at] = current
            carried = current
            continue
        if carried:
            row[purpose_at] = carried


def _complete_row_date(cell: str, line: str = "") -> str:
    """
    YYYY-MM-DD from this cell, or from a year printed on the same row band.

    Spaced parts and bidi marks are removed first. A year is used only when
    this row's own text contains it. The row above is not consulted.
    """
    prepared = _strip_bidi_date_noise(cell)
    iso = _normalize_date_cell(prepared) or ""
    if _DATE_VALUE.fullmatch(iso or ""):
        return iso
    band = _strip_bidi_date_noise(line) if line else ""
    if band:
        from_band = _dates_in_text(band) or (_iso_date_from_text(band) or "")
        if _DATE_VALUE.fullmatch(from_band or ""):
            return from_band
    fragment = iso if re.fullmatch(r"\d{2}-\d{2}", iso or "") else _date_fragment(prepared)
    if fragment and band:
        years: list[str] = []
        for token in re.findall(r"(?:19|20)\d{2}", band):
            if token not in years:
                years.append(token)
        if len(years) == 1:
            built = _iso_from_parts(years[0], fragment[:2], fragment[3:])
            if built:
                return built
    return fragment or iso


def _repair_group_dates(table: ExtractedTable) -> None:
    """
    Normalize a date that is already complete, or complete it from this row's band.

    A missing year is not taken from the licence group. A fragment with no
    year on its own line stays a fragment.
    """
    fields = _vehicle_field_at(table.columns)
    date_at = fields.get("date")
    if date_at is None:
        return
    lines = list(table.lines or [])
    for index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        while len(row) <= date_at:
            row.append("")
        original = str(row[date_at]).strip()
        line = lines[index] if index < len(lines) else ""
        iso = _complete_row_date(original, line)
        if iso and iso != original:
            row[date_at] = iso


def neighbourhood_make(columns: list[str], rows: list[list[str]], pages: list[int], index: int) -> str:
    """
    The make shared by the neighbouring vehicle rows on the same page and group.

    One neighbour is enough at the edge of the page. Two neighbours that name
    different makes produce no suggestion.
    """
    fields = _vehicle_field_at(columns)
    make_at = fields.get("make")
    if make_at is None or index < 0 or index >= len(rows):
        return ""
    purpose_at = fields.get("purpose")
    page = int(pages[index]) if index < len(pages) else None
    purpose = ""
    if purpose_at is not None and purpose_at < len(rows[index]):
        purpose = str(rows[index][purpose_at]).strip()
    found: list[str] = []
    for step in (-1, 1):
        cursor = index + step
        while 0 <= cursor < len(rows):
            if page is not None and cursor < len(pages) and int(pages[cursor]) != page:
                break
            if _nondata_kind(rows[cursor], columns):
                cursor += step
                continue
            other = ""
            if purpose_at is not None and purpose_at < len(rows[cursor]):
                other = str(rows[cursor][purpose_at]).strip()
            if purpose and other and other != purpose:
                break
            text = str(rows[cursor][make_at]).strip() if make_at < len(rows[cursor]) else ""
            if text in _MANUFACTURERS:
                found.append(text)
                break
            cursor += step
    if found and len(set(found)) == 1:
        return found[0]
    return ""


def neighbourhood_color(columns: list[str], rows: list[list[str]], pages: list[int], index: int) -> str:
    """The colour shared by neighbouring rows on the same page and group. Not written."""
    fields = _vehicle_field_at(columns)
    color_at = fields.get("color")
    if color_at is None or index < 0 or index >= len(rows):
        return ""
    purpose_at = fields.get("purpose")
    page = int(pages[index]) if index < len(pages) else None
    purpose = ""
    if purpose_at is not None and purpose_at < len(rows[index]):
        purpose = str(rows[index][purpose_at]).strip()
    found: list[str] = []
    for step in (-1, 1):
        cursor = index + step
        while 0 <= cursor < len(rows):
            if page is not None and cursor < len(pages) and int(pages[cursor]) != page:
                break
            if _nondata_kind(rows[cursor], columns):
                cursor += step
                continue
            other = ""
            if purpose_at is not None and purpose_at < len(rows[cursor]):
                other = str(rows[cursor][purpose_at]).strip()
            if purpose and other and other != purpose:
                break
            text = str(rows[cursor][color_at]).strip() if color_at < len(rows[cursor]) else ""
            if text in _OFFICIAL_COLORS:
                found.append(text)
                break
            cursor += step
    if found and len(set(found)) == 1:
        return found[0]
    return ""


def _finalize_name_columns(table: ExtractedTable) -> None:
    """
    Write an official name only when this row's raw reading is a close slip.

    توبو may become تويوتا, and that is logged as normalized. An empty cell,
    a weak cell, or a cell whose text is already the official name is not
    filled from the catalog. Unknown Arabic stays so it can be flagged.
    """
    fields = _vehicle_field_at(table.columns)
    pages = list(table.sources or [])
    for index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        page = int(pages[index]) if index < len(pages) else 0
        plate = _plate_token([row[fields["plate"]]]) if "plate" in fields and fields["plate"] < len(row) else ""
        for field, official in (("make", _MANUFACTURERS), ("color", _OFFICIAL_COLORS)):
            column = fields.get(field)
            if column is None:
                continue
            while len(row) <= column:
                row.append("")
            current = str(row[column]).strip()
            key = _source_key(page, plate, field)
            raw = table.raw_readings.get(key) or current
            if raw and key not in table.raw_readings:
                table.raw_readings[key] = raw
            source = table.provenance.get(key, "")
            if source == "manual":
                continue
            if source == "agreed" and current in official:
                continue
            stored = table.raw_readings.get(key, "")
            reading = stored or current
            if field == "make" and _short_make_fragment(reading):
                row[column] = ""
                _remember_source(table, page, plate, field, "suggested", "")
                continue
            accepted = _accept_make(reading) if field == "make" else _accept_color(reading)
            if accepted and _raw_is_slip(field, reading, accepted):
                row[column] = accepted
                if stored and key not in table.raw_readings:
                    table.raw_readings[key] = stored
                _remember_source(table, page, plate, field, "normalized", f"{accepted} — تطبيع: {reading}")
                continue
            # An official name already read on this row stays. It is not replaced
            # by the top of the catalog, and an empty cell is not filled with it.
            if current in official and (not stored or stored == current or stored in official):
                if source not in {"agreed", "manual", "normalized"}:
                    _remember_source(table, page, plate, field, "one-engine", "")
                continue
            if current in official:
                row[column] = ""
            kept = _keep_arabic_value(field, stored or current) if (stored or current) else ""
            if kept in official:
                kept = ""
            row[column] = kept
            if kept:
                _remember_source(table, page, plate, field, "one-engine", "")
                continue
            if stored in official:
                _add_reasoned_suggestion(table, page, plate, field, stored, "قراءة المقطع")
            if plate and not row[column]:
                _remember_source(
                    table,
                    page,
                    plate,
                    field,
                    "suggested",
                    table.suggestions.get(key, ""),
                )


def _repair_text_columns(table: ExtractedTable) -> None:
    """Normalize a close slip. A neighbour is a suggestion with its reason, never a value."""
    _finalize_name_columns(table)
    fields = _vehicle_field_at(table.columns)
    pages = list(table.sources or [])
    for index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        page = int(pages[index]) if index < len(pages) else 0
        plate = _plate_token([row[fields["plate"]]]) if "plate" in fields and fields["plate"] < len(row) else ""
        for field, neighbour_of in (
            ("make", neighbourhood_make),
            ("color", neighbourhood_color),
        ):
            column = fields.get(field)
            if column is None:
                continue
            current = str(row[column]).strip() if column < len(row) else ""
            neighbour = neighbour_of(table.columns, table.rows, pages, index)
            if neighbour and neighbour != current:
                _add_reasoned_suggestion(table, page, plate, field, neighbour, "جوار نفس الصفحة والمجموعة")
                key = _source_key(page, plate, field)
                if table.provenance.get(key) not in {"agreed", "manual", "normalized", "one-engine"}:
                    _remember_source(table, page, plate, field, "suggested", table.suggestions.get(key, ""))
    _release_adjacent_rare(table)


def _release_adjacent_rare(table: ExtractedTable) -> None:
    """
    A rare colour, make, or date printed on two neighbouring rows is not written.

    The glyph was between the lines and got attached to both. Both cells stay
    empty and are flagged. The page's common value is left as it is.
    """
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    if plate_at is None:
        return
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    sources = list(table.sources or [])
    for field, note in (
        ("color", "لون نادر مكرر على صفين متجاورين"),
        ("make", "صنع نادر مكرر على صفين متجاورين"),
        ("date", "تاريخ نادر مكرر على صفين متجاورين"),
    ):
        column = fields.get(field)
        if column is None:
            continue
        by_page: dict[int, list[tuple[int, str, str]]] = {}
        for index, row in enumerate(table.rows):
            if _nondata_kind(row, table.columns):
                continue
            plate = _plate_token([row[plate_at]]) if plate_at < len(row) else ""
            if not plate:
                continue
            page = int(sources[index]) if index < len(sources) else 0
            value = str(row[column]).strip() if column < len(row) else ""
            by_page.setdefault(page, []).append((index, plate, value))
        for page, items in by_page.items():
            counts: dict[str, int] = {}
            for _index, _plate, value in items:
                if value:
                    counts[value] = counts.get(value, 0) + 1
            if len(counts) < 2:
                continue
            dominant = max(counts, key=lambda key: (counts[key], key))
            for pos in range(len(items) - 1):
                index_a, plate_a, value_a = items[pos]
                index_b, plate_b, value_b = items[pos + 1]
                if not value_a or value_a != value_b or value_a == dominant or counts[value_a] > 2:
                    continue
                for index, plate in ((index_a, plate_a), (index_b, plate_b)):
                    row = table.rows[index]
                    while len(row) <= column:
                        row.append("")
                    row[column] = ""
                    _add_reasoned_suggestion(table, page, plate, field, value_a, "قيمة نادرة على صفين متجاورين")
                    _remember_source(table, page, plate, field, "suggested", table.suggestions.get(_source_key(page, plate, field), ""))
                    if notes_at is not None:
                        while len(row) <= notes_at:
                            row.append("")
                        current_note = str(row[notes_at]).strip()
                        if note not in current_note:
                            row[notes_at] = f"{current_note}؛ {note}".strip("؛ ").strip()


def _learned_path() -> Path:
    """Where manual corrections are kept for the next file."""
    try:
        from engine.paths import learned_file

        return learned_file()
    except Exception:
        override = os.environ.get("PDF_TOOL_LEARNED", "").strip()
        if override:
            return Path(override)
        return Path(__file__).resolve().parent.parent / "learned_corrections.json"


def _stored_value(entry: object) -> str:
    """The corrected text of one dictionary entry. Older files stored a string."""
    if isinstance(entry, dict):
        return str(entry.get("value") or "").strip()
    return str(entry or "").strip()


def load_corrections() -> dict[str, dict[str, dict[str, str]]]:
    """The learning dictionary. Each entry keeps the value, the date, and the source."""
    path = _learned_path()
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    cleaned: dict[str, dict[str, dict[str, str]]] = {}
    for field, pairs in data.items():
        if not isinstance(pairs, dict):
            continue
        bucket: dict[str, dict[str, str]] = {}
        for key, entry in pairs.items():
            value = _stored_value(entry)
            if not str(key).strip() or not value:
                continue
            when = str(entry.get("date") or "") if isinstance(entry, dict) else ""
            origin = str(entry.get("source") or "") if isinstance(entry, dict) else ""
            bucket[str(key)] = {"value": value, "date": when, "source": origin}
        if bucket:
            cleaned[str(field)] = bucket
    return cleaned


def save_corrections(data: dict) -> None:
    """Write the learning dictionary. The caller can also download this file."""
    path = _learned_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def remember_correction(field: str, original: str, corrected: str, source: str = "manual") -> None:
    """Save one manual correction and reuse it on later files."""
    written = str(original).strip()
    target = str(corrected).strip()
    if not field or not target or written == target:
        return
    data = load_corrections()
    data.setdefault(field, {})[written or "فراغ"] = {
        "value": target,
        "date": date.today().isoformat(),
        "source": source or "manual",
    }
    save_corrections(data)


def forget_correction(field: str, original: str) -> None:
    """Remove one saved correction."""
    data = load_corrections()
    bucket = data.get(field) or {}
    key = str(original).strip() or "فراغ"
    if key not in bucket:
        return
    bucket.pop(key, None)
    if bucket:
        data[field] = bucket
    else:
        data.pop(field, None)
    save_corrections(data)


def _apply_learned(table: ExtractedTable) -> None:
    """Reuse a saved correction. The note records original -> corrected."""
    learned = load_corrections()
    if not learned:
        return
    fields = _vehicle_field_at(table.columns)
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    pages = list(table.sources or [])
    plate_at = fields.get("plate")
    for row_index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        page = int(pages[row_index]) if row_index < len(pages) else 0
        plate = _plate_token([row[plate_at]]) if plate_at is not None and plate_at < len(row) else ""
        for field, index in fields.items():
            pairs = learned.get(field) or {}
            if index >= len(row) or not pairs:
                continue
            original = str(row[index]).strip()
            # A saved correction must not replay a prior. Dates, blanks, and
            # make/colour guesses stay out of the cell; the review list flags them.
            if field in {"date", "make", "color"} or not original:
                continue
            entry = pairs.get(original)
            corrected = _stored_value(entry)
            if corrected and corrected != original:
                _append_change(row, notes_at, original, corrected)
                row[index] = corrected
                _remember_raw(table, page, plate, field, original)
                _remember_source(table, page, plate, field, "manual", corrected)


def _exact_make_in_text(text: str) -> str:
    """One exact dictionary make in this text. A fuzzy resemblance is ignored."""
    tokens = [part for part in str(text).split() if part]
    hits: list[str] = []
    for size in (4, 3, 2, 1):
        if size > len(tokens):
            continue
        for start in range(0, len(tokens) - size + 1):
            hit = _exact_manufacturer(" ".join(tokens[start : start + size]))
            if hit and hit not in hits:
                hits.append(hit)
    if len(hits) == 1:
        return hits[0]
    return ""


def _exact_color_in_text(text: str) -> str:
    """One exact dictionary colour in this text."""
    words = re.findall(r"[\u0600-\u06FF]+", str(text))
    hits: list[str] = []
    for size in (2, 1):
        if size > len(words):
            continue
        for start in range(0, len(words) - size + 1):
            hit = _exact_color(" ".join(words[start : start + size]))
            if hit and hit not in hits:
                hits.append(hit)
    if len(hits) == 1:
        return hits[0]
    return ""


def _canonical_reading(field: str, text: str) -> str:
    """The official or well-formed value of one OCR reading, or empty."""
    raw = str(text or "").strip()
    if not raw:
        return ""
    if field == "make":
        return _exact_make_in_text(raw)
    if field == "color":
        return _exact_color_in_text(raw)
    if field == "year":
        token = re.sub(r"\D", "", raw.translate(_INDIC_DIGITS))
        return token if _YEAR_VALUE.fullmatch(token or "") else ""
    if field == "date":
        return _dates_in_text(raw)
    return ""


def vote_readings(
    field: str,
    left: str,
    right: str,
    left_confidence: float = 0,
    right_confidence: float = 0,
) -> tuple[str, str]:
    """
    Accept a value only when both readings normalize to the same text.

    One engine, even when it names an official make, is a suggestion.
    The cell is not filled from that single reading.
    """
    del left_confidence, right_confidence
    a = _canonical_reading(field, left)
    b = _canonical_reading(field, right)
    if a and b and a == b:
        return a, "اتفاق المحركين"
    if a and b and a != b:
        return "", f"لم يتفق المحركان: {a} | {b}"
    candidate = a or b
    if candidate:
        return "", f"محرك واحد: {candidate}"
    return "", "لم يتفق المحركان"


def _unique_official_make(text: str) -> str:
    """One official make named in this text. Two different makes are left alone."""
    tokens = [part for part in str(text).split() if part]
    hits: list[str] = []
    for size in (4, 3, 2, 1):
        if size > len(tokens):
            continue
        for start in range(0, len(tokens) - size + 1):
            hit = _match_manufacturer(" ".join(tokens[start : start + size]))
            if hit in _MANUFACTURERS and hit not in hits:
                hits.append(hit)
    if len(hits) == 1:
        return hits[0]
    return ""


def _unique_official_color(text: str) -> str:
    """One official colour named in this text."""
    hits: list[str] = []
    words = re.findall(r"[\u0600-\u06FF]+", str(text))
    for size in (2, 1):
        if size > len(words):
            continue
        for start in range(0, len(words) - size + 1):
            hit = _match_official_color(" ".join(words[start : start + size]))
            if hit in _OFFICIAL_COLORS and hit not in hits:
                hits.append(hit)
    if len(hits) == 1:
        return hits[0]
    return ""


def _dates_in_text(text: str, plate: str = "") -> str:
    """One ISO date in the line. A plate and a lone model year are not a date.

    The printed order is year, then month, then day. The same date copied from
    the row above produces an overlapping day-month-year window; that window
    does not cancel the year-first date.
    """
    cleaned = str(text).translate(_INDIC_DIGITS)
    if plate:
        cleaned = cleaned.replace(plate, " ")
    groups = re.findall(r"\d+", cleaned)
    year_first: list[str] = []
    year_last: list[str] = []
    for start in range(0, max(len(groups) - 2, 0)):
        window = groups[start : start + 3]
        if len(window[0]) == 4 and len(window[1]) <= 2 and len(window[2]) <= 2:
            iso = _iso_from_parts(window[0], window[1], window[2])
            if iso and iso not in year_first:
                year_first.append(iso)
        elif len(window[2]) == 4 and len(window[0]) <= 2 and len(window[1]) <= 2:
            iso = _iso_from_parts(window[2], window[1], window[0])
            if iso and iso not in year_last:
                year_last.append(iso)
    # A model year has no month and day, so it never becomes a second date.
    # Two different year-first dates are left blank instead of picking one.
    if len(year_first) == 1:
        return year_first[0]
    if not year_first and len(year_last) == 1:
        return year_last[0]
    return ""


def _year_in_text(text: str, plate: str = "", date: str = "") -> str:
    """The model year on this line, not the year inside the insurance date."""
    cleaned = str(text).translate(_INDIC_DIGITS)
    if plate:
        cleaned = cleaned.replace(plate, " ")
    years = [token for token in re.findall(r"\d{4}", cleaned) if _YEAR_VALUE.fullmatch(token)]
    if date and len(date) >= 4:
        years = [token for token in years if token != date[:4]]
    unique = list(dict.fromkeys(years))
    if len(unique) == 1:
        return unique[0]
    return ""


def _log_note(row: list[str], notes_at: int | None, text: str) -> None:
    """Append one attempt. A ترويسة or فاصل mark is left exact."""
    if notes_at is None or not text:
        return
    while len(row) <= notes_at:
        row.append("")
    current = str(row[notes_at]).strip()
    if current in _NONDATA_MARKS or text in current:
        return
    row[notes_at] = f"{current}؛ {text}".strip("؛ ")


def _fill_blanks_from_line(table: ExtractedTable) -> None:
    """
    Read make, colour, year, and date from the same printed line as the plate.

    The words were already on that baseline. Nothing is copied from the row above.
    """
    fields = _vehicle_field_at(table.columns)
    if "plate" not in fields or not table.lines or len(table.lines) != len(table.rows):
        return
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    for index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        plate = _plate_token([row[fields["plate"]] if fields["plate"] < len(row) else ""])
        if not plate:
            continue
        line = table.lines[index] or ""
        date_at = fields.get("date")
        page = int(table.sources[index]) if table.sources and index < len(table.sources) else 0
        if date_at is not None and date_at < len(row) and not str(row[date_at]).strip():
            found = _dates_in_text(line, plate)
            if found:
                _add_reasoned_suggestion(table, page, plate, "date", found, "قراءة سطر الصف")
                key = _source_key(page, plate, "date")
                if table.provenance.get(key) not in {"agreed", "manual", "normalized"}:
                    _remember_source(table, page, plate, "date", "suggested", table.suggestions.get(key, ""))
        year_at = fields.get("year")
        current_date = str(row[date_at]).strip() if date_at is not None and date_at < len(row) else ""
        if year_at is not None and year_at < len(row) and not str(row[year_at]).strip():
            found = _year_in_text(line, plate, current_date)
            if found:
                row[year_at] = found
                _append_change(row, notes_at, "", found)
        make_at = fields.get("make")
        if make_at is not None and make_at < len(row) and not str(row[make_at]).strip():
            found = _exact_make_in_text(line)
            if found:
                _add_reasoned_suggestion(table, page, plate, "make", found, "قراءة سطر الصف")
                key = _source_key(page, plate, "make")
                if table.provenance.get(key) not in {"agreed", "manual", "normalized"}:
                    _remember_source(table, page, plate, "make", "suggested", table.suggestions.get(key, ""))
        color_at = fields.get("color")
        if color_at is not None and color_at < len(row) and not str(row[color_at]).strip():
            found = _exact_color_in_text(line)
            if found:
                _add_reasoned_suggestion(table, page, plate, "color", found, "قراءة سطر الصف")
                key = _source_key(page, plate, "color")
                if table.provenance.get(key) not in {"agreed", "manual", "normalized"}:
                    _remember_source(table, page, plate, "color", "suggested", table.suggestions.get(key, ""))


def _infer_group_years(table: ExtractedTable) -> None:
    """
    Suggest the dominant year of neighbouring dates. The cell is not filled.

    04-09 stays 04-09. The provenance suggestion carries 2026-04-09 when the
    other rows on the page agree on 2026.
    """
    fields = _vehicle_field_at(table.columns)
    date_at = fields.get("date")
    plate_at = fields.get("plate")
    if date_at is None:
        return
    pages = list(table.sources or [])
    years_on_page: dict[int, list[str]] = {}
    for index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns) or date_at >= len(row):
            continue
        text = str(row[date_at]).strip()
        if not _DATE_VALUE.fullmatch(text):
            continue
        page = int(pages[index]) if index < len(pages) else 0
        years_on_page.setdefault(page, []).append(text[:4])
    for index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns) or date_at >= len(row):
            continue
        text = str(row[date_at]).strip()
        if not re.fullmatch(r"\d{2}-\d{2}", text):
            continue
        page = int(pages[index]) if index < len(pages) else 0
        counts: dict[str, int] = {}
        for year in years_on_page.get(page, []):
            counts[year] = counts.get(year, 0) + 1
        if not counts:
            continue
        ordered = sorted(counts.items(), key=lambda item: (-item[1], item[0]))
        best_year, best_count = ordered[0]
        second_count = ordered[1][1] if len(ordered) > 1 else 0
        if best_count <= second_count:
            continue
        suggestion = f"{best_year}-{text}"
        plate = _plate_token([row[plate_at]]) if plate_at is not None and plate_at < len(row) else ""
        _add_reasoned_suggestion(table, page, plate, "date", suggestion, "جوار نفس الصفحة والمجموعة")
        key = _source_key(page, plate, "date")
        if table.provenance.get(key) not in {"agreed", "manual", "normalized"}:
            _remember_source(table, page, plate, "date", "suggested", table.suggestions.get(key, ""))
        _remember_raw(table, page, plate, "date", text)


def _png_variant(image: bytes, mode: str) -> bytes:
    """Contrast, black-and-white, or a small deskew of one crop."""
    import io

    from PIL import Image, ImageOps

    picture = ImageOps.autocontrast(Image.open(io.BytesIO(image)).convert("L"))
    if mode == "binary":
        picture = picture.point(lambda pixel: 255 if pixel > 150 else 0)
    elif mode == "deskew":
        best = picture
        best_score = -1.0
        width, height = picture.size
        for angle in (-2, -1, 0, 1, 2):
            turned = picture.rotate(angle, fillcolor=255)
            pixels = list(turned.get_flattened_data() if hasattr(turned, "get_flattened_data") else turned.getdata())
            step = max(1, height // 40)
            inks: list[int] = []
            for y in range(0, height, step):
                inks.append(sum(1 for pixel in pixels[y * width : (y + 1) * width] if pixel < 140))
            if not inks:
                continue
            mean = sum(inks) / len(inks)
            score = sum((ink - mean) ** 2 for ink in inks) / len(inks)
            if score > best_score:
                best_score = score
                best = turned
        picture = best
    buffer = io.BytesIO()
    picture.save(buffer, format="PNG")
    return buffer.getvalue()


def _footer_rule_y(page: fitz.Page) -> float | None:
    """
    The horizontal rule above the footer, in PDF points.

    The line is the thin dark run in the lower part of this page. A block of
    footer words is not a rule. Pages with no such line return None.
    """
    scale = 2.0
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    except Exception:
        return None
    width, height, channels = pixmap.width, pixmap.height, max(1, pixmap.n)
    samples = pixmap.samples
    if not samples or height < 20 or width < 20:
        return None
    left = int(width * 0.08)
    right = int(width * 0.92)
    span = max(1, right - left)
    start = int(height * 0.55)
    runs: list[int] = []
    for y in range(start, height):
        best = run = 0
        row = y * width * channels
        for x in range(left, right):
            offset = row + x * channels
            pixel = samples[offset]
            if channels >= 3:
                pixel = (pixel + samples[offset + 1] + samples[offset + 2]) // 3
            if pixel < 210:
                run += 1
                if run > best:
                    best = run
            else:
                run = 0
        runs.append(best)
    best_index = None
    best_run = 0
    for index, run in enumerate(runs):
        if run < span * 0.12:
            continue
        above = runs[index - 4] if index >= 4 else 0
        below = runs[index + 4] if index + 4 < len(runs) else 0
        if run < max(above, below):
            continue
        if run > best_run:
            best_run = run
            best_index = index
    if best_index is None:
        return None
    return (start + best_index) / scale


def _collapse_baselines(ys: list[float], gap: float = 8.0) -> list[float]:
    """One baseline where two engines read the same plate a few points apart."""
    groups: list[list[float]] = []
    for y_value in sorted(ys):
        if groups and y_value - groups[-1][-1] <= gap:
            groups[-1].append(y_value)
        else:
            groups.append([y_value])
    return [sum(group) / len(group) for group in groups]


def _row_band_limits(
    y_value: float,
    row_ys: list[float],
    footer_y: float | None,
    page_height: float,
) -> tuple[float, float]:
    """
    The vertical band of one row.

    The plate box sits lower than the printed glyphs, so the crop starts half
    a row above the previous baseline and the digit tops stay inside. The
    bottom is the midpoint to the next baseline, or one point above this
    page's footer rule. A fixed offset is not used.
    """
    previous = max((item for item in row_ys if item < y_value - 0.8), default=None)
    following = min((item for item in row_ys if item > y_value + 0.8), default=None)
    pitch = 12.0
    gaps = [abs(item - y_value) for item in row_ys if abs(item - y_value) > 0.8]
    if gaps:
        pitch = min(gaps)
    # The top stops at the midpoint. Starting above the previous baseline
    # pulled that vehicle's colour and date onto this row.
    if previous is not None:
        top = (previous + y_value) / 2.0
    else:
        top = y_value - pitch * 0.45
    if following is not None:
        bottom = (y_value + following) / 2.0
    elif footer_y is not None and footer_y > y_value + 1.0:
        bottom = footer_y - 1.0
    else:
        bottom = y_value + pitch
    top = max(0.0, top)
    bottom = min(page_height, max(top + 4.0, bottom))
    return top, bottom


def _band_text(words: list[Word], y0: float, y1: float) -> tuple[str, float]:
    """Words between two baselines, and the best confidence among them."""
    near = [word for word in words if y0 <= word.center_y <= y1]
    near.sort(key=lambda word: word.x0)
    confidence = max((word.confidence for word in near), default=0.0)
    return " ".join(word.text for word in near), confidence


def _date_column_box(words: list[Word]) -> tuple[float, float] | None:
    """X span of the insurance-date column, from the leftmost year cluster on the page."""
    years = []
    for word in words:
        digits = re.sub(r"\D", "", word.text.translate(_INDIC_DIGITS))
        if len(digits) >= 4 and _YEAR_VALUE.fullmatch(digits[:4]):
            years.append(word)
    if len(years) < 3:
        return None
    years.sort(key=lambda word: word.x0)
    anchor = years[0].x0
    cluster = [word for word in years if word.x0 <= anchor + 28]
    if len(cluster) < 2:
        return None
    return min(word.x0 for word in cluster) - 6, max(word.x1 for word in cluster) + 42


def _lowest_line_date(words: list[Word], plate: str, min_y: float | None = None) -> str:
    """The date on the lowest line at or below min_y. The row above is left out."""
    if not words:
        return ""
    found = ""
    lowest = -1.0
    for line in _group_rows(words, 4.0):
        text = " ".join(word.text for word in sorted(line, key=lambda word: word.x0))
        date_text = _dates_in_text(text, plate)
        y_value = sum(word.center_y for word in line) / len(line)
        if min_y is not None and y_value < min_y:
            continue
        if date_text and y_value >= lowest:
            lowest = y_value
            found = date_text
    return found


def _short_reading(text: str) -> str:
    """One short engine reading for the notes column."""
    compact = " ".join(str(text or "").split())
    if not compact:
        return "فراغ"
    return compact[:80]


def _y_of_token(words: list[Word], token: str) -> float | None:
    hits = [word.center_y for word in words if token and token in re.sub(r"\D", "", word.text.translate(_INDIC_DIGITS))]
    if not hits:
        return None
    hits.sort()
    return hits[len(hits) // 2]


_PLATE_LOOKALIKES = str.maketrans(
    {"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "B": "8", "Z": "2", "C": "0", "c": "0"}
)


def _locate_plate_y(words: list[Word], plate: str) -> float | None:
    """The row of this plate, including a one-letter slip such as 895C5 for 89505."""
    found = _y_of_token(words, plate)
    if found is not None:
        return found
    hits = []
    for word in words:
        token = re.sub(r"\D", "", word.text.translate(_INDIC_DIGITS).translate(_PLATE_LOOKALIKES))
        if token == plate:
            hits.append(word.center_y)
    if hits:
        hits.sort()
        return hits[len(hits) // 2]
    return _y_of_near_plate(words, plate)


def _y_of_near_plate(words: list[Word], plate: str) -> float | None:
    """Baseline of one plate that differs by a single digit. Two such plates are ignored."""
    if not plate or not plate.isdigit():
        return None
    found: float | None = None
    for word in words:
        token = _plate_token([word.text])
        if not token or len(token) != len(plate) or not token.isdigit():
            continue
        if sum(left != right for left, right in zip(token, plate)) != 1:
            continue
        if found is not None:
            return None
        found = word.center_y
    return found


def _read_page_pair(page: fitz.Page, language: str) -> tuple[list[Word], list[Word]]:
    """Higher-DPI Windows OCR and Tesseract on one page."""
    scale = 4.0
    long_side = max(page.rect.width, page.rect.height) * scale
    if long_side > 4500:
        scale = 4500 / max(page.rect.width, page.rect.height)
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    width = float(pixmap.width)
    height = float(pixmap.height)
    image = _png_variant(pixmap.tobytes("png"), "contrast")
    del pixmap
    rect = fitz.Rect(page.rect)
    windows = _ocr_words(image, language, rect, width, height)
    tess_lang = "ara+eng" if language == "ar-SA" else "eng"
    tesseract = _tesseract_words(image, rect, width, height, tess_lang)
    return windows, tesseract


def _apply_voted_field(
    row: list[str],
    notes_at: int | None,
    index: int,
    field: str,
    left: str,
    right: str,
    left_confidence: float,
    right_confidence: float,
    attempt: str,
    source_table: ExtractedTable | None = None,
    source_page: int = 0,
    source_plate: str = "",
) -> bool:
    """Write a value only when both readings agree. One engine stays a suggestion."""
    while len(row) <= index:
        row.append("")
    original = str(row[index]).strip()
    if score_cell(field, original) >= 0.8 and field not in {"make", "color"}:
        return False
    chosen, reason = vote_readings(field, left, right, left_confidence, right_confidence)
    _log_note(row, notes_at, f"محاولة {attempt} {field}: {reason}")
    if chosen and chosen == original:
        if source_table is not None:
            _remember_source(source_table, source_page, source_plate, field, "agreed")
        return False
    if chosen:
        _append_change(row, notes_at, original, chosen)
        row[index] = chosen
        if source_table is not None:
            _remember_source(source_table, source_page, source_plate, field, "agreed")
        return True
    candidate = ""
    if ":" in reason and reason.split(":", 1)[0] in {"محرك واحد", "لم يتفق المحركان"}:
        candidate = reason.split(":", 1)[1].strip()
    if candidate:
        _log_note(row, notes_at, f"اقتراح {field}: {candidate}")
        if source_table is not None:
            _remember_source(source_table, source_page, source_plate, field, "one-engine", candidate)
    if field in {"make", "color"} and original:
        kept = _keep_arabic_value(field, original)
        if kept:
            row[index] = kept
    return False


def _ocr_clip_words(page: fitz.Page, clip: fitz.Rect, language: str = "ar-SA") -> list[Word]:
    """Windows OCR on one clip. Word boxes are returned in page coordinates."""
    if clip.width < 4 or clip.height < 4:
        return []
    scale = 4.0
    long_side = max(clip.width, clip.height) * scale
    if long_side > 4500:
        scale = 4500 / max(clip.width, clip.height)
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
        png = pixmap.tobytes("png")
        width = float(pixmap.width)
        height = float(pixmap.height)
        del pixmap
        words = _ocr_words(png, language, fitz.Rect(clip), width, height)
    except Exception:
        return []
    return words


def _page_anchor_words(page: fitz.Page) -> list[Word]:
    """Glyph words when the page has a text layer, otherwise one Arabic OCR of the page."""
    words = _page_words(page)
    plates = [word for word in words if _plate_token([word.text])]
    if len(plates) >= 8 and _vehicle_header_spans(words):
        return words
    try:
        image, rect, pixel_width, pixel_height = _scan_image(page)
        ocr_words = _ocr_words(image, "ar-SA", rect, pixel_width, pixel_height)
    except Exception:
        return words
    if len(ocr_words) > len(words):
        return ocr_words
    return words or ocr_words


def _header_seed(words: list[Word], field: str) -> Word | None:
    """The title word for الصنع or اللون, including the الصنج slip this scan prints."""
    hits: list[Word] = []
    for word in words:
        folded = _logical_key(word.text)
        if field == "make" and any(token in folded for token in ("صنع", "صنج", "صتع")):
            hits.append(word)
        elif field == "color" and "لون" in folded:
            hits.append(word)
    if not hits:
        return None
    top = min(word.center_y for word in hits)
    header = [word for word in hits if word.center_y <= top + 10]
    return max(header, key=lambda word: len(word.text))


def _column_from_seed(words: list[Word], seed: Word, page_width: float) -> tuple[float, float]:
    """Widen a header word across the spaced letters printed in that column."""
    near = [
        word
        for word in words
        if abs(word.center_x - seed.center_x) <= 42
        and word.center_y >= seed.center_y - 2
        and not re.search(r"\d", word.text)
    ] or [seed]
    return max(0.0, min(word.x0 for word in near) - 6), min(page_width, max(word.x1 for word in near) + 6)


def _infer_arabic_boxes(words: list[Word], page_width: float) -> dict[str, tuple[float, float]]:
    """X span of الصنع and اللون from the words themselves when the header was not read."""
    bins: dict[int, list[Word]] = {}
    for word in words:
        if re.search(r"\d", word.text) or not any(_is_arabic_letter(char) for char in word.text):
            continue
        bins.setdefault(int(word.center_x // 22), []).append(word)
    make_best: tuple[int, float, float] | None = None
    color_best: tuple[int, float, float] | None = None
    for group in bins.values():
        if len(group) < 4:
            continue
        make_hits = sum(1 for word in group if _accept_make(word.text))
        color_hits = sum(1 for word in group if _accept_color(word.text))
        x0 = max(0.0, min(word.x0 for word in group) - 4)
        x1 = min(page_width, max(word.x1 for word in group) + 4)
        if make_best is None or make_hits > make_best[0]:
            make_best = (make_hits, x0, x1)
        if color_best is None or color_hits > color_best[0]:
            color_best = (color_hits, x0, x1)
    boxes: dict[str, tuple[float, float]] = {}
    if make_best and make_best[0] >= 3:
        boxes["make"] = (make_best[1], make_best[2])
    if color_best and color_best[0] >= 3:
        boxes["color"] = (color_best[1], color_best[2])
    if "make" in boxes and "color" in boxes and abs(boxes["make"][0] - boxes["color"][0]) < 12:
        if make_best and color_best and make_best[0] >= color_best[0]:
            boxes.pop("color", None)
        else:
            boxes.pop("make", None)
    return boxes


def _arabic_column_boxes(page: fitz.Page, words: list[Word]) -> dict[str, tuple[float, float]]:
    """X span of الصنع and اللون. Ruled header edges win. Header words are the fallback."""
    boxes = dict(_attribute_column_boxes(page, words))
    width = float(page.rect.width)
    for field in ("make", "color"):
        if boxes.get(field):
            continue
        seed = _header_seed(words, field)
        if seed is not None:
            boxes[field] = _column_from_seed(words, seed, width)
    if not boxes.get("make") or not boxes.get("color"):
        inferred = _infer_arabic_boxes(words, width)
        for field in ("make", "color"):
            if not boxes.get(field) and inferred.get(field):
                boxes[field] = inferred[field]
    if boxes.get("make") and boxes.get("color"):
        return boxes
    titles = [(field, word) for word in words if (field := _vehicle_field(word.text)) in {"make", "color", "year", "plate", "date", "purpose", "admin"}]
    if len(titles) < 2:
        return boxes
    top = min(word.center_y for _field, word in titles)
    header = [(field, word) for field, word in titles if word.center_y <= top + 14] or titles
    header.sort(key=lambda item: item[1].center_x)
    width = float(page.rect.width)
    centers = [word.center_x for _field, word in header]
    for index, (field, word) in enumerate(header):
        if field not in {"make", "color"} or field in boxes:
            continue
        left = centers[index - 1] if index else max(0.0, word.x0 - 40)
        right = centers[index + 1] if index + 1 < len(centers) else min(width, word.x1 + 40)
        x0 = (left + centers[index]) / 2.0 if index else max(0.0, word.x0 - 6)
        x1 = (centers[index] + right) / 2.0 if index + 1 < len(centers) else min(width, word.x1 + 6)
        if x1 - x0 < 12:
            x0 = max(0.0, centers[index] - 28)
            x1 = min(width, centers[index] + 28)
        boxes[field] = (x0, x1)
    return boxes


def _sits_between_baselines(y_value: float, baselines: list[float]) -> bool:
    """True when the glyph sits on the cut between two plate lines."""
    ordered = sorted({round(float(item), 2) for item in baselines})
    for left, right in zip(ordered, ordered[1:]):
        if right - left < 4:
            continue
        if abs(float(y_value) - (left + right) / 2.0) <= 0.8:
            return True
    return False


def _text_line_offset(words: list[Word], baselines: list[float]) -> float:
    """How far above the plate the Arabic line sits on this page."""
    deltas: list[float] = []
    points = [float(item) for item in baselines]
    for word in words:
        if not any(_is_arabic_letter(char) for char in word.text):
            continue
        above = [point - word.center_y for point in points if 0.8 <= point - word.center_y <= 9.0]
        if above:
            deltas.append(min(above))
    if len(deltas) < 4:
        return 3.5
    deltas.sort()
    return deltas[len(deltas) // 2]


def _glyph_owns_row(
    y_value: float,
    focus: float,
    baselines: list[float],
    offset: float = 0.0,
) -> bool:
    """The glyph belongs to this row when its nearest text line is this row's line."""
    points = [float(item) - float(offset) for item in baselines if item is not None]
    if not points:
        return True
    if _sits_between_baselines(y_value, points):
        return False
    nearest = min(points, key=lambda item: abs(item - float(y_value)))
    return abs(nearest - (float(focus) - float(offset))) <= 1.25


def _arabic_band_text(
    words: list[Word],
    x0: float,
    x1: float,
    y0: float,
    y1: float,
    focus_y: float | None = None,
    baselines: list[float] | None = None,
    offset: float = 0.0,
) -> str:
    """Arabic tokens on this cell's own line. A glyph in the gap is left out."""
    titles = _column_title_keys()
    chosen = [
        word
        for word in words
        if x0 - 1.5 <= word.center_x <= x1 + 1.5 and y0 <= word.center_y <= y1 and any(_is_arabic_letter(char) for char in word.text)
    ]
    if not chosen:
        return ""
    focus = focus_y if focus_y is not None else (y0 + y1) / 2.0
    # Keep a glyph only when this plate is the nearest baseline. The old 8-point
    # window glued a colour that sat in the gap onto the lower row.
    if baselines:
        chosen = [word for word in chosen if _glyph_owns_row(word.center_y, focus, baselines, offset)]
    else:
        chosen = [word for word in chosen if -1.5 <= focus - word.center_y <= 8]
    if not chosen:
        return ""
    anchor = min(chosen, key=lambda word: abs(word.center_y - (focus - offset)))
    chosen = [word for word in chosen if abs(word.center_y - anchor.center_y) <= 4]
    chosen.sort(key=lambda word: -word.x0)
    parts = []
    for word in chosen:
        token = re.sub(r"[^\u0600-\u06FF]", "", word.text)
        if _logical_key(token) in titles:
            continue
        parts.append(word.text)
    return " ".join(parts)


def _arabic_page_job(spec: tuple) -> dict:
    """Re-read one page's make and colour columns in a worker process."""
    path, page_number, password, strict, columns, rows, sources = spec
    clock = ExtractionTimings()
    timing_token = _TIMINGS.set(clock)
    strict_token = _STRICT.set(bool(strict))
    document = None
    try:
        _prepare_ocr_apartment()
        document = open_document(path, password or "")
        table = ExtractedTable(
            [str(name) for name in columns],
            [list(row) for row in rows],
            None,
            [int(page) for page in sources],
        )
        filled = recover_arabic_columns(document, table, None, only_page=int(page_number), parallel=False)
        indexes = [index for index, page in enumerate(sources) if int(page) == int(page_number)]
        return {
            "page": int(page_number),
            "filled": int(filled),
            "rows": {index: list(table.rows[index]) for index in indexes if index < len(table.rows)},
            "baselines": dict(table.baseline_at),
            "spans": {key: dict(value) for key, value in table.column_spans.items()},
            "stages": clock.rows(),
            "error": "",
        }
    except Exception as exc:
        return {
            "page": int(page_number),
            "filled": 0,
            "rows": {},
            "baselines": {},
            "spans": {},
            "stages": clock.rows(),
            "error": str(exc),
        }
    finally:
        if document is not None:
            close_document(document)
        _STRICT.reset(strict_token)
        _TIMINGS.reset(timing_token)


def _map_in_batches(
    specs: list,
    fn: Callable,
    progress: Callable[[str], None] | None,
    total: int,
    done_start: int,
    message: str,
    on_item: Callable[[dict], None] | None = None,
) -> list[dict]:
    """Run page jobs a few at a time. The jobs themselves are unchanged."""
    if not specs:
        return []
    try:
        from engine.resources import batch_size, worker_count

        workers = worker_count(len(specs))
        limit = max(1, min(batch_size(), len(specs)))
    except Exception:
        workers = max(1, min(3, len(specs), os.cpu_count() or 2))
        limit = workers
    results: list[dict] = []
    started = time.perf_counter()
    done = done_start
    with ProcessPoolExecutor(max_workers=workers) as pool:
        inflight: dict = {}
        queue = list(specs)
        while queue or inflight:
            while queue and len(inflight) < limit:
                spec = queue.pop(0)
                inflight[pool.submit(fn, spec)] = spec
            if not inflight:
                break
            future = next(as_completed(list(inflight)))
            inflight.pop(future, None)
            item = future.result()
            results.append(item)
            if on_item is not None:
                on_item(item)
            done += 1
            if progress is not None and total:
                elapsed = time.perf_counter() - started
                finished = max(done - done_start, 1)
                left = total - done
                eta = int(round(elapsed / finished * left)) if left else 0
                progress(message.format(done=done, total=total, eta=eta))
    return results


def _conversion_stamp(pdf_path: Path, password: str, strict: bool) -> str:
    digest = hashlib.sha256()
    try:
        with Path(pdf_path).open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError:
        return ""
    extra = hashlib.sha256(
        f"{_CACHE_VERSION}|{int(bool(strict))}|{password or ''}".encode("utf-8")
    ).hexdigest()[:12]
    return f"{digest.hexdigest()}-{extra}"


def _save_partial(index: int, payload: dict) -> None:
    stamp = _CHECKPOINT.get()
    if not stamp:
        return
    try:
        from engine.checkpoint import save_page

        save_page(stamp, index, payload)
    except Exception:
        return


def _load_partial(index: int) -> dict | None:
    stamp = _CHECKPOINT.get()
    if not stamp:
        return None
    try:
        from engine.checkpoint import load_page

        return load_page(stamp, index)
    except Exception:
        return None


def _clear_partial(stamp: str) -> None:
    if not stamp:
        return
    try:
        from engine.checkpoint import clear

        clear(stamp)
    except Exception:
        return


def _box_view(box: dict) -> dict:
    return {
        "language": box.get("language"),
        "text_garbage": box.get("text_garbage"),
        "columns_unrecognised": box.get("columns_unrecognised"),
        "detected_pages": dict(box.get("detected_pages") or {}),
        "statements": list(box.get("statements") or []),
    }


def _box_delta(before: dict, after: dict) -> dict:
    delta: dict = {}
    for key in ("language", "text_garbage", "columns_unrecognised"):
        if after.get(key) != before.get(key) and after.get(key) is not None:
            delta[key] = after.get(key)
    detected = {
        key: value
        for key, value in (after.get("detected_pages") or {}).items()
        if (before.get("detected_pages") or {}).get(key) != value
    }
    if detected:
        delta["detected_pages"] = detected
    added = [line for line in (after.get("statements") or []) if line not in (before.get("statements") or [])]
    if added:
        delta["statements"] = added
    return delta


def _apply_box_delta(box: dict, delta: dict) -> None:
    for key in ("language", "text_garbage", "columns_unrecognised"):
        if delta.get(key) is not None:
            box[key] = delta[key]
    if delta.get("detected_pages"):
        box.setdefault("detected_pages", {}).update(delta["detected_pages"])
    for line in delta.get("statements") or []:
        box.setdefault("statements", []).append(line)


def _page_record(index: int, tables: list[ExtractedTable], notes: list[str], box_delta: dict) -> dict:
    return {
        "index": int(index),
        "tables": [_table_payload(table) for table in tables],
        "notes": list(notes),
        "box": box_delta,
        "error": "",
    }


def _note_benchmark(document: fitz.Document) -> None:
    clock = _TIMINGS.get()
    if clock is None:
        return
    waited = 0.0
    for item in clock.events:
        if item.get("name") == "pages_elapsed":
            waited = float(item.get("seconds") or 0)
    if waited <= 0:
        return
    try:
        from engine.benchmark import record_run

        record_run(int(document.page_count), waited)
    except Exception:
        return


def _recover_arabic_parallel(
    document: fitz.Document,
    table: ExtractedTable,
    by_page: dict[int, list[int]],
    progress: Callable[[str], None] | None,
) -> int:
    pages = sorted(page for page in by_page if 1 <= page <= document.page_count)
    specs = [
        (
            str(document.name),
            page,
            "",
            bool(_STRICT.get()),
            list(table.columns),
            [list(row) for row in table.rows],
            list(table.sources or []),
        )
        for page in pages
    ]
    results = _map_in_batches(
        specs,
        _arabic_page_job,
        progress,
        len(specs),
        0,
        "Reading Arabic columns, page {done} of {total}... ETA {eta}s",
    )
    if any(item.get("error") for item in results):
        raise RuntimeError("arabic page failed")
    clock = _TIMINGS.get()
    filled = 0
    for item in results:
        filled += int(item.get("filled") or 0)
        for index, row in (item.get("rows") or {}).items():
            table.rows[int(index)] = list(row)
        table.baseline_at.update(item.get("baselines") or {})
        for key, spans in (item.get("spans") or {}).items():
            table.column_spans.setdefault(str(key), {}).update(spans)
        if clock is not None:
            for stage in item.get("stages") or []:
                clock.add(str(stage.get("name") or "arabic"), float(stage.get("seconds") or 0))
    return filled


def recover_arabic_columns(
    document: fitz.Document,
    table: ExtractedTable,
    progress: Callable[[str], None] | None = None,
    only_page: int | None = None,
    parallel: bool = True,
) -> int:
    """
    Read الصنع and اللون again from each row's band.

    The page is rendered once to place every plate. Each of those two columns
    is then read with Windows Arabic OCR, and the words are cut by the same
    row bands. The crop is the raw reading. A later pass writes an official
    name only when that reading wins by a margin. Plates, years, and dates
    are not touched, and no row is added or removed.
    """
    if not _sheet_is_rtl(table):
        return 0
    fields = _vehicle_field_at(table.columns)
    make_at = fields.get("make")
    color_at = fields.get("color")
    plate_at = fields.get("plate")
    if plate_at is None or (make_at is None and color_at is None):
        return 0
    if not table.sources or len(table.sources) != len(table.rows):
        return 0
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    by_page: dict[int, list[int]] = {}
    for index, page_number in enumerate(table.sources):
        by_page.setdefault(int(page_number), []).append(index)
    filled = 0
    targets = [("make", make_at), ("color", color_at)]
    if (
        parallel
        and only_page is None
        and len(by_page) > 1
        and os.environ.get("PDF_TOOL_SEQUENTIAL") != "1"
        and getattr(document, "name", "")
    ):
        try:
            return _recover_arabic_parallel(document, table, by_page, progress)
        except Exception:
            pass
    for page_number, indexes in by_page.items():
        if only_page is not None and int(page_number) != int(only_page):
            continue
        if page_number < 1 or page_number > document.page_count:
            continue
        pending = []
        for index in indexes:
            row = table.rows[index]
            if _nondata_kind(row, table.columns):
                continue
            plate = _plate_token([row[plate_at] if plate_at < len(row) else ""])
            if plate:
                pending.append(index)
        if not pending:
            continue
        page = document.load_page(page_number - 1)
        if progress is not None:
            progress(f"Reading Arabic columns on page {page_number}...")
        arabic_started = time.perf_counter()
        try:
            anchors = _page_anchor_words(page)
        except Exception:
            _remember_time(f"page_{page_number}_arabic_columns", time.perf_counter() - arabic_started)
            continue
        boxes = _arabic_column_boxes(page, anchors)
        if not boxes.get("make") and not boxes.get("color"):
            _remember_time(f"page_{page_number}_arabic_columns", time.perf_counter() - arabic_started)
            continue
        footer_y = _footer_rule_y(page)
        plate_ys = _collapse_baselines([word.center_y for word in anchors if _plate_token([word.text])])
        height = float(page.rect.height)
        column_words: dict[str, list[Word]] = {}
        for field, _column in targets:
            box = boxes.get(field)
            if not box:
                continue
            x0, x1 = box
            pad = 2.0
            clip = fitz.Rect(
                max(0.0, x0 - pad),
                0.0,
                min(float(page.rect.width), x1 + pad),
                height,
            )
            column_words[field] = _ocr_clip_words(page, clip, "ar-SA")
            table.column_spans.setdefault(str(page_number), {})[field] = (x0, x1)
        line_offset = {
            field: _text_line_offset(words, plate_ys)
            for field, words in column_words.items()
        }
        located: dict[int, float] = {}
        for index in pending:
            row = table.rows[index]
            plate = _plate_token([row[plate_at] if plate_at < len(row) else ""])
            y_value = _locate_plate_y(anchors, plate) if plate else None
            if y_value is None and plate:
                stored = table.baseline_at.get(f"{page_number}|{plate}")
                y_value = float(stored) if stored else None
            if y_value is not None:
                located[index] = float(y_value)
        for index in pending:
            if index in located:
                continue
            earlier = [item for item in located if item < index]
            later = [item for item in located if item > index]
            if not earlier or not later:
                continue
            before, after = max(earlier), min(later)
            span = after - before
            located[index] = located[before] + (located[after] - located[before]) * ((index - before) / span)
        for index in pending:
            row = table.rows[index]
            plate = _plate_token([row[plate_at] if plate_at < len(row) else ""])
            if not plate or index not in located:
                continue
            y_value = located[index]
            table.baseline_at[f"{page_number}|{plate}"] = float(y_value)
            top, bottom = _row_band_limits(y_value, plate_ys or [y_value], footer_y, height)
            for field, column in targets:
                if column is None or field not in boxes:
                    continue
                while len(row) <= column:
                    row.append("")
                current = str(row[column]).strip()
                x0, x1 = boxes[field]
                shift = line_offset.get(field, 3.5)
                raw = _arabic_band_text(
                    column_words.get(field) or [], x0, x1, top, bottom, y_value, plate_ys, shift
                )
                if not raw:
                    raw = _arabic_band_text(anchors, x0, x1, top, bottom, y_value, plate_ys, shift)
                value = _keep_arabic_value(field, raw)
                incoming = (_accept_make(value) if field == "make" else _accept_color(value)) if value else ""
                held = (_accept_make(current) if field == "make" else _accept_color(current)) if current else ""
                # A weak crop must not knock out an official name. جي إم سي stays
                # unless the crop itself names a different official make.
                if held and not incoming:
                    continue
                if raw:
                    _remember_raw(table, page_number, plate, field, raw)
                if not value or value == current:
                    continue
                _append_change(row, notes_at, current, value)
                row[column] = value
                filled += 1
        _store_visual_baselines(table, page_number, page, anchors)
        _remember_time(f"page_{page_number}_arabic_columns", time.perf_counter() - arabic_started)
    return filled


def _plate_column_centers(page: fitz.Page, anchors: list[Word]) -> dict[str, float]:
    """Read the plate column by itself.

    A full-page OCR box on this register can sit on the next vehicle's digits.
    The same number, read from a clip of the plate column, stays on its own ink.
    """
    plates = [word for word in anchors if _plate_token([word.text])]
    if len(plates) < 4:
        return {}
    x0 = max(0.0, min(word.x0 for word in plates) - 4.0)
    x1 = min(float(page.rect.width), max(word.x1 for word in plates) + 36.0)
    if x1 - x0 < 12.0:
        return {}
    clip = fitz.Rect(x0, 0.0, x1, float(page.rect.height))
    found: dict[str, list[float]] = {}
    for language in ("en-US", "ar-SA"):
        for word in _ocr_clip_words(page, clip, language):
            token = _plate_token([word.text])
            if token:
                found.setdefault(token, []).append(float(word.center_y))
        if len(found) >= max(4, len(plates) // 2):
            break
    return {token: sorted(values)[len(values) // 2] for token, values in found.items()}


def _visual_plate_ys(page: fitz.Page, anchors: list[Word]) -> tuple[dict[str, float], float]:
    """Plate to the y of its own digits, plus the page's uniform correction.

    The correction is kept only when every matched plate moves by the same
    amount. A page whose boxes already sit on the digits gets a zero move.
    """
    full: dict[str, float] = {}
    for word in anchors:
        token = _plate_token([word.text])
        if token and token not in full:
            full[token] = float(word.center_y)
    column = _plate_column_centers(page, anchors)
    deltas = [column[token] - full[token] for token in column if token in full]
    if len(deltas) < 4:
        return {}, 0.0
    deltas.sort()
    if deltas[-1] - deltas[0] > 3.0:
        return {}, 0.0
    offset = _median(deltas)
    # A column hit is kept only when it lands on the same row as the full-page
    # plate. A stray reading in the header must not move that vehicle.
    agreed = {
        token: column[token]
        for token in column
        if token in full and abs(column[token] - (full[token] + offset)) <= 4.0
    }
    if len(agreed) < 4:
        return {}, 0.0
    return agreed, float(offset)


def _store_visual_baselines(table: ExtractedTable, page_number: int, page: fitz.Page, anchors: list[Word]) -> None:
    """Replace this page's stored row positions with the digit positions.

    The Arabic read has already used the anchor grid. Only the position kept
    for the review crop changes, and a plate that was placed by interpolation
    moves by the same amount as the plates that were read.
    """
    visual, offset = _visual_plate_ys(page, anchors)
    if not visual and not offset:
        return
    prefix = f"{int(page_number)}|"
    for key, value in list(table.baseline_at.items()):
        if not str(key).startswith(prefix):
            continue
        plate = str(key).split("|", 1)[1]
        if plate in visual:
            table.baseline_at[key] = float(visual[plate])
        elif offset:
            table.baseline_at[key] = float(value) + float(offset)


def _weak_cell(field: str, text: str) -> bool:
    """An official name is already strong. Only a low score is worth another read."""
    return score_cell(field, str(text or "").strip()) < 0.8


def _reading_from_words(field: str, words: list[Word], plate: str) -> str:
    text = " ".join(word.text.strip() for word in words if word.text and word.text.strip())
    if not text:
        return ""
    if field == "date":
        found = _dates_in_text(text, plate)
        return found or (_normalize_date_cell(text) or "")
    if field == "year":
        match = re.search(r"(?:19|20)\d{2}", text)
        return match.group(0) if match else ""
    if field in {"make", "color"}:
        kept = _keep_arabic_value(field, text)
        official = _accept_make(kept) if field == "make" else _accept_color(kept)
        return official or kept
    return text.strip()


def _write_confident(
    row: list[str],
    index: int,
    field: str,
    candidate: str,
    table: ExtractedTable,
    page: int,
    plate: str,
) -> bool:
    """
    Keep a clip only when it is safe to write.

    Two readings that match are handled by the caller and marked agreed.
    One clip may normalize a slip (توبو → تويوتا). It does not write an
    official name that is already the catalog string, and it never writes a date.
    """
    while len(row) <= index:
        row.append("")
    current = str(row[index]).strip()
    text = str(candidate or "").strip()
    if not text or text == current:
        return False
    if field == "date":
        iso = text if _DATE_VALUE.fullmatch(text) else ""
        if iso:
            _add_reasoned_suggestion(table, page, plate, field, iso, "قراءة المقطع")
            key = _source_key(page, plate, field)
            if table.provenance.get(key) not in {"agreed", "manual"}:
                _remember_source(table, page, plate, field, "suggested", table.suggestions.get(key, ""))
        return False
    if field == "year":
        if not _YEAR_VALUE.fullmatch(text):
            return False
        row[index] = text
        _remember_raw(table, page, plate, field, text)
        _remember_source(table, page, plate, field, "one-engine", "")
        return True
    if field not in {"make", "color"}:
        return False
    if field == "make" and _short_make_fragment(text):
        return False
    official = _MANUFACTURERS if field == "make" else _OFFICIAL_COLORS
    accepted = _accept_make(text) if field == "make" else _accept_color(text)
    if accepted and _raw_is_slip(field, text, accepted):
        row[index] = accepted
        _remember_raw(table, page, plate, field, text)
        _remember_source(table, page, plate, field, "normalized", f"{accepted} — تطبيع: {text}")
        return True
    if text in official:
        _add_reasoned_suggestion(table, page, plate, field, text, "قراءة المقطع")
        key = _source_key(page, plate, field)
        if table.provenance.get(key) not in {"agreed", "manual", "normalized"}:
            _remember_source(table, page, plate, field, "suggested", table.suggestions.get(key, ""))
        return False
    if not current:
        kept = _keep_arabic_value(field, text)
        if kept and kept not in official:
            row[index] = kept
            _remember_raw(table, page, plate, field, text)
            _remember_source(table, page, plate, field, "one-engine", "")
            return True
    return False


def _agree_clip(
    row: list[str],
    index: int,
    field: str,
    left: str,
    right: str,
    table: ExtractedTable,
    page: int,
    plate: str,
) -> bool:
    """Write a date or year when both clips read it. A make or colour is not filled here."""
    if field in {"make", "color"}:
        return False
    first = str(left or "").strip()
    second = str(right or "").strip()
    if not first or first != second or score_cell(field, first) < 0.8:
        return False
    while len(row) <= index:
        row.append("")
    row[index] = first
    _remember_raw(table, page, plate, field, first)
    _remember_source(table, page, plate, field, "agreed")
    return True


def _cell_clip(page: fitz.Page, y_value: float, box: tuple[float, float] | None, field: str) -> fitz.Rect:
    width = float(page.rect.width)
    height = float(page.rect.height)
    if box:
        left = max(0.0, float(box[0]) - 3.0)
        right = min(width, float(box[1]) + 3.0)
    else:
        left, right = 0.0, width
    if right - left < 8:
        left, right = 0.0, width
    above = 22.0 if field == "date" else 12.0
    top = max(0.0, y_value - above)
    bottom = min(height, y_value + 8.0)
    return fitz.Rect(left, top, right, bottom)


def _tesseract_clip_text(page: fitz.Page, clip: fitz.Rect, language: str, field: str, plate: str) -> str:
    if clip.width < 4 or clip.height < 4:
        return ""
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=clip, alpha=False)
        words = _tesseract_words(
            pixmap.tobytes("png"),
            fitz.Rect(0, 0, clip.width, clip.height),
            float(pixmap.width),
            float(pixmap.height),
            "ara+eng" if language == "ar-SA" else "eng",
        )
    except Exception:
        return ""
    return _reading_from_words(field, words, plate)


def _ladder_clip_text(page: fitz.Page, clip: fitz.Rect, language: str, field: str, plate: str, mode: str) -> str:
    if clip.width < 4 or clip.height < 4:
        return ""
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=clip, alpha=False)
        png = pixmap.tobytes("png")
        width = float(pixmap.width)
        height = float(pixmap.height)
        del pixmap
        image = _png_variant(png, mode)
        words = _ocr_words(image, language, fitz.Rect(clip), width, height)
    except Exception:
        return ""
    return _reading_from_words(field, words, plate)


def recover_sparse_rows(
    document: fitz.Document,
    table: ExtractedTable,
    progress: Callable[[str], None] | None = None,
) -> int:
    """
    Re-read a weak make, colour, year, or date from that cell only.

    An official name is left alone. A cell that still scores below 0.8 gets
    one Windows crop. Tesseract, then contrast, binary, and deskew, run on
    that same crop only when the first crop is still weak.
    """
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    if plate_at is None or not table.sources or len(table.sources) != len(table.rows):
        return 0
    targets = [name for name in ("make", "color", "year", "date") if name in fields]
    if not targets:
        return 0
    pending: dict[int, list[int]] = {}
    for index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        plate = _plate_token([row[plate_at] if plate_at < len(row) else ""])
        if not plate:
            continue
        weak = any(
            _weak_cell(field, row[fields[field]] if fields[field] < len(row) else "")
            for field in targets
        )
        if weak:
            pending.setdefault(int(table.sources[index]), []).append(index)
    if not pending:
        return 0
    language = "ar-SA" if _sheet_is_rtl(table) else "en-US"
    filled = 0
    anchors_by_page: dict[int, list[Word]] = {}
    for page_number, indexes in pending.items():
        if page_number < 1 or page_number > document.page_count:
            continue
        page = document.load_page(page_number - 1)
        try:
            if len((page.get_text("text") or "").strip()) >= 20:
                continue
        except Exception:
            continue
        if progress is not None:
            progress(f"Retrying weak cells on page {page_number}...")
        spans = (table.column_spans or {}).get(str(page_number), {})
        for index in indexes:
            row = table.rows[index]
            plate = _plate_token([row[plate_at] if plate_at < len(row) else ""])
            if not plate:
                continue
            stored = table.baseline_at.get(f"{page_number}|{plate}")
            y_value = float(stored) if stored else None
            if y_value is None:
                if page_number not in anchors_by_page:
                    started = time.perf_counter()
                    try:
                        anchors_by_page[page_number] = _page_anchor_words(page)
                    except Exception:
                        anchors_by_page[page_number] = []
                    _remember_time(f"page_{page_number}_locate", time.perf_counter() - started)
                located = _locate_plate_y(anchors_by_page[page_number], plate)
                if located is None:
                    continue
                y_value = float(located)
                table.baseline_at[f"{page_number}|{plate}"] = y_value
            for field in targets:
                column = fields[field]
                current = row[column] if column < len(row) else ""
                if not _weak_cell(field, str(current or "")):
                    continue
                box = spans.get(field)
                pair = box if isinstance(box, (tuple, list)) and len(box) == 2 else None
                clip = _cell_clip(page, y_value, pair, field)
                label = f"page_{page_number}_{plate}"
                with _stage(f"{label}_retry_{field}"):
                    windows = _ocr_clip_words(page, clip, language)
                windows_text = _reading_from_words(field, windows, plate)
                windows_raw = " ".join(word.text.strip() for word in windows if word.text and word.text.strip())
                with _stage(f"{label}_second_{field}"):
                    second = _tesseract_clip_text(page, clip, language, field, plate)
                if _agree_clip(
                    row,
                    column,
                    field,
                    windows_text,
                    second,
                    table,
                    page_number,
                    plate,
                ):
                    filled += 1
                    continue
                if _write_confident(row, column, field, windows_raw or windows_text, table, page_number, plate):
                    filled += 1
                    continue
                if _write_confident(row, column, field, second, table, page_number, plate):
                    filled += 1
                    continue
                held = row[column] if column < len(row) else ""
                if field != "date" or score_cell("date", str(held or "")) >= 0.5:
                    continue
                for mode in ("contrast", "binary", "deskew"):
                    with _stage(f"{label}_ladder_{mode}"):
                        ladder = _ladder_clip_text(page, clip, language, field, plate, mode)
                    if _write_confident(row, column, field, ladder, table, page_number, plate):
                        filled += 1
                        break
    return filled


def suggest_cell(
    field: str,
    value: str,
    columns: list[str] | None = None,
    row: list[str] | None = None,
    neighbour: str = "",
) -> list[str]:
    """Up to three values for the review queue. The current text is not repeated."""
    del columns, row
    ideas: list[str] = []
    current = str(value or "").strip()
    if neighbour and neighbour not in ideas:
        ideas.append(neighbour)
    learned = _stored_value((load_corrections().get(field) or {}).get(current or "فراغ", ""))
    if learned and learned not in ideas:
        ideas.append(learned)
    if field == "make":
        hit = _unique_official_make(current) if current else ""
        if hit and hit not in ideas and hit != neighbour:
            ideas.append(hit)
        elif hit and hit not in ideas:
            ideas.append(hit)
        for name in _MANUFACTURERS:
            if len(ideas) >= 3:
                break
            if name not in ideas:
                ideas.append(name)
    elif field == "color":
        hit = _unique_official_color(current) if current else ""
        if hit and hit not in ideas:
            ideas.append(hit)
        for name in _OFFICIAL_COLORS:
            if len(ideas) >= 3:
                break
            if name not in ideas:
                ideas.append(name)
    return [item for item in ideas if item and item != current][:3]


def interpolate_row_y(
    baselines: dict[str, float],
    page: int,
    plates_on_page: list[str],
    plate: str,
) -> float:
    """Y of this plate, or a Y taken from the other rows on the same page.

    A missing baseline, such as plate 98689, uses the rows above and below.
    A page with no baselines at all still gets a band so the review card has a strip.
    """
    direct = float((baselines or {}).get(f"{int(page)}|{plate}") or 0)
    if direct > 0:
        return direct
    try:
        here = list(plates_on_page).index(plate)
    except ValueError:
        return 0.0
    known: list[tuple[int, float]] = []
    for index, other in enumerate(plates_on_page):
        y_value = float((baselines or {}).get(f"{int(page)}|{other}") or 0)
        if y_value > 0:
            known.append((index, y_value))
    earlier = [item for item in known if item[0] < here]
    later = [item for item in known if item[0] > here]
    if earlier and later:
        left, right = earlier[-1], later[0]
        span = right[0] - left[0]
        if span:
            return left[1] + (right[1] - left[1]) * ((here - left[0]) / span)
    gap = 18.0
    ordered = sorted(known)
    gaps = [
        (right[1] - left[1]) / (right[0] - left[0])
        for left, right in zip(ordered, ordered[1:])
        if right[0] > left[0] and right[1] != left[1]
    ]
    if gaps:
        gap = sorted(gaps)[len(gaps) // 2]
    if earlier:
        return earlier[-1][1] + gap * (here - earlier[-1][0])
    if later:
        return max(36.0, later[0][1] - gap * (later[0][0] - here))
    return 80.0 + here * gap


def crop_context_png(
    pdf_bytes: bytes,
    page_number: int,
    y_value: float,
    x0: float | None = None,
    x1: float | None = None,
    contrast: bool = False,
) -> bytes:
    """
    Full page width around one row, plus the rows above and below, at about 150 DPI.

    The review strip is padded and raised in contrast. A known column can still
    be outlined. This does not run OCR.
    """
    if not pdf_bytes or page_number < 1 or y_value <= 0:
        return b""
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        if page_number > document.page_count:
            return b""
        page = document.load_page(page_number - 1)
        pad = 36.0 if contrast else 22.0
        top = max(0.0, y_value - pad)
        bottom = min(float(page.rect.height), y_value + pad)
        if bottom - top < 24:
            bottom = min(float(page.rect.height), top + 24)
        clip = fitz.Rect(page.rect.x0, top, page.rect.x1, bottom)
        scale = 150.0 / 72.0
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
        if contrast or (x0 is not None and x1 is not None):
            import io

            from PIL import Image, ImageDraw, ImageOps

            image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            if contrast:
                image = ImageOps.autocontrast(image)
            if x0 is not None and x1 is not None:
                draw = ImageDraw.Draw(image)
                left = (max(clip.x0, x0) - clip.x0) * scale
                right = (min(clip.x1, x1) - clip.x0) * scale
                draw.rectangle([left, 1, right, image.height - 2], outline=(176, 42, 42), width=3)
            output = io.BytesIO()
            image.save(output, format="PNG")
            return output.getvalue()
        return pixmap.tobytes("png")
    except Exception:
        return b""
    finally:
        close_document(document)


def crop_row_png(pdf_bytes: bytes, page_number: int, plate: str) -> bytes:
    """A PNG of the row that contains this plate. Empty when the plate is not seen."""
    if not pdf_bytes or page_number < 1 or not plate:
        return b""
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        if page_number > document.page_count:
            return b""
        page = document.load_page(page_number - 1)
        image, rect, width, height = _scan_image(page)
        words = _ocr_words(image, "ar-SA", rect, width, height)
        if not words:
            words = _ocr_words(image, "en-US", rect, width, height)
        y_value = _y_of_token(words, plate)
        if y_value is None:
            return b""
        clip = fitz.Rect(page.rect.x0, max(0, y_value - 10), page.rect.x1, min(page.rect.y1, y_value + 10))
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2.4, 2.4), clip=clip, alpha=False)
        return pixmap.tobytes("png")
    except Exception:
        return b""
    finally:
        close_document(document)


def vision_read_crop(png: bytes, field: str, api_key: str) -> str:
    """
    Optional Claude reading of one red crop.

    No key, no image, or a network failure returns an empty string.
    The normal conversion does not call this.
    """
    if os.environ.get("PDF_TOOL_OFFLINE") == "1":
        return ""
    key = str(api_key or "").strip()
    if not key or not png:
        return ""
    import base64
    import urllib.error
    import urllib.request

    payload = {
        "model": "claude-sonnet-4-5",
        "max_tokens": 40,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": base64.b64encode(png).decode("ascii"),
                        },
                    },
                    {
                        "type": "text",
                        "text": (
                            f"This is one cropped table cell for the field {field}. "
                            "Reply with only the printed value, with no explanation."
                        ),
                    },
                ],
            }
        ],
    }
    request = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "content-type": "application/json",
            "x-api-key": key,
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = json.loads(response.read().decode("utf-8"))
    except (OSError, urllib.error.URLError, json.JSONDecodeError, TimeoutError):
        return ""
    parts = body.get("content") if isinstance(body, dict) else None
    if not isinstance(parts, list):
        return ""
    text = " ".join(str(part.get("text", "")) for part in parts if isinstance(part, dict)).strip()
    return text.splitlines()[0].strip() if text else ""


def partition_rows(
    columns: list[str],
    rows: list[list[str]],
    pages: list[int],
) -> tuple[list[list[str]], list[int], list[list[str]], list[int]]:
    """Vehicle rows, then header and fragment rows. Nothing is dropped."""
    data_rows: list[list[str]] = []
    data_pages: list[int] = []
    excluded_rows: list[list[str]] = []
    excluded_pages: list[int] = []
    for row, page in zip(rows, pages):
        if _nondata_kind(list(row), columns):
            excluded_rows.append(list(row))
            excluded_pages.append(int(page))
        else:
            data_rows.append(list(row))
            data_pages.append(int(page))
    return data_rows, data_pages, excluded_rows, excluded_pages


def _polish_register(table: ExtractedTable) -> None:
    """Group purpose, then confident plate, name, colour, and date repairs."""
    if "plate" not in _vehicle_field_at(table.columns):
        return
    _fill_purpose_down(table)
    _repair_plates(table)
    _retire_plate_outliers(table)
    _fill_blanks_from_line(table)
    _apply_learned(table)
    _repair_text_columns(table)
    _repair_group_dates(table)
    _infer_group_years(table)


def _printed_statements(lines: list[str]) -> tuple[int | None, dict[str, int]]:
    """Grand total and per-purpose subtotals printed outside the vehicle rows."""
    total: int | None = None
    groups: dict[str, int] = {}
    for line in lines:
        folded = _normalize(line).translate(_INDIC_DIGITS)
        compact = re.sub(r"\s+", "", folded)
        numbers = []
        for token in re.findall(r"\d{1,6}", folded):
            if _YEAR_VALUE.fullmatch(token) or (len(token) >= 4 and not token.startswith("0")):
                continue
            value = int(token)
            if 1 <= value <= 100000:
                numbers.append(value)
        purpose = _purpose_token([line])
        mentions_total = any(token in compact for token in ("اجمالي", "جمالي", "المجموع")) or (
            "عدد" in compact and ("مركبات" in compact or "سيارات" in compact)
        )
        if mentions_total and numbers:
            total = max([total] + numbers) if total is not None else max(numbers)
            continue
        if purpose and len(numbers) == 1:
            current = groups.get(purpose, 0)
            groups[purpose] = max(current, numbers[0])
    return total, groups


def validate_register(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    printed_total: int | None = None,
    printed_groups: dict[str, int] | None = None,
    detected_pages: dict | None = None,
) -> list[dict[str, str]]:
    """
    Checks for the preview and the validation sheet.

    A plated row is a vehicle. A ترويسة or فاصل row is not. Status is ok,
    fail, or skip when the PDF did not print the figure.
    """
    checks: list[dict[str, str]] = []
    fields = _vehicle_field_at(columns)
    plate_at = fields.get("plate")
    purpose_at = fields.get("purpose")
    plated_rows: list[tuple[list[str], int]] = []
    for row, page in zip(rows, page_numbers):
        if _nondata_kind(list(row), columns):
            continue
        token = ""
        if plate_at is not None and plate_at < len(row):
            token = _normalize_plate_cell(row[plate_at])
        if _PLATE_VALUE.fullmatch(token or "") and not _YEAR_VALUE.fullmatch(token or ""):
            plated_rows.append((list(row), int(page)))
    actual_total = len(plated_rows)
    if printed_total:
        checks.append(
            {
                "name": "المجموع",
                "status": "ok" if actual_total == int(printed_total) else "fail",
                "expected": str(int(printed_total)),
                "actual": str(actual_total),
                "detail": "المركبات ذات اللوحة مقابل العدد المطبوع",
            }
        )
    else:
        checks.append(
            {
                "name": "المجموع",
                "status": "skip",
                "expected": "",
                "actual": str(actual_total),
                "detail": "لا يوجد مجموع مطبوع",
            }
        )
    counted: dict[str, int] = {}
    for row, _page in plated_rows:
        purpose = ""
        if purpose_at is not None and purpose_at < len(row):
            purpose = _purpose_token([row[purpose_at]]) or str(row[purpose_at]).strip()
        if purpose:
            counted[purpose] = counted.get(purpose, 0) + 1
    for purpose, expected in (printed_groups or {}).items():
        actual = counted.get(purpose, 0)
        checks.append(
            {
                "name": purpose,
                "status": "ok" if actual == int(expected) else "fail",
                "expected": str(int(expected)),
                "actual": str(actual),
                "detail": "عدد المجموعة مقابل المجموع المطبوع",
            }
        )
    by_page: dict[int, int] = {}
    for _row, page in plated_rows:
        by_page[page] = by_page.get(page, 0) + 1
    detected = {int(page): int(count) for page, count in (detected_pages or {}).items()}
    pages = sorted(set(by_page) | set(detected))
    for page in pages:
        actual = by_page.get(page, 0)
        if page not in detected:
            checks.append(
                {
                    "name": f"صفحة {page}",
                    "status": "skip",
                    "expected": "",
                    "actual": str(actual),
                    "detail": "لم يُحسب عدد الصفوف على الصفحة",
                }
            )
            continue
        expected = detected[page]
        if actual == expected:
            status, detail = "ok", "المركبات المستخرجة مقابل اللوحات المقروءة على الصفحة"
        elif actual > expected:
            # The kept rows came from a later read. The first count was short.
            status, detail = "weak", f"المرور الأول وجد {expected}، والباقي استُعيد بإعادة القراءة"
        else:
            status, detail = "fail", f"المرور الأول وجد {expected}، والناتج {actual}"
        checks.append(
            {
                "name": f"صفحة {page}",
                "status": status,
                "expected": str(expected),
                "actual": str(actual),
                "detail": detail,
            }
        )
    return checks


def _confidence_band(score: float) -> str:
    """green, amber, or red for one cell score."""
    if score >= 0.8:
        return "green"
    if score >= 0.5:
        return "amber"
    return "red"


def score_cell(field: str, value: str) -> float:
    """
    How sure this one cell is.

    An empty cell is 0. An official make or colour, a digit plate, an ISO
    date, or a four-digit year scores high. A leftover OCR string scores low.
    """
    text = str(value or "").strip()
    if field == "notes" or field == "page":
        return 1.0
    if not text:
        return 0.0
    if field == "plate":
        token = _normalize_plate_cell(text)
        if _PLATE_VALUE.fullmatch(token or "") and not _YEAR_VALUE.fullmatch(token or ""):
            return 0.95
        return 0.35
    if field == "date":
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
            return 0.92
        if re.fullmatch(r"\d{2}-\d{2}", text):
            return 0.45
        return 0.3
    if field == "year":
        return 0.9 if _YEAR_VALUE.fullmatch(text) else 0.3
    if field == "make":
        if text in _MANUFACTURERS:
            return 0.95
        if _match_manufacturer(text):
            return 0.7
        return 0.35
    if field == "color":
        if text in _OFFICIAL_COLORS:
            return 0.95
        if _match_official_color(text):
            return 0.7
        return 0.35
    if field == "purpose":
        if text in {canon for _raw, canon in _PURPOSE_CANON}:
            return 0.9
        return 0.4
    if field == "admin":
        if not text or re.fullmatch(r"\d{1,4}", text):
            return 0.85
        return 0.4
    return 0.6 if text else 0.0


def score_row(columns: list[str], row: list[str]) -> dict[str, float]:
    """One score per named field on this row."""
    fields = _vehicle_field_at(columns)
    scores: dict[str, float] = {}
    for field, index in fields.items():
        value = row[index] if index < len(row) else ""
        scores[field] = score_cell(field, value)
    return scores


def score_register(columns: list[str], rows: list[list[str]]) -> list[dict[str, float]]:
    """Stored beside the rows. The preview paints green, amber, and red from it."""
    return [score_row(columns, list(row)) for row in rows]


def sanitize_tables(tables: list[ExtractedTable]) -> list[ExtractedTable]:
    """The cleaning pass that runs before the workbook is written."""
    for table in tables:
        _sanitize_table(table)
        _apply_color_dictionary(table)
        _apply_manufacturer_dictionary(table)
        _forward_fill_attributes(table)
        _polish_register(table)
        annotate_vehicle_flags(table)
    return [table for table in tables if table.rows]


def _stated_vehicle_total(document: fitz.Document) -> int | None:
    """
    The official vehicle count printed on the last pages.

    A line such as 'عدد المركبات 564' or 'الإجمالي 564' is the target.
    A bare page number is not.
    """
    if document.page_count <= 0:
        return None
    start = max(0, document.page_count - 2)
    phrases = ("عددالمركبات", "عددالسيارات", "اجماليالمركبات", "المجموع", "الاجمالي")
    # get_text sometimes returns the Arabic in visual order. Accept that too.
    phrases = phrases + tuple(phrase[::-1] for phrase in phrases)
    found: int | None = None
    for index in range(start, document.page_count):
        try:
            page = document.load_page(index)
            raw = page.get_text("text") or ""
        except Exception:
            continue
        folded = _normalize(raw).translate(_INDIC_DIGITS)
        compact = re.sub(r"\s+", "", folded)
        for phrase in phrases:
            for match in re.finditer(re.escape(phrase) + r"(\d{2,6})", compact):
                value = int(match.group(1))
                if 10 <= value <= 100000:
                    found = value
            for match in re.finditer(r"(\d{2,6})" + re.escape(phrase), compact):
                value = int(match.group(1))
                if 10 <= value <= 100000:
                    found = value
    return found


def _column_index(table: ExtractedTable, *concepts: str) -> int | None:
    """The first column whose header names one of these fields."""
    wanted = set(concepts)
    for index, name in enumerate(table.columns):
        concept = _concept_for(name)
        if concept == "insurance":
            concept = "date"
        if concept in wanted or _logical_key(name) in wanted:
            return index
    return None


def _join_wrapped_cell(cell: str) -> str:
    """A line break inside one vehicle is a wrap, not a new row."""
    if "\n" not in cell:
        return cell
    return " ".join(part.strip() for part in cell.split("\n") if part.strip())


def _split_multiline_rows(table: ExtractedTable) -> None:
    """
    Two plate numbers in one cell are two vehicles.

    A wrapped colour or make stays on the same row. A line break is joined
    unless the plate cell itself holds two plates.
    """
    plate_at = _column_index(table, "plate", "رقماللوحه")
    rebuilt: list[list[str]] = []
    sources: list[int] = []
    source_rows = table.sources if table.sources and len(table.sources) == len(table.rows) else None
    for row_index, row in enumerate(table.rows):
        cells = [_join_wrapped_cell(cell) for cell in row]
        plate_parts: list[str] = []
        if plate_at is not None and plate_at < len(cells):
            plate_parts = re.findall(r"(?<!\d)(?:(?!19|20)\d{4,5})(?!\d)", cells[plate_at].translate(_INDIC_DIGITS))
        if len(plate_parts) <= 1:
            rebuilt.append(cells)
            if source_rows is not None:
                sources.append(source_rows[row_index])
            continue
        pieces = [[cell] for cell in cells]
        blocks = len(plate_parts)
        for block in range(blocks):
            made = []
            for part in pieces:
                made.append(part[block].strip() if block < len(part) else "")
            if plate_parts and plate_at is not None:
                while len(made) <= plate_at:
                    made.append("")
                if block < len(plate_parts):
                    made[plate_at] = plate_parts[block]
            if any(made):
                rebuilt.append(made)
                if source_rows is not None:
                    sources.append(source_rows[row_index])
    table.rows = rebuilt
    if source_rows is not None:
        table.sources = sources


def _plates_in(table: ExtractedTable) -> set[str]:
    plate_at = _column_index(table, "plate", "رقماللوحه")
    if plate_at is None:
        return set()
    found = set()
    for row in table.rows:
        if plate_at < len(row) and row[plate_at].strip():
            found.add(_normalize_plate_cell(row[plate_at]))
    return found


def _recover_wrapped_rows(document: fitz.Document, table: ExtractedTable) -> int:
    """Read each page again on a tighter line spacing and add vehicles that were missed."""
    known = _plates_in(table)
    added = 0
    width = len(table.columns)
    for index in range(document.page_count):
        try:
            page = document.load_page(index)
        except Exception:
            continue
        parsed = _vehicle_table_from_matrix(_vehicle_matrix_from_grid(page, row_tolerance=1.5), index)
        if parsed is None:
            continue
        _pass_word_coordinates(page, parsed)
        labels = [name for _key, name in _VEHICLE_COLUMNS]
        header_at = {name: column for column, name in enumerate(table.columns)}
        for row in parsed.rows:
            plate = row[5] if len(row) > 5 else ""
            token = _normalize_plate_cell(plate)
            if not token or token in known:
                continue
            known.add(token)
            placed = [""] * width
            for column, value in enumerate(row):
                if column >= len(labels):
                    break
                target = header_at.get(labels[column])
                if target is not None and target < width:
                    placed[target] = value
            table.rows.append(placed)
            if table.sources is not None:
                table.sources.append(index + 1)
            added += 1
    return added


def reconcile_register(
    document: fitz.Document,
    tables: list[ExtractedTable],
    progress: Callable[[str], None] | None = None,
) -> list[ExtractedTable]:
    """
    Compare the extracted vehicles with the total printed at the end of the PDF.

    When rows are missing, multi-line cells are split and the pages are read
    again. The warning stays even if the second pass cannot close the gap.
    """
    registers = [table for table in tables if _column_index(table, "plate", "رقماللوحه") is not None]
    if not registers:
        return tables

    def warn(message: str) -> None:
        print(message, file=sys.stderr)
        if progress is not None:
            progress(message)

    # The word scan and the page-local fill run even when the counts already agree.
    for table in registers:
        _rescan_blank_attributes(document, table)
        _forward_fill_attributes(table)
    target = _stated_vehicle_total(document)
    extracted = 0
    for table in registers:
        plate_at = _column_index(table, "plate", "رقماللوحه")
        extracted += sum(
            1
            for row in table.rows
            if not _nondata_kind(row, table.columns) and _plated_row(row, plate_at)
        )
    if target is not None and extracted != target:
        warn(
            f"Row count mismatch: the PDF states {target} vehicles and the sheet has {extracted}. "
            "Reading multi-line cells again."
        )
        for table in registers:
            _split_multiline_rows(table)
            _recover_wrapped_rows(document, table)
            _rescan_blank_attributes(document, table)
            _forward_fill_attributes(table)
        extracted = 0
        for table in registers:
            plate_at = _column_index(table, "plate", "رقماللوحه")
            extracted += sum(
                1
                for row in table.rows
                if not _nondata_kind(row, table.columns) and _plated_row(row, plate_at)
            )
        if extracted != target:
            warn(f"Row count is still {extracted} after the second pass. The stated total is {target}.")
        else:
            warn(f"Row count now matches the stated total of {target}.")
    blanks = sum(_blank_repeat_cells(table) for table in registers)
    if blanks:
        if strict_values():
            warn(f"{blanks} manufacturer, year, or colour cells are blank. They were left blank.")
        else:
            warn(f"{blanks} manufacturer, year, or colour cells are still blank after imputation.")
    return tables


def _slice_table(table: ExtractedTable, start: int, end: int) -> ExtractedTable:
    """One block of a long sheet. Columns and cell positions stay the same."""
    rows = [list(row) for row in table.rows[start:end]]
    sources = table.sources[start:end] if table.sources and len(table.sources) == len(table.rows) else None
    lines = table.lines[start:end] if table.lines and len(table.lines) == len(table.rows) else None
    piece = ExtractedTable(list(table.columns), rows, None, list(sources) if sources is not None else None)
    piece.lines = list(lines) if lines is not None else None
    piece.kind = table.kind
    return piece


def _summary_table(tables: list[ExtractedTable], kind: str) -> ExtractedTable:
    """A first sheet that lists each exported table."""
    arabic = any(_sheet_is_rtl(table) for table in tables)
    if arabic:
        columns = ["الورقة", "عدد الصفوف", "الأعمدة", "نوع المستند"]
    else:
        columns = ["Sheet", "Rows", "Columns", "Document"]
    label = _DOC_LABELS.get(kind, kind)
    rows = []
    for index, table in enumerate(tables, start=1):
        names = "، ".join(table.columns[:8]) if arabic else ", ".join(table.columns[:8])
        rows.append([str(index), str(len(table.rows)), names, label])
    summary = ExtractedTable(columns, rows)
    summary.kind = kind
    return summary


def _sheets_for_workbook(tables: list[ExtractedTable]) -> list[ExtractedTable]:
    """Split a very long table, and add a summary when the workbook has several sheets."""
    pieces: list[ExtractedTable] = []
    split = False
    for table in tables:
        if len(table.rows) <= _SHEET_ROW_LIMIT:
            pieces.append(table)
            continue
        split = True
        for start in range(0, len(table.rows), _SHEET_ROW_LIMIT):
            pieces.append(_slice_table(table, start, start + _SHEET_ROW_LIMIT))
    if len(pieces) > 1 or split:
        kind = next((table.kind for table in pieces if table.kind), "")
        return [_summary_table(pieces, kind)] + pieces
    return pieces


def _accept_page_tables(
    found: list[ExtractedTable],
    page_num: int,
    scans: list[ExtractedTable],
    grouped: dict[tuple[str, ...], ExtractedTable],
    order: list[tuple[str, ...]],
) -> None:
    """Finish one page before its rows are allowed to meet another page."""
    for table in found:
        _isolate_nondata_rows(table)
        _smooth_spellings(table)
        _clean_table(table)
        _validate_patterns(table)
        if not table.rows:
            continue
        table.sources = [page_num] * len(table.rows)
        if table.anchors:
            scans.append(table)
            continue
        match: tuple[str, ...] | None = None
        for key in order:
            current = grouped[key]
            if not table.anchors and _header_key(current) == _header_key(table):
                match = key
                break
        if match is None:
            key = _header_key(table)
            while key in grouped:
                key = key + (str(len(order)),)
            grouped[key] = table
            order.append(key)
            continue
        current = grouped[match]
        current.rows.extend(table.rows)
        if current.sources is None:
            current.sources = [0] * (len(current.rows) - len(table.rows))
        if table.sources is not None:
            current.sources.extend(table.sources)
        elif current.sources is not None:
            current.sources.extend([0] * len(table.rows))


def _table_payload(table: ExtractedTable) -> dict:
    return {
        "columns": list(table.columns),
        "rows": [list(row) for row in table.rows],
        "anchors": list(table.anchors) if table.anchors else None,
        "lines": list(table.lines) if table.lines else None,
        "kind": table.kind,
    }


def _table_from_payload(payload: dict) -> ExtractedTable:
    return ExtractedTable(
        list(payload.get("columns") or []),
        [list(row) for row in payload.get("rows") or []],
        payload.get("anchors"),
        None,
        payload.get("lines"),
        str(payload.get("kind") or ""),
    )


def _page_job(spec: tuple) -> dict:
    """Read one scanned page in its own process. The document is opened here."""
    path, page_index, password, strict, language, kind = spec
    clock = ExtractionTimings()
    timing_token = _TIMINGS.set(clock)
    strict_token = _STRICT.set(bool(strict))
    document = None
    try:
        _prepare_ocr_apartment()
        document = open_document(path, password or "")
        page = document.load_page(int(page_index))
        box: dict = {}
        if language:
            box["language"] = language
        notes: list[str] = []
        found = _tables_on_page(page, int(page_index), box, None, notes, kind or "register")
        return {
            "index": int(page_index),
            "tables": [_table_payload(table) for table in found],
            "notes": notes,
            "detected": dict(box.get("detected_pages") or {}),
            "statements": list(box.get("statements") or []),
            "stages": clock.rows(),
            "error": "",
        }
    except Exception as exc:
        return {
            "index": int(page_index),
            "tables": [],
            "notes": [],
            "detected": {},
            "statements": [],
            "stages": clock.rows(),
            "error": str(exc),
        }
    finally:
        if document is not None:
            close_document(document)
        _STRICT.reset(strict_token)
        _TIMINGS.reset(timing_token)


def _document_is_all_scans(document: fitz.Document) -> bool:
    if document.page_count < 2 or os.environ.get("PDF_TOOL_SEQUENTIAL") == "1":
        return False
    for index in range(document.page_count):
        try:
            if _text_char_count(document.load_page(index)) >= 20:
                return False
        except Exception:
            return False
    return True


def _read_pages_sequential(
    document: fitz.Document,
    pdf_path: Path,
    password: str,
    progress: Callable[[str], None] | None,
    kind: str,
    grouped: dict[tuple[str, ...], ExtractedTable],
    order: list[tuple[str, ...]],
    scans: list[ExtractedTable],
    language_box: dict,
) -> Exception | None:
    page_error: Exception | None = None
    started = time.perf_counter()
    completed = 0
    plumber = _open_pdfplumber(pdf_path, password)
    try:
        try:
            from engine.resources import batch_size

            step = batch_size()
        except Exception:
            step = 5
        for batch_start, batch_pages in iter_page_batches(document, step):
            for offset, page in enumerate(batch_pages):
                index = batch_start + offset
                page_num = index + 1
                elapsed = time.perf_counter() - started
                eta = int(round(elapsed / completed * (document.page_count - completed))) if completed else 0
                if progress is not None:
                    progress(f"Reading page {page_num} of {document.page_count}... ETA {eta}s")
                try:
                    saved = _load_partial(index)
                    if saved is not None:
                        notes = list(saved.get("notes") or [])
                        found = [_table_from_payload(item) for item in saved.get("tables") or []]
                        _apply_box_delta(language_box, saved.get("box") or {})
                    else:
                        notes = []
                        before = _box_view(language_box)
                        found = _tables_on_page(page, index, language_box, plumber, notes, kind)
                        _save_partial(
                            index,
                            _page_record(index, found, notes, _box_delta(before, _box_view(language_box))),
                        )
                    for message in notes:
                        if progress is not None:
                            progress(message)
                except ConversionError:
                    raise
                except Exception as exc:
                    page_error = page_error or exc
                    continue
                _accept_page_tables(found, page_num, scans, grouped, order)
                completed += 1
    finally:
        if plumber is not None:
            try:
                plumber.close()
            except Exception:
                pass
        _remember_time("pages_elapsed", time.perf_counter() - started)
    return page_error


def _read_scans_in_parallel(
    document: fitz.Document,
    pdf_path: Path,
    password: str,
    strict: bool,
    progress: Callable[[str], None] | None,
    kind: str,
    grouped: dict[tuple[str, ...], ExtractedTable],
    order: list[tuple[str, ...]],
    scans: list[ExtractedTable],
    language_box: dict,
) -> Exception | None:
    """Page one chooses the language. The other pages are read together."""
    started = time.perf_counter()
    total = int(document.page_count)
    page_error: Exception | None = None
    if progress is not None:
        progress(f"Reading page 1 of {total}... ETA 0s")
    saved_first = _load_partial(0)
    if saved_first is not None:
        notes = list(saved_first.get("notes") or [])
        found = [_table_from_payload(item) for item in saved_first.get("tables") or []]
        _apply_box_delta(language_box, saved_first.get("box") or {})
    else:
        notes = []
        before = _box_view(language_box)
        found = _tables_on_page(document.load_page(0), 0, language_box, None, notes, kind)
        _save_partial(0, _page_record(0, found, notes, _box_delta(before, _box_view(language_box))))
    for message in notes:
        if progress is not None:
            progress(message)
    _accept_page_tables(found, 1, scans, grouped, order)
    language = str(language_box.get("language") or "ar-SA")
    ready: list[dict] = []
    specs = []
    for index in range(1, total):
        saved = _load_partial(index)
        if saved is None:
            specs.append((str(pdf_path), index, password or "", bool(strict), language, kind))
            continue
        box = saved.get("box") or {}
        ready.append(
            {
                "index": index,
                "tables": saved.get("tables") or [],
                "notes": saved.get("notes") or [],
                "detected": box.get("detected_pages") or {},
                "statements": box.get("statements") or [],
                "stages": [],
                "error": "",
            }
        )
    if ready and progress is not None:
        progress(f"Resuming saved pages, {len(ready)} already read.")

    def _keep(item: dict) -> None:
        if item.get("error") and not item.get("tables"):
            return
        _save_partial(
            int(item.get("index") or 0),
            {
                "index": int(item.get("index") or 0),
                "tables": item.get("tables") or [],
                "notes": list(item.get("notes") or []),
                "box": {
                    "detected_pages": item.get("detected") or {},
                    "statements": list(item.get("statements") or []),
                },
                "error": "",
            },
        )

    fresh = _map_in_batches(
        specs,
        _page_job,
        progress,
        total,
        1 + len(ready),
        "Reading page {done} of {total}... ETA {eta}s",
        _keep,
    )
    payloads = ready + fresh
    clock = _TIMINGS.get()
    by_index: dict[int, list[ExtractedTable]] = {}
    for payload in payloads:
        if clock is not None:
            for stage in payload.get("stages") or []:
                clock.add(str(stage.get("name") or "page"), float(stage.get("seconds") or 0))
        index = int(payload.get("index") or 0)
        if payload.get("error") and not payload.get("tables"):
            page_error = page_error or RuntimeError(str(payload.get("error")))
            notes = []
            before = _box_view(language_box)
            try:
                by_index[index] = _tables_on_page(document.load_page(index), index, language_box, None, notes, kind)
                _save_partial(
                    index,
                    _page_record(index, by_index[index], notes, _box_delta(before, _box_view(language_box))),
                )
            except Exception as exc:
                page_error = page_error or exc
                by_index[index] = []
            continue
        detected = payload.get("detected") or {}
        language_box.setdefault("detected_pages", {}).update(detected)
        for line in payload.get("statements") or []:
            language_box.setdefault("statements", []).append(line)
        by_index[index] = [_table_from_payload(item) for item in payload.get("tables") or []]
    for index in sorted(by_index):
        _accept_page_tables(by_index[index], index + 1, scans, grouped, order)
    _remember_time("pages_elapsed", time.perf_counter() - started)
    return page_error


def extract_tables(
    pdf_path: Path,
    password: str = "",
    progress: Callable[[str], None] | None = None,
    strict: bool = False,
) -> list[ExtractedTable]:
    """
    Read one page at a time. Every scanned page is then listed, in order.

    A page is validated before the next page is read. Its rows are matched to
    the sheet by what the cells contain, so a page with a different column
    count is still included. The page number is the loop counter.
    """
    stamp = _conversion_stamp(Path(pdf_path), password, bool(strict))
    cached_tables = _load_cached_tables(Path(pdf_path), password, bool(strict))
    if cached_tables is not None:
        _clear_partial(stamp)
        return cached_tables
    try:
        opened = time.perf_counter()
        document = open_document(pdf_path, password)
        _remember_time("open_pdf", time.perf_counter() - opened)
    except (FileNotFoundError, RuntimeError) as exc:
        raise ConversionError("The PDF could not be opened. The file may be damaged.") from exc
    except PermissionError as exc:
        raise ConversionError("This PDF is password protected. Enter the password and try again.") from exc
    stamp_token = _CHECKPOINT.set(stamp)
    strict_token = _STRICT.set(bool(strict))
    try:
        if document.page_count <= 0:
            raise ConversionError("The PDF has no pages.")
        kind = detect_document_type(document)
        if progress is not None:
            progress(f"Document type: {_DOC_LABELS.get(kind, kind)}.")
        grouped: dict[tuple[str, ...], ExtractedTable] = {}
        order: list[tuple[str, ...]] = []
        scans: list[ExtractedTable] = []
        language_box: dict = {}
        page_error: Exception | None = None
        if _document_is_all_scans(document):
            try:
                page_error = _read_scans_in_parallel(
                    document,
                    Path(pdf_path),
                    password,
                    bool(strict),
                    progress,
                    kind,
                    grouped,
                    order,
                    scans,
                    language_box,
                )
            except Exception as exc:
                scans.clear()
                grouped.clear()
                order.clear()
                language_box.clear()
                page_error = exc
                page_error = _read_pages_sequential(
                    document, Path(pdf_path), password, progress, kind, grouped, order, scans, language_box
                ) or page_error
        else:
            page_error = _read_pages_sequential(
                document, Path(pdf_path), password, progress, kind, grouped, order, scans, language_box
            )
        # Every scanned page, in order, on one sheet.
        tables: list[ExtractedTable] = []
        stacked = _stack_scan_pages(scans)
        if stacked is not None:
            tables.append(stacked)
        elif scans:
            tables.extend(scans)
        tables.extend(grouped[key] for key in order)
        for table in tables:
            _drop_repeated_rows(table)
            _apply_bilingual_headers(table)
            # Pass 2: the sheet about to be written. Headers now name the plate
            # and the colour, so an empty plate is filled from the row text and
            # the colour is normalized again. The page counter is applied after this.
            _validate_patterns(table)
            _add_page_column(table)
            _apply_manufacturer_dictionary(table)
            table.kind = kind
        tables = sanitize_tables(tables)
        if kind == "register":
            tables = reconcile_register(document, tables, progress)
            tables = sanitize_tables(tables)
        for table in tables:
            recover_arabic_columns(document, table, progress)
            recover_sparse_rows(document, table, progress)
            _repair_text_columns(table)
            _repair_group_dates(table)
            _infer_group_years(table)
            annotate_vehicle_flags(table)
        if tables:
            tables[0].detected_pages = dict(language_box.get("detected_pages") or {})
            tables[0].statements = list(language_box.get("statements") or [])
        garbage = bool(language_box.get("text_garbage")) or any(
            _headers_look_garbage(list(table.columns)) for table in tables
        )
        forced = bool(language_box.get("columns_unrecognised"))
        if garbage or forced:
            _mark_ocr_attempted(tables)
        for table in tables:
            table.column_warning = build_column_warning(
                list(table.columns),
                garbage=garbage or _headers_look_garbage(list(table.columns)),
                kind=table.kind or kind,
                forced=forced,
            )
        if not tables and page_error is not None:
            raise ConversionError(f"The pages could not be read. {page_error}") from page_error
        _note_benchmark(document)
        if tables:
            _store_cached_tables(Path(pdf_path), password, bool(strict), tables)
        _clear_partial(stamp)
        return tables
    finally:
        _CHECKPOINT.reset(stamp_token)
        _STRICT.reset(strict_token)
        close_document(document)


def _group_rows(words: list[Word], tolerance: float = 7) -> list[list[Word]]:
    if not words:
        return []
    ordered = sorted(words, key=lambda word: (word.center_y, word.x0))
    rows: list[list[Word]] = [[ordered[0]]]
    current_y = ordered[0].center_y
    for word in ordered[1:]:
        if abs(word.center_y - current_y) <= tolerance:
            rows[-1].append(word)
            current_y = sum(item.center_y for item in rows[-1]) / len(rows[-1])
        else:
            rows.append([word])
            current_y = word.center_y
    return rows


def _sheet_is_rtl(table: ExtractedTable) -> bool:
    """Arabic invoices open from the right. English invoices stay left to right."""
    sample = " ".join(table.columns)
    for row in table.rows[:12]:
        sample += " " + " ".join(row)
    arabic = sum(1 for char in sample if _is_arabic_letter(char))
    latin = sum(1 for char in sample if char.isascii() and char.isalpha())
    return arabic > latin


def _sheet_name(index: int, used: set[str]) -> str:
    base = "Table" if index == 1 else f"Table {index}"
    name = base
    suffix = 2
    while name in used:
        name = f"{base} {suffix}"
        suffix += 1
    used.add(name)
    return name[:31]


def _column_is_rtl(table: ExtractedTable, index: int, sheet_rtl: bool) -> bool:
    """A column follows its own letters. A column of numbers follows the sheet."""
    parts = [table.columns[index]] if index < len(table.columns) else []
    for row in table.rows[:15]:
        if index < len(row) and row[index]:
            parts.append(row[index])
    arabic, latin = _letter_counts(" ".join(parts))
    if arabic == latin:
        return sheet_rtl
    return arabic > latin


def _set_sheet_direction(worksheet, rtl: bool) -> None:
    """
    Open an Arabic sheet from the right.

    The flag has to be the sheet view Excel actually loads. A frozen header
    is stored as the bottom-left pane, and Excel then throws that view away
    and opens the file from the left. On an Arabic sheet the frozen pane is
    the bottom-right one, because column A sits on the right.
    """
    view = worksheet.views.sheetView[0]
    view.rightToLeft = True if rtl else False
    if not rtl or view.pane is None:
        return
    view.pane.activePane = "bottomRight"
    for selection in view.selection:
        if selection.pane == "bottomLeft":
            selection.pane = "bottomRight"


# A leading left-to-right mark keeps 2026-04-09 in that order inside an RTL cell.
# A typed apostrophe does the opposite: Excel folds it into the middle and the
# date is drawn as 09 '04 -.
_DATE_LTR = "\u200e"
_NUMBER_CONCEPTS = {"plate", "year", "admin", "quantity", "price", "total"}
_DATE_CONCEPTS = {"date", "insurance"}
_MONEY_CONCEPTS = {"price", "total"}


def _export_column_kind(header: str) -> str:
    """Date, number, or text, from the header. A colour column stays text."""
    concept = _concept_for(header)
    if concept in _DATE_CONCEPTS:
        return "date"
    if concept in _NUMBER_CONCEPTS:
        return "number"
    return "text"


def _protect_excel_date(value) -> str | None:
    """
    ISO date stored as text.

    The @ format is Excel's apostrophe: the cell is text, so Excel does not
    parse the dashes and mirror them. The mark at the front keeps the year
    on the left of the month.
    """
    if value is None or isinstance(value, bool):
        return None
    if hasattr(value, "strftime"):
        text = value.strftime("%Y-%m-%d")
    else:
        text = _blank_text(value)
    if not text:
        return None
    for mark in ("\u200e", "\u200f", "\u2066", "\u2067", "\u2069", "'", "’"):
        text = text.replace(mark, "")
    normalized = _normalize_date_cell(text.strip())
    if _DATE_VALUE.fullmatch(normalized) or re.fullmatch(r"\d{2}-\d{2}", normalized or ""):
        return _DATE_LTR + normalized
    return None


def _coerce_excel_number(value):
    """An integer or a float. A plate with letters, and a blank, stay out."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value != value:
            return None
        return int(value) if value.is_integer() else value
    text = _blank_text(value)
    if not text:
        return None
    for mark in ("\u200e", "\u200f", "\u2066", "\u2067", "\u2069"):
        text = text.replace(mark, "")
    text = text.translate(_INDIC_DIGITS).replace(",", "").replace(" ", "")
    if re.fullmatch(r"-?\d+", text):
        return int(text)
    if re.fullmatch(r"-?\d+\.\d+", text):
        number = float(text)
        return int(number) if number.is_integer() else number
    return None


def _number_format_for(header: str, number) -> str:
    """Plates and years stay plain. Money keeps two decimals."""
    concept = _concept_for(header)
    if isinstance(number, float):
        if concept in _MONEY_CONCEPTS:
            return "#,##0.00"
        return "0.##"
    if concept in _MONEY_CONCEPTS:
        return "#,##0"
    return "0"


def _align_text_cell(cell, arabic_column: bool) -> None:
    """Arabic text sits on the right. A Latin column stays on the left."""
    if arabic_column:
        cell.alignment = Alignment(horizontal="right", vertical="center", readingOrder=2)
    else:
        cell.alignment = Alignment(horizontal="left", vertical="center", readingOrder=1)


_SOURCE_FILLS = {
    "agreed": PatternFill("solid", fgColor="C6EFCE"),
    "one-engine": PatternFill("solid", fgColor="F4B183"),
    "normalized": PatternFill("solid", fgColor="E2D5F1"),
    "suggested": PatternFill("solid", fgColor="FFF2CC"),
    "manual": PatternFill("solid", fgColor="D6EAF8"),
}


def _provenance_fill(table: ExtractedTable, row_index: int, column_index: int):
    """Colour of one cell from its provenance. Empty when this cell was not tagged."""
    if row_index < 0 or row_index >= len(table.rows) or not table.provenance:
        return None
    field_name = _vehicle_field(table.columns[column_index]) if column_index < len(table.columns) else ""
    if field_name not in {"make", "color", "year", "date", "admin"}:
        return None
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    row = table.rows[row_index]
    plate = _plate_token([row[plate_at]]) if plate_at is not None and plate_at < len(row) else ""
    page = int(table.sources[row_index]) if table.sources and row_index < len(table.sources) else 0
    source = table.provenance.get(_source_key(page, plate, field_name), "")
    if not source:
        return None
    if source == "one-engine" and column_index < len(row) and str(row[column_index]).strip():
        return PatternFill("solid", fgColor="F4B183")
    return _SOURCE_FILLS.get(source)


def _provenance_sheet_rows(table: ExtractedTable) -> list[list[str]]:
    """One row per tagged cell. القيمة is the raw OCR reading, not the Table cell."""
    labels = {"make": "الصنع", "color": "اللون", "year": "السنة", "date": "انتهاء التأمين", "admin": "رقم إداري"}
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    rows: list[list[str]] = []
    sources = list(table.sources or [])
    for index, row in enumerate(table.rows):
        plate = _plate_token([row[plate_at]]) if plate_at is not None and plate_at < len(row) else ""
        if not plate:
            continue
        page = int(sources[index]) if index < len(sources) else 0
        for field_name, label in labels.items():
            key = _source_key(page, plate, field_name)
            source = table.provenance.get(key, "")
            suggestion = table.suggestions.get(key, "")
            raw = table.raw_readings.get(key, "")
            if not source and not suggestion and not raw:
                continue
            rows.append([str(page), plate, label, raw, source, suggestion])
    return rows


def _exception_sheet_rows(table: ExtractedTable) -> list[list[str]]:
    """Rows whose colour or make is not the page's common value, or whose date is rare."""
    fields = _vehicle_field_at(table.columns)
    plate_at = fields.get("plate")
    if plate_at is None:
        return []
    labels = {"color": "اللون", "make": "الصنع", "date": "انتهاء التأمين"}
    sources = list(table.sources or [])
    grouped: dict[tuple[int, str], list[tuple[int, str, str]]] = {}
    for index, row in enumerate(table.rows):
        if _nondata_kind(row, table.columns):
            continue
        plate = _plate_token([row[plate_at]]) if plate_at < len(row) else ""
        if not plate:
            continue
        page = int(sources[index]) if index < len(sources) else 0
        for field in ("color", "make", "date"):
            column = fields.get(field)
            if column is None or column >= len(row):
                continue
            value = str(row[column]).replace("\u200e", "").strip()
            if value:
                grouped.setdefault((page, field), []).append((index, plate, value))
    found: list[list[str]] = []
    for (page, field), items in grouped.items():
        counts: dict[str, int] = {}
        for _index, _plate, value in items:
            counts[value] = counts.get(value, 0) + 1
        if not counts:
            continue
        dominant = max(counts, key=lambda key: (counts[key], key))
        for _index, plate, value in items:
            if field in {"color", "make"} and value != dominant:
                found.append([str(page), plate, labels[field], value, dominant, "يختلف عن السائد في الصفحة"])
            elif field == "date" and value != dominant and counts[value] <= 2:
                found.append([str(page), plate, labels[field], value, dominant, "تاريخ نادر"])
    return found


def _type_register_cells(worksheet, table: ExtractedTable) -> None:
    """Plate, year and admin number are numbers. The page is a number shown as صفحة N."""
    if table.kind != "register" and not any("لوحة" in str(name) for name in table.columns):
        return
    fields = _vehicle_field_at(table.columns)
    page_at = next((index for index, name in enumerate(table.columns) if "صفحة" in str(name) or str(name).strip().lower() == "page"), None)
    for row_offset, row in enumerate(table.rows):
        excel_row = row_offset + 2
        for field in ("plate", "year", "admin"):
            index = fields.get(field)
            if index is None or index >= len(row):
                continue
            number = _coerce_excel_number(row[index])
            if number is None:
                continue
            cell = worksheet.cell(excel_row, index + 1, number)
            cell.number_format = "0"
        if page_at is None or page_at >= len(row):
            continue
        raw_page = str(row[page_at] or "")
        for mark in ("\u200e", "\u200f", "\u2066", "\u2067", "\u2069"):
            raw_page = raw_page.replace(mark, "")
        digits = "".join(char for char in raw_page if char.isdigit())
        if not digits:
            continue
        cell = worksheet.cell(excel_row, page_at + 1, int(digits))
        cell.number_format = '"صفحة "0'


def _add_sort_date(worksheet, table: ExtractedTable) -> None:
    """A hidden real date, so the sheet can be sorted without touching the visible text."""
    fields = _vehicle_field_at(table.columns)
    date_at = fields.get("date")
    if date_at is None:
        return
    column = worksheet.max_column + 1
    header = worksheet.cell(1, column, "تاريخ للفرز")
    header.font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    header.fill = PatternFill("solid", fgColor="1F4E79")
    for row_offset, row in enumerate(table.rows):
        raw = row[date_at] if date_at < len(row) else ""
        text = str(_protect_excel_date(raw) or "").replace("\u200e", "")
        if not _DATE_VALUE.fullmatch(text):
            continue
        year, month, day = (int(part) for part in text.split("-"))
        try:
            written = date(year, month, day)
        except ValueError:
            continue
        cell = worksheet.cell(row_offset + 2, column, written)
        cell.number_format = "YYYY-MM-DD"
    worksheet.column_dimensions[get_column_letter(column)].hidden = True


def _write_summary(workbook, table: ExtractedTable, validation: list[dict] | None) -> None:
    """Counts by purpose and make, next to the printed total, plus the colour legend."""
    fields = _vehicle_field_at(table.columns)
    purpose_at = fields.get("purpose")
    make_at = fields.get("make")
    plate_at = fields.get("plate")
    purposes: dict[str, int] = {}
    makes: dict[str, int] = {}
    plated = 0
    for row in table.rows:
        if _nondata_kind(row, table.columns):
            continue
        plate = _plate_token([row[plate_at]]) if plate_at is not None and plate_at < len(row) else ""
        if not plate:
            continue
        plated += 1
        if purpose_at is not None and purpose_at < len(row) and str(row[purpose_at]).strip():
            name = str(row[purpose_at]).strip()
            purposes[name] = purposes.get(name, 0) + 1
        if make_at is not None and make_at < len(row) and str(row[make_at]).strip():
            name = str(row[make_at]).strip()
            makes[name] = makes.get(name, 0) + 1
    printed = {}
    for item in validation or []:
        printed[str(item.get("name") or "")] = str(item.get("expected") or "")
    name = "ملخص"
    if name in workbook.sheetnames:
        name = "الملخص"
    sheet = workbook.create_sheet(name)
    headers = ("البند", "العدد", "المطبوع")
    for index, name in enumerate(headers, start=1):
        cell = sheet.cell(1, index, name)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=14)
        cell.fill = PatternFill("solid", fgColor="1F4E79")
    row_number = 2
    sheet.cell(row_number, 1, "المجموع")
    sheet.cell(row_number, 2, plated)
    sheet.cell(row_number, 3, printed.get("المجموع", ""))
    row_number += 1
    for name, count in sorted(purposes.items(), key=lambda item: (-item[1], item[0])):
        sheet.cell(row_number, 1, name)
        sheet.cell(row_number, 2, count)
        printed_value = printed.get(name, "")
        sheet.cell(row_number, 3, printed_value if str(printed_value).strip() else "لم يُقرأ")
        row_number += 1
    row_number += 1
    sheet.cell(row_number, 1, "الصنع").font = Font(name="Arial", bold=True, size=14)
    row_number += 1
    known: dict[str, int] = {}
    unknown = 0
    for name, count in makes.items():
        if name in _MANUFACTURERS:
            known[name] = known.get(name, 0) + count
        else:
            unknown += count
    for name, count in sorted(known.items(), key=lambda item: (-item[1], item[0])):
        sheet.cell(row_number, 1, name)
        sheet.cell(row_number, 2, count)
        row_number += 1
    if unknown:
        sheet.cell(row_number, 1, f"غير معروف: {unknown}")
        sheet.cell(row_number, 2, unknown)
        row_number += 1
    legend = (
        ("F4B183", "برتقالي", "قراءة واحدة"),
        ("E2D5F1", "بنفسجي", "تصحيح إملائي"),
        ("FFF2CC", "أصفر", "اقتراح"),
    )
    row_number += 2
    sheet.cell(row_number, 1, "دليل الألوان").font = Font(name="Arial", bold=True, size=14)
    for color, swatch, meaning in legend:
        row_number += 1
        chip = sheet.cell(row_number, 1, swatch)
        chip.fill = PatternFill("solid", fgColor=color)
        chip.font = Font(name="Arial", size=14)
        sheet.cell(row_number, 2, meaning).font = Font(name="Arial", size=14)
    sheet.column_dimensions["A"].width = 28
    sheet.column_dimensions["B"].width = 22
    sheet.column_dimensions["C"].width = 16
    sheet.sheet_view.zoomScale = 130
    _set_sheet_direction(sheet, True)


def _format_sheet(worksheet, table: ExtractedTable) -> None:
    """Bold header, filter, cell types, and right-to-left layout when the table is Arabic."""
    column_count = len(table.columns)
    # A register, and any sheet with Arabic headers, opens from the right.
    # The saved view is set again after writing so Excel keeps that direction.
    forced = str(getattr(table, "direction", "") or "")
    if forced == "rtl":
        rtl = True
    elif forced == "ltr":
        rtl = False
    else:
        rtl = table.kind == "register" or _sheet_is_rtl(table) or any(
            _is_arabic_letter(char) for char in " ".join(table.columns)
        )
    if forced in {"rtl", "ltr"}:
        directions = [rtl] * column_count
    else:
        directions = [_column_is_rtl(table, index, rtl) for index in range(column_count)]
    worksheet.freeze_panes = "A2"
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToPage = True
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    thin = Border(
        left=Side(style="thin", color="BFCBD9"),
        right=Side(style="thin", color="BFCBD9"),
        top=Side(style="thin", color="BFCBD9"),
        bottom=Side(style="thin", color="BFCBD9"),
    )
    worksheet.row_dimensions[1].height = 24
    for index in range(1, column_count + 1):
        reading = 2 if directions[index - 1] else 1
        horizontal = "right" if directions[index - 1] else "left"
        cell = worksheet.cell(1, index)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=True, readingOrder=reading)
        cell.border = thin
        width = min(42, max(12, len(str(cell.value or "")) + 4))
        worksheet.column_dimensions[get_column_letter(index)].width = width
    # Body cells keep the values pandas wrote. Styling every cell, and giving
    # every cell its own number format, was the slow part of the workbook.
    # One column width and one header style cover the sheet. A provenance
    # colour is applied only to a cell that actually has a source.
    for col_index in range(1, column_count + 1):
        letter = get_column_letter(col_index)
        width = 14
        for row_index in range(2, min(worksheet.max_row + 1, 24)):
            width = max(width, min(42, len(str(worksheet.cell(row_index, col_index).value or "")) + 2))
        worksheet.column_dimensions[letter].width = max(worksheet.column_dimensions[letter].width or 0, width)
    # Field names are resolved once. Doing it per cell re-parses every header
    # thousands of times and was the whole of the workbook wait.
    if table.provenance:
        fields = _vehicle_field_at(table.columns)
        plate_at = fields.get("plate")
        column_fields = [_vehicle_field(name) for name in table.columns]
        one_engine_fill = PatternFill("solid", fgColor="F4B183")
        tagged = {"make", "color", "year", "date", "admin"}
        sources = list(table.sources or [])
        for row_offset, row in enumerate(table.rows):
            plate = _plate_token([row[plate_at]]) if plate_at is not None and plate_at < len(row) else ""
            if not plate:
                continue
            page = int(sources[row_offset]) if row_offset < len(sources) else 0
            for col_index, field_name in enumerate(column_fields):
                if field_name not in tagged:
                    continue
                source = table.provenance.get(_source_key(page, plate, field_name), "")
                if not source:
                    continue
                if source == "one-engine" and col_index < len(row) and str(row[col_index]).strip():
                    fill = one_engine_fill
                else:
                    fill = _SOURCE_FILLS.get(source)
                if fill is not None:
                    worksheet.cell(row_offset + 2, col_index + 1).fill = fill
    last = get_column_letter(column_count)
    worksheet.auto_filter.ref = f"A1:{last}{worksheet.max_row}"
    worksheet.sheet_view.zoomScale = 120
    _type_register_cells(worksheet, table)
    _add_sort_date(worksheet, table)
    # Last, so the freeze and the zoom cannot replace this view.
    _set_sheet_direction(worksheet, rtl)


@dataclass
class ExtractedFigure:
    """One photo, chart, or diagram, still the original bytes when the PDF stored them."""

    page: int
    x: float
    y: float
    width: float
    height: float
    data: bytes
    page_height: float


def _pixmap_png(page: fitz.Page, clip: fitz.Rect) -> bytes | None:
    """Rasterize a vector chart once. The result is a picture, not recognized text."""
    try:
        box = fitz.Rect(clip) & page.rect
        if box.width < 2 or box.height < 2:
            return None
        scale = 2.0
        long_side = max(box.width, box.height) * scale
        if long_side > 2400:
            scale = 2400 / max(box.width, box.height)
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=box, alpha=False)
        return pixmap.tobytes("png")
    except Exception:
        return None


def _embedded_figure_bytes(doc: fitz.Document, xref: int) -> bytes | None:
    """Prefer the file stored in the PDF. Rasterize only when Excel cannot embed it."""
    try:
        extracted = doc.extract_image(xref)
        data = extracted.get("image") if isinstance(extracted, dict) else None
    except Exception:
        data = None
    if data and _image_kind(data):
        return data
    try:
        pixmap = fitz.Pixmap(doc, xref)
        if pixmap.n - pixmap.alpha > 3:
            pixmap = fitz.Pixmap(fitz.csRGB, pixmap)
        return pixmap.tobytes("png")
    except Exception:
        return None


def _ruled_table_boxes(page: fitz.Page) -> list[fitz.Rect]:
    """Borders of a real grid. Those strokes are the table, not a chart."""
    try:
        found = page.find_tables(vertical_strategy="lines", horizontal_strategy="lines")
        tables = list(getattr(found, "tables", []) or [])
    except Exception:
        return []
    boxes: list[fitz.Rect] = []
    for table in tables:
        try:
            boxes.append(fitz.Rect(table.bbox))
        except Exception:
            continue
    return boxes


def _vector_charts(page: fitz.Page, image_boxes: list[fitz.Rect]) -> list[tuple[fitz.Rect, bytes]]:
    """A dense cluster of strokes is a chart or diagram, not a table rule."""
    try:
        drawings = page.get_drawings() or []
    except Exception:
        return []
    if not drawings or len(drawings) > 2500:
        return []
    pieces: list[fitz.Rect] = []
    for drawing in drawings:
        raw = drawing.get("rect")
        if raw is None:
            continue
        rect = fitz.Rect(raw)
        if rect.width < 0.4 or rect.height < 0.4:
            continue
        if max(rect.width, rect.height) > page.rect.width * 0.92 and min(rect.width, rect.height) < 3:
            continue
        pieces.append(rect)
    charts: list[tuple[fitz.Rect, bytes]] = []
    table_boxes = _ruled_table_boxes(page)
    # Padding has to cross the gap between neighbouring bars.
    for cluster in _merge_boxes(pieces, padding=14):
        if cluster.width < 90 or cluster.height < 60:
            continue
        if cluster.get_area() > page.rect.get_area() * 0.72:
            continue
        if _looks_like_seal(cluster, page.rect):
            continue
        if any((fitz.Rect(cluster) & box).get_area() > cluster.get_area() * 0.45 for box in table_boxes):
            continue
        if any(cluster.intersects(box) and _intersection_ratio(cluster, box) > 0.4 for box in image_boxes):
            continue
        count = 0
        for piece in pieces:
            if cluster.x0 - 1 <= piece.x0 and piece.x1 <= cluster.x1 + 1 and cluster.y0 - 1 <= piece.y0 and piece.y1 <= cluster.y1 + 1:
                count += 1
        if count < 10:
            continue
        rendered = _pixmap_png(page, cluster)
        if rendered:
            charts.append((cluster, rendered))
    return charts


def _intersection_ratio(left: fitz.Rect, right: fitz.Rect) -> float:
    overlap = fitz.Rect(left) & fitz.Rect(right)
    if overlap.is_empty or left.get_area() <= 0:
        return 0.0
    return overlap.get_area() / left.get_area()


def _figures_on_page(doc: fitz.Document, page: fitz.Page, page_index: int) -> list[ExtractedFigure]:
    """Photos and charts on one page. Stamps and a full-page scan are not figures."""
    figures: list[ExtractedFigure] = []
    image_boxes: list[fitz.Rect] = []
    try:
        infos = page.get_image_info(xrefs=True) or []
    except Exception:
        infos = []
    seen: set[tuple] = set()
    for info in infos:
        raw = info.get("bbox")
        xref = int(info.get("xref") or 0)
        if not raw or xref <= 0:
            continue
        rect = fitz.Rect(raw)
        if rect.width < 24 or rect.height < 24 or _covers_page(rect, page):
            continue
        key = (xref, round(rect.x0, 1), round(rect.y0, 1))
        if key in seen:
            continue
        seen.add(key)
        data = _embedded_figure_bytes(doc, xref)
        if not data:
            continue
        image_boxes.append(rect)
        figures.append(
            ExtractedFigure(
                page_index + 1,
                rect.x0,
                rect.y0,
                rect.width,
                rect.height,
                data,
                page.rect.height,
            )
        )
    for rect, data in _vector_charts(page, image_boxes):
        figures.append(
            ExtractedFigure(
                page_index + 1,
                rect.x0,
                rect.y0,
                rect.width,
                rect.height,
                data,
                page.rect.height,
            )
        )
    return figures


def extract_figures(pdf_path: Path, password: str = "") -> list[ExtractedFigure]:
    """Read pictures from the PDF. A failure here still leaves the tables usable."""
    try:
        document = open_document(pdf_path, password)
    except (FileNotFoundError, PermissionError, RuntimeError):
        return []
    try:
        figures: list[ExtractedFigure] = []
        for batch_start, batch_pages in iter_page_batches(document):
            for offset, page in enumerate(batch_pages):
                index = batch_start + offset
                try:
                    figures.extend(_figures_on_page(document, page, index))
                except Exception:
                    continue
        return figures
    finally:
        close_document(document)


def _add_figure_sheet(workbook, figures: list[ExtractedFigure]) -> None:
    """Place each picture where it sat on the page. Pages are stacked down the sheet."""
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
    from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D

    sheet = workbook.create_sheet("Figures")
    sheet.sheet_view.showGridLines = False
    sheet["A1"] = "Photos and charts, copied at their original positions. They were not rewritten or read as text."
    sheet.column_dimensions["A"].width = 88
    page_height: dict[int, float] = {}
    for figure in figures:
        page_height[figure.page] = max(page_height.get(figure.page, 0.0), figure.page_height)
    offset: dict[int, float] = {}
    cursor = 36.0
    for page in sorted(page_height):
        offset[page] = cursor
        cursor += page_height[page] + 28.0
    for figure in figures:
        try:
            image = XLImage(BytesIO(figure.data))
            image.width = max(1, int(round(figure.width * 96 / 72)))
            image.height = max(1, int(round(figure.height * 96 / 72)))
            image.anchor = AbsoluteAnchor(
                pos=XDRPoint2D(int(round(figure.x * 12700)), int(round((offset[figure.page] + figure.y) * 12700))),
                ext=XDRPositiveSize2D(int(round(figure.width * 12700)), int(round(figure.height * 12700))),
            )
            sheet.add_image(image)
        except Exception:
            continue
    if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
        default = workbook["Sheet"]
        if default["A1"].value is None and default.max_row <= 1:
            workbook.remove(default)


def _validation_from_tables(tables: list[ExtractedTable]) -> list[dict[str, str]]:
    """Build the checks from the sheets about to be written."""
    if not tables:
        return []
    columns, rows, pages = records_from_tables(tables)
    detected: dict[int, int] = {}
    statements: list[str] = []
    for table in tables:
        detected.update(dict(table.detected_pages or {}))
        for line in table.statements or []:
            if line not in statements:
                statements.append(line)
    printed_total, printed_groups = _printed_statements(statements)
    return validate_register(columns, rows, pages, printed_total, printed_groups, detected)


def _validation_sheet_rows(checks: list[dict[str, str]], arabic: bool) -> tuple[list[str], list[list[str]]]:
    if arabic:
        header = ["التحقق", "الحالة", "المتوقع", "الناتج", "التفصيل"]
    else:
        header = ["Check", "Status", "Expected", "Actual", "Detail"]
    rows = [
        [item.get("name", ""), item.get("status", ""), item.get("expected", ""), item.get("actual", ""), item.get("detail", "")]
        for item in checks
    ]
    return header, rows


def _write_audit(workbook, rows: list) -> None:
    """One row per review edit: when, who, page, plate, field, before, after."""
    if not rows:
        return
    sheet = workbook.create_sheet("سجل")
    sheet.sheet_view.rightToLeft = True
    sheet.append(["الوقت", "المراجع", "الصفحة", "اللوحة", "الحقل", "قبل", "بعد"])
    for item in rows:
        if not isinstance(item, dict):
            continue
        sheet.append(
            [
                str(item.get("when") or ""),
                str(item.get("who") or ""),
                item.get("page") if item.get("page") is not None else "",
                str(item.get("plate") or ""),
                str(item.get("field") or ""),
                str(item.get("before") or ""),
                str(item.get("after") or ""),
            ]
        )
    for cell in sheet[1]:
        cell.font = Font(bold=True)


def write_excel(
    tables: list[ExtractedTable],
    output_path: Path,
    figures: list[ExtractedFigure] | None = None,
    preserve_edits: bool = False,
    strict: bool = False,
    validation: list[dict[str, str]] | None = None,
    direction: str = "auto",
    developer: bool = False,
) -> None:
    """pandas writes each table. openpyxl sets the header, filter, direction, and figures."""
    strict_token = _STRICT.set(bool(strict))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    # The temporary name has to end in .xlsx. openpyxl refuses to reopen
    # any other suffix when the right-to-left view is applied.
    partial = output_path.with_name(f"~{output_path.stem}.partial.xlsx")
    used_names: set[str] = set()
    figures = figures or []
    try:
        # Edited rows are written as the user left them. A fresh extract is cleaned first.
        # Strict reading does not copy a blank make, colour, or year from the row above.
        if preserve_edits:
            tables = _sheets_for_workbook(list(tables))
        else:
            tables = _sheets_for_workbook(sanitize_tables(tables))
        chosen_direction = str(direction or "auto")
        for table in tables:
            if chosen_direction in {"rtl", "ltr"}:
                table.direction = chosen_direction
        sheet_rtl = chosen_direction == "rtl" or (
            chosen_direction != "ltr"
            and any(table.kind == "register" or _sheet_is_rtl(table) for table in tables)
        )
        with pd.ExcelWriter(partial, engine="openpyxl") as writer:
            data_index = 0
            excluded_frames: list[pd.DataFrame] = []
            for table in tables:
                # A hand edit is written as stored. Cleaning runs only on a fresh extract.
                if not preserve_edits:
                    _normalize_table_dates(table)
                    annotate_vehicle_flags(table)
                    _merge_wrapped_fragment_rows(table)
                    _forward_fill_attributes(table)
                columns = [_clean_cell(name) for name in table.columns]
                rows = []
                for row in table.rows:
                    padded = list(row) + [""] * max(0, len(columns) - len(row))
                    rows.append(
                        [
                            _normalize_date_cell(cell) if _export_column_kind(columns[index]) == "date" else _clean_cell(cell)
                            for index, cell in enumerate(padded[: len(columns)])
                        ]
                    )
                table.columns = columns
                sources = list(table.sources or [])
                if len(sources) != len(rows):
                    sources = [1] * len(rows)
                data_rows = []
                data_sources = []
                excluded_rows = []
                for row, source in zip(rows, sources):
                    if _nondata_kind(row, columns):
                        excluded_rows.append(row)
                    else:
                        data_rows.append(row)
                        data_sources.append(source)
                table.rows = data_rows
                table.sources = data_sources
                frame = pd.DataFrame(data_rows, columns=columns).fillna("")
                if excluded_rows:
                    excluded_frames.append(pd.DataFrame(excluded_rows, columns=columns).fillna(""))
                if columns and columns[0] in {"Sheet", "الورقة"}:
                    name = "ملخص" if columns[0] == "الورقة" else "Summary"
                    if name in used_names:
                        name = _sheet_name(1, used_names)
                    else:
                        used_names.add(name)
                else:
                    data_index += 1
                    name = _sheet_name(data_index, used_names)
                frame.to_excel(writer, index=False, sheet_name=name)
                _format_sheet(writer.sheets[name], table)
            if developer and excluded_frames:
                excluded = pd.concat(excluded_frames, ignore_index=True)
                excluded.to_excel(writer, index=False, sheet_name="excluded")
            if developer:
                provenance_rows: list[list[str]] = []
                for table in tables:
                    provenance_rows.extend(_provenance_sheet_rows(table))
                if provenance_rows:
                    pd.DataFrame(
                        provenance_rows,
                        columns=["الصفحة", "اللوحة", "الحقل", "القيمة", "المصدر", "الاقتراح"],
                    ).to_excel(writer, index=False, sheet_name="provenance")
            exception_rows: list[list[str]] = []
            for table in tables:
                exception_rows.extend(_exception_sheet_rows(table))
            if exception_rows:
                pd.DataFrame(
                    exception_rows,
                    columns=["الصفحة", "اللوحة", "الحقل", "القيمة", "السائد في الصفحة", "السبب"],
                ).to_excel(writer, index=False, sheet_name="مراجعة")
            if validation is None:
                checked = time.perf_counter()
                validation = _validation_from_tables(tables)
                _remember_time("validation", time.perf_counter() - checked)
            if developer and validation:
                arabic = any(_is_arabic_letter(char) for char in " ".join(tables[0].columns)) if tables else True
                header, sheet_rows = _validation_sheet_rows(validation, arabic)
                pd.DataFrame(sheet_rows, columns=header).to_excel(writer, index=False, sheet_name="validation")
            for table in tables:
                if table.kind == "register" or any("لوحة" in str(name) for name in table.columns):
                    _write_summary(writer.book, table, validation)
                    break
            audit_rows: list = []
            for table in tables:
                audit_rows.extend(list(getattr(table, "audit", None) or []))
            if audit_rows:
                _write_audit(writer.book, audit_rows)
            if figures:
                _add_figure_sheet(writer.book, figures)
            elif not tables:
                raise ConversionError(
                    "No table was found. A digital PDF needs text in rows and columns. A scan is read from the picture on this PC."
                )
            # Direction is set on the workbook pandas already holds, then saved once.
            # Reopening the file to paint every cell was most of the wait.
            for sheet in writer.book.worksheets:
                _set_sheet_direction(sheet, sheet_rtl)
        partial.replace(output_path)
    except PermissionError as exc:
        raise ConversionError("Close the Excel file if it is open, then try again.") from exc
    except OSError as exc:
        raise ConversionError(f"The Excel file could not be saved. {exc}") from exc
    finally:
        _STRICT.reset(strict_token)
        if partial.exists():
            try:
                partial.unlink()
            except OSError:
                pass


_PAGE_MODES = ("auto", "grid", "words", "ocr")


def _fit_row(source_columns: list[str], row: list[str], target_columns: list[str]) -> list[str]:
    """Place one row's cells under a shared header list. A missing column stays blank."""
    lookup = {name: (row[index] if index < len(row) else "") for index, name in enumerate(source_columns)}
    return [str(lookup.get(name, "") or "") for name in target_columns]


def records_from_tables(tables: list[ExtractedTable]) -> tuple[list[str], list[list[str]], list[int]]:
    """
    One header, one row list, and the source page of each row.

    The page number is the counter assigned when that page was read.
    """
    columns: list[str] = []
    rows: list[list[str]] = []
    pages: list[int] = []
    for table in tables:
        incoming = [str(name) for name in table.columns]
        for name in incoming:
            if name not in columns:
                columns.append(name)
                for row in rows:
                    row.append("")
        sources = table.sources if table.sources and len(table.sources) == len(table.rows) else [1] * len(table.rows)
        for row, page in zip(table.rows, sources):
            cells = ["" if cell is None else str(cell) for cell in row]
            rows.append(_fit_row(incoming, cells, columns))
            pages.append(int(page))
    return columns, rows, pages


def replace_page_rows(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    tables: list[ExtractedTable],
    page_number: int,
) -> tuple[list[str], list[list[str]], list[int]]:
    """Drop one page's rows and put the new reading of that page in their place."""
    page_number = int(page_number)
    old_columns = list(columns)
    fresh_columns, fresh_rows, _fresh_pages = records_from_tables(tables)
    merged = list(old_columns)
    for name in fresh_columns:
        if name not in merged:
            merged.append(name)
    fresh_rows = [_fit_row(fresh_columns, row, merged) for row in fresh_rows]
    output_rows: list[list[str]] = []
    output_pages: list[int] = []
    placed = False
    for row, page in zip(rows, page_numbers):
        if int(page) == page_number:
            if not placed:
                output_rows.extend(fresh_rows)
                output_pages.extend([page_number] * len(fresh_rows))
                placed = True
            continue
        output_rows.append(_fit_row(old_columns, row, merged))
        output_pages.append(int(page))
    if not placed:
        output_rows.extend(fresh_rows)
        output_pages.extend([page_number] * len(fresh_rows))
    return merged, output_rows, output_pages


def _page_row_indexes(page_numbers: list[int], page_number: int) -> list[int]:
    """Positions, in the full sheet, of one page's rows."""
    return [index for index, page in enumerate(page_numbers) if int(page) == int(page_number)]


def replace_row_range(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    page_number: int,
    start: int,
    end: int,
    tables: list[ExtractedTable],
) -> tuple[list[str], list[list[str]], list[int]]:
    """
    Replace one inclusive span of rows on a single page.

    start and end count from 1 within that page. Other pages stay where they are.
    """
    indexes = _page_row_indexes(page_numbers, page_number)
    if not indexes:
        return replace_page_rows(columns, rows, page_numbers, tables, page_number)
    start = max(1, int(start))
    end = min(max(start, int(end)), len(indexes))
    drop = set(indexes[start - 1 : end])
    insert_at = indexes[start - 1]
    fresh_columns, fresh_rows, _fresh_pages = records_from_tables(tables)
    merged = list(columns)
    for name in fresh_columns:
        if name not in merged:
            merged.append(name)
    fresh_rows = [_fit_row(fresh_columns, row, merged) for row in fresh_rows]
    output_rows: list[list[str]] = []
    output_pages: list[int] = []
    placed = False
    for index, (row, page) in enumerate(zip(rows, page_numbers)):
        if index in drop:
            if not placed:
                output_rows.extend(fresh_rows)
                output_pages.extend([int(page_number)] * len(fresh_rows))
                placed = True
            continue
        output_rows.append(_fit_row(list(columns), list(row), merged))
        output_pages.append(int(page))
    if not placed:
        output_rows.extend(fresh_rows)
        output_pages.extend([int(page_number)] * len(fresh_rows))
    return merged, output_rows, output_pages


def repair_date_fragment(text: str, year_hint: str = "") -> str:
    """
    Clean one انتهاء التأمين cell.

    A full date becomes YYYY-MM-DD. '- 04 '09' stays 04-09. year_hint is
    not written in, because that year was not part of this cell.
    """
    del year_hint
    if not _clean_cell(text):
        return ""
    return _normalize_date_cell(text)


def _year_hint(columns: list[str], row: list[str], above: list[str] | None) -> str:
    """A 4-digit year from this row only. The row above is not used."""
    del above
    probe = ExtractedTable(list(columns), [list(row)])
    year_at = _column_index(probe, "year", "السنه")
    date_at = set(_date_column_indexes(probe))
    ordered: list[str] = []
    if year_at is not None and year_at < len(row):
        ordered.append(str(row[year_at]))
    for index in date_at:
        if index < len(row):
            ordered.append(str(row[index]))
    for text in ordered:
        match = re.search(r"(?:19|20)\d{2}", _clean_cell(text).translate(_INDIC_DIGITS))
        if match and _YEAR_VALUE.fullmatch(match.group()):
            return match.group()
    return ""


def repair_dates_in_range(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    page_number: int,
    start: int,
    end: int,
) -> tuple[list[list[str]], int]:
    """Repair date cells on one page's row span. Returns the rows and how many changed."""
    indexes = _page_row_indexes(page_numbers, page_number)
    if not indexes:
        return rows, 0
    start = max(1, int(start))
    end = min(max(start, int(end)), len(indexes))
    probe = ExtractedTable(list(columns), [list(rows[indexes[0]])])
    date_at = _date_column_indexes(probe)
    if not date_at:
        return rows, 0
    updated = [list(row) for row in rows]
    changed = 0
    for position in range(start - 1, end):
        global_index = indexes[position]
        above = updated[indexes[position - 1]] if position else None
        hint = _year_hint(columns, updated[global_index], above)
        for column in date_at:
            while len(updated[global_index]) <= column:
                updated[global_index].append("")
            current = updated[global_index][column]
            repaired = repair_date_fragment(current, hint)
            if repaired != current:
                updated[global_index][column] = repaired
                changed += 1
    return updated, changed


def fill_attributes_from_above(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    page_number: int,
    start: int,
    end: int,
) -> tuple[list[list[str]], int]:
    """Copy a blank colour, make, or year from the previous row on the same page."""
    indexes = _page_row_indexes(page_numbers, page_number)
    if len(indexes) < 2:
        return rows, 0
    start = max(1, int(start))
    end = min(max(start, int(end)), len(indexes))
    probe = ExtractedTable(list(columns), [list(rows[indexes[0]])])
    fields = [(index, field) for index, field in _repeat_columns(probe) if field in {"make", "color", "year"}]
    if not fields:
        return rows, 0
    updated = [list(row) for row in rows]
    changed = 0
    for position in range(start - 1, end):
        if position == 0:
            continue
        current = updated[indexes[position]]
        previous = updated[indexes[position - 1]]
        if _nondata_kind(current, columns) or _nondata_kind(previous, columns):
            continue
        for column, field in fields:
            while len(current) <= column:
                current.append("")
            while len(previous) <= column:
                previous.append("")
            if str(current[column]).strip():
                continue
            value = _inheritable_attribute(field, previous[column])
            if not value:
                continue
            current[column] = value
            changed += 1
    return updated, changed


def fill_blank_fields_from_above(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    fields: tuple[str, ...] | list[str],
    indexes: list[int] | None = None,
) -> tuple[list[list[str]], int]:
    """
    Copy a blank الصنع or اللون from the nearest filled cell above it.

    Only the named fields are copied, and only into a cell that is still
    blank. The walk stops at a header, a gap, or the previous page. Nothing
    is copied until the caller asks.
    """
    if not rows:
        return rows, 0
    wanted = {str(field) for field in fields}
    probe = ExtractedTable(list(columns), [list(rows[0])])
    targets = [(index, field) for index, field in _repeat_columns(probe) if field in wanted and field in {"make", "color"}]
    if not targets:
        return rows, 0
    selected = None if indexes is None else {int(index) for index in indexes}
    updated = [list(row) for row in rows]
    changed = 0
    for position, current in enumerate(updated):
        if selected is not None and position not in selected:
            continue
        if _nondata_kind(current, columns):
            continue
        page = int(page_numbers[position]) if position < len(page_numbers) else None
        for column, field in targets:
            while len(current) <= column:
                current.append("")
            if str(current[column]).strip():
                continue
            value = ""
            for earlier in range(position - 1, -1, -1):
                if page is not None and earlier < len(page_numbers) and int(page_numbers[earlier]) != page:
                    break
                previous = updated[earlier]
                if _nondata_kind(previous, columns):
                    break
                while len(previous) <= column:
                    previous.append("")
                candidate = _inheritable_attribute(field, previous[column])
                if candidate:
                    value = candidate
                    break
            if not value:
                continue
            current[column] = value
            changed += 1
    return updated, changed


def normalize_edited_dates(
    columns: list[str],
    rows: list[list[str]],
    indexes: list[int] | None = None,
) -> tuple[list[list[str]], int]:
    """
    Rewrite dates that already contain a year to YYYY-MM-DD.

    A fragment such as 04-09 stays 04-09. A missing year is not invented
    from the row above or from the model year.
    """
    if not rows:
        return rows, 0
    probe = ExtractedTable(list(columns), [list(rows[0])])
    date_at = _date_column_indexes(probe)
    if not date_at:
        return rows, 0
    chosen = range(len(rows)) if indexes is None else indexes
    updated = [list(row) for row in rows]
    changed = 0
    for position in chosen:
        if position < 0 or position >= len(updated):
            continue
        row = updated[position]
        if _nondata_kind(row, columns):
            continue
        for column in date_at:
            while len(row) <= column:
                row.append("")
            current = str(row[column])
            repaired = repair_date_fragment(current, "")
            if repaired != current:
                row[column] = repaired
                changed += 1
    return updated, changed


def row_needs_review(columns: list[str], row: list[str]) -> bool:
    """A vehicle row with an empty make, colour, year, or date, or a date that is not YYYY-MM-DD."""
    if _nondata_kind(list(row), columns):
        return False
    fields = _vehicle_field_at(columns)
    for key in ("make", "color", "year", "date"):
        index = fields.get(key)
        if index is None:
            continue
        value = str(row[index]).strip() if index < len(row) else ""
        if not value:
            return True
        if key == "date" and not _DATE_VALUE.fullmatch(value):
            return True
    return False


def repair_status(columns: list[str], rows: list[list[str]]) -> dict[str, int]:
    """Remaining blank make/colour cells, invalid dates, and rows that still need review."""
    fields = _vehicle_field_at(columns)
    make_at = fields.get("make")
    color_at = fields.get("color")
    date_at = fields.get("date")
    blanks = 0
    invalid_dates = 0
    review_rows = 0
    for row in rows:
        if _nondata_kind(list(row), columns):
            continue
        if row_needs_review(columns, list(row)):
            review_rows += 1
        for index in (make_at, color_at):
            if index is None:
                continue
            if index >= len(row) or not str(row[index]).strip():
                blanks += 1
        if date_at is None:
            continue
        raw = str(row[date_at]).strip() if date_at < len(row) else ""
        if not _DATE_VALUE.fullmatch(raw):
            invalid_dates += 1
    return {"blanks": blanks, "invalid_dates": invalid_dates, "review_rows": review_rows}


def merge_fragment_rows(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    page_number: int,
    start: int,
    end: int,
) -> tuple[list[list[str]], list[int], int]:
    """
    Fold a gap fragment into the row above it, then remove the fragment.

    A fragment that has its own plate, different from the row above, stays.
    """
    indexes = _page_row_indexes(page_numbers, page_number)
    if len(indexes) < 2:
        return rows, page_numbers, 0
    start = max(1, int(start))
    end = min(max(start, int(end)), len(indexes))
    if start <= 1:
        start = 2
    if start > end:
        return rows, page_numbers, 0
    probe = ExtractedTable(list(columns), [list(rows[indexes[0]])])
    plate_at = _column_index(probe, "plate", "رقماللوحه")
    fold = [(index, field) for index, field in _repeat_columns(probe)]
    purpose_at = _column_index(probe, "purpose")
    if purpose_at is not None:
        fold.append((purpose_at, "purpose"))
    updated = [list(row) for row in rows]
    anchor = updated[indexes[start - 2]]
    removed: set[int] = set()
    merged = 0
    for position in range(start - 1, end):
        fragment = updated[indexes[position]]
        if (
            plate_at is not None
            and _plated_row(fragment, plate_at)
            and _plated_row(anchor, plate_at)
            and _normalize_plate_cell(fragment[plate_at]) != _normalize_plate_cell(anchor[plate_at])
        ):
            continue
        for column, field in fold:
            while len(fragment) <= column:
                fragment.append("")
            while len(anchor) <= column:
                anchor.append("")
            extra = str(fragment[column]).strip()
            if not extra:
                continue
            current = str(anchor[column]).strip()
            if field == "year":
                if not current:
                    anchor[column] = _inheritable_attribute(field, extra) or extra
                continue
            if not current:
                anchor[column] = extra
            elif extra not in current:
                anchor[column] = f"{current} {extra}".strip()
        removed.add(indexes[position])
        merged += 1
    if not removed:
        return rows, page_numbers, 0
    kept_rows = [row for index, row in enumerate(updated) if index not in removed]
    kept_pages = [page for index, page in enumerate(page_numbers) if index not in removed]
    return kept_rows, kept_pages, merged


def delete_row_range(
    rows: list[list[str]],
    page_numbers: list[int],
    page_number: int,
    start: int,
    end: int,
) -> tuple[list[list[str]], list[int], int]:
    """Remove an inclusive span of rows from one page."""
    indexes = _page_row_indexes(page_numbers, page_number)
    if not indexes:
        return rows, page_numbers, 0
    start = max(1, int(start))
    end = min(max(start, int(end)), len(indexes))
    drop = set(indexes[start - 1 : end])
    kept_rows = [row for index, row in enumerate(rows) if index not in drop]
    kept_pages = [page for index, page in enumerate(page_numbers) if index not in drop]
    return kept_rows, kept_pages, len(drop)


def _locate_row_y(page: fitz.Page, row: list[str]) -> float | None:
    """The vertical centre of a row, found from a plate or another distinctive token."""
    words = _page_words(page)
    if not words:
        return None
    needles: list[str] = []
    for cell in row:
        digits = re.sub(r"\D", "", str(cell).translate(_INDIC_DIGITS))
        if 4 <= len(digits) <= 8:
            needles.append(digits)
    if not needles:
        return None
    hits: list[float] = []
    for word in words:
        compact = re.sub(r"\D", "", word.text.translate(_INDIC_DIGITS))
        if any(needle and needle in compact for needle in needles):
            hits.append(word.center_y)
    if not hits:
        return None
    hits.sort()
    return hits[len(hits) // 2]


def _band_for_rows(page: fitz.Page, page_rows: list[list[str]], start: int, end: int) -> tuple[float, float]:
    """The vertical span of one inclusive row range on the page."""
    height = float(page.rect.height)
    count = max(len(page_rows), 1)
    top = height * 0.08
    bottom = height * 0.92
    span = (bottom - top) / count
    centres = [_locate_row_y(page, row) for row in page_rows]
    start_index = max(0, int(start) - 1)
    end_index = min(count - 1, max(start_index, int(end) - 1))
    y0 = centres[start_index]
    y1 = centres[end_index]
    if y0 is None:
        y0 = top + start_index * span
    if y1 is None:
        y1 = top + (end_index + 1) * span
    else:
        following = centres[end_index + 1] if end_index + 1 < len(centres) else None
        y1 = (following - 2.0) if following and following > y1 else y1 + max(16.0, span * 0.65)
    y0 = max(0.0, y0 - 8.0)
    y1 = min(height, max(y0 + 12.0, y1))
    return y0, y1


def _preview_column_boxes(page: fitz.Page, columns: list[str]) -> list[tuple[float, float] | None]:
    """X span of each preview column, from the ruled header when it can be read."""
    boxes: list[tuple[float, float] | None] = [None] * len(columns)
    words = _page_words(page)
    spans = _vehicle_header_spans(words) if words else None
    if spans:
        vertical, _horizontal = _page_rules(page)
        edges = _vehicle_column_edges(spans, vertical, float(page.rect.width))
        if len(edges) == len(spans) + 1:
            for index, (field, _word) in enumerate(spans):
                for column, name in enumerate(columns):
                    if _vehicle_field(name) == field:
                        boxes[column] = (edges[index], edges[index + 1])
                        break
    matched = [(index, box) for index, box in enumerate(boxes) if box is not None]
    data_columns = [
        index
        for index, name in enumerate(columns)
        if not _header_is_printed_page(name) and not _is_notes_header(name)
    ]
    if len(matched) >= 2:
        known = sorted(matched, key=lambda item: item[1][0])
        for (left_index, left_box), (right_index, right_box) in zip(known, known[1:]):
            gap_indexes = [index for index in data_columns if left_index < index < right_index and boxes[index] is None]
            if not gap_indexes:
                continue
            span = right_box[0] - left_box[1]
            slot = span / (len(gap_indexes) + 1)
            for offset, index in enumerate(gap_indexes, start=1):
                boxes[index] = (left_box[1] + (offset - 1) * slot, left_box[1] + offset * slot)
    elif len(matched) < max(1, len(data_columns) // 2):
        width = float(page.rect.width)
        slot = width / max(len(data_columns), 1)
        boxes = [None] * len(columns)
        for offset, index in enumerate(data_columns):
            boxes[index] = (offset * slot, (offset + 1) * slot)
    return boxes


def _rows_from_word_band(
    page: fitz.Page,
    y0: float,
    y1: float,
    columns: list[str],
    tolerance: float,
) -> list[list[str]]:
    """Group words inside a vertical band into the preview columns."""
    words = [word for word in _page_words(page) if y0 <= word.center_y <= y1]
    if not words:
        return []
    boxes = _preview_column_boxes(page, columns)
    built: list[list[str]] = []
    for line in _group_rows(words, tolerance):
        cells = [""] * len(columns)
        for word in sorted(line, key=lambda item: item.x0):
            hits = [
                index
                for index, box in enumerate(boxes)
                if box is not None and box[0] <= word.center_x <= box[1]
            ]
            if not hits:
                continue
            index = min(hits, key=lambda item: boxes[item][1] - boxes[item][0])
            cells[index] = f"{cells[index]} {word.text}".strip()
        logical = [_logical_arabic(cell) if cell else "" for cell in cells]
        if any(cell.strip() for cell in logical):
            built.append(logical)
    return built


def _rows_from_ocr_band(
    page: fitz.Page,
    y0: float,
    y1: float,
    columns: list[str],
) -> list[list[str]]:
    """OCR only the vertical band, then place the cells under the current headers."""
    clip = fitz.Rect(page.rect.x0, y0, page.rect.x1, max(y0 + 12.0, y1))
    scratch = fitz.open()
    try:
        band = scratch.new_page(width=float(page.rect.width), height=float(clip.height))
        band.show_pdf_page(band.rect, page.parent, page.number, clip=clip)
        found = _tables_from_scan(band, 0, {})
    except Exception:
        found = []
    finally:
        scratch.close()
    if not found:
        return []
    _columns, scanned, _pages = records_from_tables(found)
    source_columns = found[0].columns if found else _columns
    # records_from_tables already unioned columns; fit from that union back is lossy.
    # Fit each scanned table by its own headers.
    fitted: list[list[str]] = []
    for table in found:
        for row in table.rows:
            cells = ["" if cell is None else str(cell) for cell in row]
            fitted.append(_fit_row([str(name) for name in table.columns], cells, columns))
    if fitted:
        return fitted
    return [_fit_row(source_columns, row, columns) for row in scanned]


def _polish_band_rows(
    columns: list[str], rows: list[list[str]], page_number: int, kind: str
) -> tuple[list[str], list[list[str]]]:
    """Normalize one band, drop a header or subtotal line, and fill the review note."""
    if not rows:
        return list(columns), []
    table = ExtractedTable(
        list(columns),
        [list(row) for row in rows],
        None,
        [int(page_number)] * len(rows),
    )
    table.kind = kind or ""
    _normalize_table_dates(table)
    _apply_color_dictionary(table)
    _apply_manufacturer_dictionary(table)
    _isolate_nondata_rows(table)
    table.sources = [int(page_number)] * len(table.rows)
    annotate_vehicle_flags(table)
    word = "صفحة" if any(_is_arabic_letter(char) for char in " ".join(columns)) else "Page"
    label = f"{word} \u2066{int(page_number)}\u2069"
    for index, name in enumerate(columns):
        if not _header_is_printed_page(name):
            continue
        for row in table.rows:
            while len(row) <= index:
                row.append("")
            row[index] = label
    return list(table.columns), table.rows


def extract_row_band(
    pdf_path: str | Path,
    page_number: int,
    page_rows: list[list[str]],
    columns: list[str],
    start: int,
    end: int,
    password: str = "",
    mode: str = "auto",
    kind: str = "register",
    strict: bool = False,
) -> list[ExtractedTable]:
    """
    Re-read the vertical band of one row span.

    The rest of the page is not returned. mode is auto, grid, words, or ocr.
    """
    chosen = (mode or "auto").strip().lower()
    if chosen not in _PAGE_MODES:
        chosen = "auto"
    if not page_rows or not columns:
        return []
    source = Path(pdf_path)
    document = open_document(source, password)
    strict_token = _STRICT.set(bool(strict))
    try:
        if page_number < 1 or page_number > document.page_count:
            raise ConversionError(f"Page {page_number} is not in this PDF.")
        page = document.load_page(int(page_number) - 1)
        y0, y1 = _band_for_rows(page, page_rows, start, end)
        vertical, _horizontal = _page_rules(page)
        use_grid = chosen == "grid" or (chosen == "auto" and len(vertical) >= 4)
        if chosen == "ocr":
            band_rows = _rows_from_ocr_band(page, y0, y1, columns)
        else:
            tolerance = 3.0 if use_grid else 6.5
            band_rows = _rows_from_word_band(page, y0, y1, columns, tolerance)
            if not band_rows and chosen == "auto":
                band_rows = _rows_from_ocr_band(page, y0, y1, columns)
        band = ExtractedTable(list(columns), [list(row) for row in band_rows], None, [int(page_number)] * len(band_rows))
        _assemble_row_dates(page, band, y_floor=y0, y_limit=y1)
        columns, band_rows = _polish_band_rows(columns, band.rows, int(page_number), kind)
        if not band_rows:
            raise ConversionError("That row band did not produce any cells.")
        return [ExtractedTable(list(columns), band_rows, None, [int(page_number)] * len(band_rows))]
    finally:
        _STRICT.reset(strict_token)
        close_document(document)


def _vehicle_field_at(columns: list[str]) -> dict[str, int]:
    """Index of each vehicle field. Page and review columns are ignored."""
    found: dict[str, int] = {}
    for index, name in enumerate(columns):
        if _header_is_printed_page(name) or _is_notes_header(name):
            continue
        field_name = _vehicle_field(name)
        if field_name and field_name not in found:
            found[field_name] = index
    return found


def _vehicle_row_flags(columns: list[str], row: list[str]) -> str:
    """
    Notes for one vehicle row.

    A bad plate, year, colour, make, or date is named here. The other cells
    in the row, and the rows around it, stay where they are.
    """
    fields = _vehicle_field_at(columns)
    if "plate" not in fields:
        return ""

    def cell(field_name: str) -> str:
        index = fields.get(field_name)
        if index is None or index >= len(row):
            return ""
        return str(row[index]).strip()

    plate = _normalize_plate_cell(cell("plate"))
    year = _normalize_year_cell(cell("year"))
    date_raw = cell("date")
    iso = _normalize_date_cell(date_raw)
    color = cell("color")
    make = cell("make")
    flags: list[str] = []
    plated = bool(plate and _PLATE_VALUE.fullmatch(plate) and not _YEAR_VALUE.fullmatch(plate))
    if not plated:
        flags.append("لوحة مفقودة أو غير صالحة")
    if cell("year") and not _YEAR_VALUE.fullmatch(year or ""):
        flags.append("سنة غير صالحة")
    elif plated and not year:
        flags.append("سنة مفقودة")
    if not date_raw:
        flags.append("تاريخ مفقود")
    elif re.fullmatch(r"\d{2}-\d{2}", iso or ""):
        flags.append("تاريخ ناقص")
    elif not _DATE_VALUE.fullmatch(iso or ""):
        flags.append("تاريخ غير مقروء")
    if color and not (_match_official_color(color) or _match_color(color)):
        flags.append("لون غير معروف")
    elif plated and not color:
        flags.append("لون مفقود")
    brand = _normalize_brand(make) if make else ""
    known_make = brand in _MANUFACTURERS or brand in {item[0] for item in _BRAND_ALIASES}
    if make and not known_make:
        flags.append("صنع غير معروف")
    elif plated and not make:
        flags.append("صنع مفقود")
    if plated and year and plate == year:
        flags.append("تعارض بين اللوحة والسنة")
    if date_raw and _PLATE_VALUE.fullmatch(_normalize_plate_cell(date_raw)) and not _DATE_VALUE.fullmatch(iso or ""):
        flags.append("تعارض في التاريخ")
    return "؛ ".join(flags)


def annotate_vehicle_flags(table: ExtractedTable) -> None:
    """Fill the review column for a vehicle sheet. Other columns are not moved."""
    if "plate" not in _vehicle_field_at(table.columns):
        return
    notes_at = next((index for index, name in enumerate(table.columns) if _is_notes_header(name)), None)
    if notes_at is None:
        table.columns = list(table.columns) + [_NOTES_HEADER]
        notes_at = len(table.columns) - 1
    for row in table.rows:
        while len(row) <= notes_at:
            row.append("")
        kind = _nondata_kind(row, table.columns)
        if kind:
            row[notes_at] = kind
            continue
        flags = [part.strip() for part in _vehicle_row_flags(table.columns, row).split("؛") if part.strip()]
        row[notes_at] = "؛ ".join(flags)


def review_page(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    page_number: int,
) -> tuple[int, int, int]:
    """
    Row count, garbled dates, and blank make/colour/date cells on one page.

    A date that is not already YYYY-MM-DD, and a blank manufacturer or colour
    on a row that otherwise has text, are the two warnings.
    """
    page_number = int(page_number)
    chosen = [row for row, page in zip(rows, page_numbers) if int(page) == page_number]
    if not chosen:
        return 0, 0, 0
    if "plate" in _vehicle_field_at(columns):
        garbled = 0
        missing = 0
        for row in chosen:
            if _nondata_kind(list(row), columns):
                continue
            note = _vehicle_row_flags(columns, list(row))
            if "تاريخ" in note:
                garbled += 1
            if any(token in note for token in ("لوحة", "سنة", "لون", "صنع", "تعارض")):
                missing += 1
        return len(chosen), garbled, missing
    probe = ExtractedTable(list(columns), [list(row) for row in chosen[:1]])
    date_at = set(_date_column_indexes(probe))
    attribute_at = {index for index, _field in _repeat_columns(probe) if _field in {"make", "color"}}
    page_at = {index for index, name in enumerate(columns) if _header_is_printed_page(name) or _is_notes_header(name)}
    garbled = 0
    missing = 0
    for row in chosen:
        cells = list(row) + [""] * max(0, len(columns) - len(row))
        body = [cell.strip() for index, cell in enumerate(cells) if index not in page_at and str(cell).strip()]
        if not body:
            continue
        for index in date_at:
            text = str(cells[index]).strip() if index < len(cells) else ""
            if not text:
                missing += 1
                continue
            iso = _normalize_date_cell(text)
            if _DATE_VALUE.fullmatch(_clean_cell(text)):
                continue
            looks_reversed = bool(re.search(r"\d", text) and re.search(r"[-/'’]", text))
            if _DATE_VALUE.fullmatch(iso) or looks_reversed:
                garbled += 1
        for index in attribute_at:
            if index >= len(cells) or not str(cells[index]).strip():
                missing += 1
    return len(chosen), garbled, missing


def page_audit(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
) -> list[dict[str, int]]:
    """
    One entry per page that has rows.

    extracted is every row kept for that page. plated is a vehicle whose
    رقم اللوحة is a real plate. separators are headers and gap lines, and
    they are not counted as vehicles.
    """
    fields = _vehicle_field_at(columns)
    plate_at = fields.get("plate")
    pages = sorted({int(page) for page in page_numbers})
    report: list[dict[str, int]] = []
    for page in pages:
        chosen = [row for row, number in zip(rows, page_numbers) if int(number) == page]
        plated = 0
        separators = 0
        for row in chosen:
            cells = list(row)
            if _nondata_kind(cells, columns):
                separators += 1
                continue
            if plate_at is None or plate_at >= len(cells):
                continue
            token = _normalize_plate_cell(cells[plate_at])
            if _PLATE_VALUE.fullmatch(token or "") and not _YEAR_VALUE.fullmatch(token or ""):
                plated += 1
        report.append(
            {
                "page": page,
                "extracted": len(chosen),
                "plated": plated,
                "separators": separators,
            }
        )
    return report


def _finish_page_tables(tables: list[ExtractedTable], page_number: int, kind: str) -> list[ExtractedTable]:
    """Clean one page the same way a full extract cleans each page, then stop."""
    ready: list[ExtractedTable] = []
    for table in tables:
        _isolate_nondata_rows(table)
        _smooth_spellings(table)
        _clean_table(table)
        _validate_patterns(table)
        if not table.rows:
            continue
        table.sources = [int(page_number)] * len(table.rows)
        ready.append(table)
    for table in ready:
        _apply_bilingual_headers(table)
        _validate_patterns(table)
        _add_page_column(table)
        _apply_manufacturer_dictionary(table)
        table.kind = kind or ""
    return sanitize_tables(ready)


def extract_page_tables(
    pdf_path: str | Path,
    page_number: int,
    password: str = "",
    mode: str = "auto",
    kind: str = "register",
    strict: bool = False,
) -> list[ExtractedTable]:
    """
    Read one page again.

    mode is auto, grid, words, or ocr. Other pages are not opened for extraction.
    """
    chosen = (mode or "auto").strip().lower()
    if chosen not in _PAGE_MODES:
        chosen = "auto"
    source = Path(pdf_path)
    document = open_document(source, password)
    plumber = _open_pdfplumber(source, password)
    strict_token = _STRICT.set(bool(strict))
    try:
        if page_number < 1 or page_number > document.page_count:
            raise ConversionError(f"Page {page_number} is not in this PDF.")
        index = int(page_number) - 1
        page = document.load_page(index)

        def attempt(reader, *args) -> list[ExtractedTable]:
            try:
                found = reader(*args) or []
            except Exception:
                return []
            return [table for table in found if getattr(table, "rows", None)]

        if chosen == "grid":
            found = attempt(_vehicle_tables, page, index, plumber) or attempt(_tables_from_rules, page, index)
        elif chosen == "words":
            found = attempt(_tables_from_words, page, index) or attempt(_tables_from_aligned_words, page, index)
        elif chosen == "ocr":
            found = attempt(_tables_from_scan, page, index, {})
        else:
            found = attempt(_tables_on_page, page, index, {}, plumber, None, kind or "register")
        for table in found or []:
            _assemble_row_dates(page, table)
        finished = _finish_page_tables(list(found or []), int(page_number), kind or "")
        for table in finished:
            recover_arabic_columns(document, table)
            _repair_text_columns(table)
        return finished
    finally:
        _STRICT.reset(strict_token)
        if plumber is not None:
            try:
                plumber.close()
            except Exception:
                pass
        close_document(document)


def read_pdf_workbook(
    pdf_path: str | Path,
    password: str = "",
    progress: Callable[[str], None] | None = None,
    strict: bool = False,
) -> dict:
    """Extract every page and return rows tagged with the page they came from."""
    global _LAST_TIMINGS
    source = Path(pdf_path)
    clock = ExtractionTimings()
    token = _TIMINGS.set(clock)
    try:
        return _read_pdf_workbook(source, password, progress, bool(strict))
    finally:
        _LAST_TIMINGS = clock.rows()
        _TIMINGS.reset(token)


def _read_pdf_workbook(
    source: Path,
    password: str,
    progress: Callable[[str], None] | None,
    strict: bool,
) -> dict:
    tables = extract_tables(source, password, progress, strict=bool(strict))
    figure_started = time.perf_counter()
    figures = extract_figures(source, password)
    _remember_time("figures", time.perf_counter() - figure_started)
    document = open_document(source, password)
    try:
        page_count = int(document.page_count)
        kind = next((table.kind for table in tables if table.kind), "")
        if not kind:
            kind = detect_document_type(document)
        stated = _stated_vehicle_total(document) if kind == "register" else None
    finally:
        close_document(document)
    columns, rows, pages = records_from_tables(tables)
    detected: dict[int, int] = {}
    statements: list[str] = []
    for table in tables:
        detected.update(dict(table.detected_pages or {}))
        for line in table.statements or []:
            if line not in statements:
                statements.append(line)
    printed_total, printed_groups = _printed_statements(statements)
    if stated is None:
        stated = printed_total
    checked = time.perf_counter()
    validation = validate_register(columns, rows, pages, stated, printed_groups, detected)
    _remember_time("validation", time.perf_counter() - checked)
    clock = _TIMINGS.get()
    return {
        "columns": columns,
        "rows": rows,
        "page_numbers": pages,
        "page_count": page_count,
        "kind": kind,
        "figures": figures,
        "stated_total": stated,
        "strict": bool(strict),
        "validation": validation,
        "detected_pages": detected,
        "printed_groups": printed_groups,
        "statements": statements,
        "confidence": score_register(columns, rows),
        "baselines": {key: value for table in tables for key, value in (table.baseline_at or {}).items()},
        "column_spans": {key: value for table in tables for key, value in (table.column_spans or {}).items()},
        "provenance": {key: value for table in tables for key, value in (table.provenance or {}).items()},
        "suggestions": {key: value for table in tables for key, value in (table.suggestions or {}).items()},
        "raw_readings": {key: value for table in tables for key, value in (table.raw_readings or {}).items()},
        "timings": clock.rows() if clock is not None else [],
        "column_warning": next((table.column_warning for table in tables if table.column_warning), ""),
        "ocr_attempted": any(bool(getattr(table, "ocr_attempted", False)) for table in tables),
    }


def write_edited_excel(
    columns: list[str],
    rows: list[list[str]],
    page_numbers: list[int],
    output_path: str | Path,
    kind: str = "",
    figures: list[ExtractedFigure] | None = None,
    strict: bool = False,
    validation: list[dict[str, str]] | None = None,
    direction: str = "auto",
    provenance: dict[str, str] | None = None,
    suggestions: dict[str, str] | None = None,
    raw_readings: dict[str, str] | None = None,
    developer: bool = False,
    audit: list | None = None,
) -> None:
    """Write the edited rows. Sheet direction, filter, and borders stay in place."""
    target = Path(output_path)
    if not rows:
        write_excel([], target, figures, preserve_edits=True, strict=bool(strict), validation=validation, direction=direction, developer=developer)
        return
    table = ExtractedTable(
        [str(name) for name in columns],
        [["" if cell is None else str(cell) for cell in row] for row in rows],
        None,
        [int(page) for page in page_numbers],
    )
    table.kind = kind or ""
    table.direction = direction or "auto"
    table.provenance = dict(provenance or {})
    table.suggestions = dict(suggestions or {})
    table.raw_readings = dict(raw_readings or {})
    table.audit = [dict(item) for item in (audit or []) if isinstance(item, dict)]
    write_excel(
        [table],
        target,
        figures,
        preserve_edits=True,
        strict=bool(strict),
        validation=validation,
        direction=direction,
        developer=developer,
    )


def convert_pdf_to_excel(
    pdf_path: str | Path,
    excel_path: str | Path,
    password: str = "",
    progress: Callable[[str], None] | None = None,
    strict: bool = False,
    developer: bool = False,
) -> ExcelWriteCount:
    """Convert one PDF. Returns the number of data rows, and how many figures were stored."""
    source = Path(pdf_path).expanduser()
    target = Path(excel_path).expanduser()
    if target.suffix.lower() != ".xlsx":
        target = target.with_suffix(".xlsx")
    if not source.is_file():
        raise ConversionError("Choose a PDF file that exists.")
    if source.stat().st_size <= 0:
        raise ConversionError("The PDF is empty.")
    try:
        with source.open("rb") as handle:
            sample = handle.read(1024)
    except OSError as exc:
        raise ConversionError("The PDF could not be read.") from exc
    if b"%PDF" not in sample:
        raise ConversionError("That file is not a PDF.")
    if not target.parent.is_dir():
        raise ConversionError(f"The destination folder does not exist: {target.parent}")
    try:
        same = source.resolve() == target.resolve()
    except OSError:
        same = False
    if same:
        raise ConversionError("Choose a different name for the Excel file.")

    global _LAST_TIMINGS
    clock = ExtractionTimings()
    token = _TIMINGS.set(clock)
    try:
        tables = extract_tables(source, password, progress, strict=bool(strict))
        figure_started = time.perf_counter()
        figures = extract_figures(source, password)
        _remember_time("figures", time.perf_counter() - figure_started)
        if not tables and not figures:
            raise ConversionError(
                "No table was found. A digital PDF needs text in rows and columns. A scan is read from the picture on this PC."
            )
        write_started = time.perf_counter()
        marked = len(clock.events)
        write_excel(tables, target, figures, strict=bool(strict), developer=bool(developer))
        accounted = sum(
            float(item["seconds"])
            for item in clock.events[marked:]
            if item.get("name") == "validation"
        )
        _remember_time("build_excel", max(0.0, time.perf_counter() - write_started - accounted))
        _LAST_TIMINGS = clock.rows()
        return ExcelWriteCount(sum(len(table.rows) for table in tables), len(figures))
    finally:
        _TIMINGS.reset(token)


# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------

def _run_window() -> None:
    """The desktop window. Review clicks stay on the in-memory table."""
    from desktop_review import run_desktop

    run_desktop()


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        _run_window()
        return 0
    if len(args) not in (1, 2):
        print("Usage: python pdf_to_excel.py input.pdf [output.xlsx]", file=sys.stderr)
        return 2
    source = Path(args[0])
    target = Path(args[1]) if len(args) == 2 else source.with_suffix(".xlsx")
    try:
        count = convert_pdf_to_excel(source, target)
    except ConversionError as exc:
        print(exc, file=sys.stderr)
        return 1
    figures = getattr(count, "figures", 0)
    extra = f" and {figures} figures" if figures else ""
    print(f"Saved {int(count)} rows{extra} to {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
