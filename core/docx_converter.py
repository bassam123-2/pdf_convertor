# -*- coding: utf-8 -*-
"""
Convert an Arabic PDF into a Word document without scrambling the letters.

The PDF is read on this computer with PyMuPDF and pdfplumber. Nothing is
uploaded. Words are written into Word as ordinary Unicode Arabic. They are
not reshaped, and python-bidi is not used. Word applies right-to-left once,
through the paragraph bidi flag.

The letter scrambling in academic PDFs comes from the file itself. The
glyphs are stored from the left of the line toward the right, and a lam-alef
ligature (لا لأ لإ لآ) is stored as two characters on one point, lam then
alef. Reversing that pair with the rest of the line turns لا into ال.
This script turns the visual line back into logical Unicode and keeps each
ligature as one unit, so Word receives العلاقة rather than العالقة.

    pip install pymupdf pdfplumber python-docx
    python arabic_pdf_to_docx.py input.pdf output.docx

A password-protected file:

    python arabic_pdf_to_docx.py input.pdf output.docx --password secret
"""

from __future__ import annotations

from core.safe_tables import find_tables as _safe_find_tables

import argparse
import contextvars
import copy
import os
import functools
import re
import sys
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Inches, Pt, RGBColor

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.arabic_fixer import _Atom, _arabic_count, _is_arabic, _logical_text, plain_arabic
from core.layout_check import DISCREPANCY_LIMIT, layout_discrepancy, page_is_ruled
from core.pdf_loader import PAGE_BATCH as _PAGE_BATCH
from core.pdf_loader import close_document, iter_page_batches, open_document


# Word stores the complex-script face separately from the Latin face.
# Traditional Arabic draws the Arabic. Arial draws digits and English words
# that sit inside the same sentence.
_ARABIC_FONT = "Traditional Arabic"
_LATIN_FONT = "Arial"
# One family for every run of the Word file, Arabic and Latin alike. Empty
# keeps the PDF's own families. Set for the length of one write.
_FONT = contextvars.ContextVar("pdf_tool_font", default="")


def output_fonts() -> tuple[str, str]:
    """The Latin and the Arabic family the Word file is written in."""
    chosen = (_FONT.get() or "").strip()
    if chosen:
        return chosen, chosen
    return _LATIN_FONT, _ARABIC_FONT

# One PDF page is written as one Word section. These match the margins in
# _start_section, in points, so a height estimate uses the same frame Word will.
_MARGIN_X_PT = 1.6 / 2.54 * 72.0
_MARGIN_Y_PT = 1.5 / 2.54 * 72.0
# Running logos and page numbers live in this fraction of the sheet.
# They are written into the section header and footer, not the body.
_CHROME_BAND = 0.05
# The running footer sits lower than the header band. Anything that starts
# here belongs to the footer, not to a body text box.
_FOOTER_ZONE = 0.12
# Pages are converted in groups of PAGE_BATCH. A repeated header is judged
# inside the group, and a failed page does not keep the group's pictures alive.

class ConversionError(Exception):
    """The PDF could not be opened or the Word file could not be saved."""


# Set for the RTL repair pass. It drops a glyph that sits inside another glyph.
_PAGE_REPAIR = contextvars.ContextVar("pdf_tool_page_repair", default="")
# rtl, ltr, or auto. Auto decides from the letters in each paragraph.
_DIRECTION = contextvars.ContextVar("pdf_tool_direction", default="rtl")

_A4_WIDTH = 595.28
_A4_HEIGHT = 841.89


@dataclass
class WordOptions:
    """Choices made before a conversion. Defaults match the command-line reader."""

    preserve_first_page: bool = False
    lock_header: bool = True
    lock_footer: bool = True
    page_spec: str = ""
    merge_paragraphs: bool = True
    tables_as_images: bool = False
    fit_a4: bool = False
    direction: str = "rtl"
    # "faithful": the Word page looks and prints like the PDF page. Text boxes
    # sit at their PDF positions over one background picture of the graphics.
    # "editable": the flowing layout, with paragraphs, tables, and a Word
    # header and footer. New fields go last so existing calls keep working.
    mode: str = "faithful"
    # Optional per-page override: {1: "faithful", 2: "editable", ...}.
    # The page classifier fills the plan; these keys win when present.
    page_strategies: dict | None = None
    # Optional progress(current, total, message) and cancel() -> bool hooks.
    progress: object | None = None
    cancel: object | None = None
    # Parallel page readers (1 = sequential). Each worker opens its own PDF.
    workers: int = 1
    # The family every run is written in (Arabic and Latin). Times New Roman
    # ships with Windows and Office and has full Arabic, so the file looks the
    # same on every machine. "" keeps each PDF font's own family.
    font: str = "Times New Roman"
    # Editable mode only. True writes ordinary flowing paragraphs (columns as
    # borderless table cells); False pins every paragraph in a floating frame
    # at its PDF position, as earlier versions did.
    flow: bool = True


def _is_faithful(options: "WordOptions") -> bool:
    """True unless the caller asked for the flowing layout."""
    mode = str(getattr(options, "mode", "faithful") or "faithful").strip().lower()
    if mode in {"faithful", "exact"}:
        return True
    if mode in {"editable", "flow", "flowing"}:
        return False
    raise ConversionError(f"Unknown mode: {mode}. Use faithful or editable.")


@dataclass
class WordPage:
    """One PDF page after it has been read, ready to write or replace."""

    number: int
    blocks: list
    width: float
    height: float
    kind: str = "digital"
    show_header: bool = True
    show_footer: bool = True
    cover: bool = False
    undo: dict | None = None
    # Faithful mode keeps the page here: a background picture and text lines.
    # None means the page is written by the flowing layout.
    faithful: object | None = None
    # True when an editable page was read the earlier way and is written in
    # positioned frames instead of flowing paragraphs.
    frames: bool = False


@dataclass
class WordBuild:
    """The pages kept in memory so one page can be replaced without rereading the others."""

    pages: list[WordPage]
    page_count: int
    counts: dict
    chrome: "_PageChrome"
    options: WordOptions


@dataclass
class TextBlock:
    """One paragraph or heading, still in PDF coordinates."""

    kind: str  # "paragraph" or "heading"
    bbox: tuple[float, float, float, float]
    text: str
    size: float = 12.0
    bold: bool = False
    align: str = "right"
    gap_before: float = 0.0
    role: str = ""
    toc_page: str = ""
    fill: str = ""
    color: str = ""


@dataclass
class TableBlock:
    """One grid. Rows are top to bottom. Columns are left to right, as printed."""

    bbox: tuple[float, float, float, float]
    rows: list[list[str]] = field(default_factory=list)
    fills: list[list[str]] = field(default_factory=list)
    widths: list[float] = field(default_factory=list)
    rules: bool = True
    fit_scale: float = 1.0
    # Cells that one cell covers, per row: (first column, last column).
    spans: list[list[tuple[int, int]]] = field(default_factory=list)
    # Pictures inside cells: (row, column, image bytes, width pt, height pt).
    images: list[tuple] = field(default_factory=list)
    # Styled runs per cell, {(row, column): [(text, bold, italic, colour)]}; "\n" breaks the line.
    styles: dict = field(default_factory=dict)
    centered: bool = False
    # How the printed cell sets its text, {(row, column): (horizontal, vertical)}.
    aligns: dict = field(default_factory=dict)


@dataclass
class ImageBlock:
    """A figure painted on the page. The full-page border is not one of these."""

    bbox: tuple[float, float, float, float]
    data: bytes
    width: float
    height: float
    align: str = "center"
    gap_before: float = 0.0
    fit: str = ""
    graphic: bool = False


@dataclass
class _Line:
    """One visual line after its letters have been put back in logical order."""

    text: str
    bbox: tuple[float, float, float, float]
    size: float
    bold: bool
    origin_y: float
    color: str = ""


def _hex_color(value) -> str:
    """An sRGB color as six hex digits, or nothing for missing ink."""
    if not value:
        return ""
    try:
        channels = [int(round(float(channel) * 255)) for channel in list(value)[:3]]
    except (TypeError, ValueError):
        return ""
    if len(channels) < 3:
        return ""
    return "".join(f"{channel:02X}" for channel in channels)


def _upright(trace: dict) -> bool:
    """True when the span runs left to right along a level baseline."""
    direction = trace.get("dir") or (1.0, 0.0)
    try:
        return float(direction[0]) > 0.99 and abs(float(direction[1])) < 0.05
    except (TypeError, ValueError, IndexError):
        return True


def _gap_spaces(atoms: list[_Atom]) -> list[_Atom]:
    """
    Spaces the PDF left out.

    Some programs paint each word as its own span and move the pen between
    words, so the line has no space glyph. Read as it is stored, that line is
    one long word. A space is added where the gap between two letters is
    wider than half of the font's own space, measured on top of the gap the
    line normally leaves between letters. Letter-spaced text therefore keeps
    its words, and a line that already has its space glyphs is not changed.

    A font that paints its own spaces on this page has less need of a guess.
    In such a font only a gap as wide as a whole space becomes one, so a pen
    move inside an address or a number is not read as a word break.
    """
    found: list[tuple[_Atom, _Atom, float, float, float, bool]] = []
    painted: dict[str, int] = {}
    for band in _cluster_baselines(atoms):
        ordered = sorted(band, key=lambda atom: (atom.x0, atom.origin_x))
        letters = [
            atom
            for atom in ordered
            if atom.text.strip()
            and atom.x1 - atom.x0 > 0.05
            and not all(unicodedata.category(char) == "Mn" for char in atom.text)
        ]
        if len(letters) < 2:
            continue
        gaps: list[float] = []
        edge = letters[0].x1
        for atom in letters[1:]:
            gaps.append(max(0.0, atom.x0 - edge) / max(atom.size, 1.0))
            edge = max(edge, atom.x1)
        typical = sorted(gaps)[len(gaps) // 2]
        counted = {id(atom) for atom in letters}
        edge = None
        previous: _Atom | None = None
        centers: list[float] = []
        for atom in ordered:
            if not atom.text.strip():
                if atom.x1 > atom.x0:
                    centers.append((atom.x0 + atom.x1) / 2.0)
                continue
            if id(atom) not in counted:
                continue
            if previous is not None and edge is not None:
                # A painted space between the two letters: nothing is added.
                # A space painted on top of a letter is a duplicate that the
                # line reader drops, so it does not stand in for a real gap.
                if any(edge + 0.4 < center < atom.x0 - 0.4 for center in centers):
                    painted[previous.font] = painted.get(previous.font, 0) + 1
                else:
                    size = max(min(previous.size, atom.size), 1.0)
                    space = atom.space or previous.space
                    if not 0.15 * size <= space <= 0.6 * size:
                        space = 0.25 * size
                    gap = atom.x0 - edge
                    if gap >= typical * size + max(0.5 * space, 0.12 * size):
                        found.append((previous, atom, edge, gap, space, gap >= typical * size + space))
            edge = atom.x1 if edge is None else max(edge, atom.x1)
            previous = atom
            centers = []
    guessed: dict[str, int] = {}
    for previous, _atom, _edge, _gap, _space, _whole in found:
        guessed[previous.font] = guessed.get(previous.font, 0) + 1
    added: list[_Atom] = []
    for previous, atom, edge, gap, space, whole in found:
        if not whole and painted.get(previous.font, 0) > guessed[previous.font]:
            continue
        odd = previous.level % 2 == 1 or atom.level % 2 == 1
        # The space is as wide as a painted one and stands where a painted one
        # would: after the word it follows. In a line that reads from the
        # right, that word is the right one. A wide gap stays open, so a
        # column gutter is still seen.
        host = atom if odd else previous
        reach = min(gap, space)
        left = atom.x0 - reach if odd else edge
        added.append(
            _Atom(
                " ",
                left,
                min(previous.y0, atom.y0),
                left + reach,
                max(previous.y1, atom.y1),
                left,
                host.origin_y,
                1 if odd else 0,
                host.size,
                host.bold,
                host.color,
                host.font,
                host.italic,
                host.flags,
                host.space,
            )
        )
    return added


def _atoms_on_page(page: fitz.Page, upright_only: bool = False) -> list[_Atom]:
    """
    Read every glyph in the order it is painted.

    A lam-alef ligature comes back as two Unicode characters with the same
    origin: lam first, then the alef, and the alef has no glyph of its own.
    Those two stay one unit. Splitting them and reversing the line is what
    turned العلاقة into العالقة and الخلاصة into خالصة.

    Each glyph keeps the PDF's font name, weight, slant, and color. A space
    the PDF made by moving the pen is added after the glyphs. With
    `upright_only`, text set at an angle is left out.
    """
    atoms: list[_Atom] = []
    try:
        traces = page.get_texttrace() or []
    except Exception:
        return atoms
    traces = _shape_repaired(page, traces)
    for trace in traces:
        if int(trace.get("type") or 0) != 0:
            continue
        if upright_only and not _upright(trace):
            continue
        level = int(trace.get("bidi_lvl") or 0)
        try:
            size = float(trace.get("size") or 12)
        except (TypeError, ValueError):
            size = 12.0
        flags = int(trace.get("flags") or 0)
        bold = bool(flags & 16)
        italic = bool(flags & 2)
        font = str(trace.get("font") or "")
        try:
            space = float(trace.get("spacewidth") or 0.0)
        except (TypeError, ValueError):
            space = 0.0
        color = _hex_color(trace.get("color"))
        current: _Atom | None = None
        for raw in trace.get("chars") or []:
            origin = raw[2]
            box = raw[3]
            origin_x = float(origin[0])
            glyph_id = int(raw[1])
            if (
                current is not None
                and glyph_id < 0
                and abs(origin_x - current.origin_x) <= 0.8
            ):
                current.text += chr(raw[0])
                current.x0 = min(current.x0, float(box[0]))
                current.y0 = min(current.y0, float(box[1]))
                current.x1 = max(current.x1, float(box[2]))
                current.y1 = max(current.y1, float(box[3]))
                continue
            current = _Atom(
                chr(raw[0]),
                float(box[0]),
                float(box[1]),
                float(box[2]),
                float(box[3]),
                origin_x,
                float(origin[1]),
                level,
                size,
                bold,
                color,
                font,
                italic,
                flags,
                space,
            )
            atoms.append(current)
    _repair_ligatures(page, atoms)
    atoms.extend(_gap_spaces(atoms))
    return atoms


def _shape_repaired(page: fitz.Page, traces: list) -> list:
    """Traces whose wrongly mapped glyphs are read from their outlines (see font_repair)."""
    if os.environ.get("PDF_TOOL_SHAPE_REPAIR", "1") == "0":
        return traces
    try:
        from core.font_repair import repaired_traces

        return repaired_traces(page, traces)
    except Exception:
        return traces


def _repair_ligatures(page: fitz.Page, atoms: list[_Atom]) -> None:
    """Lam-alef glyphs whose Unicode value the PDF stored wrongly."""
    from core import ligature_repair

    if not any(len(atom.text) == 1 and ord(atom.text) > 0x36F and not ("\u0600" <= atom.text <= "\u06FF")
               and (ligature_repair.truncated_ligature(atom.text) or ligature_repair._FOREIGN.match(atom.text))
               for atom in atoms):
        return
    try:
        known = ligature_repair.document_vocabulary(page)
    except Exception:
        known = None
    ligature_repair.repair_atoms(atoms, known)


def _cluster_baselines(atoms: list[_Atom]) -> list[list[_Atom]]:
    """Stack glyphs that share a baseline. A 16-point gap is the next line."""
    import unicodedata

    def is_mark(atom) -> bool:
        text = atom.text or ""
        return bool(text) and all(unicodedata.category(ch) == "Mn" for ch in text)

    # A raised mark (shadda, a tanween) can sit on its own baseline. Lines are
    # built from the letters, against the line's first letter; marks then
    # join the nearest line.
    ordered = sorted(atoms, key=lambda atom: (atom.origin_y, atom.origin_x))
    letters = [atom for atom in ordered if not is_mark(atom)]
    marks = [atom for atom in ordered if is_mark(atom)]
    if not letters:
        letters, marks = ordered, []
    lines: list[list[_Atom]] = []
    for atom in letters:
        limit = max(2.4, atom.size * 0.28)
        if lines and abs(atom.origin_y - lines[-1][0].origin_y) <= limit:
            lines[-1].append(atom)
        else:
            lines.append([atom])
    for atom in marks:
        limit = max(2.4, atom.size * 0.28) * 2.0
        best = min(lines, key=lambda line: abs(atom.origin_y - line[0].origin_y))
        if abs(atom.origin_y - best[0].origin_y) <= limit:
            best.append(atom)
        else:
            lines.append([atom])
    if marks:
        lines.sort(key=lambda line: line[0].origin_y)
    return lines


def _split_gutter(atoms: list[_Atom], page_width: float) -> list[list[_Atom]]:
    """
    Split a baseline where a wide empty gap falls between two columns.

    The cut is driven by the gap itself, not by where it sits on the page. A
    narrow sidebar puts the gutter left of the page centre; requiring the hole
    near mid-page left the phone and email glued to the body beside them.
    Two ems, or a gap that spans a clear share of the page, is enough.
    """
    ordered = sorted(atoms, key=lambda atom: atom.x0)
    if not ordered:
        return []
    parts: list[list[_Atom]] = [[ordered[0]]]
    for atom in ordered[1:]:
        right_edge = max(item.x1 for item in parts[-1])
        gap = atom.x0 - right_edge
        size = max(
            max((item.size for item in parts[-1] if item.text.strip()), default=1.0),
            float(atom.size or 1.0),
            1.0,
        )
        wide = gap >= max(22.0, 2.0 * size)
        sidebar = gap >= max(14.0, 0.9 * size) and gap >= page_width * 0.08
        if wide or sidebar:
            parts.append([atom])
        else:
            parts[-1].append(atom)
    # A contents line is a title on the right and a page number on the left,
    # with the dotted leader left unread. Put the number back on that title.
    def part_text(part: list[_Atom]) -> str:
        return "".join(atom.text for atom in part)

    def number_part(part: list[_Atom]) -> bool:
        text = part_text(part)
        digits = re.sub(r"\D", "", text)
        letters = sum(1 for char in text if char.isalpha() or _is_arabic(char))
        return bool(digits) and letters == 0 and len(digits) <= 3

    def is_leader(part: list[_Atom]) -> bool:
        text = part_text(part).strip()
        return len(text) >= 2 and all(char in ".·…•‧" or char.isspace() for char in text)

    leaders = [part for part in parts if is_leader(part)]
    rest = [part for part in parts if part not in leaders]
    numbers = [part for part in rest if number_part(part)]
    titles = [part for part in rest if _arabic_count(part_text(part)) >= 2]
    if len(numbers) == 1 and titles and len(rest) == len(numbers) + len(titles):
        host = max(titles, key=lambda part: max(atom.x1 for atom in part))
        host.extend(numbers[0])
        rest = [part for part in rest if part is not numbers[0]]
        return rest
    return parts


def _box(raw) -> tuple[float, float, float, float]:
    return (float(raw[0]), float(raw[1]), float(raw[2]), float(raw[3]))


def _overlap_ratio(inner, outer) -> float:
    """How much of `inner` sits inside `outer`. 1 means the inner box is covered."""
    width = max(0.0, min(inner[2], outer[2]) - max(inner[0], outer[0]))
    height = max(0.0, min(inner[3], outer[3]) - max(inner[1], outer[1]))
    area = max(0.0, inner[2] - inner[0]) * max(0.0, inner[3] - inner[1])
    if area <= 0:
        return 0.0
    return (width * height) / area


def _mark_headings(blocks: list[TextBlock]) -> None:
    """A short line set clearly larger than the body is a heading."""
    sizes = [block.size for block in blocks if block.text]
    if not sizes:
        return
    body = sorted(sizes)[len(sizes) // 2]
    for block in blocks:
        if block.text.startswith(("•", "●", "▪", "◦", "-", "–")):
            continue
        if block.size >= max(15.0, body * 1.15) and len(block.text) <= 70:
            block.kind = "heading"
            block.bold = True


def _paint_banners(page: fitz.Page, blocks: list[TextBlock]) -> None:
    """A short heading that sits on a colored bar keeps that color and the bar's width."""
    paints = _drawing_paints(page)
    if not paints:
        return
    page_width = float(page.rect.width)
    for block in blocks:
        if block.role == "toc" or len(block.text) > 80:
            continue
        center_x = (block.bbox[0] + block.bbox[2]) / 2.0
        center_y = (block.bbox[1] + block.bbox[3]) / 2.0
        cover = [
            paint
            for paint in paints
            if paint[0] - 1.5 <= center_x <= paint[2] + 1.5 and paint[1] - 1.5 <= center_y <= paint[3] + 1.5
        ]
        if not cover:
            continue
        left, top, right, bottom, fill, _area = min(cover, key=lambda paint: paint[5])
        block.fill = fill
        bar_width = right - left
        bar_height = bottom - top
        text_width = block.bbox[2] - block.bbox[0]
        if bar_height <= 40 and text_width + 12 < bar_width < page_width * 0.96:
            block.bbox = (
                max(0.0, left),
                min(block.bbox[1], top),
                min(page_width, right),
                max(block.bbox[3], bottom),
            )


def _clean_rows(raw_rows) -> list[list[str]]:
    """Pad every row to the same width and keep the cell text in stored order."""
    rows = []
    width = 0
    for raw in raw_rows or []:
        if raw is None:
            continue
        cells = [plain_arabic("" if cell is None else str(cell).replace("\n", " ")) for cell in raw]
        width = max(width, len(cells))
        rows.append(cells)
    if width == 0:
        return []
    rectangle = [row + [""] * (width - len(row)) for row in rows]
    return [row for row in rectangle if any(cell for cell in row)]


def _usable_table(rows: list[list[str]]) -> list[list[str]] | None:
    """
    Drop a one-line rule that is only the page number, and drop columns
    that have no text. A real table has at least two rows and two filled columns.
    """
    if len(rows) < 2 or not rows or len(rows[0]) < 2:
        return None
    width = max(len(row) for row in rows)
    keep = [index for index in range(width) if any(index < len(row) and row[index].strip() for row in rows)]
    if len(keep) < 2:
        return None
    trimmed = [[row[index] if index < len(row) else "" for index in keep] for row in rows]
    filled = [cell.strip() for row in trimmed for cell in row if cell.strip()]
    if filled and all(cell.isdigit() and len(cell) <= 4 for cell in filled):
        return None
    return trimmed


def _atoms_in(atoms: list[_Atom], box: tuple[float, float, float, float]) -> list[_Atom]:
    """Letters whose centers sit in the rectangle."""
    inside = []
    for atom in atoms:
        x = (atom.x0 + atom.x1) / 2.0
        y = (atom.y0 + atom.y1) / 2.0
        if box[0] - 1 <= x <= box[2] + 1 and box[1] - 1 <= y <= box[3] + 1:
            inside.append(atom)
    return inside


def _read_atoms(atoms: list[_Atom]) -> str:
    """Logical text of one cell, one visual line at a time."""
    parts = []
    for line in _cluster_baselines(atoms):
        text = _logical_text(line)
        if text:
            parts.append(text)
    return plain_arabic(" ".join(parts))


def _text_outside(atoms: list[_Atom], box, covered: list) -> str:
    """Letters in `box` that are not already inside a detected cell."""
    free = []
    for atom in _atoms_in(atoms, box):
        center = ((atom.x0 + atom.x1) / 2.0, (atom.y0 + atom.y1) / 2.0)
        if any(
            other[0] <= center[0] <= other[2] and other[1] <= center[1] <= other[3]
            for other in covered
        ):
            continue
        free.append(atom)
    return _read_atoms(free)


def _text_inside(atoms: list[_Atom], box: tuple[float, float, float, float]) -> str:
    """
    Read the letters whose centers sit in one cell.

    The cell uses the same logical reading as a paragraph, including the
    lam-alef ligature, so the cell is not reversed a second time in Word.
    """
    return _read_atoms(_atoms_in(atoms, box))


def _grid_text(atoms: list[_Atom], grid: list[list]) -> list[list[str]]:
    """Replace a reader's cell strings with the letters actually inside each box."""
    rows = []
    for row_boxes in grid:
        row = []
        for box in row_boxes:
            if not box:
                row.append("")
                continue
            row.append(_text_inside(atoms, _box(box)))
        rows.append(row)
    return _clean_rows(rows)


def _union_box(left, right) -> tuple[float, float, float, float]:
    return (
        min(left[0], right[0]),
        min(left[1], right[1]),
        max(left[2], right[2]),
        max(left[3], right[3]),
    )


def _drawing_paints(page: fitz.Page) -> list[tuple]:
    """Colored rectangles on the page, skipping white paper."""
    paints = []
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        return paints
    from core.table_rebuild import is_box

    for drawing in drawings:
        fill = drawing.get("fill")
        rect = drawing.get("rect")
        if not fill or rect is None or not is_box(drawing):
            continue
        try:
            red, green, blue = (int(round(float(channel) * 255)) for channel in list(fill)[:3])
        except (TypeError, ValueError):
            continue
        if red > 242 and green > 242 and blue > 242:
            continue
        area = max(0.0, float(rect.x1) - float(rect.x0)) * max(0.0, float(rect.y1) - float(rect.y0))
        if area < 20:
            continue
        if min(float(rect.width), float(rect.height)) < 3.0:
            continue  # a rule or a cell border, not a bar behind text
        if float(rect.width) < 14.0 or area < 200:
            continue  # a stroke of a letter drawn as a box, not a bar
        paints.append(
            (
                float(rect.x0),
                float(rect.y0),
                float(rect.x1),
                float(rect.y1),
                f"{red:02X}{green:02X}{blue:02X}",
                area,
            )
        )
    return paints


def _paint_at(paints: list[tuple], x: float, y: float) -> str:
    """The smallest colored rectangle that covers this point."""
    color = ""
    best = 1e18
    for left, top, right, bottom, paint, area in paints:
        if left - 1.5 <= x <= right + 1.5 and top - 1.5 <= y <= bottom + 1.5 and area < best:
            color = paint
            best = area
    return color


def _tables_from_pymupdf(page: fitz.Page, chars: list[_Atom]) -> list[TableBlock]:
    """
    Ruled tables.

    One visual table is sometimes reported as two grids that share the title
    column. Those grids are joined, then each cell is placed by its rectangle
    so a wrapped line stays in its own column.
    """
    try:
        found = _safe_find_tables(page, vertical_strategy="lines", horizontal_strategy="lines")
        tables = list(getattr(found, "tables", None) or found)
    except Exception:
        return []
    groups: list[dict] = []
    for table in tables:
        try:
            boxes = [cell for cell in table.cells if cell]
            bbox = _box(table.bbox)
        except Exception:
            continue
        if not boxes:
            continue
        host = None
        for group in groups:
            if _overlap_ratio(bbox, group["bbox"]) >= 0.15 or _overlap_ratio(group["bbox"], bbox) >= 0.15:
                host = group
                break
        if host is None:
            groups.append({"boxes": list(boxes), "bbox": bbox})
        else:
            host["boxes"].extend(boxes)
            host["bbox"] = _union_box(host["bbox"], bbox)
    blocks: list[TableBlock] = []
    for group in groups:
        finished = _matrix_from_boxes(group["boxes"], chars, page)
        if finished:
            rows, fills = finished
            blocks.append(TableBlock(group["bbox"], rows, fills))
    return blocks


def _cell_paragraphs(atoms: list[_Atom], box: tuple[float, float, float, float]) -> str:
    """A cell's text, with a line break where the printed cell starts a new paragraph."""
    return _cell_runs(atoms, box)[0]


def _style_of(atoms: list[_Atom]) -> tuple[bool, bool, str]:
    """The weight, slant and colour most of these letters are printed in."""
    letters = [atom for atom in atoms if atom.text.strip()] or atoms
    if not letters:
        return False, False, ""
    bold = sum(1 for atom in letters if atom.bold) * 2 > len(letters)
    italic = sum(1 for atom in letters if getattr(atom, "italic", False)) * 2 > len(letters)
    colours: dict[str, int] = {}
    for atom in letters:
        colours[atom.color or ""] = colours.get(atom.color or "", 0) + 1
    return bold, italic, max(colours, key=colours.get)


def _line_runs(line: list[_Atom], text: str) -> list[tuple[str, bool, bool, str]]:
    """
    One printed line as runs of one weight, slant and colour.

    A left-to-right line is cut into words at the gaps between letters, so a
    bold name followed by plain text keeps both. A line with Arabic keeps
    one style, the one most of its letters use: its letters are reordered
    and cannot be matched to words one by one.
    """
    if _arabic_count(text):
        bold, italic, colour = _style_of(line)
        return [(text, bold, italic, colour)]
    ordered = sorted(line, key=lambda atom: atom.x0)
    words: list[list[_Atom]] = []
    for atom in ordered:
        if not atom.text.strip():
            if words and words[-1]:
                words.append([])
            continue
        gap = atom.x0 - words[-1][-1].x1 if words and words[-1] else 0.0
        if words and words[-1] and gap > max(1.0, atom.size * 0.2):
            words.append([])
        if not words:
            words.append([])
        words[-1].append(atom)
    words = [word for word in words if word]
    joined = " ".join("".join(atom.text for atom in word) for word in words)
    if re.sub(r"\s+", "", joined) != re.sub(r"\s+", "", text):
        bold, italic, colour = _style_of(line)
        return [(text, bold, italic, colour)]
    runs: list[list] = []
    for word in words:
        piece = "".join(atom.text for atom in word)
        style = _style_of(word)
        if runs and tuple(runs[-1][1:]) == style:
            runs[-1][0] += " " + piece
        else:
            if runs:
                runs[-1][0] += " "
            runs.append([piece, *style])
    return [tuple(run) for run in runs]


def _cell_runs(atoms: list[_Atom], box: tuple[float, float, float, float]) -> tuple[str, list[tuple[str, bool, bool, str]]]:
    """
    A cell's text and its styled runs; a run of "\n" is a line break.

    Lines that wrap are joined. A new paragraph begins after a line that
    stops well short of the cell's right side, after a wide gap, or at a
    line that starts with a bullet or a dash.
    """
    lines = _cluster_baselines(_atoms_in(atoms, box))
    width = max(1.0, box[2] - box[0])
    parts: list[str] = []
    runs: list[tuple[str, bool, bool, str]] = []
    previous = None
    for line in lines:
        text = plain_arabic(_logical_text(line)).strip()
        if not text:
            continue
        left = min(atom.x0 for atom in line)
        right = max(atom.x1 for atom in line)
        top = min(atom.y0 for atom in line)
        bottom = max(atom.y1 for atom in line)
        if previous is not None:
            p_left, p_right, p_top, p_bottom = previous
            pitch = max(1.0, p_bottom - p_top)
            rtl = _arabic_count(text) > len(re.findall(r"[A-Za-z]", text))
            short = (p_left - box[0] > width * 0.25) if rtl else (box[2] - p_right > width * 0.25)
            # A bold title over plain text is two paragraphs, however long.
            restyled = _style_of(line)[:2] != previous_style[:2]
            new = short or restyled or top - p_bottom > pitch * 0.9 or re.match(r"[-–•▪●·◦]\s", text)
            parts.append("\n" if new else " ")
            runs.append(("\n", False, False, "") if new else (" ", False, False, ""))
        parts.append(text)
        runs.extend(_line_runs(line, text))
        previous = (left, right, top, bottom)
        previous_style = _style_of(line)
    return "".join(parts).strip(), runs


def _cell_alignment(atoms: list[_Atom], box: tuple[float, float, float, float]) -> tuple[str, str]:
    """
    Where a printed cell sets its text: ("center" or "", "center" or "top").

    Lines whose left and right gaps match are centred; a block whose space
    above and below match sits in the middle of the cell.
    """
    lines = [line for line in _cluster_baselines(_atoms_in(atoms, box)) if any(a.text.strip() for a in line)]
    if not lines:
        return "", "center"
    width = max(1.0, box[2] - box[0])
    height = max(1.0, box[3] - box[1])
    centred = 0
    for line in lines:
        letters = [a for a in line if a.text.strip()]
        left = min(a.x0 for a in letters) - box[0]
        right = box[2] - max(a.x1 for a in letters)
        if left > 6 and right > 6 and abs(left - right) <= max(6.0, width * 0.08):
            centred += 1
    top = min(a.y0 for line in lines for a in line) - box[1]
    bottom = box[3] - max(a.y1 for line in lines for a in line)
    horizontal = "center" if centred * 2 > len(lines) and (len(lines) > 1 or centred) else ""
    vertical = "center" if abs(top - bottom) <= max(6.0, height * 0.2) else "top"
    return horizontal, vertical


def _rule_tables(page: fitz.Page, atoms: list[_Atom]) -> list[TableBlock]:
    """Tables whose thin drawn borders close every row (see table_rebuild.rule_grids)."""
    try:
        from core.table_rebuild import rule_grids

        grids = rule_grids(page)
    except Exception:
        return []
    paints = _drawing_paints(page)
    blocks: list[TableBlock] = []
    from core.table_rebuild import cell_pictures, split_open_spans

    for columns, rows_y, cells in grids:
        cells = [split_open_spans(row, columns, rows_y[r], rows_y[r + 1], atoms) for r, row in enumerate(cells)]
        count = len(columns) - 1
        rows: list[list[str]] = []
        fills: list[list[str]] = []
        styles: dict = {}
        aligns: dict = {}
        for r, row_cells in enumerate(cells):
            texts = [""] * count
            colours = [""] * count
            for first, last in row_cells:
                box = (columns[first], rows_y[r], columns[last + 1], rows_y[r + 1])
                texts[first], styles[(r, first)] = _cell_runs(atoms, box)
                aligns[(r, first)] = _cell_alignment(atoms, box)
                colour = _paint_at(paints, (box[0] + box[2]) / 2.0, box[1] + min(2.5, (box[3] - box[1]) / 2.0)) if paints else ""
                for k in range(first, last + 1):
                    colours[k] = colour
            rows.append(texts)
            fills.append(colours)
        if sum(1 for row in rows for cell in row if cell) < 3:
            continue
        bbox = (columns[0], rows_y[0], columns[-1], rows_y[-1])
        blocks.append(TableBlock(
            bbox,
            rows,
            fills,
            widths=[columns[k + 1] - columns[k] for k in range(count)],
            rules=True,
            spans=[list(row) for row in cells],
            images=cell_pictures(page, columns, rows_y, cells),
            styles=styles,
            aligns=aligns,
        ))
    return blocks


def _rows_from_boxes(cells) -> list[list]:
    """
    Group cell boxes into visual rows.

    PyMuPDF's cell list is stored down each column. Slicing it as rows puts
    three lines of the left column into one Word row. Rows are the boxes that
    share a top edge, ordered from the left of the page to the right.
    """
    boxes = [cell for cell in cells if cell]
    if not boxes:
        return []
    boxes.sort(key=lambda box: (float(box[1]), float(box[0])))
    rows: list[list] = []
    for box in boxes:
        top = float(box[1])
        if rows and abs(top - float(rows[-1][0][1])) <= 2.0:
            rows[-1].append(box)
        else:
            rows.append([box])
    for row in rows:
        row.sort(key=lambda box: float(box[0]))
    return rows


def _cluster_edges(values: list[float], tolerance: float) -> list[float]:
    """Collapse edges that are the same column or the same row within a few points."""
    groups: list[list[float]] = []
    for value in sorted(values):
        if groups and value - groups[-1][-1] <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [sum(group) / len(group) for group in groups]


def _merge_continuations(rows: list[list[str]]) -> list[list[str]]:
    """
    Join a wrapped line back onto the cell above it.

    A results table stores الفرضية and الأولى as two ruled rows. The second
    row has no digits and only continues cells that already have text, so it
    belongs to the row above. A real data row carries a number, so it stays.
    """
    if not rows:
        return rows
    merged = [list(rows[0])]
    for row in rows[1:]:
        previous = merged[-1]
        filled = [index for index, cell in enumerate(row) if cell.strip()]
        previous_filled = [index for index, cell in enumerate(previous) if cell.strip()]
        continues = (
            filled
            and len(filled) < len(previous_filled)
            and all(previous[index].strip() for index in filled)
            and all(not any(char.isdigit() for char in row[index]) for index in filled)
        )
        if continues:
            for index in filled:
                previous[index] = f"{previous[index].rstrip()} {row[index].strip()}".strip()
            continue
        merged.append(list(row))
    return merged


def _finish_grid(
    rows: list[list[str]], fills: list[list[str]]
) -> tuple[list[list[str]], list[list[str]]] | None:
    """Drop empty columns and empty rows, and keep each cell's color with it."""
    if _usable_table(rows) is None:
        return None
    width = max(len(row) for row in rows)
    keep = [
        index
        for index in range(width)
        if any(index < len(row) and row[index].strip() for row in rows)
    ]
    kept_rows: list[list[str]] = []
    kept_fills: list[list[str]] = []
    for row_index, row in enumerate(rows):
        fill = fills[row_index] if row_index < len(fills) else []
        cells = [row[index] if index < len(row) else "" for index in keep]
        paints = [fill[index] if index < len(fill) else "" for index in keep]
        if any(cell.strip() for cell in cells):
            kept_rows.append(cells)
            kept_fills.append(paints)
    if _usable_table(kept_rows) is None:
        return None
    return kept_rows, kept_fills


def _matrix_from_boxes(
    boxes, atoms: list[_Atom], page: fitz.Page | None = None
) -> tuple[list[list[str]], list[list[str]]] | None:
    """
    Build a grid from cell rectangles, using each rectangle's place on the page.

    A short row used to be padded from the left, so a word that sits in the
    right-hand column landed in the first Word column. Edges are clustered
    instead, and a box that repeats one already kept is ignored.
    """
    wide = [box for box in boxes if box and (box[2] - box[0]) >= 12 and (box[3] - box[1]) >= 6]
    if len(wide) < 4:
        return None
    kept = []
    for box in sorted(wide, key=lambda item: (item[2] - item[0]) * (item[3] - item[1]), reverse=True):
        if any(_overlap_ratio(box, other) >= 0.6 for other in kept):
            continue
        kept.append(box)
    columns = _cluster_edges([box[0] for box in kept], 8)
    row_lines = _cluster_edges([box[1] for box in kept], 3)
    if len(columns) < 2 or len(row_lines) < 2:
        return None
    # The right and bottom edges close the last column and the last row.
    # Text that sits in that rectangle is read even when the detector left
    # the rectangle without its own cell box.
    right = max(box[2] for box in kept)
    bottom = max(box[3] for box in kept)
    covered = [_box(box) for box in kept]
    grid = [[""] * len(columns) for _ in row_lines]
    fills = [[""] * len(columns) for _ in row_lines]
    paints = _drawing_paints(page) if page is not None else []
    for box in kept:
        column = min(range(len(columns)), key=lambda index: abs(box[0] - columns[index]))
        row = min(range(len(row_lines)), key=lambda index: abs(box[1] - row_lines[index]))
        text = _text_inside(atoms, _box(box))
        if text:
            grid[row][column] = f"{grid[row][column]} {text}".strip() if grid[row][column] else text
        if paints and not fills[row][column]:
            fills[row][column] = _paint_at(
                paints, (float(box[0]) + float(box[2])) / 2.0, (float(box[1]) + float(box[3])) / 2.0
            )
    # A ruled cell can contain a line the detector never boxed. Read that
    # rectangle only where the box pass left the cell empty, and skip letters
    # that already belong to a box so a tall cell is not copied into the row below.
    bounds_x = columns + [right]
    bounds_y = row_lines + [bottom]
    for row_index in range(len(row_lines)):
        for column_index in range(len(columns)):
            rectangle = (
                bounds_x[column_index],
                bounds_y[row_index],
                bounds_x[column_index + 1],
                bounds_y[row_index + 1],
            )
            extra = _text_outside(atoms, rectangle, covered)
            if extra:
                current = grid[row_index][column_index]
                grid[row_index][column_index] = f"{current} {extra}".strip() if current else extra
            if paints and not fills[row_index][column_index]:
                fills[row_index][column_index] = _paint_at(
                    paints,
                    (rectangle[0] + rectangle[2]) / 2.0,
                    (rectangle[1] + rectangle[3]) / 2.0,
                )
    merged_rows = _merge_continuations(grid)
    # Replay the same continuation test so fill rows stay aligned with text rows.
    cursor = 0
    merged_fills = [list(fills[0])] if fills else []
    for index, row in enumerate(grid[1:], start=1):
        if cursor >= len(merged_rows):
            break
        previous = merged_rows[cursor]
        filled = [cell_index for cell_index, cell in enumerate(row) if cell.strip()]
        previous_filled = [cell_index for cell_index, cell in enumerate(previous) if cell.strip()]
        continues = (
            filled
            and len(filled) < len(previous_filled)
            and all(previous[cell_index].strip() for cell_index in filled)
            and all(not any(char.isdigit() for char in row[cell_index]) for cell_index in filled)
        )
        if continues:
            continue
        cursor += 1
        merged_fills.append(list(fills[index]) if index < len(fills) else [""] * len(columns))
    if len(merged_fills) != len(merged_rows):
        merged_fills = [[""] * len(row) for row in merged_rows]
    return _finish_grid(merged_rows, merged_fills)


def _tables_from_pdfplumber(plumber_page, chars: list[dict]) -> list[TableBlock]:
    """Ruled tables pdfplumber can see. Cell text is not reordered."""
    if plumber_page is None:
        return []
    try:
        found = plumber_page.find_tables() or []
    except Exception:
        return []
    blocks: list[TableBlock] = []
    for table in found:
        try:
            rows = _clean_rows(table.extract())
            bbox = _box(table.bbox)
        except Exception:
            continue
        flat = []
        for row in getattr(table, "rows", None) or []:
            flat.extend(getattr(row, "cells", None) or [])
        grid = _rows_from_boxes(flat)
        if grid and chars:
            rows = _grid_text(chars, grid) or rows
        rows = _usable_table(rows)
        if rows:
            blocks.append(TableBlock(bbox, rows))
    return blocks


def _line_from_atoms(atoms: list[_Atom]) -> _Line | None:
    """Build one logical line from the glyphs that share its baseline and column."""
    text = _logical_text(atoms)
    if not text:
        return None
    letters = [atom for atom in atoms if atom.text.strip()]
    if not letters:
        return None
    bbox = (
        min(atom.x0 for atom in atoms),
        min(atom.y0 for atom in atoms),
        max(atom.x1 for atom in atoms),
        max(atom.y1 for atom in atoms),
    )
    colors: dict[str, int] = {}
    for atom in letters:
        ink = getattr(atom, "color", "") or ""
        if not ink or ink.upper() in {"000000", "FFFFFF"}:
            continue
        colors[ink] = colors.get(ink, 0) + max(1, len(atom.text.strip()))
    ink = max(colors, key=colors.get) if colors else ""
    return _Line(
        text,
        bbox,
        max(atom.size for atom in letters),
        sum(1 for atom in letters if atom.bold) * 2 >= len(letters),
        sum(atom.origin_y for atom in letters) / len(letters),
        ink,
    )


def _atom_center(atom: _Atom) -> tuple[float, float]:
    return (atom.x0 + atom.x1) / 2.0, (atom.y0 + atom.y1) / 2.0


def _center_inside(atom: _Atom, box: tuple[float, float, float, float]) -> bool:
    x, y = _atom_center(atom)
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


def _lines_on_page(
    page: fitz.Page,
    atoms: list[_Atom],
    tables: list[TableBlock],
    figures: list[ImageBlock] | None = None,
) -> list[_Line]:
    """Lines of body text. Glyphs inside a table or a figure stay there."""
    figures = figures or []
    kept = []
    for atom in atoms:
        box = (atom.x0, atom.y0, atom.x1, atom.y1)
        if any(_overlap_ratio(box, table.bbox) >= 0.5 for table in tables):
            continue
        # Words painted on a chart or diagram stay inside that picture.
        if any(_center_inside(atom, figure.bbox) for figure in figures):
            continue
        kept.append(atom)
    if _PAGE_REPAIR.get() == "rtl":
        kept = _separate_overlapping_atoms(kept)
    lines: list[_Line] = []
    for band in _cluster_baselines(kept):
        for part in _split_gutter(band, page.rect.width):
            line = _line_from_atoms(part)
            if line is not None:
                lines.append(line)
    return lines


def _is_page_number(line: _Line, page_height: float) -> bool:
    """A bare page number in the top or bottom margin is furniture, not a paragraph."""
    if not re.fullmatch(r"\d{1,4}", line.text.strip()):
        return False
    return line.bbox[1] > page_height * 0.88 or line.bbox[3] < page_height * 0.08


def _gutter(lines: list[_Line], page: fitz.Page) -> float | None:
    """The empty gap between two columns, or nothing when the page has one column."""
    body = [line for line in lines if 50 < line.origin_y < page.rect.height - 45]
    if len(body) < 4:
        return None
    # Use left and right edges, not centres: a narrow sidebar's centre can sit
    # outside the middle third of the page and still leave a clear gutter.
    edges = sorted((line.bbox[0], line.bbox[2]) for line in body)
    best_gap = 0.0
    gutter = None
    limit = float(page.rect.width)
    rights = sorted({edge[1] for edge in edges})
    lefts = sorted({edge[0] for edge in edges})
    for right in rights:
        for left in lefts:
            if left <= right:
                continue
            gap = left - right
            middle = (left + right) / 2.0
            if gap > best_gap and gap >= max(28.0, limit * 0.06) and limit * 0.08 < middle < limit * 0.92:
                best_gap = gap
                gutter = middle
    if best_gap < 28:
        return None
    return gutter


def _column_segments(
    lines: list[_Line], page: fitz.Page, tables: list[TableBlock] | None = None
) -> list[list[_Line]]:
    """
    Read the right column from top to bottom, then the left column.

    A line that stretches across both columns stays at its own height, so a
    title is not pushed to the bottom of the page. A table that crosses the
    gutter is the same kind of break: the left column above the table is read
    before the right column below it.
    """
    usable = [line for line in lines if not _is_page_number(line, page.rect.height)]
    gutter = _gutter(usable, page)
    if gutter is None:
        return [sorted(usable, key=lambda line: line.origin_y)] if usable else []
    wide: list[_Line] = []
    right: list[_Line] = []
    left: list[_Line] = []
    for line in usable:
        width = line.bbox[2] - line.bbox[0]
        center = (line.bbox[0] + line.bbox[2]) / 2.0
        if width >= page.rect.width * 0.62:
            wide.append(line)
        elif center >= gutter:
            right.append(line)
        else:
            left.append(line)
    markers: list[tuple[float, _Line | None]] = [(line.origin_y, line) for line in wide]
    for table in tables or []:
        if table.bbox[2] - table.bbox[0] >= page.rect.width * 0.45:
            markers.append((table.bbox[1], None))
    markers.sort(key=lambda item: item[0])
    cuts = [-1e9] + [y for y, _line in markers] + [1e9]
    segments: list[list[_Line]] = []
    for index in range(len(cuts) - 1):
        start, end = cuts[index], cuts[index + 1]
        for column in (right, left):
            chunk = [line for line in column if start < line.origin_y <= end]
            chunk.sort(key=lambda line: line.origin_y)
            if chunk:
                segments.append(chunk)
        if index < len(markers) and markers[index][1] is not None:
            segments.append([markers[index][1]])
    return segments


def _page_number_text(text: str) -> str:
    """The page number when a line is only digits and a dotted leader."""
    compact = re.sub(r"[.\s\u00b7\u2022\u2024\u2026‧·]+", "", text)
    if re.fullmatch(r"\d{1,3}", compact):
        return compact
    return ""


def _attach_page_numbers(lines: list[_Line]) -> list[_Line]:
    """
    Put a page number that sits to the left of a title back on that title.

    A contents row is one baseline in the PDF, but the number and the title
    are often read as two lines. Joining them keeps the entry in one paragraph.
    """
    used: set[int] = set()
    for line in lines:
        if _arabic_count(line.text) < 2 or _page_number_text(line.text):
            continue
        center = (line.bbox[1] + line.bbox[3]) / 2.0
        best: int | None = None
        best_dy = 8.0
        for index, other in enumerate(lines):
            if index in used or other is line:
                continue
            number = _page_number_text(other.text)
            if not number or (other.bbox[2] - other.bbox[0]) > 180:
                continue
            other_center = (other.bbox[1] + other.bbox[3]) / 2.0
            dy = abs(center - other_center)
            if dy <= best_dy and other.bbox[0] < line.bbox[0]:
                best = index
                best_dy = dy
        if best is None:
            continue
        used.add(best)
        other = lines[best]
        line.text = f"{line.text.rstrip()} {_page_number_text(other.text)}"
        line.bbox = (
            min(line.bbox[0], other.bbox[0]),
            min(line.bbox[1], other.bbox[1]),
            max(line.bbox[2], other.bbox[2]),
            max(line.bbox[3], other.bbox[3]),
        )
    return [line for index, line in enumerate(lines) if index not in used and line.text.strip()]


def _toc_entry(text: str) -> tuple[str, str] | None:
    """A contents line is a title, a dotted leader, and a page number."""
    if not text:
        return None
    cleaned = re.sub(r"(?:\.\s*){2,}|[.\u00b7\u2022\u2024\u2026‧]{2,}", " ", text)
    # The dotted leader makes a contents line long; its words are short.
    if len(cleaned) > 120:
        return None
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .\t")
    match = re.match(r"^(?P<title>.+?)\s+(?P<page>\d{1,3})$", cleaned)
    if match is None:
        match = re.match(r"^(?P<page>\d{1,3})\s+(?P<title>.+)$", cleaned)
    if match is None:
        return None
    title = match.group("title").strip(" .\t")
    page = match.group("page")
    if _arabic_count(title) < 2 or len(title) > 90:
        return None
    return title, page


def _mark_contents(blocks: list[TextBlock]) -> None:
    """Keep each contents entry as one line with its page number."""
    hits = [block for block in blocks if _toc_entry(block.text)]
    if len(hits) < 4:
        return
    named = any(
        "المحتويات" in block.text or "الفهرس" in block.text for block in blocks
    )
    led = sum(1 for block in hits if re.search(r"(?:\.\s*){3,}|[.\u00b7\u2026]{3,}", block.text))
    if not named and led < 3 and len(hits) < 5:
        return
    for block in hits:
        title, number = _toc_entry(block.text) or ("", "")
        if not title:
            continue
        block.text = title
        block.toc_page = number
        block.role = "toc"
        block.kind = "paragraph"


def _clean_flow_text(text: str) -> str:
    """One paragraph, with blank lines and repeated spaces removed."""
    cleaned = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    cleaned = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", cleaned)
    cleaned = re.sub(r"[ \t]{2,}", " ", cleaned)
    return cleaned.strip()


def _same_line_edge(previous: _Line, line: _Line) -> bool:
    """Two lines share a margin, so they are the same alignment."""
    return abs(previous.bbox[2] - line.bbox[2]) <= 24 or abs(previous.bbox[0] - line.bbox[0]) <= 24


def _continues_paragraph(previous: _Line, line: _Line) -> bool:
    """
    True when this line is the next line of the same paragraph.

    Same size, same alignment, and ordinary line spacing. A bullet, a contents
    entry, or a wider gap starts a new paragraph.
    """
    if abs(line.size - previous.size) > 1.6 or line.bold != previous.bold:
        return False
    if line.text.startswith(("•", "●", "▪", "◦", "فرضية", "الفرض")):
        return False
    if _toc_entry(previous.text) is not None or _toc_entry(line.text) is not None:
        return False
    if not _same_line_edge(previous, line):
        return False
    gap = line.bbox[1] - previous.bbox[3]
    baseline = line.origin_y - previous.origin_y
    if baseline <= 0 and gap < -2:
        return False
    close_baseline = 0 < baseline <= max(previous.size * 1.65, 14.0)
    close_gap = -1.0 <= gap <= max(previous.size * 0.85, 7.0)
    return close_baseline or close_gap


def _separate_overlapping_atoms(atoms: list[_Atom]) -> list[_Atom]:
    """
    Drop a glyph that sits inside another glyph and repeats its letters.

    Glyphs are filed in a grid of cells by line and by the x range they
    cover, so each glyph is only compared with glyphs on its own line that
    it touches. The result is the same as comparing every pair.
    """
    ordered = sorted(atoms, key=lambda atom: (round(atom.origin_y, 1), -atom.x1))
    kept: list[_Atom] = []
    dropped: set[int] = set()
    cell_x, cell_y = 12.0, 6.0
    grid: dict[tuple[int, int], list[_Atom]] = {}

    def cells(atom: _Atom):
        gy = int(atom.origin_y // cell_y)
        for gx in range(int(atom.x0 // cell_x), int(atom.x1 // cell_x) + 1):
            yield gx, gy

    for atom in ordered:
        if not atom.text.strip():
            # A space is kept, and filed too: a letter drawn over it removes it.
            atom._order = len(kept)
            kept.append(atom)
            for key in cells(atom):
                grid.setdefault(key, []).append(atom)
            continue
        swallowed = False
        seen: set[int] = set()
        reach = max(2.4, atom.size * 0.35)
        rows = range(int((atom.origin_y - reach) // cell_y), int((atom.origin_y + reach) // cell_y) + 1)
        columns = range(int(atom.x0 // cell_x), int(atom.x1 // cell_x) + 1)
        candidates = []
        for gy in rows:
            for gx in columns:
                for other in grid.get((gx, gy), ()):
                    if id(other) in seen or id(other) in dropped:
                        continue
                    seen.add(id(other))
                    candidates.append(other)
        # Earlier glyphs first, as in a plain scan of the kept list.
        candidates.sort(key=lambda other: other._order)
        for other in candidates:
            if abs(atom.origin_y - other.origin_y) > reach:
                continue
            overlap = min(atom.x1, other.x1) - max(atom.x0, other.x0)
            narrow = max(0.1, min(atom.x1 - atom.x0, other.x1 - other.x0))
            if overlap < narrow * 0.65:
                continue
            shorter = atom.text.strip()
            longer = other.text.strip()
            if shorter not in longer and longer not in shorter:
                continue
            if len(shorter) > len(longer):
                dropped.add(id(other))
                continue
            swallowed = True
            break
        if not swallowed:
            atom._order = len(kept)
            kept.append(atom)
            for key in cells(atom):
                grid.setdefault(key, []).append(atom)
    return [atom for atom in kept if id(atom) not in dropped]


def _join_segment(lines: list[_Line]) -> list[TextBlock]:
    """
    Join consecutive lines of one paragraph.

    Lines that share alignment, size, and line spacing become one paragraph.
    A short last line is kept with the lines above it. A bullet, a contents
    row, or a clear gap stays a separate paragraph.
    """
    if not lines:
        return []
    from core.html_layout import MERGE_PARAGRAPHS

    if not MERGE_PARAGRAPHS.get():
        blocks: list[TextBlock] = []
        for line in lines:
            text = _clean_flow_text(plain_arabic(line.text))
            if not text or not re.search(r"[0-9A-Za-z\u0600-\u06FF]", text):
                continue
            blocks.append(
                TextBlock("paragraph", line.bbox, text, line.size, line.bold, color=line.color)
            )
        return blocks
    groups: list[list[_Line]] = [[lines[0]]]
    for line in lines[1:]:
        previous = groups[-1][-1]
        if _continues_paragraph(previous, line):
            groups[-1].append(line)
        else:
            groups.append([line])
    blocks: list[TextBlock] = []
    for group in groups:
        text = _clean_flow_text(plain_arabic(" ".join(line.text for line in group if line.text)))
        if not text or not re.search(r"[0-9A-Za-z\u0600-\u06FF]", text):
            continue
        boxes = [line.bbox for line in group]
        bbox = (
            min(box[0] for box in boxes),
            min(box[1] for box in boxes),
            max(box[2] for box in boxes),
            max(box[3] for box in boxes),
        )
        colors = [line.color for line in group if line.color]
        ink = colors[0] if colors and all(color == colors[0] for color in colors) else ""
        blocks.append(
            TextBlock(
                "paragraph",
                bbox,
                text,
                max(line.size for line in group),
                any(line.bold for line in group),
                color=ink,
            )
        )
    return blocks


def _paragraphs_from_atoms(
    page: fitz.Page,
    atoms: list[_Atom],
    tables: list[TableBlock],
    figures: list[ImageBlock] | None = None,
) -> list[TextBlock]:
    """Body paragraphs in Arabic reading order: right column, then left column."""
    blocks: list[TextBlock] = []
    lines = _attach_page_numbers(_lines_on_page(page, atoms, tables, figures))
    for segment in _column_segments(lines, page, tables):
        blocks.extend(_join_segment(segment))
    _mark_contents(blocks)
    return blocks


def _page_strings(blocks: list) -> list[str]:
    """Every paragraph and cell string on the page, for finding a repeated stamp."""
    strings = []
    for block in blocks:
        if isinstance(block, TableBlock):
            strings.extend(cell for row in block.rows for cell in row)
        elif isinstance(block, TextBlock):
            strings.append(block.text)
    return strings


def _opening_text(blocks: list) -> str:
    """The title, when this page has one. A stamp is never the title."""
    for block in blocks:
        if isinstance(block, TextBlock) and block.text:
            return block.text
    return ""


def _strip_stamps(blocks: list) -> None:
    """
    Remove a phrase that is stamped into sentences and table cells.

    A real name that repeats, such as the market in the title, is kept.
    A term such as "نظام المعلومات" is also kept, because its words appear
    in other sentences and not only inside one repeated phrase.
    """
    strings = [text for text in _page_strings(blocks) if text]
    if len(strings) < 3:
        return
    blob = "\n".join(strings)
    opening = _opening_text(blocks)
    words_by_string = [text.split() for text in strings]
    banned: list[str] = []
    seen: set[str] = set()
    for size in (5, 4, 3):
        for words in words_by_string:
            for index in range(len(words) - size + 1):
                phrase = " ".join(words[index : index + size])
                if phrase in seen:
                    continue
                seen.add(phrase)
                if phrase in opening:
                    continue
                count = blob.count(phrase)
                if count < 3 or sum(_arabic_count(part) for part in phrase.split()) < 8:
                    continue
                # Every copy of these words is this stamp, including a copy
                # that was printed against the previous word with no space.
                if any(blob.count(part) != count for part in phrase.split()):
                    continue
                banned.append(phrase)
    if not banned:
        return

    def cleaned(text: str) -> str:
        for phrase in banned:
            text = text.replace(phrase, " ")
        return plain_arabic(text)

    for block in blocks:
        if isinstance(block, TableBlock):
            block.rows = [[cleaned(cell) for cell in row] for row in block.rows]
        elif isinstance(block, TextBlock):
            block.text = cleaned(block.text)


def _encoding_broken(page: fitz.Page, atoms: list[_Atom]) -> bool:
    """
    True when the font's Unicode map mixes Latin letters into Arabic words.

    A citation such as (Smith, 2020) is a separate Latin word and does not
    count. A word such as تتaون does: the glyph was drawn as an Arabic letter
    and stored as a Latin one. Those pages are read from the rendered picture.
    """
    if len(atoms) < 40:
        return False
    lines = _lines_on_page(page, atoms, [])
    tokens = re.findall(r"\S+", "\n".join(line.text for line in lines))
    arabic_tokens = 0
    mixed = 0
    for token in tokens:
        has_arabic = any(_is_arabic(char) for char in token)
        has_latin = any(char.isascii() and char.isalpha() for char in token)
        if not has_arabic:
            continue
        arabic_tokens += 1
        if has_latin:
            mixed += 1
    if arabic_tokens < 12:
        return False
    return mixed >= 5 and mixed / arabic_tokens >= 0.06


def _recognize_png(data: bytes):
    """Read one rendered page with Windows Arabic OCR."""
    try:
        from winrt.runtime import ApartmentType, init_apartment
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import BitmapAlphaMode, BitmapDecoder, BitmapPixelFormat
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    except Exception:
        return None
    try:
        init_apartment(ApartmentType.MTA)
    except Exception:
        pass
    engine = OcrEngine.try_create_from_language(Language("ar-SA"))
    if engine is None:
        return None
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(data)
    writer.store_async().get()
    writer.detach_stream()
    stream.seek(0)
    decoder = BitmapDecoder.create_async(stream).get()
    bitmap = decoder.get_software_bitmap_async(
        BitmapPixelFormat.BGRA8, BitmapAlphaMode.PREMULTIPLIED
    ).get()
    return engine.recognize_async(bitmap).get()


def _trace_sizes(page: fitz.Page) -> list[tuple[float, float, str]]:
    """Font size and ink color of each text span, keyed by its baseline."""
    rows: list[tuple[float, float, str]] = []
    try:
        traces = page.get_texttrace() or []
    except Exception:
        return rows
    for trace in traces:
        if int(trace.get("type") or 0) != 0:
            continue
        try:
            size = float(trace.get("size") or 0)
        except (TypeError, ValueError):
            continue
        chars = trace.get("chars") or []
        if not chars or not 7.0 <= size <= 36.0:
            continue
        origin = chars[0][2]
        try:
            baseline = float(origin[1])
        except (TypeError, ValueError, IndexError):
            continue
        rows.append((baseline, size, _hex_color(trace.get("color"))))
    return rows


def _size_near(rows: list[tuple[float, float, str]], y: float, fallback: float) -> float:
    """The PDF font size on this line. A scanned page has no spans and keeps the word height."""
    if not rows:
        return fallback
    baseline, size, _color = min(rows, key=lambda row: abs(row[0] - y))
    if abs(baseline - y) > 10:
        return fallback
    return size


def _color_near(rows: list[tuple[float, float, str]], y: float) -> str:
    """Ink that is not black, taken from the span on this line."""
    near = [row for row in rows if abs(row[0] - y) <= 8 and row[2] and row[2].upper() not in {"000000", "FFFFFF"}]
    if not near:
        return ""
    counts: dict[str, int] = {}
    for _baseline, _size, color in near:
        counts[color] = counts.get(color, 0) + 1
    return max(counts, key=counts.get)


# One letter inside quotes. The picture reader drops these; the text layer still has them.
_QUOTED_LETTER = re.compile(r'["«“]\s*([ء-يA-Za-z])\s*["»”]')
_SEPARATED_NUMBER = re.compile(r"\d{1,4}(?:[/\-.:]\d{1,4})+")


def _quote_only(text: str) -> bool:
    stripped = text.strip()
    return bool(stripped) and all(char in "\"'«»“”" for char in stripped)


def _restore_quoted_letters(page: fitz.Page, atoms: list[_Atom]) -> list[_Atom]:
    """
    Put back a quoted single letter the picture reader dropped.

    Used only after the picture reader, and only when the text layer still
    has the letter as its own glyph inside quotes. The glyph's own box is
    kept, so the letter is not parked at a guessed offset. A page with no
    text layer is left as the picture reader returned it.
    """
    if not atoms:
        return atoms
    recovered: list[tuple] = []
    try:
        bands = _cluster_baselines(_atoms_on_page(page))
    except Exception:
        return atoms
    for band in bands:
        line = _line_from_atoms(band)
        if line is None:
            continue
        for match in _QUOTED_LETTER.finditer(line.text):
            letter = match.group(1)
            hosts = [atom for atom in band if plain_arabic(atom.text) == letter]
            if not hosts:
                continue
            host = hosts[0]
            recovered.append((host.origin_y, letter, host))
    for baseline, letter, host in recovered:
        line_atoms = [
            atom
            for atom in atoms
            if abs(atom.origin_y - baseline) <= 14 or abs(((atom.y0 + atom.y1) / 2.0) - baseline) <= 14
        ]
        blob = "".join(atom.text for atom in line_atoms)
        if f'"{letter}"' in blob or f"«{letter}»" in blob:
            continue

        def bare(atom: _Atom) -> bool:
            stripped = plain_arabic(atom.text).strip(" \"'«»“”")
            return stripped == letter and len(stripped) == 1

        kept = next((atom for atom in line_atoms if bare(atom)), None)
        if kept is not None:
            kept.text = f'"{letter}" '
            atoms[:] = [
                atom
                for atom in atoms
                if atom is kept or atom not in line_atoms or not (bare(atom) or _quote_only(atom.text))
            ]
            continue
        glued = None
        for atom in line_atoms:
            stripped = atom.text.rstrip()
            if stripped.endswith(letter) and _arabic_count(stripped) > 1:
                glued = atom
                break
        token = f'"{letter}" '
        if glued is not None:
            stripped = glued.text.rstrip()
            glued.text = stripped[: -len(letter)].rstrip() + " "
        atoms[:] = [atom for atom in atoms if atom not in line_atoms or not _quote_only(atom.text)]
        others = [
            atom
            for atom in band
            if atom is not host and (_arabic_count(atom.text) or any(char.isascii() and char.isalpha() for char in atom.text))
        ]
        side = "left"
        if others:
            middle = (min(atom.x0 for atom in others) + max(atom.x1 for atom in others)) / 2.0
            side = "left" if (host.x0 + host.x1) / 2.0 < middle else "right"
        words = [atom for atom in line_atoms if len(plain_arabic(atom.text).strip()) > 1]
        word = glued if glued is not None else (min(words, key=lambda atom: abs(atom.origin_y - baseline)) if words else None)
        if word is None:
            continue
        gap = max(2.0, float(word.size) * 0.25)
        x = word.x0 - gap if side == "left" else word.x1 + gap
        atoms.append(
            _Atom(
                token,
                x,
                word.y0,
                x + gap,
                word.y1,
                x,
                word.origin_y,
                1 if _arabic_count(letter) else 0,
                word.size,
                False,
            )
        )
    return atoms


def _glyph_touches(left: _Atom, right: _Atom, gap: float = 0.8) -> bool:
    dx = max(0.0, max(left.x0, right.x0) - min(left.x1, right.x1))
    dy = max(0.0, max(left.y0, right.y0) - min(left.y1, right.y1))
    return dx <= gap and dy <= gap


def _layer_numbers(atoms: list[_Atom]) -> list[tuple[float, str, _Atom]]:
    """
    Separated numbers from a text layer, with one cmap-glued digit removed.

    A broken map stores an Arabic letter as a digit and paints it against
    the number, so 1440/06/29 is read as 61440/06/29. The extra digit is the
    one that touches an Arabic glyph and leaves a group longer than four
    digits. A real number such as 2021 is not separated, so it is not cut.
    """
    found: list[tuple[float, str, _Atom]] = []
    for band in _cluster_baselines(atoms):
        ordered = sorted(band, key=lambda atom: atom.x0)
        run: list[_Atom] = []

        def flush() -> None:
            if not run:
                return
            text = "".join(atom.text.strip() for atom in run)
            kept = list(run)
            edge = kept[0]
            if (
                re.fullmatch(r"\d", edge.text.strip())
                and any(_glyph_touches(edge, other) and any(_is_arabic(char) for char in other.text) for other in ordered)
                and re.fullmatch(r"\d{5,}(?:[/\-.:]\d{1,4})+", text)
            ):
                kept = kept[1:]
                text = "".join(atom.text.strip() for atom in kept)
            edge = kept[-1] if kept else None
            if (
                edge is not None
                and re.fullmatch(r"\d", edge.text.strip())
                and any(_glyph_touches(edge, other) and any(_is_arabic(char) for char in other.text) for other in ordered)
                and re.fullmatch(r"(?:\d{1,4}[/\-.:])+\d{5,}", text)
            ):
                kept = kept[:-1]
                text = "".join(atom.text.strip() for atom in kept)
            if kept and _SEPARATED_NUMBER.fullmatch(text):
                span = _Atom(
                    text,
                    min(atom.x0 for atom in kept),
                    min(atom.y0 for atom in kept),
                    max(atom.x1 for atom in kept),
                    max(atom.y1 for atom in kept),
                    min(atom.origin_x for atom in kept),
                    sum(atom.origin_y for atom in kept) / len(kept),
                    0,
                    max(atom.size for atom in kept),
                    False,
                )
                found.append((span.origin_y, text, span))

        for atom in ordered:
            if re.fullmatch(r"[\d/\-.:]", atom.text.strip() or " "):
                run.append(atom)
            else:
                flush()
                run = []
        flush()
    return found


def _restore_layer_numbers(page: fitz.Page, atoms: list[_Atom]) -> list[_Atom]:
    """
    Put back a number the picture reader skipped.

    The text layer's digits are used only when the picture reader has no
    copy of that number on the same line. A page with no text layer changes
    nothing. The number keeps the text layer's position.
    """
    if not atoms:
        return atoms
    try:
        layer = _atoms_on_page(page)
    except Exception:
        return atoms
    if not layer:
        return atoms
    height = float(page.rect.height)
    for baseline, text, host in _layer_numbers(layer):
        if host.y0 < height * 0.08 or host.y0 > height * 0.88:
            continue
        line_atoms = [
            atom
            for atom in atoms
            if abs(atom.origin_y - baseline) <= 14 or abs(((atom.y0 + atom.y1) / 2.0) - baseline) <= 14
        ]
        if text in "".join(atom.text for atom in line_atoms):
            continue
        anchor = min(line_atoms, key=lambda atom: abs(atom.origin_y - baseline)) if line_atoms else host
        atoms.append(
            _Atom(
                f"{text} ",
                host.x0,
                anchor.y0,
                max(host.x1, host.x0 + 4.0),
                anchor.y1,
                host.x0,
                anchor.origin_y,
                0,
                anchor.size or host.size,
                False,
            )
        )
    return atoms


def _ocr_word_atoms(page: fitz.Page) -> list[_Atom]:
    """
    Words read from the rendered page, already in logical Unicode.

    Each Arabic word is one unit, so the line reorder flips the words and
    leaves the letters inside a word as Windows read them. Logo noise such as
    the Latin letters inside a picture of a wordmark is left out.
    """
    scale = 2.0
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    try:
        result = _recognize_png(pixmap.tobytes("png"))
    except Exception:
        return []
    if result is None:
        return []
    words = []
    for line in result.lines:
        for item in line.words:
            text = plain_arabic(getattr(item, "text", "") or "")
            text = "".join(
                char
                for char in text
                if not (
                    unicodedata.category(char).startswith("L")
                    and not _is_arabic(char)
                    and not (char.isascii() and char.isalpha())
                )
            )
            text = plain_arabic(text)
            if not text:
                continue
            has_arabic = any(_is_arabic(char) for char in text)
            quoted = any(char in "\"'«»“”" for char in text)
            # A wordmark read as "2i?3O" or "30#2" is the logo, not a sentence.
            # A quote or a one-letter Arabic token is not a wordmark.
            if not has_arabic and not quoted and not re.fullmatch(r"[\d\s.,:;/()\-–—%+]+", text):
                continue
            if any(char.isascii() and char.isalpha() for char in text) and not has_arabic:
                continue
            box = item.bounding_rect
            x0 = float(box.x) / scale
            y0 = float(box.y) / scale
            x1 = (float(box.x) + float(box.width)) / scale
            y1 = (float(box.y) + float(box.height)) / scale
            short = quoted or (has_arabic and sum(1 for char in text if _is_arabic(char)) <= 1)
            if (x1 - x0 < 1 or y1 - y0 < 6) and not short:
                continue
            if y1 - y0 < 3 and x1 - x0 < 2:
                continue
            # The authority name is real text. A two-letter scrap in that band is the logo.
            if y1 < 72 and _arabic_count(text) < 4 and not text[:1].isdigit():
                continue
            # The dotted leader makes the page number read as ".4." or "، 5".
            # The real digits are taken from the left margin of the PDF.
            if x1 < 150 and not any(
                0x0621 <= ord(char) <= 0x064A or 0x0671 <= ord(char) <= 0x06D3 for char in text
            ):
                continue
            words.append((text, x0, y0, x1, y1))
    if not words:
        return []
    words.sort(key=lambda item: ((item[2] + item[4]) / 2.0, item[1]))
    bands: list[list] = []
    for word in words:
        center = (word[2] + word[4]) / 2.0
        height = max(8.0, word[4] - word[2])
        if bands and abs(center - bands[-1][0]) <= max(6.0, height * 0.45):
            bands[-1][1].append(word)
        else:
            bands.append((center, [word]))
    sizes = _trace_sizes(page)
    atoms: list[_Atom] = []
    for _center, band in bands:
        baseline = max(word[4] for word in band)
        for text, x0, y0, x1, y1 in band:
            arabic = _arabic_count(text) > 0 or any(char in "،؛؟\"'«»“”" for char in text)
            atoms.append(
                _Atom(
                    f"{text} ",
                    x0,
                    y0,
                    x1,
                    y1,
                    x0,
                    baseline,
                    1 if arabic else 0,
                    _size_near(sizes, (y0 + y1) / 2.0, max(8.0, y1 - y0)),
                    False,
                    _color_near(sizes, (y0 + y1) / 2.0),
                )
            )
    atoms.extend(_left_margin_numbers(page, atoms))
    return _restore_layer_numbers(page, _restore_quoted_letters(page, atoms))


def _left_margin_numbers(page: fitz.Page, atoms: list[_Atom]) -> list[_Atom]:
    """
    Page numbers printed in the left margin.

    The picture reader turns the dotted leader into a broken digit. The PDF
    stored that margin number as a real digit, away from the broken letters.
    """
    if not atoms:
        return []
    digits = []
    try:
        traces = page.get_texttrace() or []
    except Exception:
        return []
    for trace in traces:
        if int(trace.get("type") or 0) != 0:
            continue
        for raw in trace.get("chars") or []:
            value = int(raw[0] or 0)
            if value < 48 or value > 57:
                continue
            box = raw[3]
            x0, y0, x1, y1 = (float(box[0]), float(box[1]), float(box[2]), float(box[3]))
            if (x0 + x1) / 2.0 > 100:
                continue
            digits.append((chr(value), x0, y0, x1, y1))
    digits.sort(key=lambda item: (item[2], item[1]))
    groups: list[list] = []
    for digit in digits:
        if groups and abs(digit[2] - groups[-1][-1][2]) <= 3 and digit[1] - groups[-1][-1][3] <= 10:
            groups[-1].append(digit)
        else:
            groups.append([digit])
    baselines = [atom.origin_y for atom in atoms]
    numbers: list[_Atom] = []
    for group in groups:
        group.sort(key=lambda item: item[1])
        text = "".join(item[0] for item in group)
        if not text or len(text) > 3:
            continue
        x0 = min(item[1] for item in group)
        y0 = min(item[2] for item in group)
        x1 = max(item[3] for item in group)
        y1 = max(item[4] for item in group)
        baseline = min(baselines, key=lambda value: abs(value - y1))
        if abs(baseline - y1) > 12:
            continue
        numbers.append(_Atom(f"{text} ", x0, y0, x1, y1, x0, baseline, 0, max(8.0, y1 - y0), False))
    return numbers


def _near_white(color) -> bool:
    """Paper, a white cell, and a hairline highlight are not part of a diagram."""
    if not color:
        return True
    try:
        red, green, blue = (float(color[0]), float(color[1]), float(color[2]))
    except (TypeError, ValueError, IndexError):
        return True
    return red >= 0.93 and green >= 0.93 and blue >= 0.93


def _drawing_shapes(page: fitz.Page) -> list[tuple]:
    """
    Colored shapes on the page.

    Each item is the box, the fill color, and whether the path is only a rectangle.
    Table rules are thin strokes and are not included. A white cell is not included.
    """
    shapes = []
    page_area = float(page.rect.width * page.rect.height) or 1.0
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        return shapes
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None:
            continue
        box = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
        width = box[2] - box[0]
        height = box[3] - box[1]
        if width < 12 or height < 12 or width * height > page_area * 0.7:
            continue
        fill = drawing.get("fill")
        if _near_white(fill):
            continue
        items = drawing.get("items") or []
        kinds = {item[0] for item in items if item}
        rectangular = bool(kinds) and kinds <= {"re", "l"}
        try:
            color = tuple(round(float(channel), 2) for channel in list(fill)[:3])
        except (TypeError, ValueError):
            color = None
        shapes.append((box, color, rectangular))
    return shapes


def _shape_bands(shapes: list[tuple]) -> list[list[tuple]]:
    """Shapes that share a horizontal band, such as a row of process arrows."""
    pending = sorted(shapes, key=lambda shape: shape[0][1])
    bands: list[list[tuple]] = []
    for shape in pending:
        box = shape[0]
        placed = False
        for band in bands:
            top = min(item[0][1] for item in band)
            bottom = max(item[0][3] for item in band)
            overlap = min(box[3], bottom) - max(box[1], top)
            smaller = min(box[3] - box[1], bottom - top)
            if smaller > 0 and overlap >= smaller * 0.45:
                band.append(shape)
                placed = True
                break
        if not placed:
            bands.append([shape])
    return bands


def _is_shape_lattice(bands: list[list[tuple]]) -> bool:
    """
    A data table's filled cells: several rows, aligned columns, every cell a rectangle.

    A row of arrows is one band, or the paths are not rectangles, so it is not a lattice.
    """
    if len(bands) < 2:
        return False
    counts = [len(band) for band in bands]
    if min(counts) < 2 or max(counts) - min(counts) > 1:
        return False
    if any(not shape[2] for band in bands for shape in band):
        return False
    rows = [sorted(shape[0][0] for shape in band) for band in bands]
    for left, right in zip(rows, rows[1:]):
        width = min(len(left), len(right))
        if any(abs(a - b) > 12 for a, b in zip(left[:width], right[:width])):
            return False
    return True


def _lattice_band_indexes(bands: list[list[tuple]]) -> set[int]:
    """Bands that together are a filled data grid, and so must stay a table."""
    used: set[int] = set()
    for start in range(len(bands)):
        for end in range(start + 2, len(bands) + 1):
            if _is_shape_lattice(bands[start:end]):
                used.update(range(start, end))
    return used


def _diagram_parts(page: fitz.Page) -> list[tuple[tuple[float, float, float, float], tuple | None]]:
    """
    Filled shapes and stroked icons that can belong to a diagram.

    A heading bar is a short full-width rectangle and stays with the heading.
    Strokes with no fill are the circles and icons. White paper is not a shape.
    """
    height = float(page.rect.height)
    width = float(page.rect.width)
    page_area = width * height or 1.0
    parts: list[tuple[tuple[float, float, float, float], tuple | None]] = []
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        return parts
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None:
            continue
        box = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
        shape_width = box[2] - box[0]
        shape_height = box[3] - box[1]
        if shape_width < 8 or shape_height < 8:
            continue
        if shape_width * shape_height > page_area * 0.55:
            continue
        if box[3] < height * 0.12 or box[1] > height * 0.90:
            continue
        if shape_height <= 3.0 and shape_width > width * 0.45:
            continue
        if shape_height < 32 and shape_width > width * 0.82:
            continue
        fill = drawing.get("fill")
        stroke = drawing.get("color")
        if _near_white(fill) and _near_white(stroke):
            continue
        color = None
        source = fill if fill else stroke
        if source:
            try:
                color = tuple(round(float(channel), 2) for channel in list(source)[:3])
            except (TypeError, ValueError):
                color = None
        parts.append((box, color))
    return parts


def _boxes_touch(left: tuple[float, float, float, float], right: tuple[float, float, float, float], gap: float) -> bool:
    """True when two boxes meet or the space between them is smaller than `gap`."""
    dx = max(0.0, max(left[0], right[0]) - min(left[2], right[2]))
    dy = max(0.0, max(left[1], right[1]) - min(left[3], right[3]))
    return dx <= gap and dy <= gap


def _union_boxes(boxes: list[tuple[float, float, float, float]]) -> tuple[float, float, float, float]:
    return (
        min(box[0] for box in boxes),
        min(box[1] for box in boxes),
        max(box[2] for box in boxes),
        max(box[3] for box in boxes),
    )


def _vector_figures(page: fitz.Page) -> list[ImageBlock]:
    """
    Infographics, arrows, and colored diagrams, each kept as one picture.

    Cards in two columns are one figure even when a gap separates the columns.
    The clip includes the icons and the words painted on the cards. A filled
    data grid is left for the table reader. A heading bar is not a figure.
    """
    parts = _diagram_parts(page)
    if len(parts) < 2:
        return []
    clusters: list[list[tuple]] = [[part] for part in parts]
    merged = True
    while merged:
        merged = False
        for index in range(len(clusters)):
            for other in range(len(clusters) - 1, index, -1):
                if any(_boxes_touch(left[0], right[0], 36.0) for left in clusters[index] for right in clusters[other]):
                    clusters[index].extend(clusters[other])
                    del clusters[other]
                    merged = True
                    break
            if merged:
                break
    merged = True
    while merged:
        merged = False
        for index in range(len(clusters)):
            for other in range(index + 1, len(clusters)):
                if len(clusters[index]) < 2 or len(clusters[other]) < 2:
                    continue
                left = _union_boxes([part[0] for part in clusters[index]])
                right = _union_boxes([part[0] for part in clusters[other]])
                overlap = min(left[3], right[3]) - max(left[1], right[1])
                smaller = min(left[3] - left[1], right[3] - right[1])
                gap = max(0.0, max(left[0], right[0]) - min(left[2], right[2]))
                # Two columns of one diagram sit closer than a third of the page.
                if smaller > 36 and overlap >= smaller * 0.55 and gap < float(page.rect.width) * 0.34:
                    clusters[index].extend(clusters[other])
                    del clusters[other]
                    merged = True
                    break
            if merged:
                break
    figures: list[ImageBlock] = []
    page_area = float(page.rect.width * page.rect.height) or 1.0
    page_width = float(page.rect.width)
    try:
        text_atoms = _atoms_on_page(page)
    except Exception:
        text_atoms = []
    seen: list[tuple[float, float, float, float]] = []
    for cluster in clusters:
        colors = {part[1] for part in cluster if part[1] is not None}
        group = [part[0] for part in cluster]
        probe = _union_boxes(group)
        overlaps_text = any(
            probe[0] <= (atom.x0 + atom.x1) / 2.0 <= probe[2]
            and probe[1] <= (atom.y0 + atom.y1) / 2.0 <= probe[3]
            for atom in text_atoms
            if atom.text.strip()
        )
        # A pair of shapes is a diagram when it has more than one color and
        # words sit on it. A pair with no words is a bullet or a logo.
        if len(cluster) < 3 and not (len(colors) >= 2 and overlaps_text):
            continue
        box = (
            max(0.0, min(item[0] for item in group) - 4),
            max(0.0, min(item[1] for item in group) - 4),
            min(float(page.rect.width), max(item[2] for item in group) + 4),
            min(float(page.rect.height), max(item[3] for item in group) + 4),
        )
        if any(_overlap_ratio(box, other) >= 0.6 for other in seen):
            continue
        width = box[2] - box[0]
        height = box[3] - box[1]
        area = width * height
        if width < max(70.0, page_width * 0.12) or height < 24 or area > page_area * 0.62:
            continue
        try:
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=fitz.Rect(*box), alpha=False)
            data = pixmap.tobytes("png")
        except Exception:
            continue
        if not data:
            continue
        seen.append(box)
        figures.append(ImageBlock(box, data, width, height, graphic=True))
    return figures


def _without_figure_tables(tables: list[TableBlock], figures: list[ImageBlock]) -> list[TableBlock]:
    """A row of colored cards is a figure, even when the line finder calls it a grid."""
    kept = []
    for table in tables:
        covered = any(
            _overlap_ratio(table.bbox, figure.bbox) >= 0.45
            or _overlap_ratio(figure.bbox, table.bbox) >= 0.45
            for figure in figures
        )
        if not covered:
            kept.append(table)
    return kept


class _ColumnSchema:
    """
    The column count of the first real table in the file.

    Later tables may be narrower. They may not grow extra columns: a ruled
    line that the detector treats as a new column is folded back into the
    last real column. A later table that is genuinely wider, with filled
    cells, is left as it was read.
    """

    def __init__(self) -> None:
        self.columns: int | None = None


def _content_frame(page: fitz.Page) -> tuple[float, float]:
    """Usable width and height, in points, after the Word margins."""
    return (
        max(72.0, float(page.rect.width) - 2 * _MARGIN_X_PT),
        max(72.0, float(page.rect.height) - 2 * _MARGIN_Y_PT),
    )


def _phrase_key(text: str) -> str:
    return re.sub(r"\s+", " ", plain_arabic(text))


def _is_page_label(text: str) -> bool:
    """A page number, with or without the word 'page', is not body text."""
    compact = _phrase_key(text)
    # Collapse letter-spaced labels such as "P a g e" from some Word exports.
    joined = re.sub(r"(?<=[A-Za-z])\s+(?=[A-Za-z])", "", compact)
    pattern = (
        r"(?:page|صفحة|الصفحة)?\s*[\d٠-٩]{1,4}\s*(?:of|من)?\s*[\d٠-٩]{0,4}"
        r"|[\d٠-٩]{1,4}\s*[|]\s*page"
        r"|[\d٠-٩]{1,4}\s+page"
        r"|page\s*[\d٠-٩]{1,4}"
    )
    return bool(re.fullmatch(pattern, compact, flags=re.IGNORECASE) or re.fullmatch(pattern, joined, flags=re.IGNORECASE))


def _estimated_table_height(table: TableBlock, width_pt: float) -> float:
    """How tall this grid will be once Word wraps each cell into the page width."""
    columns = max((len(row) for row in table.rows), default=1)
    column_width = max(width_pt / max(columns, 1), 36.0)
    chars_per_line = max(6.0, column_width / 6.4)
    total = 0.0
    for row in table.rows:
        lines = 1
        for cell in row:
            lines = max(lines, max(1, int((len(cell) + chars_per_line - 1) // chars_per_line)))
        total += min(lines, 12) * 15.0 + 4.0
    return total


def _table_rank(page: fitz.Page, table: TableBlock) -> tuple[float, int]:
    """Prefer the grid whose wrapped height still matches the rectangle on the page."""
    width, _height = _content_frame(page)
    estimate = _estimated_table_height(table, width)
    bbox_height = max(1.0, table.bbox[3] - table.bbox[1])
    distortion = abs(estimate - bbox_height) / bbox_height
    filled = sum(len(cell) for row in table.rows for cell in row)
    return (distortion, -filled)


def _is_false_grid(page: fitz.Page, table: TableBlock, atoms: list[_Atom] | None = None) -> bool:
    """
    A grid that cannot sit back on the page it was read from.

    Decorative rules are sometimes reported as one row per line, including
    rows that have no words. Rebuilding that grid is what turns one PDF page
    into many Word pages. The letters are kept and read again as paragraphs.
    A real table has about one row per line of text inside it.
    """
    width, _height = _content_frame(page)
    estimate = _estimated_table_height(table, width)
    bbox_height = max(1.0, table.bbox[3] - table.bbox[1])
    page_height = float(page.rect.height)
    if len(table.rows) > max(48, int(page_height / 11)):
        return True
    if atoms:
        box = (
            table.bbox[0] - 2,
            table.bbox[1] - 2,
            table.bbox[2] + 2,
            table.bbox[3] + 2,
        )
        inside = [atom for atom in atoms if atom.text.strip() and _center_inside(atom, box)]
        baselines = _cluster_baselines(inside)
        if len(baselines) >= 3 and len(table.rows) > max(len(baselines) * 1.35, len(baselines) + 3):
            return True
    return estimate > page_height * 1.35 and estimate > bbox_height * 1.7


def _margin_table(page: fitz.Page, table: TableBlock) -> bool:
    """A grid that lives entirely in the header or footer band is running furniture."""
    return _chrome_edge(table.bbox, float(page.rect.height)) is not None


def _dedupe_tables(page: fitz.Page, tables: list[TableBlock]) -> list[TableBlock]:
    """Two detectors describing the same rectangle: keep the one that fits the page."""
    kept: list[TableBlock] = []
    for table in sorted(tables, key=lambda item: _table_rank(page, item)):
        if any(
            _overlap_ratio(table.bbox, other.bbox) >= 0.45
            or _overlap_ratio(other.bbox, table.bbox) >= 0.45
            for other in kept
        ):
            continue
        kept.append(table)
    kept.sort(key=lambda item: item.bbox[1])
    return kept


def _margin_phrases(page: fitz.Page) -> list[tuple[str, str]]:
    """Short lines painted in the top or bottom band of this page only."""
    atoms = _atoms_on_page(page)
    if len(atoms) < 4:
        return []
    height = float(page.rect.height)
    found: list[tuple[str, str]] = []
    for band in _cluster_baselines(atoms):
        line = _line_from_atoms(band)
        if line is None:
            continue
        text = _phrase_key(line.text)
        if not text or len(text) > 90:
            continue
        edge = _chrome_edge(line.bbox, height)
        if edge == "top":
            found.append(("top", text))
        elif edge == "bottom":
            found.append(("bottom", text))
    return found


def _furniture_need(page_count: int) -> int:
    """How many pages in this batch must share a margin line before it is furniture."""
    if page_count <= 1:
        return 99
    if page_count == 2:
        return 2
    return max(3, (page_count * 2 + 4) // 5)


def _furniture_for(pages: list[fitz.Page]) -> set[tuple[str, str]]:
    """Margin lines repeated inside this batch. Body text is never part of the set."""
    counts: dict[tuple[str, str], int] = {}
    for page in pages:
        seen: set[tuple[str, str]] = set()
        for item in _margin_phrases(page):
            if item in seen:
                continue
            seen.add(item)
            counts[item] = counts.get(item, 0) + 1
    need = _furniture_need(len(pages))
    return {item for item, count in counts.items() if count >= need}


def _chrome_edge(bbox: tuple[float, float, float, float], page_height: float) -> str | None:
    """Top or bottom band. A tall body box that merely crosses the band is kept."""
    top, bottom = float(bbox[1]), float(bbox[3])
    band = page_height * _CHROME_BAND
    height = max(0.0, bottom - top)
    center = (top + bottom) / 2.0
    if bottom <= band or (center <= band and height <= page_height * 0.08):
        return "top"
    if top >= page_height - band or (center >= page_height - band and height <= page_height * 0.08):
        return "bottom"
    return None


def _block_band(bbox: tuple[float, float, float, float], page_height: float) -> str | None:
    return _chrome_edge(bbox, page_height)


def _text_block_height(block: TextBlock, width_pt: float) -> float:
    size = 16.0 if block.kind == "heading" else min(max(block.size, 11.0), 14.0)
    chars_per_line = max(8.0, width_pt / (size * 0.5))
    lines = max(1, int((len(block.text) + chars_per_line - 1) // chars_per_line))
    return lines * size * 1.08 + (8.0 if block.kind == "heading" else 2.0)


def _fold_to_schema(table: TableBlock, schema: _ColumnSchema) -> None:
    """Lock the first data table's width. Fold only columns that look invented."""
    if not table.rows:
        return
    width = max(len(row) for row in table.rows)
    if width < 2:
        return
    if schema.columns is None:
        if len(table.rows) >= 3:
            cells = [cell for row in table.rows for cell in row if cell.strip()]
            if cells and sum(len(cell) for cell in cells) / len(cells) <= 60:
                schema.columns = width
        return
    if width <= schema.columns:
        return
    extra: list[str] = []
    for row in table.rows:
        padded = list(row) + [""] * (width - len(row))
        extra.extend(cell.strip() for cell in padded[schema.columns :] if cell.strip())
    if extra:
        median = sorted(len(cell) for cell in extra)[len(extra) // 2]
        filled_rows = 0
        for row in table.rows:
            padded = list(row) + [""] * (width - len(row))
            if any(cell.strip() for cell in padded[schema.columns :]):
                filled_rows += 1
        if median >= 4 and filled_rows >= max(2, int(len(table.rows) * 0.5)):
            return
    folded: list[list[str]] = []
    keep = schema.columns
    for row in table.rows:
        padded = list(row) + [""] * (width - len(row))
        head = padded[: keep - 1]
        tail = " ".join(cell.strip() for cell in padded[keep - 1 :] if cell.strip())
        folded.append(head + [tail])
    table.rows = folded


def _squeeze_images(blocks: list, width_pt: float, height_pt: float) -> None:
    """
    Keep a figure at its PDF size unless the page cannot hold it.

    An infographic stays at the box it was drawn in. Only a page that is
    already taller than one sheet scales the pictures down, together.
    """
    images = [block for block in blocks if isinstance(block, ImageBlock) and block.fit != "page"]
    if not images:
        return
    text_height = 0.0
    for block in blocks:
        if isinstance(block, TextBlock):
            text_height += _text_block_height(block, width_pt)
        elif isinstance(block, TableBlock):
            text_height += min(_estimated_table_height(block, width_pt), height_pt)
    picture_height = sum(max(0.0, image.height) for image in images)
    if text_height + picture_height <= height_pt * 0.92:
        return
    room = max(36.0, height_pt * 0.92 - text_height)
    scale = room / picture_height if picture_height else 1.0
    scale = min(1.0, max(0.45, scale))
    for image in images:
        if image.width <= 0 or image.height <= 0:
            continue
        image.height = max(24.0, image.height * scale)
        image.width = max(24.0, image.width * scale)
        if image.width > width_pt:
            image.height = image.height * (width_pt / image.width)
            image.width = width_pt


def _settle_page(
    page: fitz.Page,
    blocks: list,
    furniture: set[tuple[str, str]],
    schema: _ColumnSchema,
) -> list:
    """
    Make one PDF page fit its own Word section.

    Nothing here reads the previous or the next page. A repeated margin line
    is removed, a paragraph already written on this page is not written again,
    and a picture is scaled to the room that remains.
    """
    height = float(page.rect.height)
    width_pt, height_pt = _content_frame(page)
    kept: list = []
    seen: set[str] = set()
    for block in blocks:
        if isinstance(block, TextBlock):
            key = _phrase_key(block.text)
            if not key:
                continue
            # The top and bottom bands are the section header and footer.
            if _chrome_edge(block.bbox, height) is not None:
                continue
            if block.bbox[1] >= height * (1.0 - _FOOTER_ZONE):
                continue
            if _is_page_label(key) and block.bbox[1] >= height * 0.86:
                continue
            band = _block_band(block.bbox, height)
            if band is not None and (band, key) in furniture:
                continue
            if key in seen and len(key) >= 8:
                continue
            seen.add(key)
            kept.append(block)
            continue
        if isinstance(block, TableBlock):
            _fold_to_schema(block, schema)
            if block.rows:
                kept.append(block)
            continue
        if isinstance(block, ImageBlock) and block.fit != "page":
            # A logo is written in the header. Corner art in the footer is not a figure.
            if _chrome_edge(block.bbox, height) is not None:
                continue
            if block.bbox[1] >= height * (1.0 - _FOOTER_ZONE):
                continue
            if block.bbox[1] < height * _CHROME_BAND and block.bbox[3] <= height * 0.14:
                continue
        kept.append(block)
    _squeeze_images(kept, width_pt, height_pt)
    return kept


def _container_boxes(page: fitz.Page) -> list[tuple[float, float, float, float]]:
    """
    Panels, cards, and framed regions on this page.

    A full-page border is not a container: every line would fall inside it
    and the page would be treated as one block. A hairline rule is not a
    container either. Text is later pinned to the smallest panel that
    actually holds it, so the words stay with that panel.
    """
    page_area = float(page.rect.width * page.rect.height) or 1.0
    boxes: list[tuple[float, float, float, float]] = []
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        drawings = []
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None:
            continue
        box = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
        width = box[2] - box[0]
        height = box[3] - box[1]
        if width < 36 or height < 16:
            continue
        if width * height > page_area * 0.72:
            continue
        boxes.append(box)
    return boxes


def _release_text_panels(figures: list[ImageBlock], atoms: list[_Atom]) -> list[ImageBlock]:
    """
    Keep a diagram even when a label is painted on it.

    An infographic is one picture: the words on the arrows stay in that picture
    and are not also written as paragraphs. A large image that is really a
    page of body text is still released, so the paragraph reader can take it.
    """
    kept: list[ImageBlock] = []
    for figure in figures:
        if figure.graphic:
            kept.append(figure)
            continue
        letters = 0
        for atom in atoms:
            if atom.text.strip() and _center_inside(atom, figure.bbox):
                letters += len(atom.text.strip())
                if letters >= 80:
                    break
        if letters >= 80:
            continue
        kept.append(figure)
    return kept


def _clip_to_page(page: fitz.Page, blocks: list) -> list:
    """Drop anything whose center falls outside this page. A block cannot cross the sheet."""
    width = float(page.rect.width)
    height = float(page.rect.height)
    kept = []
    for block in blocks:
        box = block.bbox
        clipped = (
            min(max(float(box[0]), 0.0), width),
            min(max(float(box[1]), 0.0), height),
            min(max(float(box[2]), 0.0), width),
            min(max(float(box[3]), 0.0), height),
        )
        if clipped[2] - clipped[0] < 1 or clipped[3] - clipped[1] < 1:
            continue
        center_y = (clipped[1] + clipped[3]) / 2.0
        if center_y < 0 or center_y > height:
            continue
        block.bbox = clipped
        kept.append(block)
    return kept


def _group_with_containers(page: fitz.Page, blocks: list) -> list:
    """
    Pin each paragraph to the panel that contains its center.

    Paragraphs that share a panel are written together, in top-to-bottom
    order, at the position of that panel. A table or a picture keeps its own
    place. Nothing is taken from another page: the boxes are clipped to this
    sheet first.
    """
    blocks = _clip_to_page(page, blocks)
    containers = _container_boxes(page)
    if not containers:
        return blocks

    def container_of(block: TextBlock) -> int | None:
        center_x = (block.bbox[0] + block.bbox[2]) / 2.0
        center_y = (block.bbox[1] + block.bbox[3]) / 2.0
        hits: list[tuple[float, int]] = []
        for index, box in enumerate(containers):
            if box[0] - 1 <= center_x <= box[2] + 1 and box[1] - 1 <= center_y <= box[3] + 1:
                area = (box[2] - box[0]) * (box[3] - box[1])
                hits.append((area, index))
        if not hits:
            return None
        hits.sort()
        return hits[0][1]

    groups: dict[int, list[TextBlock]] = {}
    free: list = []
    for block in blocks:
        if isinstance(block, TextBlock):
            container = container_of(block)
            if container is not None:
                groups.setdefault(container, []).append(block)
                continue
        free.append(block)
    events: list[tuple[float, int, list]] = []
    for block in free:
        events.append((block.bbox[1], 1, [block]))
    for members in groups.values():
        members.sort(key=lambda block: (block.bbox[1], -block.bbox[2]))
        events.append((min(block.bbox[1] for block in members), 0, members))
    events.sort(key=lambda item: (item[0], item[1]))
    ordered: list = []
    for _top, _priority, members in events:
        ordered.extend(members)
    return ordered


def _letter_bag(text: str) -> Counter:
    """Letters and digits only, so punctuation and spaces do not change the score."""
    bag: Counter = Counter()
    for char in text:
        if _is_arabic(char) or char.isdigit() or (char.isascii() and char.isalpha()):
            bag[char] += 1
    return bag


def _coverage(source: str, output: str) -> float:
    """Share of the page's letters that the output still contains, ignoring order."""
    bag = _letter_bag(source)
    total = sum(bag.values())
    if total < 12:
        return 1.0
    found = _letter_bag(output)
    hit = sum(min(found[char], count) for char, count in bag.items())
    return hit / total


def _fragmented(text: str) -> bool:
    """True when Arabic words have been split into single letters."""
    tokens = [token for token in re.findall(r"\S+", text) if any(_is_arabic(char) for char in token)]
    if len(tokens) < 10:
        return False
    singles = sum(1 for token in tokens if _arabic_count(token) <= 1)
    return singles / len(tokens) >= 0.35


def _body_atoms(page: fitz.Page, atoms: list[_Atom]) -> list[_Atom]:
    """Letters outside the header and footer bands. Those bands are removed on purpose."""
    height = float(page.rect.height)
    kept = []
    for atom in atoms:
        center = (atom.y0 + atom.y1) / 2.0
        if center <= height * _CHROME_BAND or center >= height * (1.0 - _CHROME_BAND):
            continue
        kept.append(atom)
    return kept or list(atoms)


def _output_text(blocks: list) -> str:
    parts: list[str] = []
    for block in blocks:
        if isinstance(block, TextBlock):
            parts.append(block.text)
        elif isinstance(block, TableBlock):
            parts.extend(cell for row in block.rows for cell in row)
    return "\n".join(parts)


def _layout_hit(atoms: list[_Atom], blocks: list) -> float:
    """
    Share of source letters whose center still sits inside a written block.

    Letters painted on a picture we kept are not required as text: they stay
    inside that picture.
    """
    figures = [block.bbox for block in blocks if isinstance(block, ImageBlock)]
    boxes = [block.bbox for block in blocks if isinstance(block, (TextBlock, TableBlock))]
    letters = []
    for atom in atoms:
        if not atom.text.strip():
            continue
        if any(_center_inside(atom, box) for box in figures):
            continue
        letters.append(atom)
    if len(letters) < 8:
        return 1.0
    if not boxes:
        return 0.0
    hit = 0
    for atom in letters:
        if any(_center_inside(atom, box) for box in boxes):
            hit += 1
    return hit / len(letters)


def _assess_page(page: fitz.Page, blocks: list, source: list[_Atom], trust_letters: bool) -> dict:
    """
    Compare this page's reading with the letters on the PDF page.

    The score is used only to choose between the first reading and one retry.
    A passing page is saved as it is. A failing page is read again by the
    other method, and the higher score is the one that is written.
    """
    output = _output_text(blocks)
    source_text = "".join(atom.text for atom in source)
    coverage = _coverage(source_text, output) if trust_letters else 1.0
    layout = _layout_hit(source, blocks)
    fragmented = _fragmented(output)
    rows = sum(len(block.rows) for block in blocks if isinstance(block, TableBlock))
    baselines = len(_cluster_baselines([atom for atom in source if atom.text.strip()])) or 1
    extra = max(0, rows - int(baselines * 1.35) - 2)
    source_letters = sum(_letter_bag(source_text).values())
    if not trust_letters and source_letters >= 30:
        output_letters = sum(_letter_bag(output).values())
        if output_letters < source_letters * 0.35:
            coverage = output_letters / source_letters
    score = coverage * 100.0 + layout * 25.0 - (30.0 if fragmented else 0.0) - extra * 4.0
    if source_letters < 12 and trust_letters:
        ok = extra == 0
        score = 100.0 if ok else 40.0
    else:
        ok = coverage >= 0.78 and layout >= 0.60 and not fragmented and extra == 0
    return {
        "ok": ok,
        "score": score,
        "coverage": coverage,
        "layout": layout,
        "fragmented": fragmented,
        "extra": extra,
    }


def _table_has_words(table: TableBlock) -> bool:
    """A ruled grid with no words is a decoration, not a table."""
    return any(cell.strip() for row in table.rows for cell in row)


def _paragraphs_from_regions(
    page: fitz.Page,
    atoms: list[_Atom],
    tables: list[TableBlock],
    figures: list[ImageBlock],
) -> list[TextBlock]:
    """
    One PDF text block stays one paragraph.

    This is the block reader. It does not join a line from one block onto
    the next block, which is what pulls a side-by-side title into the body.
    """
    try:
        data = page.get_text("dict") or {}
    except Exception:
        data = {}
    regions: list[TextBlock] = []
    used: set[int] = set()
    for raw in data.get("blocks") or []:
        if raw.get("type") != 0:
            continue
        raw_box = raw.get("bbox")
        if not raw_box:
            continue
        box = _box(raw_box)
        if any(_overlap_ratio(box, table.bbox) >= 0.45 for table in tables):
            continue
        if any(_overlap_ratio(box, figure.bbox) >= 0.45 for figure in figures):
            continue
        inside = _atoms_in(atoms, box)
        if len(inside) < 2:
            continue
        for atom in inside:
            used.add(id(atom))
        pieces = _paragraphs_from_atoms(page, inside, [], [])
        if not pieces:
            continue
        if any(piece.role == "toc" for piece in pieces) or len(pieces) == 1:
            for piece in pieces:
                regions.append(piece)
            continue
        text = plain_arabic(" ".join(piece.text for piece in pieces if piece.text))
        if not text:
            continue
        regions.append(
            TextBlock(
                "heading" if any(piece.kind == "heading" for piece in pieces) else "paragraph",
                box,
                text,
                max(piece.size for piece in pieces),
                any(piece.bold for piece in pieces),
            )
        )
    rest = [atom for atom in atoms if id(atom) not in used]
    if rest:
        regions.extend(_paragraphs_from_atoms(page, rest, tables, figures))
    if len(regions) < 2:
        return _paragraphs_from_atoms(page, atoms, tables, figures)
    regions.sort(key=lambda block: (round(block.bbox[1], 0), -block.bbox[2]))
    return regions


def _align_for_box(box: tuple[float, float, float, float], page_width: float, text: str) -> str:
    """Center, left, or right from where the box sits on the sheet."""
    width = page_width or 1.0
    center = (box[0] + box[2]) / 2.0
    line_width = box[2] - box[0]
    if abs(center - width / 2.0) <= width * 0.12 and line_width <= width * 0.72:
        return "center"
    latin = _arabic_count(text) * 2 < len(text.strip())
    if center <= width * 0.36 and line_width <= width * 0.5 and latin:
        return "left"
    return "right"


def _block_height(block, page: fitz.Page) -> float:
    if isinstance(block, TextBlock):
        return max(block.size, 11.0) * 1.35
    if isinstance(block, TableBlock):
        return max(len(block.rows), 1) * 18.0
    if isinstance(block, ImageBlock):
        if block.fit == "page":
            return float(page.rect.height) * 0.9
        return min(block.height or 80.0, float(page.rect.height) * 0.85)
    return 14.0


def _same_row(left, right) -> bool:
    """Two boxes sit beside each other rather than one under the other."""
    top = max(left.bbox[1], right.bbox[1])
    bottom = min(left.bbox[3], right.bbox[3])
    overlap = bottom - top
    smaller = min(left.bbox[3] - left.bbox[1], right.bbox[3] - right.bbox[1])
    if smaller < 6 or overlap < smaller * 0.55:
        return False
    if left.bbox[2] < right.bbox[0]:
        gap = right.bbox[0] - left.bbox[2]
    elif right.bbox[2] < left.bbox[0]:
        gap = left.bbox[0] - right.bbox[2]
    else:
        return False
    return 8.0 <= gap <= 260.0


def _pair_side_by_side(blocks: list) -> list:
    """
    Keep two boxes that share a row side by side.

    Word has no page coordinates. A borderless one-row table is what holds
    the left box on the left and the right box on the right.
    """
    pending = list(blocks)
    paired: list = []
    used: set[int] = set()
    for index, block in enumerate(pending):
        if index in used or not isinstance(block, TextBlock) or block.role == "toc":
            continue
        partner = None
        for other_index in range(index + 1, min(index + 4, len(pending))):
            other = pending[other_index]
            if other_index in used or not isinstance(other, TextBlock) or other.role == "toc":
                continue
            if _same_row(block, other):
                partner = other_index
                break
        if partner is None:
            continue
        other = pending[partner]
        cells = sorted((block, other), key=lambda item: item.bbox[0])
        used.add(index)
        used.add(partner)
        box = (
            min(item.bbox[0] for item in cells),
            min(item.bbox[1] for item in cells),
            max(item.bbox[2] for item in cells),
            max(item.bbox[3] for item in cells),
        )
        paired.append(
            (
                index,
                TableBlock(
                    box,
                    [[item.text for item in cells]],
                    widths=[max(8.0, item.bbox[2] - item.bbox[0]) for item in cells],
                    rules=False,
                ),
            )
        )
    if not paired:
        return blocks
    replaced = {index: table for index, table in paired}
    skip = used - set(replaced)
    flow = []
    for index, block in enumerate(pending):
        if index in replaced:
            flow.append(replaced[index])
        elif index not in skip:
            flow.append(block)
    return flow


def _assign_positions(page: fitz.Page, blocks: list) -> None:
    """
    Copy each box's place on the PDF onto the block.

    A line in the middle of the sheet keeps the gap above it, scaled so the
    page still ends on one sheet. A short line in the center stays centered.
    """
    if not blocks:
        return
    width = float(page.rect.width) or 1.0
    for block in blocks:
        if isinstance(block, TextBlock) and block.role not in {"toc", "title", "subtitle", "meta"}:
            block.align = _align_for_box(block.bbox, width, block.text)
        elif isinstance(block, TextBlock) and block.role in {"title", "subtitle", "meta"} and not block.align:
            block.align = _align_for_box(block.bbox, width, block.text)
    previous = None
    for block in blocks:
        if getattr(block, "fit", "") == "page":
            block.gap_before = 0.0
            previous = block
            continue
        if previous is None:
            gap = block.bbox[1] - float(page.rect.height) * 0.05
        else:
            gap = block.bbox[1] - previous.bbox[3]
        block.gap_before = max(0.0, gap)
        previous = block
    _width, frame = _content_frame(page)
    occupied = sum(_block_height(block, page) for block in blocks)
    if occupied >= frame * 0.92:
        for block in blocks:
            if getattr(block, "fit", "") != "page":
                block.gap_before = 0.0
        return
    spare = max(0.0, frame - occupied)
    raw = sum(max(0.0, block.gap_before) for block in blocks)
    scale = 1.0 if raw <= spare or raw <= 0 else spare / raw
    # Keep the PDF gap, scaled only when the sheet has no room for it.
    for block in blocks:
        block.gap_before = max(0.0, block.gap_before) * scale


def _placed_boxes(page: fitz.Page, blocks: list) -> list[tuple[float, float, float, float]]:
    """
    The rectangles Word will actually paint, after gaps and alignment.

    These are rendered to a temporary image and compared with the PDF page.
    Stacking every line at the top, which the PDF did not do, fails that check.
    """
    if not blocks:
        return []
    page_width = float(page.rect.width) or 1.0
    page_height = float(page.rect.height) or 1.0
    cursor = page_height * 0.05
    placed: list[tuple[float, float, float, float]] = []
    for block in blocks:
        cursor += max(0.0, float(getattr(block, "gap_before", 0) or 0))
        height = _block_height(block, page)
        box = block.bbox
        box_width = max(8.0, float(box[2]) - float(box[0]))
        align = getattr(block, "align", "right")
        full = isinstance(block, TableBlock) or getattr(block, "fit", "") == "page"
        if full and getattr(block, "fit", "") == "page":
            left, right = 0.0, page_width
        elif full:
            left, right = _MARGIN_X_PT, page_width - _MARGIN_X_PT
        elif align == "center":
            left = max(0.0, (page_width - box_width) / 2.0)
            right = left + box_width
        elif align == "left":
            left = _MARGIN_X_PT
            right = left + box_width
        else:
            right = page_width - _MARGIN_X_PT
            left = max(0.0, right - box_width)
        bottom = min(page_height, cursor + height)
        placed.append((left, cursor, right, bottom))
        cursor += height
    return placed


def _kept_ruled_table(blocks: list) -> bool:
    return any(isinstance(block, TableBlock) and block.rules and len(block.rows) >= 2 for block in blocks)


def _ornament_images(page: fitz.Page, avoid: list[tuple] | None = None) -> list[ImageBlock]:
    """
    Colored corner art that is not a text column.

    The mark is painted as a masked image that runs off the page, so the
    figure reader drops it. The colorful pixels are clipped and kept.
    Pixels already inside a diagram are left to that diagram.
    """
    height = float(page.rect.height)
    width = float(page.rect.width)
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(0.4, 0.4), alpha=False)
    except Exception:
        return []
    try:
        from PIL import Image
    except Exception:
        return []
    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    scale = width / float(pixmap.width or 1)
    pixels = image.load()
    # Both bottom corners. A mark that repeats there is page chrome, not a column.
    span_x = int(pixmap.width * 0.26)
    span_y = int(pixmap.height * 0.18)
    windows = (
        (0, pixmap.height - span_y, span_x, pixmap.height),
        (pixmap.width - span_x, pixmap.height - span_y, pixmap.width, pixmap.height),
    )
    blocked = [tuple(box) for box in (avoid or [])]
    blocks: list[ImageBlock] = []
    for left, top, right, bottom in windows:
        xs: list[int] = []
        ys: list[int] = []
        for y in range(top, bottom, 2):
            for x in range(left, right, 2):
                point_x = x * scale
                point_y = y * scale
                if any(
                    box[0] - 6 <= point_x <= box[2] + 6 and box[1] - 6 <= point_y <= box[3] + 6
                    for box in blocked
                ):
                    continue
                red, green, blue = pixels[x, y]
                if max(red, green, blue) - min(red, green, blue) < 46:
                    continue
                if red > 232 and green > 232 and blue > 232:
                    continue
                xs.append(x)
                ys.append(y)
        if len(xs) < 24:
            continue
        box = (
            max(0.0, min(xs) * scale - 3),
            max(0.0, min(ys) * scale - 3),
            min(width, (max(xs) + 1) * scale + 3),
            min(height, (max(ys) + 1) * scale + 3),
        )
        if box[2] - box[0] < 16 or box[3] - box[1] < 16:
            continue
        if box[2] - box[0] > width * 0.55 or box[3] - box[1] > height * 0.28:
            continue
        snap = _snapshot_box(page, box)
        if snap is None:
            continue
        snap.graphic = True
        blocks.append(snap)
    return blocks


def extract_page(
    page: fitz.Page,
    plumber_page=None,
    background_xrefs: set[int] | None = None,
    kind_out: dict | None = None,
    mode: str = "auto",
) -> list:
    """
    Read one page in Arabic reading order.

    `mode` chooses the reader:
    auto — the text layer, or the picture when the layer is missing or the
    font map is wrong. The two readings are never written together.
    digital — the text layer only.
    ocr — the rendered page only.
    plain — words in their boxes, with no table detector. Used when a grid
    has invented rows or pulled text away from its panel.
    blocks — one PDF text block stays one paragraph, in its own box.
    grid — a ruled page keeps its cells instead of being flattened.

    Text that belongs to a kept table is left out of the paragraphs. An
    infographic is kept as one picture, labels and all. Paragraphs above and
    below it stay text. Every box is clipped to this page.
    """
    atoms = _atoms_on_page(page)
    letters = sum(len(atom.text.strip()) for atom in atoms)
    # A short text layer on a page with no picture is still the real text:
    # there are no pixels to read it from instead.
    digital_layer = letters >= 40 or (letters >= 3 and not _page_has_images(page))
    reading = "digital" if digital_layer else "scan"
    if mode == "ocr":
        recovered = _ocr_word_atoms(page)
        if len(recovered) >= 8:
            atoms = recovered
        reading = "scan"
    elif mode == "plain":
        if not digital_layer:
            recovered = _ocr_word_atoms(page)
            if len(recovered) >= 8:
                atoms = recovered
                reading = "scan"
    elif mode == "digital":
        reading = "digital"
    elif not digital_layer:
        recovered = _ocr_word_atoms(page)
        if len(recovered) >= 8:
            atoms = recovered
        reading = "scan"
    elif _encoding_broken(page, atoms):
        recovered = _ocr_word_atoms(page)
        if len(recovered) >= 8:
            atoms = recovered
            reading = "scan"
    if kind_out is not None:
        kind_out["kind"] = "scan" if reading == "scan" else "digital"
    if mode == "plain":
        tables = []
        drawn = _vector_figures(page)
        figures = _release_text_panels(
            drawn
            + _figures(page, atoms, background_xrefs or set())
            + _ornament_images(page, [block.bbox for block in drawn]),
            atoms,
        )
    else:
        ruled = _rule_tables(page, atoms) if reading == "digital" else []
        tables = _tables_from_pymupdf(page, atoms)
        if not tables:
            tables = _tables_from_pdfplumber(plumber_page, atoms)
        tables = _dedupe_tables(page, tables)
        if ruled:
            # Borders that close every row say where the cells are; a line
            # finder over the same area also takes each shading box's edge
            # for a rule and splits the rows.
            tables = [table for table in tables if not _covers_any(table.bbox, ruled)]
        if mode == "grid":
            # A ruled page keeps its cells even when the grid is dense.
            # A margin stamp and a grid with no words are still dropped.
            tables = [
                table
                for table in tables
                if not _margin_table(page, table) and _table_has_words(table)
            ]
        else:
            tables = [
                table
                for table in tables
                if not _is_false_grid(page, table, atoms) and not _margin_table(page, table)
            ]
        drawn = _vector_figures(page)
        figures = (
            drawn
            + _figures(page, atoms, background_xrefs or set())
            + _ornament_images(page, [block.bbox for block in drawn])
        )
        figures = _release_text_panels(figures, atoms)
        tables = _without_figure_tables(tables, figures)
        if ruled:
            figures = [figure for figure in figures if not any(_overlap_ratio(figure.bbox, table.bbox) > 0.5 for table in ruled)]
            tables = ruled + tables
        if reading == "digital":
            coloured = [c for c in _coloured_tables(page, []) if not _covers_any(c.bbox, ruled)]
            if coloured:
                # Over the same area the coloured-cell reading wins: a line
                # finder takes its white gaps for a grid and splits its rows.
                tables = [
                    table for table in tables
                    if not any(_overlap_ratio(table.bbox, c.bbox) > 0.3 or _overlap_ratio(c.bbox, table.bbox) > 0.3 for c in coloured)
                ] + coloured
        else:
            # Text read from the page picture: tables are read from it too.
            tables += _picture_tables(page, tables)
    if mode == "blocks":
        paragraphs = _paragraphs_from_regions(page, atoms, tables, figures)
    else:
        paragraphs = _paragraphs_from_atoms(page, atoms, tables, figures)
    _mark_headings(paragraphs)
    _paint_banners(page, paragraphs)
    flow = []
    pending = sorted([*tables, *figures], key=lambda block: block.bbox[1])
    for paragraph in paragraphs:
        while pending and pending[0].bbox[1] <= paragraph.bbox[1]:
            flow.append(pending.pop(0))
        flow.append(paragraph)
    flow.extend(pending)
    _strip_stamps(flow)
    return _group_with_containers(page, flow)


def _picture_tables(page: fitz.Page, found: list) -> list:
    """Ruled tables on a page with no usable text layer, read cell by cell."""
    try:
        from core import image_tables
    except Exception:
        return []
    extra = []
    try:
        boxes = image_tables.tables_on_page(page)
    except Exception:
        return []
    for box in boxes:
        if any(_overlap_ratio(box, other.bbox) > 0.3 for other in found):
            continue
        try:
            table = image_tables.read_table(page, box)
        except Exception:
            table = None
        if table is None:
            continue
        extra.append(TableBlock(bbox=table.bbox, rows=table.rows, widths=table.widths, rules=True))
    return extra


def _covers_any(box, tables: list) -> bool:
    return any(_overlap_ratio(box, t.bbox) > 0.3 or _overlap_ratio(t.bbox, box) > 0.3 for t in tables)


def _coloured_tables(page: fitz.Page, found: list) -> list:
    """
    Tables drawn as a column of filled label cells with no rules.

    The ruled-grid finders above cannot see them; the coloured-cell reader
    of table_rebuild can. A table already found over the same area wins.
    """
    try:
        from core import table_rebuild
    except Exception:
        return []
    extra = []
    try:
        rebuilt = table_rebuild.fill_tables(page)
    except Exception:
        return []
    for table in rebuilt:
        box = tuple(float(v) for v in table.bbox)
        if any(_overlap_ratio(box, other.bbox) > 0.3 for other in found):
            continue
        extra.append(TableBlock(
            bbox=box,
            rows=[list(row) for row in table.rows],
            fills=[list(row) for row in (table.fills or [])],
            widths=list(table.widths or []),
            rules=False,
        ))
    return extra


def _retry_mode(report: dict, digital_layer: bool) -> str:
    """
    The other reader.

    Missing or split letters on a text page are read from the picture.
    Invented rows, or words that no longer sit in their boxes, are read
    again without the table detector.
    """
    if digital_layer and (report.get("coverage", 1) < 0.78 or report.get("fragmented")):
        return "ocr"
    return "plain"


def _page_has_images(page: fitz.Page) -> bool:
    try:
        return bool(page.get_images(full=True))
    except Exception:
        return True


def _cover_lines(page: fitz.Page, atoms: list[_Atom]) -> list[_Line]:
    """One visual line at a time. A cover is not joined into a paragraph."""
    lines: list[_Line] = []
    for band in _cluster_baselines(atoms):
        for part in _split_gutter(band, page.rect.width):
            line = _line_from_atoms(part)
            if line is None or _is_page_number(line, page.rect.height):
                continue
            if len(line.text) < 2:
                continue
            lines.append(line)
    lines.sort(key=lambda line: (line.origin_y, -line.bbox[2]))
    return lines


def _is_cover_page(page: fitz.Page, atoms: list[_Atom]) -> bool:
    """
    The first page is a cover when it is mostly space and a few large titles.

    A normal opening page, with a title and then a run of body lines, is left
    on the ordinary reader.
    """
    lines = _cover_lines(page, atoms)
    if not lines or len(lines) > 10:
        return False
    letters = sum(len(line.text) for line in lines)
    if letters > 420:
        return False
    sizes = sorted(line.size for line in lines)
    largest = sizes[-1]
    median = sizes[len(sizes) // 2]
    has_title = largest >= 18 or (median > 0 and largest >= median * 1.45 and largest >= 15)
    if not has_title:
        return False
    ordered = sorted(lines, key=lambda line: line.origin_y)
    span = ordered[-1].bbox[3] - ordered[0].bbox[1]
    gaps = [right.bbox[1] - left.bbox[3] for left, right in zip(ordered, ordered[1:])]
    roomy = span < page.rect.height * 0.62 or any(gap >= 40 for gap in gaps)
    return roomy


def _cover_align(line: _Line, page: fitz.Page) -> str:
    """A line near the middle of the sheet is centered. Other Arabic lines start on the right."""
    width = float(page.rect.width) or 1.0
    center = (line.bbox[0] + line.bbox[2]) / 2.0
    line_width = line.bbox[2] - line.bbox[0]
    if abs(center - width / 2.0) <= width * 0.14 and line_width <= width * 0.78:
        return "center"
    return "right"


def _cover_logos(page: fitz.Page, atoms: list[_Atom], background_xrefs: set[int]) -> list[ImageBlock]:
    """
    Logos and compact marks on a cover.

    A full-page plate stays out, so the background is not pasted over the
    title. A small picture, or a cluster of shapes in the upper part of the
    page, is kept and written above the words that sit under it.
    """
    page_area = float(page.rect.width * page.rect.height) or 1.0
    logos: list[ImageBlock] = []
    seen: set[int] = set()
    for image in _image_entries(page):
        bbox = image.get("bbox")
        xref = int(image.get("xref") or 0)
        if not bbox or not xref or xref in seen or xref in background_xrefs:
            continue
        if _full_bleed(bbox, page):
            continue
        box = tuple(float(value) for value in bbox)
        width = box[2] - box[0]
        height = box[3] - box[1]
        area = width * height
        if width < 28 or height < 28 or area > page_area * 0.28 or area < page_area * 0.004:
            continue
        try:
            data = _image_bytes(page, xref)
        except Exception:
            data = None
        if not data or _pale_image(data):
            continue
        seen.add(xref)
        center = (box[0] + box[2]) / 2.0
        align = "center" if abs(center - page.rect.width / 2.0) <= page.rect.width * 0.16 else "right"
        logos.append(ImageBlock(box, data, width, height, align, 0.0))
    if logos:
        return logos
    marks: list[tuple[float, float, float, float]] = []
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        drawings = []
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None:
            continue
        box = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
        width = box[2] - box[0]
        height = box[3] - box[1]
        if width < 24 or height < 16:
            continue
        if width > page.rect.width * 0.8 and height > page.rect.height * 0.35:
            continue
        if box[1] > page.rect.height * 0.4:
            continue
        marks.append(box)
    if not marks:
        return []
    clip = (
        max(0.0, min(box[0] for box in marks) - 4),
        max(0.0, min(box[1] for box in marks) - 4),
        min(page.rect.width, max(box[2] for box in marks) + 4),
        min(page.rect.height, max(box[3] for box in marks) + 4),
    )
    if (clip[2] - clip[0]) * (clip[3] - clip[1]) > page_area * 0.28:
        return []
    # Words inside the mark are part of the logo. They are not also written as a title.
    if any(_center_inside(atom, clip) and atom.text.strip() for atom in atoms):
        return []
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=fitz.Rect(*clip), alpha=False)
        data = pixmap.tobytes("png")
    except Exception:
        return []
    if not data:
        return []
    center = (clip[0] + clip[2]) / 2.0
    align = "center" if abs(center - page.rect.width / 2.0) <= page.rect.width * 0.16 else "right"
    return [ImageBlock(clip, data, clip[2] - clip[0], clip[3] - clip[1], align, 0.0)]


def _fit_cover_gaps(blocks: list, content_height: float) -> None:
    """Keep the cover's vertical gaps, scaled so the page still closes on one sheet."""
    occupied = 0.0
    for block in blocks:
        if isinstance(block, ImageBlock):
            occupied += min(block.height, content_height * 0.28)
        elif isinstance(block, TextBlock):
            occupied += min(max(block.size, 12.0), 36.0) * 1.25
    spare = content_height * 0.92 - occupied
    raw = sum(max(0.0, block.gap_before) for block in blocks)
    if raw <= 0:
        return
    if spare < raw:
        scale = max(0.0, spare) / raw
        for block in blocks:
            block.gap_before = block.gap_before * scale


def _read_cover(
    page: fitz.Page,
    atoms: list[_Atom],
    background_xrefs: set[int],
) -> list | None:
    """
    Title, subtitle, and the short lines under them, in the order they were printed.

    Each line keeps the gap above it and is centered only when it sat in the
    middle of the page. Arabic lines stay right to left.
    """
    if not _is_cover_page(page, atoms):
        return None
    # A designed cover is vectors and logos. Rebuilding it from the text
    # layer drops the art and the colors, so the page is kept as printed.
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.7, 1.7), alpha=False)
        data = pixmap.tobytes("png")
    except Exception:
        data = b""
    if data:
        return [
            ImageBlock(
                (0.0, 0.0, float(page.rect.width), float(page.rect.height)),
                data,
                float(page.rect.width),
                float(page.rect.height),
                "center",
                0.0,
                "page",
            )
        ]
    lines = _cover_lines(page, atoms)
    logos = _cover_logos(page, atoms, background_xrefs)
    logo_boxes = [logo.bbox for logo in logos]
    kept_lines = []
    for line in lines:
        center_y = (line.bbox[1] + line.bbox[3]) / 2.0
        center_x = (line.bbox[0] + line.bbox[2]) / 2.0
        if any(box[0] <= center_x <= box[2] and box[1] <= center_y <= box[3] for box in logo_boxes):
            continue
        kept_lines.append(line)
    if not kept_lines and not logos:
        return None
    sizes = [line.size for line in kept_lines] or [12.0]
    largest = max(sizes)
    blocks: list = []
    for line in kept_lines:
        if line.size >= max(16.0, largest * 0.82) and len(line.text) <= 90:
            role = "title"
            kind = "heading"
        elif line.size >= max(14.0, largest * 0.5) and len(line.text) <= 120:
            role = "subtitle"
            kind = "heading"
        else:
            role = "meta"
            kind = "paragraph"
        blocks.append(
            TextBlock(
                kind,
                line.bbox,
                line.text,
                line.size,
                line.bold or role == "title",
                _cover_align(line, page),
                0.0,
                role,
            )
        )
    flow: list = [*logos, *blocks]
    flow.sort(key=lambda block: (block.bbox[1], -block.bbox[2]))
    previous_bottom = 0.0
    for block in flow:
        gap = block.bbox[1] - previous_bottom
        block.gap_before = min(160.0, max(0.0, gap - 4.0))
        previous_bottom = max(previous_bottom, block.bbox[3])
    _width, height = _content_frame(page)
    _fit_cover_gaps(flow, height)
    return flow


def _read_page(
    page: fitz.Page,
    plumber_page,
    background_xrefs: set[int],
    furniture: set[tuple[str, str]],
    schema: _ColumnSchema,
    kind_out: dict,
    page_number: int,
    notes: list[str],
) -> list:
    """
    Read one page, compare its layout with the PDF, and retry up to twice.

    The first reading follows the text stream. A page whose ink does not sit
    in the same bands, or a ruled page that came back without a table, is
    read again from PDF blocks and then from the ruled grid. The closest
    reading is written. A retry cannot see the next page.
    """
    source = _body_atoms(page, _atoms_on_page(page))
    raw = _atoms_on_page(page)
    letters = sum(len(atom.text.strip()) for atom in raw)
    digital_layer = letters >= 40 or (letters >= 3 and not _page_has_images(page))
    trust_letters = digital_layer and not _encoding_broken(page, raw)
    # A cover kept as one picture loses every word on it. Only a cover with
    # no trustworthy text layer is kept that way; otherwise it is read like
    # any page, and its art is carried by the page background.
    if page_number == 1 and not trust_letters:
        cover = _read_cover(page, raw, background_xrefs)
        if cover is not None:
            kind_out["kind"] = "digital" if digital_layer else "scan"
            return cover

    def attempt(mode: str) -> tuple[list, dict]:
        local_kind: dict[str, str] = {}
        blocks = extract_page(page, plumber_page, background_xrefs, local_kind, mode)
        trial = _ColumnSchema()
        trial.columns = schema.columns
        blocks = _settle_page(page, blocks, furniture, trial)
        blocks = _pair_side_by_side(blocks)
        _assign_positions(page, blocks)
        report = _assess_page(page, blocks, source, trust_letters)
        visual = layout_discrepancy(page, _placed_boxes(page, blocks), _kept_ruled_table(blocks))
        report["visual"] = visual
        report["ruled"] = page_is_ruled(page)
        if visual > DISCREPANCY_LIMIT:
            report["ok"] = False
        report["score"] = report["score"] + (1.0 - visual) * 40.0
        report["kind"] = local_kind.get("kind", "digital")
        report["columns"] = trial.columns
        report["mode"] = mode
        return blocks, report

    def next_mode(used: set[str], report: dict) -> str | None:
        """The other reader. At most two retries, each a different strategy."""
        if report.get("ruled") and not report.get("kept_table") and "grid" not in used:
            return "grid"
        if report.get("visual", 0) > DISCREPANCY_LIMIT and "blocks" not in used:
            return "blocks"
        fallback = _retry_mode(report, digital_layer)
        if fallback not in used:
            return fallback
        for candidate in ("blocks", "grid", "plain", "ocr"):
            if candidate not in used:
                return candidate
        return None

    readings: list[tuple[list, dict]] = []
    used: set[str] = set()
    mode = "auto"
    for _retry in range(3):
        try:
            blocks, report = attempt(mode)
        except Exception as exc:
            notes.append(f"Page {page_number}: the {mode} reading failed. {exc}")
            report = {
                "ok": False,
                "score": -1,
                "coverage": 0.0,
                "layout": 0.0,
                "visual": 1.0,
                "fragmented": False,
                "extra": 0,
                "kind": "digital",
                "columns": schema.columns,
                "mode": mode,
            }
            blocks = []
        report["kept_table"] = _kept_ruled_table(blocks)
        readings.append((blocks, report))
        used.add(mode)
        if report.get("ok") and blocks:
            break
        if len(used) > 2:
            break
        mode = next_mode(used, report)
        if mode is None:
            break

    def _rank(item: tuple[list, dict]) -> tuple:
        blocks, report = item
        # A real grid that still has the page's words beats a paragraph dump
        # of the same page. Flattening the grid is what scattered its rows.
        kept = (
            _kept_ruled_table(blocks)
            and report.get("coverage", 0) >= 0.78
            and not report.get("extra")
        )
        return (report.get("score", -1) + (12.0 if kept else 0.0), 1 if kept else 0, 1 if blocks else 0)

    chosen_blocks, chosen = max(readings, key=_rank)
    if chosen.get("columns") is not None or schema.columns is None:
        schema.columns = chosen.get("columns", schema.columns)
    kind_out["kind"] = chosen.get("kind", "digital")
    if len(readings) == 1 and chosen.get("ok"):
        return chosen_blocks
    first = readings[0][1]
    if first.get("extra"):
        why = f"{first['extra']} extra table rows"
    elif first.get("fragmented"):
        why = "letters were split"
    elif first.get("coverage", 1) < 0.78:
        why = f"coverage {first.get('coverage', 0):.0%}"
    elif first.get("visual", 0) > DISCREPANCY_LIMIT:
        why = f"layout discrepancy {first.get('visual', 0):.0%}"
    elif first.get("layout", 1) < 0.60:
        why = f"layout match {first.get('layout', 0):.0%}"
    else:
        why = "the reading did not match the page"
    tried = ", ".join(item[1].get("mode", "") for item in readings)
    notes.append(
        f"Page {page_number}: {why}. Tried {tried} and kept {chosen.get('mode', 'auto')} "
        f"(coverage {chosen.get('coverage', 0):.0%}, layout discrepancy {chosen.get('visual', 0):.0%})."
    )
    return chosen_blocks


def _image_entries(page: fitz.Page) -> list[dict]:
    try:
        info = page.get_image_info(xrefs=True)
    except TypeError:
        info = page.get_image_info()
    return list(info or [])


def _full_bleed(bbox, page: fitz.Page) -> bool:
    """An image that runs out to the edges of the sheet is a background plate."""
    pad = 28.0
    return (
        float(bbox[0]) <= pad
        and float(bbox[1]) <= pad
        and float(bbox[2]) >= page.rect.width - pad
        and float(bbox[3]) >= page.rect.height - pad
    )


def _pale_image(data: bytes) -> bool:
    """A faint pattern or watermark is almost white. A chart is not."""
    try:
        from PIL import Image

        picture = Image.open(BytesIO(data)).convert("RGB")
    except Exception:
        return False
    picture.thumbnail((36, 36))
    pixels = list(picture.getdata())
    if not pixels:
        return False
    pale = sum(1 for red, green, blue in pixels if red >= 218 and green >= 218 and blue >= 218)
    return pale / len(pixels) >= 0.82


def _running_backgrounds(document: fitz.Document) -> set[int]:
    """
    The same decorative image on many pages is a background, not a figure.

    A watermark, a geometric pattern, or a header plate is stored once and
    painted on page after page. A diagram used once is left alone.
    """
    if document.page_count < 3:
        return set()
    counts: dict[int, int] = {}
    for index in range(document.page_count):
        page = document.load_page(index)
        seen: set[int] = set()
        for image in _image_entries(page):
            xref = int(image.get("xref") or 0)
            if not xref or xref in seen:
                continue
            seen.add(xref)
            counts[xref] = counts.get(xref, 0) + 1
    minimum = max(3, document.page_count // 3)
    return {xref for xref, count in counts.items() if count >= minimum}


def _figures(page: fitz.Page, atoms: list[_Atom], background_xrefs: set[int]) -> list[ImageBlock]:
    """
    Charts and diagrams, each kept whole.

    A full-page plate, a pale pattern, and a picture repeated across the
    file are backgrounds and are not written. Words that sit on a real
    figure are left in the picture instead of being copied out as paragraphs.
    """
    page_area = page.rect.width * page.rect.height or 1.0
    body_text = sum(len(atom.text.strip()) for atom in atoms) >= 24
    figures: list[ImageBlock] = []
    seen: set[int] = set()
    for image in _image_entries(page):
        bbox = image.get("bbox")
        if not bbox:
            continue
        width = float(bbox[2]) - float(bbox[0])
        height = float(bbox[3]) - float(bbox[1])
        area = width * height
        xref = int(image.get("xref") or 0)
        if not xref or xref in seen:
            continue
        bleed = _full_bleed(bbox, page)
        wide = area >= page_area * 0.45 or (
            width >= page.rect.width * 0.72 and height >= page.rect.height * 0.35
        )
        if xref in background_xrefs:
            continue
        # A corner ornament is painted past the edge of the sheet. A chart is not.
        if (
            body_text
            and (
                float(bbox[0]) < -8
                or float(bbox[1]) < -8
                or float(bbox[2]) > page.rect.width + 8
                or float(bbox[3]) > page.rect.height + 8
            )
        ):
            continue
        source_pixels = int(image.get("width") or 0) * int(image.get("height") or 0)
        if body_text and source_pixels and source_pixels < 80 and area > page_area * 0.04:
            continue
        if body_text and (bleed or wide):
            continue
        if width < 80 or height < 80 or area < page_area * 0.04:
            continue
        if area > page_area * 0.92 and body_text:
            continue
        try:
            data = _image_bytes(page, xref)
        except Exception:
            data = None
        if not data:
            continue
        if area >= page_area * 0.08 and _pale_image(data):
            continue
        seen.add(xref)
        figures.append(ImageBlock(tuple(float(value) for value in bbox), data, width, height))
    return figures


def _image_bytes(page: fitz.Page, xref: int) -> bytes | None:
    """
    One embedded picture as file bytes, with its transparency.

    A picture with a soft mask stores its see-through parts in the mask.
    Taking the raw picture alone paints those parts black (a logo on a black
    square). The mask is joined back on and the picture saved as PNG.
    """
    document = page.parent
    smask = 0
    try:
        for entry in page.get_images(full=True):
            if int(entry[0]) == int(xref):
                smask = int(entry[1] or 0)
                break
    except Exception:
        smask = 0
    if smask:
        try:
            base = fitz.Pixmap(document, xref)
            if base.alpha:
                base = fitz.Pixmap(base, 0)
            if base.colorspace is not None and base.colorspace.n not in (1, 3):
                base = fitz.Pixmap(fitz.csRGB, base)
            mask = fitz.Pixmap(document, smask)
            return fitz.Pixmap(base, mask).tobytes("png")
        except Exception:
            pass
        try:
            # Older PyMuPDF builds cannot join a mask to a picture; Pillow can.
            from PIL import Image

            base = fitz.Pixmap(document, xref)
            if base.alpha:
                base = fitz.Pixmap(base, 0)
            if base.colorspace is None or base.colorspace.n != 3:
                base = fitz.Pixmap(fitz.csRGB, base)
            mask = fitz.Pixmap(document, smask)
            if mask.alpha or (mask.colorspace is not None and mask.colorspace.n != 1):
                mask = fitz.Pixmap(fitz.csGRAY, mask)
            picture = Image.frombytes("RGB", (base.width, base.height), base.samples)
            alpha = Image.frombytes("L", (mask.width, mask.height), mask.samples)
            if alpha.size != picture.size:
                alpha = alpha.resize(picture.size)
            picture.putalpha(alpha)
            out = BytesIO()
            picture.save(out, format="PNG")
            return out.getvalue()
        except Exception:
            pass
    try:
        data = document.extract_image(xref).get("image")
    except Exception:
        return None
    return _word_readable(data, document, xref)


def _word_readable(data: bytes | None, document=None, xref: int = 0) -> bytes | None:
    """
    Picture bytes Word's writer can size.

    A JPEG that says it has 0 dots per inch, a CMYK JPEG or a format Word
    does not take (JPX, JBIG2) stops python-docx. Such a picture is drawn
    again as PNG.
    """
    if not data:
        return data
    try:
        from docx.image.image import Image as _DocxImage

        probe = _DocxImage.from_blob(data)
        probe.width, probe.height  # noqa: B018
        if probe.horz_dpi and probe.vert_dpi:
            return data
    except Exception:
        pass
    try:
        pixmap = fitz.Pixmap(document, xref) if document is not None and xref else fitz.Pixmap(data)
        if pixmap.colorspace is not None and pixmap.colorspace.n not in (1, 3):
            pixmap = fitz.Pixmap(fitz.csRGB, pixmap)
        return pixmap.tobytes("png")
    except Exception:
        return data


def text_direction(text: str) -> str:
    """rtl or ltr for this string, after the conversion's language choice."""
    mode = _DIRECTION.get()
    if mode == "ltr":
        return "ltr"
    if mode == "rtl":
        return "rtl"
    arabic = _arabic_count(text or "")
    latin = len(re.findall(r"[A-Za-z]", text or ""))
    if latin > arabic:
        return "ltr"
    return "rtl"


def _set_run_font(run, size_pt: float, bold: bool, text: str | None = None) -> None:
    """
    Point the complex-script slot at Traditional Arabic when the run is Arabic.

    The letters stay in logical Unicode. The font only decides how Word draws
    them. An English run stays on Arial and is not marked bidirectional.
    """
    sample = text if text is not None else (run.text or "")
    rtl = text_direction(sample) == "rtl"
    run.bold = bold
    run.font.size = Pt(size_pt)
    run.font.rtl = rtl
    run.font.complex_script = rtl
    run.font.cs_bold = bold
    properties = run._element.get_or_add_rPr()
    fonts = properties.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        properties.append(fonts)
    latin_font, arabic_font = output_fonts()
    fonts.set(qn("w:ascii"), latin_font)
    fonts.set(qn("w:hAnsi"), latin_font)
    fonts.set(qn("w:cs"), arabic_font)
    fonts.set(qn("w:eastAsia"), arabic_font)
    size_cs = properties.find(qn("w:szCs"))
    if size_cs is None:
        size_cs = OxmlElement("w:szCs")
        properties.append(size_cs)
    size_cs.set(qn("w:val"), str(int(round(size_pt * 2))))
    language = properties.find(qn("w:lang"))
    if language is None:
        language = OxmlElement("w:lang")
        properties.append(language)
    language.set(qn("w:bidi"), "ar-SA" if rtl else "en-US")
    language.set(qn("w:val"), "ar-SA" if rtl else "en-US")


def _set_paragraph_ltr(paragraph) -> None:
    """An English paragraph starts at the left and is not bidirectional."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    properties = paragraph._p.get_or_add_pPr()
    # An explicit "off": a removed w:bidi would inherit the style's or the
    # document default's right-to-left setting and mirror the line.
    bidi = properties.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        anchor = properties.find(qn("w:jc"))
        if anchor is not None:
            anchor.addprevious(bidi)
        else:
            properties.append(bidi)
    bidi.set(qn("w:val"), "0")
    justification = properties.find(qn("w:jc"))
    if justification is None:
        justification = OxmlElement("w:jc")
        properties.append(justification)
    justification.set(qn("w:val"), "left")
    mark = properties.find(qn("w:rPr"))
    if mark is not None and mark.find(qn("w:rtl")) is not None:
        mark.remove(mark.find(qn("w:rtl")))
    rtl_property = getattr(type(paragraph.paragraph_format), "rtl", None)
    if isinstance(rtl_property, property):
        paragraph.paragraph_format.rtl = False
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.widow_control = False


def _set_paragraph_rtl(paragraph, text: str = "") -> None:
    """
    Make the paragraph start at the right margin.

    Word reverses the words left and right once a paragraph is bidirectional.
    Storing jc=right is what pushed these lines to the left edge. jc=left is
    the value Word then draws on the right, which is where an Arabic line
    starts. w:bidi is the paragraph direction. The paragraph mark is marked
    rtl as well, which is the flag Word stores when the direction button is on.
    """
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    properties = paragraph._p.get_or_add_pPr()
    justification = properties.find(qn("w:jc"))
    if justification is None:
        justification = OxmlElement("w:jc")
        properties.append(justification)
    justification.set(qn("w:val"), "left")
    bidi = properties.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        properties.append(bidi)
    bidi.set(qn("w:val"), "1")
    mark = properties.find(qn("w:rPr"))
    if mark is None:
        mark = OxmlElement("w:rPr")
        properties.append(mark)
    if mark.find(qn("w:rtl")) is None:
        mark.append(OxmlElement("w:rtl"))
    rtl_property = getattr(type(paragraph.paragraph_format), "rtl", None)
    if isinstance(rtl_property, property):
        paragraph.paragraph_format.rtl = True
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.widow_control = False


def _set_block_align(paragraph, align: str) -> None:
    """
    Place the line where its box sat.

    With the paragraph set bidirectional, jc=left is the right edge and
    jc=right is the left edge. Center stays center, so Arabic still shapes.
    """
    if align == "center":
        _set_cover_center(paragraph)
        return
    if align != "left":
        return
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    properties = paragraph._p.get_or_add_pPr()
    justification = properties.find(qn("w:jc"))
    if justification is None:
        justification = OxmlElement("w:jc")
        properties.append(justification)
    justification.set(qn("w:val"), "right")


def _set_cover_center(paragraph) -> None:
    """Center the line and keep the paragraph bidirectional, so Arabic still shapes."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    properties = paragraph._p.get_or_add_pPr()
    justification = properties.find(qn("w:jc"))
    if justification is None:
        justification = OxmlElement("w:jc")
        properties.append(justification)
    justification.set(qn("w:val"), "center")


def _shade_box(properties, color: str) -> None:
    """Paint a paragraph or a cell with one solid color from the PDF."""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), color)
    existing = properties.find(qn("w:shd"))
    if existing is not None:
        properties.remove(existing)
    properties.append(shading)


def _ink_on(color: str) -> RGBColor | None:
    """White letters on a dark bar. A light bar keeps the normal ink."""
    try:
        red = int(color[0:2], 16)
        green = int(color[2:4], 16)
        blue = int(color[4:6], 16)
    except (TypeError, ValueError):
        return None
    if red * 0.30 + green * 0.59 + blue * 0.11 < 150:
        return RGBColor(255, 255, 255)
    return None


def _add_page_field(paragraph) -> None:
    """A Word page number that updates on every sheet."""
    begin = paragraph.add_run()
    start = OxmlElement("w:fldChar")
    start.set(qn("w:fldCharType"), "begin")
    begin._r.append(start)
    instruction = paragraph.add_run()
    text = OxmlElement("w:instrText")
    text.set(qn("xml:space"), "preserve")
    text.text = " PAGE "
    instruction._r.append(text)
    end = paragraph.add_run()
    stop = OxmlElement("w:fldChar")
    stop.set(qn("w:fldCharType"), "end")
    end._r.append(stop)


def _apply_faithful_page_footer(
    section,
    chrome: "_PageChrome",
    label_format: str,
    side: str,
    section_bidi: bool = False,
    place: str = "footer",
    start: int | None = None,
) -> None:
    """
    Write one consistent PAGE field footer (or header) for faithful pages.

    `place` is where the PDF printed the label. `start` is the number the PDF
    printed on the section's first page, so the live field shows the same
    number instead of counting from 1.

    Formats such as '{PAGE} | Page' match thesis exports where the PDF printed
    '8 | Page'. Every page of the section gets the same separator and direction.
    """
    # Keep room for the footer line under the positioned content.
    try:
        section.bottom_margin = Emu(max(int(section.bottom_margin), int(36 * 12700)))
        section.footer_distance = Emu(int(14 * 12700))
    except Exception:
        pass
    if place == "header":
        try:
            section.header_distance = Emu(int(14 * 12700))
        except Exception:
            pass
        section.header.is_linked_to_previous = False
        footer = section.header
    else:
        footer = section.footer
    _clear_story(footer)
    if start is not None:
        _set_page_start(section, start)
    paragraph = footer.paragraphs[0]
    paragraph.text = ""
    _clear_paragraph_borders(paragraph)
    if section_bidi and side == "left":
        # RTL section: left slot is still the left edge of the page.
        _set_paragraph_ltr(paragraph)
    else:
        _set_paragraph_ltr(paragraph)
    if side == "right":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif side == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fmt = label_format or "{PAGE} | Page"
    # Split around the PAGE placeholder; write a live field in its place.
    parts = re.split(r"\{PAGE\}", fmt, maxsplit=1)
    before = parts[0] if parts else ""
    after = parts[1] if len(parts) > 1 else ""
    if before:
        run = paragraph.add_run(before)
        _set_run_font(run, 9.0, False, before)
    _add_page_field(paragraph)
    if after:
        run = paragraph.add_run(after)
        _set_run_font(run, 9.0, False, after)
    for run in paragraph.runs:
        _set_run_font(run, 9.0, False)


def _set_page_start(section, number: int) -> None:
    """Make the section's PAGE field count from `number`."""
    properties = section._sectPr
    numbering = properties.find(qn("w:pgNumType"))
    if numbering is None:
        numbering = OxmlElement("w:pgNumType")
        # pgNumType sits before cols/docGrid in sectPr.
        anchor = properties.find(qn("w:cols"))
        if anchor is None:
            anchor = properties.find(qn("w:docGrid"))
        if anchor is not None:
            anchor.addprevious(numbering)
        else:
            properties.append(numbering)
    numbering.set(qn("w:start"), str(max(0, int(number))))


def _add_text(document: Document, block: TextBlock, content_width_pt: float = 460.0) -> None:
    """Write one paragraph. Headings are the same style, only larger and bold."""
    if block.role == "toc" and block.toc_page:
        title_text = _clean_flow_text(block.text)
        title_text = re.sub(r"(?:\.\s*){2,}|[.\u00b7\u2022\u2024\u2026‧]{2,}", " ", title_text)
        title_text = re.sub(r"\s+", " ", title_text).strip(" .\t")
        number_text = _clean_flow_text(block.toc_page)
        if not title_text or not number_text:
            return
    else:
        title_text = ""
        number_text = ""
        written = _clean_flow_text(block.text)
        if not written or not re.search(r"[0-9A-Za-z\u0600-\u06FF]", written):
            return
        block.text = written
    paragraph = document.add_paragraph()
    paragraph.style = document.styles["Normal"]
    _set_paragraph_rtl(paragraph, block.text)
    if block.role == "toc" and block.toc_page:
        size = min(max(block.size, 11.0), 14.0)
        title = paragraph.add_run(title_text)
        _set_run_font(title, size, block.bold)
        paragraph.add_run("\t")
        number = paragraph.add_run(number_text)
        _set_run_font(number, size, False)
        # One right-aligned stop, the same distance on every contents line.
        # A bidirectional paragraph measures that stop from the right, so the
        # page numbers share a single vertical edge and the dots fill the gap.
        paragraph.paragraph_format.tab_stops.add_tab_stop(
            Pt(max(120.0, content_width_pt - 6.0)),
            WD_TAB_ALIGNMENT.RIGHT,
            WD_TAB_LEADER.DOTS,
        )
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.keep_together = True
        return
    _set_block_align(paragraph, block.align)
    if block.role:
        size = min(max(block.size, 8.0), 28.0)
    elif block.kind == "heading":
        size = min(max(block.size, 8.0), 16.0)
    else:
        size = min(max(block.size, 8.0), 14.0)
    written = block.text
    if written[:1] in "•●▪◦" and not written[1:2].isspace():
        written = written[:1] + " " + written[1:]
    run = paragraph.add_run(written)
    _set_run_font(run, size, block.bold or block.kind == "heading" or block.role == "title")
    if block.role == "title":
        run.font.color.rgb = RGBColor(0x1F, 0x2A, 0x37)
    elif block.role == "subtitle":
        run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    elif block.kind == "heading":
        run.font.color.rgb = RGBColor(0x1F, 0x2A, 0x37)
    paragraph.paragraph_format.space_before = Pt(max(0.0, block.gap_before))
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    if block.fill:
        _shade_box(paragraph._p.get_or_add_pPr(), block.fill)
        ink = _ink_on(block.fill)
        if ink is not None:
            run.font.color.rgb = ink


def _shade_header(cell) -> None:
    """Give the first row a light background so the header is easy to see."""
    properties = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), "E8EEF4")
    existing = properties.find(qn("w:shd"))
    if existing is not None:
        properties.remove(existing)
    properties.append(shading)


def _set_cell_width(cell, width_dxa: int, valign: str = "center") -> None:
    properties = cell._tc.get_or_add_tcPr()
    width = OxmlElement("w:tcW")
    width.set(qn("w:w"), str(width_dxa))
    width.set(qn("w:type"), "dxa")
    existing = properties.find(qn("w:tcW"))
    if existing is not None:
        properties.remove(existing)
    properties.append(width)
    align = OxmlElement("w:vAlign")
    align.set(qn("w:val"), valign)
    existing_align = properties.find(qn("w:vAlign"))
    if existing_align is not None:
        properties.remove(existing_align)
    properties.append(align)


def _border(name: str, show: bool = True):
    edge = OxmlElement(name)
    edge.set(qn("w:val"), "single" if show else "nil")
    edge.set(qn("w:sz"), "8" if show else "0")
    edge.set(qn("w:space"), "0")
    edge.set(qn("w:color"), "8AA0B4" if show else "auto")
    return edge


def _set_table_borders(table, show: bool = True) -> None:
    """Draw a full grid. The style alone is not enough in every Word version."""
    properties = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge_name in ("w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"):
        borders.append(_border(edge_name, show))
    existing = properties.find(qn("w:tblBorders"))
    if existing is not None:
        properties.remove(existing)
    properties.append(borders)


def _definition_weights(rows: list[list[str]], widths: list[float], columns: int) -> list[float] | None:
    """
    A term-and-definition grid: the short column is 30 percent, the long one 70.

    A balanced two-column grid keeps the widths it was drawn with.
    """
    if columns != 2 or not rows:
        return None
    left_lengths: list[int] = []
    right_lengths: list[int] = []
    for row in rows:
        padded = list(row) + [""] * (2 - len(row))
        left_lengths.append(len(str(padded[0]).strip()))
        right_lengths.append(len(str(padded[1]).strip()))
    if not right_lengths or sum(right_lengths) <= 0:
        return None
    left_avg = sum(left_lengths) / len(left_lengths)
    right_avg = sum(right_lengths) / len(right_lengths)
    narrow = left_avg <= max(8.0, right_avg * 0.55) and left_avg <= 42
    measured = False
    if widths and len(widths) >= 2 and sum(widths[:2]) > 0:
        measured = widths[0] / sum(widths[:2]) <= 0.42
    if narrow or measured:
        return [30.0, 70.0]
    return None


def _pad_table_cells(table, horizontal: int = 80, vertical: int = 50) -> None:
    """Keep Arabic inside the cell border. Values are twips, about four and two and a half points."""
    edges = (("w:top", vertical), ("w:left", horizontal), ("w:bottom", vertical), ("w:right", horizontal))
    properties = table._tbl.tblPr
    margins = OxmlElement("w:tblCellMar")
    for edge, value in edges:
        node = OxmlElement(edge)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    existing = properties.find(qn("w:tblCellMar"))
    if existing is not None:
        properties.remove(existing)
    properties.append(margins)
    for cell in table._tbl.findall(".//" + qn("w:tc")):
        cell_properties = cell.find(qn("w:tcPr"))
        if cell_properties is None:
            cell_properties = OxmlElement("w:tcPr")
            cell.insert(0, cell_properties)
        padding = OxmlElement("w:tcMar")
        for edge, value in edges:
            node = OxmlElement(edge)
            node.set(qn("w:w"), str(value))
            node.set(qn("w:type"), "dxa")
            padding.append(node)
        old = cell_properties.find(qn("w:tcMar"))
        if old is not None:
            cell_properties.remove(old)
        cell_properties.append(padding)


def _add_table(
    document: Document,
    block: TableBlock,
    content_width_emu: int,
    max_height_pt: float | None = None,
) -> None:
    """
    Rebuild the grid with bordered cells.

    Columns stay in the order they were printed, from the left of the page to
    the right. Setting the table itself to bidiVisual would mirror that grid.
    Each cell's paragraph is still right-to-left, so the Arabic inside a cell
    starts at the right edge of that cell.
    """
    rows = block.rows
    if not rows:
        return
    columns = max(len(row) for row in rows)
    rows = [list(row) + [""] * (columns - len(row)) for row in rows]
    table = document.add_table(rows=len(rows), cols=columns)
    table.autofit = False
    table.allow_autofit = False
    if block.rules:
        try:
            table.style = "Table Grid"
        except KeyError:
            pass
    _set_table_borders(table, show=block.rules)

    # A definition pair is locked at 30/70. Any other grid keeps its own widths.
    defined = _definition_weights(rows, block.widths, columns)
    if defined:
        weights = defined
    elif block.widths and len(block.widths) == columns:
        weights = [max(float(value), 1.0) for value in block.widths]
    else:
        weights = []
        for index in range(columns):
            longest = max((len(row[index]) for row in rows), default=1)
            weights.append(float(max(longest, 6)))
    total = sum(weights) or 1.0
    width_dxa = max(1, int(content_width_emu / 635))
    column_dxa = [max(600, int(width_dxa * weight / total)) for weight in weights]
    scale = width_dxa / max(sum(column_dxa), 1)
    column_dxa = [max(500, int(value * scale)) for value in column_dxa]

    properties = table._tbl.tblPr
    table_width = OxmlElement("w:tblW")
    table_width.set(qn("w:w"), str(sum(column_dxa)))
    table_width.set(qn("w:type"), "dxa")
    existing_width = properties.find(qn("w:tblW"))
    if existing_width is not None:
        properties.remove(existing_width)
    properties.append(table_width)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    existing_layout = properties.find(qn("w:tblLayout"))
    if existing_layout is not None:
        properties.remove(existing_layout)
    properties.append(layout)

    # Printed order, left to right, said outright: in a right-to-left
    # section some readers mirror a table that does not say.
    visual = OxmlElement("w:bidiVisual")
    visual.set(qn("w:val"), "0")
    properties.insert(0, visual)
    try:
        from core.faithful import _order_tbl_pr

        _order_tbl_pr(properties)
    except Exception:
        pass

    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for node, dxa in zip(grid.findall(qn("w:gridCol")), column_dxa):
            node.set(qn("w:w"), str(dxa))

    # The first row is a header only when its cells read like headings; a
    # table continued from a picture can start with an ordinary data row.
    first = [(_clean_flow_text(cell) or "") for cell in rows[0]] if rows else []
    header_row = bool(first) and max((len(cell) for cell in first), default=0) <= 60
    for row_index, row in enumerate(rows):
        # A row of long or several-line cells reads from the top, as printed.
        tall = any(len(text or "") > 90 or (text or "").count("\n") >= 2 for text in row)
        for col_index, text in enumerate(row):
            cell = table.cell(row_index, col_index)
            horizontal, vertical = (block.aligns or {}).get((row_index, col_index), ("", "top" if tall else "center"))
            _set_cell_width(cell, column_dxa[col_index], vertical)
            paragraph = cell.paragraphs[0]
            paragraph.style = document.styles["Normal"]
            # Each cell runs in the direction of its own letters: an English
            # cell starts at the left even in an Arabic document.
            latin_cell = len(re.findall(r"[A-Za-z]", text or "")) > _arabic_count(text or "") or not _arabic_count(text or "")
            if text and latin_cell and _DIRECTION.get() != "rtl" or text and len(re.findall(r"[A-Za-z]", text)) > max(3, _arabic_count(text) * 3):
                _set_paragraph_ltr(paragraph)
            else:
                _set_paragraph_rtl(paragraph, text)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            if block.centered or horizontal == "center":
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # A cell's own paragraphs stay apart (a title, a name, "Venue:").
            parts = [_clean_flow_text(part) for part in (text or "").split("\n")]
            parts = [part for part in parts if part]
            text = " ".join(parts)
            font_size = max(7.5, 11.0 * block.fit_scale)
            if getattr(block, "text_font", ""):
                from core.flow_layout import _width_factor

                font_size *= _width_factor(text or "", block.text_font)
            run = None
            runs = []
            styled = (block.styles or {}).get((row_index, col_index))
            plain_bold = row_index == 0 and header_row and block.rules and not block.fills
            if styled:
                for piece, bold, italic, colour in styled:
                    if piece == "\n":
                        paragraph.add_run().add_break()
                        continue
                    run = paragraph.add_run(piece)
                    runs.append(run)
                    _set_run_font(run, font_size, bool(bold), piece)
                    if italic:
                        run.italic = True
                    if colour and re.fullmatch(r"#?[0-9A-Fa-f]{6}", colour) and colour.lstrip("#").upper() not in {"000000"}:
                        run.font.color.rgb = RGBColor.from_string(colour.lstrip("#").upper())
            else:
                for part_index, part in enumerate(parts):
                    if part_index:
                        paragraph.add_run().add_break()
                    run = paragraph.add_run(part)
                    runs.append(run)
                    _set_run_font(run, font_size, plain_bold)
            if text and re.fullmatch(r"[\d٠-٩]{1,4}", text) and horizontal != "center" and not block.centered:
                # A page number stays on the right edge of its cell.
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                justification = paragraph._p.get_or_add_pPr().find(qn("w:jc"))
                if justification is not None:
                    justification.set(qn("w:val"), "left")
            fill = ""
            if row_index < len(block.fills) and col_index < len(block.fills[row_index]):
                fill = block.fills[row_index][col_index]
            colored = any(color for paints in block.fills for color in paints)
            if fill:
                _shade_box(cell._tc.get_or_add_tcPr(), fill)
                ink = _ink_on(fill)
                if ink is not None:
                    for each in runs:
                        each.font.color.rgb = ink
            elif row_index == 0 and header_row and block.rules and not colored:
                _shade_header(cell)
        _lock_row(table.rows[row_index])
    if block.centered and not block.rules:
        # Cards: each cell framed on its own, in a light dashed line.
        for row in table.rows:
            for cell in row.cells:
                properties = cell._tc.get_or_add_tcPr()
                borders = OxmlElement("w:tcBorders")
                for side in ("top", "left", "bottom", "right"):
                    edge = OxmlElement(f"w:{side}")
                    edge.set(qn("w:val"), "dashed")
                    edge.set(qn("w:sz"), "8")
                    edge.set(qn("w:space"), "0")
                    edge.set(qn("w:color"), "A6A6A6")
                    borders.append(edge)
                old_borders = properties.find(qn("w:tcBorders"))
                if old_borders is not None:
                    properties.remove(old_borders)
                width_node = properties.find(qn("w:tcW"))
                if width_node is not None:
                    width_node.addnext(borders)
                else:
                    properties.insert(0, borders)
    for row_index, column, data, width_pt, height_pt in block.images or []:
        if row_index < len(rows) and column < columns:
            _cell_picture(table.cell(row_index, column), data, width_pt, column_dxa[column] / 20.0)
    for row_index, row_spans in enumerate(block.spans or []):
        if row_index >= len(rows):
            break
        for first, last in row_spans:
            if last > first and last < columns:
                try:
                    table.cell(row_index, first).merge(table.cell(row_index, last))
                except Exception:
                    pass
    _pad_table_cells(table, 100 if defined else 80, 60 if defined else 40)
    _fit_table_height(table, _row_heights(block, content_width_emu / 12700.0, max_height_pt))
    grows = {entry[0] for entry in block.images or []}
    if block.spans:
        # A table read from its drawn borders: rows keep at least their
        # printed room and grow when the Word font needs more.
        grows = set(range(len(table.rows)))
    for row_index in grows:
        # A row with a picture grows to hold it.
        if row_index < len(table.rows):
            height = table.rows[row_index]._tr.get_or_add_trPr().find(qn("w:trHeight"))
            if height is not None:
                height.set(qn("w:hRule"), "atLeast")


def _cell_picture(cell, data: bytes, width_pt: float, room_pt: float) -> None:
    """Put a picture at the top of a cell, at its printed size or the cell's width."""
    from io import BytesIO as _BytesIO

    width = max(12.0, min(width_pt, room_pt - 8.0))
    try:
        paragraph = cell.paragraphs[0]
        holder = paragraph.insert_paragraph_before() if paragraph.text.strip() else paragraph
        holder.alignment = WD_ALIGN_PARAGRAPH.CENTER
        holder.add_run().add_picture(_BytesIO(data), width=Pt(width))
    except Exception:
        pass


def _lock_row(row) -> None:
    """Keep a row in one piece so Word does not slice it onto the next page."""
    properties = row._tr.get_or_add_trPr()
    if properties.find(qn("w:cantSplit")) is None:
        properties.append(OxmlElement("w:cantSplit"))


def _line_count(text: str, width_pt: float, size: float) -> int:
    """How many lines Word will need for this text at this width."""
    compact = len(text.strip())
    if compact == 0:
        return 1
    per_line = max(6, int(width_pt / max(size * 0.55, 4.0)))
    return max(1, (compact + per_line - 1) // per_line)


def _row_heights(block: TableBlock, width_pt: float, max_height_pt: float | None) -> list[float]:
    """One height per row. The total is never taller than the room on this page."""
    font = max(7.5, 11.0 * block.fit_scale)
    columns = max((len(row) for row in block.rows), default=1)
    if block.widths and len(block.widths) == columns:
        weight_total = sum(block.widths) or 1.0
        widths = [max(24.0, width_pt * (value / weight_total)) for value in block.widths]
    else:
        widths = [width_pt / columns] * columns
    heights: list[float] = []
    for row in block.rows:
        lines = 1
        for index, cell in enumerate(row):
            column_width = widths[index] if index < len(widths) else width_pt / columns
            lines = max(lines, sum(_line_count(part, column_width * 0.92, font) for part in (cell or "").split("\n")))
        heights.append(lines * font * 1.25 + 2.0)
    total = sum(heights)
    if max_height_pt and total > max_height_pt > 0:
        factor = max_height_pt / total
        heights = [height * factor for height in heights]
    return heights


def _fit_table_height(table, heights: list[float]) -> None:
    """
    Lock each row to the height it was given.

    Word otherwise grows a wrapped cell until the grid spills onto the next page.
    """
    if not heights:
        return
    for row, height_pt in zip(table.rows, heights):
        properties = row._tr.get_or_add_trPr()
        existing = properties.find(qn("w:trHeight"))
        if existing is not None:
            properties.remove(existing)
        height = OxmlElement("w:trHeight")
        height.set(qn("w:val"), str(max(80, int(height_pt * 20))))
        height.set(qn("w:hRule"), "exact")
        properties.append(height)


def _zero_cell_margins(table) -> None:
    """No padding inside the cage, so the picture cannot shift inside the cell."""
    properties = table._tbl.tblPr
    margins = OxmlElement("w:tblCellMar")
    for edge in ("w:top", "w:left", "w:bottom", "w:right"):
        node = OxmlElement(edge)
        node.set(qn("w:w"), "0")
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    existing = properties.find(qn("w:tblCellMar"))
    if existing is not None:
        properties.remove(existing)
    properties.append(margins)


def _add_image_cage(
    document: Document,
    widths_pt: list[float],
    image_index: int,
    image: bytes,
    image_width_pt: float,
    image_height_pt: float,
) -> None:
    """
    A borderless fixed table that holds one figure in place.

    The columns sum to the content width. Empty cells on either side are the
    PDF offset, so the picture stays where it was drawn and the text above
    and below cannot slide into it.
    """
    table = document.add_table(rows=1, cols=len(widths_pt))
    table.autofit = False
    table.allow_autofit = False
    _set_table_borders(table, show=False)
    _zero_cell_margins(table)
    dxa = [max(80, int(round(width * 20))) for width in widths_pt]
    total = sum(dxa)
    properties = table._tbl.tblPr
    table_width = OxmlElement("w:tblW")
    table_width.set(qn("w:w"), str(total))
    table_width.set(qn("w:type"), "dxa")
    existing_width = properties.find(qn("w:tblW"))
    if existing_width is not None:
        properties.remove(existing_width)
    properties.append(table_width)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    existing_layout = properties.find(qn("w:tblLayout"))
    if existing_layout is not None:
        properties.remove(existing_layout)
    properties.append(layout)
    justification = properties.find(qn("w:jc"))
    if justification is None:
        justification = OxmlElement("w:jc")
        properties.append(justification)
    justification.set(qn("w:val"), "left")
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for node, width in zip(grid.findall(qn("w:gridCol")), dxa):
            node.set(qn("w:w"), str(width))
    for index, width in enumerate(dxa):
        cell = table.cell(0, index)
        _set_cell_width(cell, width)
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        if index == image_index:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            picture = max(16.0, image_width_pt - 1.0)
            paragraph.add_run().add_picture(BytesIO(image), width=Emu(int(picture * 12700)))
    _lock_row(table.rows[0])
    _fit_table_height(table, [image_height_pt])


def _reserve_height(paragraph, before_pt: float, line_pt: float) -> None:
    """Hold a vertical slot. The picture itself is positioned on the page."""
    properties = paragraph._p.get_or_add_pPr()
    spacing = properties.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        properties.append(spacing)
    spacing.set(qn("w:before"), str(max(0, int(before_pt * 20))))
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(max(20, int(line_pt * 20))))
    spacing.set(qn("w:lineRule"), "exact")


def _add_image(document: Document, block: ImageBlock, content_width_emu: int, content_height_pt: float) -> None:
    """Place one figure. A cover fills the page. Any other figure keeps its PDF box."""
    paragraph = document.add_paragraph()
    if block.align == "center" or block.fit == "page":
        _set_cover_center(paragraph)
    else:
        _set_paragraph_rtl(paragraph, "")
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    if block.fit == "page":
        paragraph.paragraph_format.line_spacing = 1.0
        width_pt = content_width_emu / 12700
        if block.width > 0 and block.height > 0 and content_height_pt > 0:
            height_pt = width_pt * block.height / block.width
            limit = content_height_pt * 0.92
            if height_pt > limit:
                width_pt = width_pt * limit / height_pt
        if width_pt >= 24:
            paragraph.add_run().add_picture(BytesIO(block.data), width=Emu(int(width_pt * 12700)))
        return
    width_pt = max(0.0, min(block.width, content_width_emu / 12700.0))
    height_pt = block.height if block.height > 0 else 0.0
    if block.width > 0 and width_pt < block.width and height_pt:
        height_pt = height_pt * (width_pt / block.width)
    if content_height_pt > 0 and height_pt > content_height_pt * 0.85:
        height_pt = content_height_pt * 0.85
        if block.height > 0:
            width_pt = min(width_pt, block.width * (height_pt / block.height))
    if width_pt < 24 or height_pt < 16:
        paragraph._p.getparent().remove(paragraph._p)
        return
    # The opening paragraph is only the gap. The cage itself holds the picture.
    if block.gap_before >= 1.0:
        _reserve_height(paragraph, block.gap_before, 1.0)
    else:
        paragraph._p.getparent().remove(paragraph._p)
    content_pt = content_width_emu / 12700.0
    left_margin = 0.0
    if document.sections:
        left_margin = float(document.sections[-1].left_margin.pt)
    offset = max(0.0, float(block.bbox[0]) - left_margin)
    if offset + width_pt > content_pt:
        offset = max(0.0, content_pt - width_pt)
    remainder = max(0.0, content_pt - offset - width_pt)
    columns: list[float] = []
    if offset >= 6.0:
        columns.append(offset)
    columns.append(max(24.0, width_pt))
    image_index = len(columns) - 1
    if remainder >= 6.0:
        columns.append(remainder)
    _add_image_cage(document, columns, image_index, block.data, width_pt, height_pt)


def _mark_document_ltr(document: Document) -> None:
    """An English document stays left to right. Paragraphs are not bidirectional."""
    normal = document.styles["Normal"]
    normal_properties = normal.element.get_or_add_pPr()
    bidi = normal_properties.find(qn("w:bidi"))
    if bidi is not None:
        normal_properties.remove(bidi)
    normal_jc = normal_properties.find(qn("w:jc"))
    if normal_jc is None:
        normal_jc = OxmlElement("w:jc")
        normal_properties.append(normal_jc)
    normal_jc.set(qn("w:val"), "left")
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.0
    normal.paragraph_format.widow_control = False
    settings = document.settings.element
    existing = settings.find(qn("w:bidi"))
    if existing is not None:
        settings.remove(existing)


def _mark_document_rtl(document: Document) -> None:
    """Tell Word the file itself is a right-to-left Arabic document."""
    normal = document.styles["Normal"]
    normal_properties = normal.element.get_or_add_pPr()
    if normal_properties.find(qn("w:bidi")) is None:
        normal_properties.append(OxmlElement("w:bidi"))
    normal_jc = normal_properties.find(qn("w:jc"))
    if normal_jc is None:
        normal_jc = OxmlElement("w:jc")
        normal_properties.append(normal_jc)
    normal_jc.set(qn("w:val"), "left")
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.0
    normal.paragraph_format.widow_control = False
    settings = document.settings.element
    if settings.find(qn("w:bidi")) is None:
        settings.append(OxmlElement("w:bidi"))
    theme = settings.find(qn("w:themeFontLang"))
    if theme is None:
        theme = OxmlElement("w:themeFontLang")
        settings.append(theme)
    theme.set(qn("w:val"), "en-US")
    theme.set(qn("w:bidi"), "ar-SA")


def _collapse_section_breaks(document: Document) -> None:
    """
    Keep the section break from opening a blank page.

    Word stores the break on an empty paragraph. The Normal style's spacing
    makes that paragraph taller than the room left on a full page, so each
    PDF page was followed by an empty Word page.
    """
    for paragraph in document.element.body.findall(qn("w:p")):
        properties = paragraph.find(qn("w:pPr"))
        if properties is None or properties.find(qn("w:sectPr")) is None:
            continue
        spacing = properties.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            properties.append(spacing)
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "20")
        spacing.set(qn("w:lineRule"), "exact")


def _clamp_margin(points: float, low: float, high: float) -> float:
    return min(high, max(low, points))


def _body_insets(page: fitz.Page) -> tuple[float, float, float, float]:
    """
    Left, top, right and bottom padding, in points.

    Left and right are the outermost body ink, which is the page margin.
    Top and bottom stay small: the header and footer occupy that edge, and
    the gap down to the first line is paragraph spacing, not a margin.
    """
    width = float(page.rect.width)
    height = float(page.rect.height)
    lefts: list[float] = []
    rights: list[float] = []
    for atom in _atoms_on_page(page):
        if not atom.text.strip():
            continue
        center_y = (atom.y0 + atom.y1) / 2.0
        if center_y <= height * _CHROME_BAND or center_y >= height * (1.0 - _CHROME_BAND):
            continue
        lefts.append(atom.x0)
        rights.append(max(0.0, width - atom.x1))
    for image in _image_entries(page):
        bbox = image.get("bbox")
        if not bbox:
            continue
        box = (float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3]))
        if _chrome_edge(box, height) is not None:
            continue
        if box[3] - box[1] < 12 or box[2] - box[0] < 12:
            continue
        lefts.append(box[0])
        rights.append(max(0.0, width - box[2]))
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        drawings = []
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None:
            continue
        box = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
        shape_width = box[2] - box[0]
        shape_height = box[3] - box[1]
        if shape_width > width * 0.92 or shape_height < 10 or shape_height > height * 0.75:
            continue
        if _chrome_edge(box, height) is not None or box[1] < height * _CHROME_BAND:
            continue
        lefts.append(box[0])
        rights.append(max(0.0, width - box[2]))
    if not lefts:
        return (46.0, 48.0, 46.0, 40.0)
    return (
        _clamp_margin(min(lefts), 28.0, 90.0),
        42.0,
        _clamp_margin(min(rights), 28.0, 90.0),
        36.0,
    )


def _header_body_top(page: fitz.Page) -> float:
    """The top of the first real line. The logo crop stops above it."""
    height = float(page.rect.height)
    cap = height * 0.125
    # Each atom is one glyph: any letter or digit of body size counts. A
    # filled box of the body (a table cell, a banner) counts too, so a table
    # that starts high on the page is never cut into the header picture.
    tops = [
        atom.y0
        for atom in _atoms_on_page(page)
        if atom.text.strip()
        and atom.y0 > height * 0.06
        and float(atom.size or 0) >= 6
        and any(char.isalnum() for char in atom.text)
    ]
    try:
        from core.table_rebuild import is_box, visible_drawings

        for drawing in visible_drawings(page):
            rect = drawing.get("rect")
            fill = drawing.get("fill")
            if (
                rect is not None and fill and is_box(drawing)
                and rect.y0 > height * 0.06 and rect.width >= 40 and rect.height >= 10
                and not all(channel >= 0.94 for channel in fill[:3])
            ):
                tops.append(float(rect.y0))
    except Exception:
        pass
    if not tops:
        return cap
    return min(cap, min(tops) - 3.0)


def _header_art(page: fitz.Page) -> tuple[float, float, float, float] | None:
    """The logo's box. A mark that starts in the top 5% and stays in the header."""
    height = float(page.rect.height)
    band = height * _CHROME_BAND
    boxes: list[tuple[float, float, float, float]] = []
    for image in _image_entries(page):
        bbox = image.get("bbox")
        if not bbox:
            continue
        box = (float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3]))
        if box[1] < band and band < box[3] <= height * 0.16 and (box[3] - box[1]) >= 8:
            boxes.append(box)
        elif box[3] <= band and (box[3] - box[1]) >= 8 and (box[2] - box[0]) >= 8:
            boxes.append(box)
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        drawings = []
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None:
            continue
        box = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
        shape_height = box[3] - box[1]
        shape_width = box[2] - box[0]
        if shape_height < 8 and shape_width > page.rect.width * 0.55:
            continue
        if shape_height < 8 or shape_width < 8:
            continue
        if box[3] > height * 0.16 or box[1] > band:
            continue
        fill = drawing.get("fill")
        if _near_white(fill) and shape_height < 14:
            continue
        boxes.append(box)
    if not boxes:
        return None
    pad = 2.0
    return (
        max(0.0, min(box[0] for box in boxes) - pad),
        max(0.0, min(box[1] for box in boxes) - pad),
        min(float(page.rect.width), max(box[2] for box in boxes) + pad),
        min(height * 0.16, _header_body_top(page), max(box[3] for box in boxes) + pad),
    )


def _header_strip(page: fitz.Page) -> tuple[bytes, float, float, str, tuple[float, float, float, float]] | None:
    """
    The whole header band, edge to edge.

    Logos and the title share one strip across the page. Cropping that strip
    down to a single mark is what turned the header into a thumbnail. The
    picture keeps the band's own height-to-width ratio.
    """
    art = _header_art(page)
    if art is None:
        return None
    page_width = float(page.rect.width)
    top = max(0.0, art[1])
    bottom = min(float(page.rect.height) * 0.16, max(art[3], top + 8.0))
    if bottom - top < 8 or page_width < 40:
        return None
    box = (0.0, top, page_width, bottom)
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=fitz.Rect(*box), alpha=False)
        data = pixmap.tobytes("png")
    except Exception:
        return None
    if not data:
        return None
    data, height = _trim_header_tail(data, bottom - top)
    return data, height, page_width, "center", (box[0], box[1], box[2], box[1] + height)


def _pixel_kind(red: int, green: int, blue: int) -> str:
    """Blank paper, a hairline rule (gray or blue), or real ink."""
    if red >= 246 and green >= 246 and blue >= 246:
        return "blank"
    if blue > 70 and blue > red + 20 and blue > green + 10 and red < 190:
        return "rule"
    if abs(red - green) < 18 and abs(green - blue) < 18 and red < 200:
        return "rule"
    return "ink"


def _row_kind(pixels, y: int, width: int, step: int) -> str:
    counts = {"blank": 0, "rule": 0, "ink": 0}
    samples = 0
    for x in range(0, width, step):
        red, green, blue = pixels[x, y]
        counts[_pixel_kind(red, green, blue)] += 1
        samples += 1
    if not samples or counts["blank"] / samples >= 0.9:
        return "blank"
    if counts["rule"] / samples >= 0.55 and counts["ink"] / samples < 0.2:
        return "rule"
    return "ink"


def _trim_header_tail(data: bytes, band_height: float) -> tuple[bytes, float]:
    """
    Drop blank rows and a stray rule under or above the band.

    A blue or gray line only a few pixels thick is removed. A filled banner,
    which is the same color for many rows, stays. The picture then ends at
    the banner instead of carrying a line into the body.
    """
    try:
        from PIL import Image
    except Exception:
        return data, band_height
    try:
        image = Image.open(BytesIO(data)).convert("RGB")
    except Exception:
        return data, band_height
    width, height = image.size
    if width < 8 or height < 8:
        return data, band_height
    pixels = image.load()
    step = max(1, width // 80)

    def hairline(y: int, toward: int) -> bool:
        if _row_kind(pixels, y, width, step) != "rule":
            return False
        probe = y + toward
        if probe < 0 or probe >= height:
            return True
        return _row_kind(pixels, probe, width, step) != "rule"

    top = 0
    bottom = height - 1
    floor = int(height * 0.35)
    while bottom > top + floor:
        kind = _row_kind(pixels, bottom, width, step)
        if kind == "blank" or hairline(bottom, -4):
            bottom -= 1
            continue
        break
    while top < bottom - floor:
        kind = _row_kind(pixels, top, width, step)
        if kind == "blank" or hairline(top, 4):
            top += 1
            continue
        break
    if top == 0 and bottom >= height - 2:
        return data, band_height
    cropped = image.crop((0, top, width, bottom + 1))
    output = BytesIO()
    cropped.save(output, format="PNG")
    return output.getvalue(), band_height * ((bottom + 1 - top) / height)


def _printed_number(text: str) -> int | None:
    """The digits of a page label, including Arabic-Indic digits."""
    translated = text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    digits = re.sub(r"\D", "", translated)
    if not digits or len(digits) > 4:
        return None
    try:
        return int(digits)
    except ValueError:
        return None


def _page_number_atoms(atoms: list[_Atom]) -> list[_Atom]:
    """
    The glyphs of the printed page number in a footer line, if it has one.

    A run of digits is the page number when it stands apart from the words
    (a wide gap), when it is the whole line, or when a page-word stands next
    to it. Digits inside the running text ("30 November 2026") are text.
    """
    def is_digit(atom) -> bool:
        return re.fullmatch(r"[\d٠-٩]{1,4}", plain_arabic(atom.text).strip() or " ") is not None

    visible = sorted((atom for atom in atoms if atom.text.strip()), key=lambda atom: atom.x0)
    if not visible:
        return []
    runs: list[list[_Atom]] = []
    for atom in visible:
        if not is_digit(atom):
            continue
        size = max(4.0, atom.y1 - atom.y0)
        if runs and atom.x0 - runs[-1][-1].x1 <= size * 0.35 and abs(atom.y0 - runs[-1][-1].y0) <= size * 0.5:
            runs[-1].append(atom)
        else:
            runs.append([atom])
    picked: list[_Atom] = []
    for run in runs:
        ids = {id(atom) for atom in run}
        others = [atom for atom in visible if id(atom) not in ids and abs(atom.y0 - run[0].y0) <= max(4.0, run[0].y1 - run[0].y0)]
        if not others:
            picked.extend(run)
            continue
        size = max(4.0, run[0].y1 - run[0].y0)
        left = run[0].x0
        right = run[-1].x1
        gap = min(
            min((left - atom.x1 for atom in others if atom.x1 <= left + 0.5), default=1e9),
            min((atom.x0 - right for atom in others if atom.x0 >= right - 0.5), default=1e9),
        )
        if gap > size * 1.5:
            picked.extend(run)
            continue
        line = _line_from_atoms([
            atom for atom in atoms
            if id(atom) not in ids and abs(atom.y0 - run[0].y0) <= max(4.0, run[0].y1 - run[0].y0)
        ])
        rest = plain_arabic(line.text).strip() if line is not None else ""
        if re.search(r"((?:ال)?صفح[ةه]|page)\s*[|:\-–]?\s*$|^\s*((?:ال)?صفح[ةه]|page)(?:\s|$)", rest, flags=re.IGNORECASE):
            picked.extend(run)
    return picked


def _read_footer(page: fitz.Page) -> tuple[str, str, int | None, float]:
    """
    The running title, a page-word such as "page", the printed number, and the gap.

    Space glyphs are part of the line. Dropping them joins the title into one
    word. The page-word stays beside the number; it is not part of the title.
    The gap is the distance from the bottom of the page up to that line.
    """
    height = float(page.rect.height)
    atoms = [
        atom
        for atom in _atoms_on_page(page)
        if atom.y0 >= height * (1.0 - _FOOTER_ZONE)
    ]
    if not any(atom.text.strip() for atom in atoms):
        return "", "", None, 18.0
    number_atoms = _page_number_atoms(atoms)
    numbered = {id(atom) for atom in number_atoms}
    title_atoms = [atom for atom in atoms if id(atom) not in numbered]
    line = _line_from_atoms(title_atoms) if any(atom.text.strip() for atom in title_atoms) else None
    title = plain_arabic(line.text) if line is not None else ""
    label = ""
    found = re.search(r"(?:^|\s)((?:ال)?صفح[ةه]|page)\s*$", title, flags=re.IGNORECASE)
    if found:
        label = found.group(1)
        title = title[: found.start()].strip()
    number_value = _printed_number("".join(atom.text for atom in number_atoms)) if number_atoms else None
    if line is None and not number_atoms:
        return "", "", None, 18.0
    anchors = [atom for atom in atoms if atom.text.strip()]
    top = min(atom.y0 for atom in anchors)
    gap = min(40.0, max(12.0, height - top))
    return title, label, number_value, gap


def _footer_rule(page: fitz.Page) -> bool:
    """A vertical stroke in the footer band, between the page-word and the number."""
    height = float(page.rect.height)
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        return False
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None:
            continue
        width = float(rect.x1) - float(rect.x0)
        tall = float(rect.y1) - float(rect.y0)
        if width <= 2.0 and tall >= 6.0 and float(rect.y0) >= height * (1.0 - _FOOTER_ZONE) - 4:
            return True
    return False


def _edge_slots(
    page: fitz.Page,
    edge: str,
    art: tuple[float, float, float, float] | None = None,
) -> tuple[dict[str, str], int | None]:
    """
    Left, center and right text in the top or bottom 5% of the page.

    A bare page number is marked so the footer can use a live PAGE field.
    Words painted on the logo are left in the logo picture.
    """
    width = float(page.rect.width)
    height = float(page.rect.height)
    slots = {"left": "", "center": "", "right": ""}
    number_value: int | None = None
    for band in _cluster_baselines(_atoms_on_page(page)):
        for part in _split_gutter(band, width):
            line = _line_from_atoms(part)
            if line is None:
                continue
            text = plain_arabic(line.text).strip()
            if not text or len(text) > 90:
                continue
            bare = _is_page_label(text) or re.fullmatch(r"[\d٠-٩]{1,4}", text) is not None
            center_y = (line.bbox[1] + line.bbox[3]) / 2.0
            if edge == "top":
                if _chrome_edge(line.bbox, height) != "top":
                    continue
            elif center_y < height * (1.0 - _CHROME_BAND) and not (bare and line.bbox[1] >= height * 0.90):
                continue
            if art is not None:
                center_x = (line.bbox[0] + line.bbox[2]) / 2.0
                if art[0] <= center_x <= art[2] and art[1] <= center_y <= art[3]:
                    continue
            center_x = (line.bbox[0] + line.bbox[2]) / 2.0
            slot = "left" if center_x < width * 0.33 else "right" if center_x > width * 0.67 else "center"
            if edge == "bottom" and bare:
                slots[slot] = "\x00PAGE"
                number_value = _printed_number(text)
            elif not slots[slot]:
                slots[slot] = text
    return slots, number_value


@dataclass
class _PageChrome:
    """The running header, footer and body padding shared by pages of one size."""

    header_png: bytes | None = None
    header_band: float = 0.0
    header_width: float = 0.0
    header_top: float = 14.0
    header_align: str = "center"
    header_text: dict[str, str] = field(default_factory=lambda: {"left": "", "center": "", "right": ""})
    footer: dict[str, str] = field(default_factory=lambda: {"left": "", "center": "", "right": ""})
    footer_png: bytes | None = None
    footer_band: float = 0.0
    footer_width: float = 0.0
    footer_number_side: str = ""
    page_numbers: dict[int, int] = field(default_factory=dict)
    footer_gap: float = 18.0
    footer_label: str = ""
    footer_rule: bool = False
    insets: tuple[float, float, float, float] = (46.0, 56.0, 46.0, 42.0)
    majority: str = "rtl"


def _thin_rule_box(box: tuple[float, float, float, float], page_width: float) -> bool:
    """A full-width hairline. It is not a banner and not a logo."""
    height = box[3] - box[1]
    width = box[2] - box[0]
    return height <= 2.8 and width >= page_width * 0.45


def _filled_band(page: fitz.Page, edge: str) -> tuple[float, float] | None:
    """
    A colored or graphic band in the header or the footer.

    A thin rule, a white cell, and a one-off shape outside the margin are left out.
    """
    height = float(page.rect.height)
    width = float(page.rect.width)
    boxes: list[tuple[float, float, float, float]] = []

    def keep(box: tuple[float, float, float, float]) -> None:
        band_height = box[3] - box[1]
        band_width = box[2] - box[0]
        if band_height < 8 or band_width < width * 0.28 or _thin_rule_box(box, width):
            return
        if edge == "top":
            if box[1] > height * 0.08 or box[3] > height * 0.16:
                return
        elif box[1] < height * 0.84 or box[3] < height * 0.90:
            return
        boxes.append(box)

    for image in _image_entries(page):
        bbox = image.get("bbox")
        if bbox:
            keep((float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])))
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        drawings = []
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None or _near_white(drawing.get("fill")):
            continue
        keep((float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1)))
    if not boxes:
        return None
    top = max(0.0, min(box[1] for box in boxes) - 1.0)
    bottom = min(height, max(box[3] for box in boxes) + 1.0)
    if bottom - top < 8 or bottom - top > height * 0.16:
        return None
    return top, bottom


def _crop_band(
    page: fitz.Page,
    top: float,
    bottom: float,
    x0: float | None = None,
    x1: float | None = None,
) -> tuple[bytes, float, float] | None:
    """Rasterize one margin band and cut a stray blue or gray line off its edge."""
    width = float(page.rect.width)
    left = 0.0 if x0 is None else max(0.0, x0)
    right = width if x1 is None else min(width, x1)
    if right - left < 40 or bottom - top < 8:
        return None
    try:
        pixmap = page.get_pixmap(
            matrix=fitz.Matrix(2, 2),
            clip=fitz.Rect(left, top, right, bottom),
            alpha=False,
        )
        data = pixmap.tobytes("png")
    except Exception:
        return None
    if not data:
        return None
    data, band = _trim_header_tail(data, bottom - top)
    if band < 6:
        return None
    return data, band, right - left


def _cluster_bands(
    bands: list[tuple[int, float, float]],
    threshold: int,
) -> tuple[int, float, float] | None:
    """The y-span that shows up on enough pages. One page's private graphic loses."""
    if not bands:
        return None
    ordered = sorted(bands, key=lambda item: item[2])
    groups: list[list[tuple[int, float, float]]] = []
    for item in ordered:
        if groups and abs(item[2] - groups[-1][-1][2]) <= 10:
            groups[-1].append(item)
        else:
            groups.append([item])
    groups.sort(key=len, reverse=True)
    best = groups[0]
    if len(best) < threshold:
        return None
    middle = len(best) // 2
    ranked = sorted(best, key=lambda item: item[2])
    chosen = ranked[middle]
    y0 = sorted(item[1] for item in best)[middle]
    y1 = sorted(item[2] for item in best)[middle]
    return chosen[0], y0, y1


def _header_key(slots: dict[str, str]) -> str:
    parts = [slots.get(name, "") for name in ("right", "center", "left")]
    text = plain_arabic(" ".join(part for part in parts if part and part != "\x00PAGE"))
    text = re.sub(r"[\d٠-٩\s]+", "", text)
    return text[:80]


def _repeat_count(count: int) -> int:
    """How many pages must share a margin before it is the running header or footer."""
    if count <= 1:
        return 1
    return max(2, (count + 1) // 2)


def _read_chrome(document: fitz.Document, page_indexes: list[int] | None = None) -> _PageChrome:
    """
    Decide the running header and footer from the pages that will be converted.

    A band or a phrase has to recur. A table border, a chart that touches the
    top of one page, and a blue line under the banner do not become the header.
    A colored band is kept as a picture. Plain text is kept as text.
    """
    chrome = _PageChrome()
    if page_indexes is None:
        indexes = list(range(document.page_count))
    else:
        indexes = [index for index in page_indexes if 0 <= index < document.page_count]
    samples = []
    for index in indexes:
        page = document.load_page(index)
        atoms = _atoms_on_page(page)
        if _is_cover_page(page, atoms):
            continue
        samples.append((index, page, atoms))
    if not samples and indexes:
        index = indexes[0]
        page = document.load_page(index)
        samples.append((index, page, _atoms_on_page(page)))
    if not samples:
        return chrome

    threshold = _repeat_count(len(samples))
    inset_rows: list[tuple[float, float, float, float]] = []
    header_bands: list[tuple[int, float, float]] = []
    footer_bands: list[tuple[int, float, float]] = []
    header_keys: list[tuple[str, dict[str, str]]] = []
    footer_titles: list[tuple[str, str, str, int | None, float]] = []
    header_strips: list[tuple] = []
    arabic_letters = 0
    latin_letters = 0
    pages_by_index = {index: page for index, page, _atoms in samples}

    for index, page, atoms in samples:
        inset_rows.append(_body_insets(page))
        height = float(page.rect.height)
        for atom in atoms:
            if atom.y0 < height * 0.16 or atom.y1 > height * 0.88:
                continue
            arabic_letters += _arabic_count(atom.text)
            latin_letters += len(re.findall(r"[A-Za-z]", atom.text))
        header_box = _filled_band(page, "top")
        if header_box is not None:
            header_bands.append((index, header_box[0], header_box[1]))
        footer_box = _filled_band(page, "bottom")
        if footer_box is not None:
            footer_bands.append((index, footer_box[0], footer_box[1]))
        slots, _number = _edge_slots(page, "top", None)
        key = _header_key(slots)
        if len(key) >= 2:
            header_keys.append((key, slots))
        title, label, number_value, gap = _read_footer(page)
        title_key = re.sub(r"[\d٠-٩\s]+", "", plain_arabic(title))[:80]
        if title_key or number_value is not None:
            footer_titles.append((title_key, title, label, number_value, gap))
        strip = _header_strip(page)
        if strip is not None and strip[1] >= 8:
            header_strips.append(strip)
        if _footer_rule(page):
            chrome.footer_rule = True
        if number_value is not None:
            chrome.page_numbers[index] = number_value

    chrome.majority = "ltr" if latin_letters > arabic_letters else "rtl"
    header_choice = _cluster_bands(header_bands, threshold)
    if header_choice is not None:
        index, top, bottom = header_choice
        cropped = _crop_band(pages_by_index[index], top, bottom)
        if cropped is not None:
            chrome.header_png, chrome.header_band, chrome.header_width = cropped
            chrome.header_align = "center"
            chrome.header_top = top
    if chrome.header_png is None and len(header_strips) >= threshold:
        # Logos are several small marks. Together they are the header band.
        ranked = sorted(header_strips, key=lambda item: item[1])
        data, band, strip_width, align, box = ranked[len(ranked) // 2]
        chrome.header_png = data
        chrome.header_band = band
        chrome.header_width = strip_width
        chrome.header_align = align
        chrome.header_top = box[1]
    if chrome.header_png is None and header_keys:
        counts = Counter(key for key, _slots in header_keys)
        key, count = counts.most_common(1)[0]
        if count >= threshold:
            chrome.header_text = next(slots for item, slots in header_keys if item == key)

    footer_choice = _cluster_bands(footer_bands, threshold)
    if footer_choice is not None:
        index, top, bottom = footer_choice
        page = pages_by_index[index]
        number_atoms = [
            atom
            for atom in _atoms_on_page(page)
            if atom.y0 >= top - 2 and re.fullmatch(r"[\d٠-٩]{1,4}", plain_arabic(atom.text).strip() or " ")
        ]
        left, right = 0.0, float(page.rect.width)
        side = ""
        if number_atoms:
            number_left = min(atom.x0 for atom in number_atoms)
            number_right = max(atom.x1 for atom in number_atoms)
            center = (number_left + number_right) / 2.0
            if center < page.rect.width * 0.22:
                left = number_right + 4.0
                side = "left"
            elif center > page.rect.width * 0.78:
                right = number_left - 4.0
                side = "right"
            if right - left < page.rect.width * 0.45:
                left, right, side = 0.0, float(page.rect.width), ""
        cropped = _crop_band(page, top, bottom, left, right)
        if cropped is not None:
            chrome.footer_png, chrome.footer_band, chrome.footer_width = cropped
            chrome.footer_number_side = side
            chrome.footer_gap = min(40.0, max(12.0, float(page.rect.height) - top))
            if side:
                chrome.footer = {"left": "\x00PAGE" if side == "left" else "", "center": "", "right": "\x00PAGE" if side == "right" else ""}
    if chrome.footer_png is None and footer_titles:
        counts = Counter(key for key, _title, _label, _number, _gap in footer_titles if key)
        chosen_key = ""
        if counts:
            chosen_key, count = counts.most_common(1)[0]
            if count < threshold:
                chosen_key = ""
        for key, title, label, number_value, gap in footer_titles:
            if chosen_key and key != chosen_key:
                continue
            if not chosen_key and number_value is None:
                continue
            chrome.footer = {
                "left": "\x00PAGE" if number_value is not None else "",
                "center": "",
                "right": title if chosen_key else "",
            }
            chrome.footer_label = label
            chrome.footer_gap = gap
            break
    if inset_rows:
        chrome.insets = tuple(sorted(column)[len(column) // 2] for column in zip(*inset_rows))  # type: ignore[assignment]
    return chrome


def _clear_story(story, collapse: bool = False) -> None:
    """Unlink a header or footer and drop the placeholder paragraph's text."""
    story.is_linked_to_previous = False
    if not story.paragraphs:
        return
    paragraph = story.paragraphs[0]
    paragraph.text = ""
    if not collapse:
        return
    properties = paragraph._p.get_or_add_pPr()
    spacing = properties.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        properties.append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "20")
    spacing.set(qn("w:lineRule"), "exact")


def _slot_alignment(paragraph, visual: str, arabic: bool) -> None:
    """
    Put a header or footer slot on the edge it occupied.

    A bidirectional paragraph mirrors left and right, so the stored value is
    the opposite of the visual edge. A Latin page number is not bidirectional.
    """
    if arabic:
        stored = {"left": "right", "right": "left", "center": "center"}.get(visual, "left")
    else:
        stored = visual if visual in {"left", "center", "right"} else "left"
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[stored]
    justification = paragraph._p.get_or_add_pPr().find(qn("w:jc"))
    if justification is None:
        justification = OxmlElement("w:jc")
        paragraph._p.get_or_add_pPr().append(justification)
    justification.set(qn("w:val"), stored)


def _fill_slots(
    story,
    slots: dict[str, str],
    content_width_emu: int,
    picture: bytes | None = None,
    picture_slot: str = "",
    picture_width_pt: float = 0.0,
) -> None:
    """Three columns: left, center, right. A PAGE mark becomes a live page field."""
    if not any(slots.values()) and not picture:
        return
    table = story.add_table(rows=1, cols=3, width=Emu(content_width_emu))
    table.autofit = False
    _set_table_borders(table, show=False)
    names = ("left", "center", "right")
    for index, name in enumerate(names):
        paragraph = table.cell(0, index).paragraphs[0]
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        text = slots.get(name) or ""
        arabic = bool(text) and text != "\x00PAGE" and text_direction(text) == "rtl"
        if arabic:
            _set_paragraph_rtl(paragraph, text)
        elif text:
            _set_paragraph_ltr(paragraph)
        _slot_alignment(paragraph, name, arabic)
        if picture and name == picture_slot and picture_width_pt >= 16:
            paragraph.add_run().add_picture(BytesIO(picture), width=Emu(int(picture_width_pt * 12700)))
        if text == "\x00PAGE":
            _add_page_field(paragraph)
            continue
        if not text:
            continue
        run = paragraph.add_run(text)
        _set_run_font(run, 9.0, False, text)


def _clear_paragraph_borders(paragraph) -> None:
    """No rule under a header paragraph. The logo meets the body with no divider."""
    properties = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    for edge in ("w:top", "w:left", "w:bottom", "w:right", "w:between"):
        node = OxmlElement(edge)
        node.set(qn("w:val"), "nil")
        node.set(qn("w:sz"), "0")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "auto")
        borders.append(node)
    existing = properties.find(qn("w:pBdr"))
    if existing is not None:
        properties.remove(existing)
    properties.append(borders)


def _fill_header(section, chrome: _PageChrome, content_width_emu: int, _page_width_pt: float) -> None:
    """Place the header band across the full printable width, at its own ratio."""
    header = section.header
    _clear_story(header)
    has_logo = bool(chrome.header_png) and chrome.header_band > 0 and chrome.header_width > 0
    has_text = any(chrome.header_text.values())
    if not has_logo and not has_text:
        return
    paragraph = header.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _clear_paragraph_borders(paragraph)
    if has_logo:
        from core.html_layout import locked_size

        content_width_pt = max(72.0, content_width_emu / 12700.0)
        display_w, display_h = locked_size(chrome.header_width, chrome.header_band, content_width_pt)
        if display_w >= 16 and display_h >= 8:
            paragraph.add_run().add_picture(BytesIO(chrome.header_png), width=Emu(int(display_w * 12700)))
        # The strip already contains the title. A second row would draw a gap under it.
        return
    _reserve_height(paragraph, 0.0, 1.0)
    if has_text:
        _fill_slots(header, chrome.header_text, content_width_emu)


def _place_footer_band(footer, chrome: _PageChrome, content_width_emu: int) -> None:
    """The recurring colored footer, full width, with no rule above it."""
    from core.html_layout import locked_size

    content_width_pt = max(72.0, content_width_emu / 12700.0)
    side = chrome.footer_number_side if chrome.footer_number_side in {"left", "right"} else ""
    display_w, display_h = locked_size(chrome.footer_width, chrome.footer_band, content_width_pt)
    if display_w < 16 or display_h < 8 or not chrome.footer_png:
        return
    paragraph = footer.paragraphs[0]
    if not side:
        paragraph.add_run().add_picture(BytesIO(chrome.footer_png), width=Emu(int(display_w * 12700)))
        return
    _clear_story(footer, collapse=True)
    number_pt = 42.0
    picture_pt = max(24.0, display_w - number_pt)
    table = footer.add_table(rows=1, cols=2, width=Emu(content_width_emu))
    table.autofit = False
    _set_table_borders(table, show=False)
    picture_first = side == "right"
    cells = [picture_pt, number_pt] if picture_first else [number_pt, picture_pt]
    for index, width_pt in enumerate(cells):
        cell = table.cell(0, index)
        _set_cell_width(cell, int(width_pt * 20))
        cell_paragraph = cell.paragraphs[0]
        _clear_paragraph_borders(cell_paragraph)
        _set_paragraph_ltr(cell_paragraph)
        cell_paragraph.paragraph_format.space_before = Pt(0)
        cell_paragraph.paragraph_format.space_after = Pt(0)
        is_picture = (index == 0) == picture_first
        if is_picture:
            cell_paragraph.add_run().add_picture(
                BytesIO(chrome.footer_png), width=Emu(int(picture_pt * 12700))
            )
            continue
        if side == "right":
            cell_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            justification = cell_paragraph._p.get_or_add_pPr().find(qn("w:jc"))
            if justification is not None:
                justification.set(qn("w:val"), "right")
        _add_page_field(cell_paragraph)
        for run in cell_paragraph.runs:
            _set_run_font(run, 8.0, False, "1")


def _fill_footer(section, chrome: _PageChrome, content_width_emu: int) -> None:
    """
    One footer line: the running title on the right, the page number on the left.

    A colored footer is the cropped band instead of a second text line. A
    bidirectional paragraph starts at the right. A right tab at the far end
    of that line is measured from the right, so the page number lands on the
    left. An English footer keeps the title on the left and the number on the right.
    """
    footer = section.footer
    _clear_story(footer)
    paragraph = footer.paragraphs[0]
    paragraph.text = ""
    _clear_paragraph_borders(paragraph)
    if chrome.footer_png and chrome.footer_band > 0 and chrome.footer_width > 0:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        _place_footer_band(footer, chrome, content_width_emu)
        return
    title = (chrome.footer.get("right") or "").strip()
    if title == "\x00PAGE":
        title = ""
    has_page = (
        chrome.footer.get("left") == "\x00PAGE"
        or chrome.footer.get("right") == "\x00PAGE"
        or bool(chrome.page_numbers)
    )
    if not title and not has_page and not chrome.footer_label:
        _clear_story(footer, collapse=True)
        return
    rtl_footer = _DIRECTION.get() != "ltr" and (not title or text_direction(title) == "rtl")
    if title and rtl_footer:
        _set_paragraph_rtl(paragraph, title)
    elif title:
        _set_paragraph_ltr(paragraph)
    elif _DIRECTION.get() == "ltr":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    spacing = paragraph._p.get_or_add_pPr().find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        paragraph._p.get_or_add_pPr().append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "220")
    spacing.set(qn("w:lineRule"), "exact")
    if title:
        run = paragraph.add_run(title)
        _set_run_font(run, 8.0, False, title)
    if has_page or chrome.footer_label:
        if title:
            paragraph.add_run("\t")
            content_pt = max(120.0, content_width_emu / 12700.0 - 6.0)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Pt(content_pt), WD_TAB_ALIGNMENT.RIGHT)
        if chrome.footer_label:
            # A vertical rule in the footer becomes a bar. A plain page-word does not.
            separator = " | " if chrome.footer_rule else " "
            label_run = paragraph.add_run(chrome.footer_label + separator)
            _set_run_font(label_run, 8.0, False, chrome.footer_label)
        if has_page:
            _add_page_field(paragraph)
        for run in paragraph.runs:
            _set_run_font(run, 8.0, False)


def _apply_page_setup(
    section,
    page: fitz.Page | None,
    chrome: _PageChrome,
    cover: bool,
    page_index: int,
    show_header: bool = True,
    show_footer: bool = True,
    lock_header: bool = True,
    lock_footer: bool = True,
    page_size: tuple[float, float] | None = None,
    section_bidi: bool = True,
) -> tuple[int, float]:
    """
    Match this section to the PDF sheet.

    Width and height are the PDF page. Margins are the body padding, grown
    just enough to hold the header band and the footer line.
    """
    if page_size is not None:
        page_width = float(page_size[0])
        page_height = float(page_size[1])
    else:
        page_width = float(page.rect.width)
        page_height = float(page.rect.height)
    section.page_width = Emu(int(page_width * 12700))
    section.page_height = Emu(int(page_height * 12700))
    existing_bidi = section._sectPr.find(qn("w:bidi"))
    if section_bidi and existing_bidi is None:
        section._sectPr.append(OxmlElement("w:bidi"))
    elif not section_bidi and existing_bidi is not None:
        section._sectPr.remove(existing_bidi)
    if cover:
        section.left_margin = Cm(0.5)
        section.right_margin = Cm(0.5)
        section.top_margin = Cm(0.4)
        section.bottom_margin = Cm(0.4)
        _clear_story(section.header)
        _clear_story(section.footer)
    else:
        left, top, right, bottom = chrome.insets
        has_header = (
            show_header
            and lock_header
            and (bool(chrome.header_png) or any(chrome.header_text.values()))
        )
        has_footer = show_footer and lock_footer and (any(chrome.footer.values()) or bool(chrome.footer_png))
        # The logo keeps the gap it had under the top edge of the PDF.
        header_pad = min(max(chrome.header_top, 8.0), 28.0) if chrome.header_png else 18.0
        footer_pad = min(max(chrome.footer_gap, 12.0), 40.0) if has_footer else 18.0
        if has_header and chrome.header_width > 0:
            layout_w = max(72.0, page_width - left - right)
            logo_h = layout_w * (chrome.header_band / chrome.header_width)
            top = max(top, header_pad + min(logo_h, 42.0) + 2.0)
        if has_footer and chrome.footer_png and chrome.footer_width > 0:
            layout_w = max(72.0, page_width - left - right)
            footer_h = layout_w * (chrome.footer_band / chrome.footer_width)
            bottom = max(bottom, footer_pad + min(footer_h, 36.0))
        elif has_footer:
            bottom = max(bottom, footer_pad + 12.0)
        # Keep a readable frame if the measured padding would swallow the page.
        if page_width - left - right < 280:
            overflow = 280 - (page_width - left - right)
            left = max(28.0, left - overflow / 2.0)
            right = max(28.0, right - overflow / 2.0)
        if page_height - top - bottom < 360:
            overflow = 360 - (page_height - top - bottom)
            top = max(header_pad, top - overflow / 2.0)
            bottom = max(footer_pad, bottom - overflow / 2.0)
        if has_header:
            top = max(top, header_pad + 8.0)
        if has_footer:
            bottom = max(bottom, footer_pad + 12.0)
        section.left_margin = Emu(int(round(left * 12700)))
        section.right_margin = Emu(int(round(right * 12700)))
        section.top_margin = Emu(int(round(top * 12700)))
        section.bottom_margin = Emu(int(round(bottom * 12700)))
        section.header_distance = Emu(int(round(header_pad * 12700)))
        section.footer_distance = Emu(int(round(footer_pad * 12700)))
        width = int(section.page_width - section.left_margin - section.right_margin)
        if has_header:
            _fill_header(section, chrome, width, page_width)
        else:
            _clear_story(section.header)
        if has_footer:
            _fill_footer(section, chrome, width)
        else:
            _clear_story(section.footer)
        start = chrome.page_numbers.get(page_index)
        if start:
            numbering = section._sectPr.find(qn("w:pgNumType"))
            if numbering is None:
                numbering = OxmlElement("w:pgNumType")
                section._sectPr.append(numbering)
            numbering.set(qn("w:start"), str(start))
    width = int(section.page_width - section.left_margin - section.right_margin)
    height_pt = float(section.page_height - section.top_margin - section.bottom_margin) / 12700
    return width, height_pt


def _start_section(
    document: Document,
    page: fitz.Page | None,
    first: bool,
    chrome: _PageChrome,
    cover: bool,
    page_index: int,
    show_header: bool = True,
    show_footer: bool = True,
    lock_header: bool = True,
    lock_footer: bool = True,
    page_size: tuple[float, float] | None = None,
    section_bidi: bool = True,
) -> tuple[int, float]:
    """Open a section only for a new sheet size or a cover, then apply the PDF frame."""
    if first:
        section = document.sections[0]
        body = document.element.body
        opening = body.find(qn("w:p"))
        if opening is not None and not opening.findall(".//" + qn("w:t")):
            properties = opening.find(qn("w:pPr"))
            if properties is None or properties.find(qn("w:sectPr")) is None:
                body.remove(opening)
    else:
        section = document.add_section(WD_SECTION.NEW_PAGE)
    return _apply_page_setup(
        section,
        page,
        chrome,
        cover,
        page_index,
        show_header,
        show_footer,
        lock_header,
        lock_footer,
        page_size,
        section_bidi,
    )


def _block_span(block, width_pt: float) -> float:
    """How tall this block will be after it has been packed onto the page."""
    if isinstance(block, TextBlock):
        if block.role:
            size = min(max(block.size, 8.0), 28.0)
        elif block.kind == "heading":
            size = min(max(block.size, 8.0), 16.0)
        else:
            size = min(max(block.size, 8.0), 14.0)
        return _line_count(block.text, width_pt, size) * size * 1.2 + max(0.0, block.gap_before)
    if isinstance(block, TableBlock):
        return sum(_row_heights(block, width_pt, None)) + max(0.0, block.gap_before)
    if isinstance(block, ImageBlock):
        if block.fit == "page":
            return 0.0
        return min(block.height or 120.0, width_pt * 0.85) + max(0.0, block.gap_before)
    return 14.0


def _pack_page(blocks: list, frame_pt: float, width_pt: float) -> None:
    """
    Shrink one PDF page until it fits on one Word page.

    Gaps go first. If the words and the grid are still taller than the sheet,
    the type and the row heights are scaled down together. Nothing is pushed
    onto a following page.
    """
    if not blocks:
        return
    if any(getattr(block, "fit", "") == "page" for block in blocks):
        for block in blocks:
            block.gap_before = 0.0
        return
    budget = max(72.0, frame_pt * 0.92)

    def content_span() -> float:
        total = 0.0
        for block in blocks:
            gap = block.gap_before
            block.gap_before = 0.0
            total += _block_span(block, width_pt)
            block.gap_before = gap
        return total

    content = content_span()
    if content > budget and content > 0:
        scale = max(0.55, min(1.0, budget / content))
        for block in blocks:
            block.gap_before = 0.0
            if isinstance(block, TextBlock):
                block.size = max(8.0, block.size * scale)
            elif isinstance(block, TableBlock):
                block.fit_scale = min(block.fit_scale, scale)
            elif isinstance(block, ImageBlock) and block.fit != "page":
                block.height = max(24.0, (block.height or 80.0) * scale)
                block.width = max(24.0, block.width * scale)
        return
    room = budget - content
    raw = sum(max(0.0, block.gap_before) for block in blocks)
    if raw <= 0:
        return
    if room <= 0:
        for block in blocks:
            block.gap_before = 0.0
        return
    if raw > room:
        scale = room / raw
        for block in blocks:
            block.gap_before = max(0.0, block.gap_before) * scale


def _is_cover_sheet(blocks: list) -> bool:
    return bool(blocks) and len(blocks) == 1 and isinstance(blocks[0], ImageBlock) and blocks[0].fit == "page"


def _sheet_key(page: fitz.Page) -> tuple[float, float]:
    return (round(float(page.rect.width), 1), round(float(page.rect.height), 1))


def _add_page_break(document: Document) -> None:
    """One source page boundary. The break paragraph itself has no height."""
    paragraph = document.add_paragraph()
    properties = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "20")
    spacing.set(qn("w:lineRule"), "exact")
    properties.append(spacing)
    run = paragraph.add_run()
    break_mark = OxmlElement("w:br")
    break_mark.set(qn("w:type"), "page")
    run._r.append(break_mark)


def _write_page(
    document: Document,
    blocks: list,
    page_width_pt: float,
    page_height_pt: float,
    counts: dict,
) -> None:
    """
    Write one PDF page from its HTML box model.

    Each block becomes an absolutely positioned frame. The frames share this
    page, so the Word page count stays equal to the PDF page count.
    """
    from core.html_layout import write_page_html

    write_page_html(document, blocks, page_width_pt, page_height_pt, counts)


def parse_page_spec(spec: str, page_count: int) -> list[int]:
    """Read `1-5` or `3,4,7`. An empty value means every page, in order."""
    text = (spec or "").strip()
    if not text or text in {"*", "all", "الكل"}:
        return list(range(1, page_count + 1))
    chosen: list[int] = []
    seen: set[int] = set()
    for part in re.split(r"[،,;]+", text):
        piece = part.strip()
        if not piece:
            continue
        span = re.fullmatch(r"(\d+)\s*[-–—]\s*(\d+)", piece)
        if span:
            start, end = int(span.group(1)), int(span.group(2))
            if start > end:
                start, end = end, start
            numbers = range(start, end + 1)
        elif re.fullmatch(r"\d+(?:\s+\d+)*", piece):
            numbers = [int(item) for item in piece.split()]
        else:
            raise ConversionError(f"The page list '{spec}' could not be read.")
        for number in numbers:
            if number < 1 or number > page_count:
                raise ConversionError(f"Page {number} is not in this PDF.")
            if number not in seen:
                seen.add(number)
                chosen.append(number)
    if not chosen:
        raise ConversionError("The page list has no numbers.")
    return chosen


def _open_readers(source: Path, password: str):
    """Open the PDF for PyMuPDF and, when it can, for pdfplumber."""
    try:
        document_pdf = open_document(source, password)
    except FileNotFoundError as exc:
        raise ConversionError(f"The PDF was not found: {source}") from exc
    except PermissionError as exc:
        raise ConversionError("This PDF needs the correct password.") from exc
    except RuntimeError as exc:
        raise ConversionError(f"The PDF could not be opened. {exc}") from exc
    if document_pdf.page_count == 0:
        close_document(document_pdf)
        raise ConversionError("The PDF has no pages.")
    plumber = None
    try:
        import pdfplumber

        kwargs = {"password": password} if password else {}
        plumber = pdfplumber.open(str(source), **kwargs)
    except Exception:
        plumber = None
    return document_pdf, plumber


def _close_readers(document_pdf, plumber) -> None:
    if plumber is not None:
        try:
            plumber.close()
        except Exception:
            pass
    close_document(document_pdf)


def _plumber_at(plumber, index: int):
    if plumber is None:
        return None
    try:
        return plumber.pages[index]
    except Exception:
        return None


def _page_picture(page: fitz.Page) -> list:
    """The page as one picture, with no body reflow and no running header."""
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.7, 1.7), alpha=False)
        data = pixmap.tobytes("png")
    except Exception:
        data = b""
    if not data:
        return []
    width = float(page.rect.width)
    height = float(page.rect.height)
    return [
        ImageBlock(
            (0.0, 0.0, width, height),
            data,
            width,
            height,
            "center",
            0.0,
            "page",
        )
    ]


def _snapshot_box(page: fitz.Page, bbox, fit: str = "") -> ImageBlock | None:
    """A high-resolution clip of one region, used for a table kept as a picture."""
    rect = fitz.Rect(*bbox) & page.rect
    if rect.width < 4 or rect.height < 4:
        return None
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), clip=rect, alpha=False)
        data = pixmap.tobytes("png")
    except Exception:
        return None
    if not data:
        return None
    return ImageBlock(
        (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1)),
        data,
        float(rect.width),
        float(rect.height),
        "center",
        0.0,
        fit,
        True,
    )


def _pictures_for_tables(page: fitz.Page, blocks: list) -> list:
    output = []
    for block in blocks:
        if not isinstance(block, TableBlock):
            output.append(block)
            continue
        output.append(_snapshot_box(page, block.bbox, "table") or block)
    return output


def _refresh_images(page: fitz.Page, blocks: list, background_xrefs: set[int]) -> list:
    """Replace figures on this page and leave the paragraphs and tables in place."""
    if any(isinstance(block, ImageBlock) and block.fit == "page" for block in blocks):
        return _page_picture(page)
    atoms = _atoms_on_page(page)
    drawn = _vector_figures(page)
    figures = _release_text_panels(
        drawn
        + _figures(page, atoms, background_xrefs)
        + _ornament_images(page, [block.bbox for block in drawn]),
        atoms,
    )
    kept = [
        block
        for block in blocks
        if not isinstance(block, ImageBlock) or block.fit == "table"
    ]
    flow = [*kept, *figures]
    flow.sort(key=lambda block: (block.bbox[1], -block.bbox[2]))
    return flow


def _refresh_tables(page: fitz.Page, plumber_page, blocks: list, as_images: bool) -> list:
    """Replace only the grids on this page."""
    atoms = _atoms_on_page(page)
    tables = _tables_from_pymupdf(page, atoms)
    if not tables:
        tables = _tables_from_pdfplumber(plumber_page, atoms)
    tables = _dedupe_tables(page, tables)
    tables = [
        table
        for table in tables
        if not _is_false_grid(page, table, atoms) and not _margin_table(page, table) and _table_has_words(table)
    ]
    if not tables:
        raise ConversionError("No table was found on that page.")
    rendered: list = []
    for table in tables:
        if as_images:
            rendered.append(_snapshot_box(page, table.bbox, "table") or table)
        else:
            rendered.append(table)
    kept = [
        block
        for block in blocks
        if not isinstance(block, TableBlock) and not (isinstance(block, ImageBlock) and block.fit == "table")
    ]
    flow = [*kept, *rendered]
    flow.sort(key=lambda block: (block.bbox[1], -block.bbox[2]))
    return flow


def _capture_page(
    page: fitz.Page,
    plumber_page,
    background_xrefs: set[int],
    furniture: set[tuple[str, str]],
    schema: _ColumnSchema,
    page_number: int,
    notes: list[str],
    options: WordOptions,
    repair: str = "",
) -> WordPage:
    """Read one page with the layout choices that apply to this conversion."""
    from core.html_layout import MERGE_PARAGRAPHS

    merge = True if repair == "rtl" else options.merge_paragraphs
    merge_token = MERGE_PARAGRAPHS.set(merge)
    repair_token = _PAGE_REPAIR.set("rtl" if repair == "rtl" else "")
    kind: dict[str, str] = {}
    try:
        if page_number == 1 and options.preserve_first_page and repair != "rtl":
            blocks = _page_picture(page)
            kind["kind"] = "digital"
        else:
            blocks = _read_page(
                page,
                plumber_page,
                background_xrefs,
                furniture,
                schema,
                kind,
                page_number,
                notes,
            )
            if repair == "rtl":
                blocks = sorted(blocks, key=lambda block: (block.bbox[1], -block.bbox[2]))
            if options.tables_as_images:
                blocks = _pictures_for_tables(page, blocks)
    finally:
        MERGE_PARAGRAPHS.reset(merge_token)
        _PAGE_REPAIR.reset(repair_token)
    cover = _is_cover_sheet(blocks) or (
        page_number == 1 and options.preserve_first_page and repair != "rtl"
    )
    return WordPage(
        number=page_number,
        blocks=blocks,
        width=float(page.rect.width),
        height=float(page.rect.height),
        kind=kind.get("kind", "digital"),
        show_header=not cover,
        show_footer=not cover,
        cover=cover,
    )


def _capture_faithful(
    page: fitz.Page,
    plumber_page,
    background_xrefs: set[int],
    furniture: set[tuple[str, str]],
    schema: _ColumnSchema,
    page_number: int,
    notes: list[str],
    options: WordOptions,
    repair: str = "",
) -> WordPage:
    """
    Read one page for faithful mode.

    A page with a text layer becomes a background picture and text boxes.
    A page whose words live only inside embedded pictures (charts, raster
    tables, figures) also stays on this path: the background keeps those
    pictures. Falling through to the scan OCR path invents stray labels and
    drops the figure — that path is reserved for true full-page scans with
    no usable embedded images.
    """
    from core import faithful

    width = float(page.rect.width)
    height = float(page.rect.height)
    if page_number == 1 and options.preserve_first_page and repair != "rtl":
        return WordPage(page_number, [], width, height, faithful=faithful.picture_page(page))
    atoms = _atoms_on_page(page)
    digital_layer = sum(len(atom.text.strip()) for atom in atoms) >= 40
    images = _image_entries(page)
    try:
        drawings = len(page.get_drawings() or [])
    except Exception:
        drawings = 0
    if not digital_layer:
        # Vector-outlined pages (glyphs as paths, no text operators) must stay
        # as a picture — with or without a small logo image. OCR invents
        # crushed tables and displaced Arabic on these pages.
        if drawings >= 20:
            try:
                from core.page_strategy import _image_coverage

                coverage = _image_coverage(page) if images else 0.0
            except Exception:
                coverage = 0.0
            if coverage < 0.92:
                notes.append(
                    f"Page {page_number}: no text layer but {drawings} vector drawings "
                    "(outlined glyphs). Kept as a page picture so layout matches the PDF."
                )
                return WordPage(
                    page_number,
                    [],
                    width,
                    height,
                    kind="picture",
                    faithful=faithful.picture_page(
                        page,
                        note=(
                            f"Page {page_number}: text is drawn as vectors with no Unicode "
                            "text layer. The page is kept as a picture."
                        ),
                    ),
                )
        if not images:
            # No text, no pictures, few drawings: a true scan / blank page.
            record = _capture_page(
                page, plumber_page, background_xrefs, furniture, schema, page_number, notes, options, repair
            )
            if record.kind == "scan":
                return record
    payload = faithful.read_page(page, repair)
    notes.extend(payload.notes)
    # Keep editable text boxes even when some glyphs were OCR-repaired.
    # Whole-page pictures are only for true outlined/no-text pages (above).
    return WordPage(page_number, [], width, height, faithful=payload)


def _capture_flow(
    page: fitz.Page,
    plumber_page,
    background_xrefs: set[int],
    furniture: set[tuple[str, str]],
    schema: _ColumnSchema,
    page_number: int,
    notes: list[str],
    options: WordOptions,
    repair: str = "",
    fallback=None,
) -> WordPage:
    """
    Read one page for the flowing editable layout.

    `fallback` is the reader earlier versions used for this page. A page
    that cannot flow without loss goes to it, so it comes out exactly as it
    did before: positioned editable frames (with OCR text when the font map
    is broken) or the faithful layout.

    The faithful reader does the reading: it splits side-by-side columns at
    their rules, drops fake-bold copies and repairs broken Arabic glyphs. Its
    lines then become ordinary paragraphs instead of positioned boxes. A page
    the faithful reader cannot read as text (a scan, outlined glyphs) goes to
    the ordinary reader, which can OCR it.
    """
    from core import faithful
    from core.flow_layout import blocks_from_faithful

    width = float(page.rect.width)
    height = float(page.rect.height)
    raw = _atoms_on_page(page)
    letters = sum(len(atom.text.strip()) for atom in raw)
    digital_layer = letters >= 40 or (letters >= 3 and not _page_has_images(page))
    if digital_layer and _encoding_broken(page, raw):
        # A font map that stores Arabic letters as Latin ones: the words are
        # read from the page picture, and that reading belongs in the
        # positioned layout. Deciding here also keeps the slow picture
        # reading to one pass per page.
        # Read with the ordinary reader, whatever the classifier chose: it
        # takes the words from the page picture (OCR) and writes them as
        # editable text. The faithful reader would leave them in the picture.
        notes.append(f"Page {page_number}: read from the page picture (the font stores Arabic letters as Latin ones).")
        return _positioned(_capture_page, page, plumber_page, background_xrefs, furniture, schema, page_number, notes, options, repair)
    # A page that is only a picture of a table has almost no text layer, yet
    # its table can be read from the picture. Such a page is read here too.
    table_picture = False
    if not digital_layer and _page_has_images(page):
        try:
            from core import image_tables, table_rebuild

            table_picture = image_tables.available() and any(
                region.kind == "raster_table" for region in table_rebuild.classify_regions(page)
            )
        except Exception:
            table_picture = False
    if (digital_layer or table_picture) and not (page_number == 1 and options.preserve_first_page):
        try:
            payload = faithful.read_page(page, repair, background=False)
        except Exception as exc:
            payload = None
            notes.append(f"Page {page_number}: the flowing reader failed ({exc}); the ordinary reader was used.")
        usable_tables = [t for t in (payload.tables if payload is not None else []) if getattr(t, "usable", True) and getattr(t, "rows", None)]
        if payload is not None and (payload.lines or payload.tables) and (digital_layer or usable_tables):
            from core.flow_layout import page_flows_well

            flows, why = page_flows_well(page, payload)
            if flows:
                notes.extend(payload.notes)
                blocks = blocks_from_faithful(page, payload)
                return WordPage(page_number, blocks, width, height, kind="digital")
            # Positioned layout for this page, so nothing that only the page
            # picture carries is lost and the page stays one page.
            notes.append(f"Page {page_number}: kept in the positioned layout ({why}).")
            return _positioned(fallback, page, plumber_page, background_xrefs, furniture, schema, page_number, notes, options, repair)
    return _positioned(fallback or _capture_page, page, plumber_page, background_xrefs, furniture, schema, page_number, notes, options, repair)


def _positioned(reader, page, plumber_page, background_xrefs, furniture, schema, page_number, notes, options, repair) -> WordPage:
    """Read a page the way earlier versions did; mark it to be written in frames."""
    reader = reader or _capture_faithful
    record = reader(page, plumber_page, background_xrefs, furniture, schema, page_number, notes, options, repair)
    record.frames = True
    return record


def _remember_page(record: WordPage) -> dict:
    return {
        "blocks": record.blocks,
        "show_header": record.show_header,
        "show_footer": record.show_footer,
        "cover": record.cover,
        "kind": record.kind,
        "faithful": getattr(record, "faithful", None),
    }


def _find_word_page(build: WordBuild, page_number: int) -> WordPage:
    for record in build.pages:
        if record.number == int(page_number):
            return record
    raise ConversionError(f"Page {page_number} is not in this document.")


def read_word_build(
    pdf_path: str | Path,
    password: str = "",
    options: WordOptions | None = None,
) -> WordBuild:
    """
    Read the requested pages and keep each page's blocks.

    A later re-scan replaces one page in this build. The other pages stay
    as they were read here.
    """
    source = Path(pdf_path).expanduser()
    if not source.is_file():
        raise ConversionError(f"The PDF was not found: {source}")
    options = options or WordOptions()
    faithful_mode = _is_faithful(options)
    document_pdf, plumber = _open_readers(source, password)
    counts = {
        "pages": 0,
        "paragraphs": 0,
        "headings": 0,
        "tables": 0,
        "digital": 0,
        "scans": 0,
        "retries": [],
        "skipped": [],
        "page_plans": [],
    }
    pages: list[WordPage] = []
    try:
        page_count = document_pdf.page_count
        selected = set(parse_page_spec(options.page_spec, page_count))
        from core.page_strategy import classify_document

        plans = {
            plan.number: plan
            for plan in classify_document(
                document_pdf,
                default_mode=options.mode,
                overrides=getattr(options, "page_strategies", None) or {},
            )
        }
        counts["page_plans"] = [plans[number].as_dict() for number in sorted(plans) if number in selected]
        counts["decision_log"] = [
            {"page": plan["number"], "mode": plan["strategy"], "reason": plan["reason"]}
            for plan in counts["page_plans"]
        ]
        background_xrefs = _running_backgrounds(document_pdf)
        chrome = _read_chrome(document_pdf, [number - 1 for number in sorted(selected)])
        schema = _ColumnSchema()
        selected_numbers = sorted(number for number in selected)
        total = len(selected_numbers)
        workers = max(1, int(getattr(options, "workers", 1) or 1))

        def _cancelled() -> bool:
            hook = getattr(options, "cancel", None)
            return bool(callable(hook) and hook())

        def _progress(done: int, message: str) -> None:
            hook = getattr(options, "progress", None)
            if callable(hook):
                try:
                    hook(done, total, message)
                except Exception:
                    pass

        if workers > 1 and total > 1 and faithful_mode:
            # Each worker opens its own PDF copy; fitz pages are not shared.
            pages = _read_pages_parallel(
                source,
                password,
                selected_numbers,
                plans,
                options,
                background_xrefs,
                chrome,
                counts,
                workers,
                _cancelled,
                _progress,
            )
        else:
            done = 0
            for batch_start, batch_pages in iter_page_batches(document_pdf, _PAGE_BATCH):
                if _cancelled():
                    counts["skipped"].append("Conversion cancelled.")
                    break
                furniture = _furniture_for(batch_pages)
                for offset, page in enumerate(batch_pages):
                    index = batch_start + offset
                    number = index + 1
                    if number not in selected:
                        continue
                    if _cancelled():
                        counts["skipped"].append("Conversion cancelled.")
                        break
                    try:
                        plan = plans.get(number)
                        strategy = plan.strategy if plan is not None else ("faithful" if faithful_mode else "editable")
                        use_faithful = strategy in {"faithful", "broken_font"} or (
                            faithful_mode and strategy not in {"editable", "table", "scanned"}
                        )
                        if strategy == "scanned":
                            use_faithful = False
                        if strategy == "editable" or strategy == "table":
                            use_faithful = False
                        # The flowing layout also takes pages the classifier keeps
                        # faithful only for their art (a logo, a ruled frame). A slide
                        # or a page whose text is drawn as outlines stays faithful.
                        reason = plan.reason if plan is not None else ""
                        # How earlier versions read this page. A page the flowing
                        # layout cannot take is read this way, unchanged.
                        fallback = _capture_faithful if use_faithful else _capture_page
                        if (
                            use_faithful
                            and not faithful_mode
                            and getattr(options, "flow", True)
                            and strategy == "faithful"
                            and not reason.startswith(("landscape page", "dense vector drawings with no text layer"))
                        ):
                            use_faithful = False
                        if use_faithful:
                            capture = _capture_faithful
                        elif getattr(options, "flow", True) and not faithful_mode:
                            # A "scanned" page may be a picture of a table; the flow
                            # reader tries that and otherwise hands it to `fallback`.
                            capture = functools.partial(_capture_flow, fallback=fallback)
                        else:
                            capture = _capture_page
                        record = capture(
                            page,
                            _plumber_at(plumber, index),
                            background_xrefs,
                            furniture,
                            schema,
                            number,
                            counts["retries"],
                            options,
                        )
                    except ConversionError as exc:
                        counts["skipped"].append(str(exc))
                        record = WordPage(number, [], float(page.rect.width), float(page.rect.height))
                    except Exception as exc:
                        counts["skipped"].append(f"Page {number} was skipped. {exc}")
                        record = WordPage(number, [], float(page.rect.width), float(page.rect.height))
                    if record.kind == "scan":
                        counts["scans"] += 1
                    else:
                        counts["digital"] += 1
                    pages.append(record)
                    done += 1
                    _progress(done, f"Read page {number}")
        _collect_font_report(pages, counts)
        _mark_continued_table_headers(pages)
    finally:
        _close_readers(document_pdf, plumber)
    return WordBuild(pages, page_count, counts, chrome, options)


def _collect_font_report(pages: list[WordPage], counts: dict) -> None:
    missing: list[str] = []
    seen: set[str] = set()
    for record in pages:
        payload = getattr(record, "faithful", None)
        for name in getattr(payload, "missing_fonts", None) or []:
            if name not in seen:
                seen.add(name)
                missing.append(name)
    counts["missing_fonts"] = missing


def _mark_continued_table_headers(pages: list[WordPage]) -> None:
    from core import table_rebuild

    by_page: dict[int, list] = {}
    for record in pages:
        payload = getattr(record, "faithful", None)
        tables = getattr(payload, "tables", None) or []
        if tables:
            by_page[record.number] = tables
    table_rebuild.mark_continued_headers(by_page)


def _read_pages_parallel(
    source: Path,
    password: str,
    selected_numbers: list[int],
    plans: dict,
    options: WordOptions,
    background_xrefs: set[int],
    chrome: "_PageChrome",
    counts: dict,
    workers: int,
    cancelled,
    progress,
) -> list[WordPage]:
    """Read pages concurrently; each worker opens its own document handle."""
    from concurrent.futures import ThreadPoolExecutor, as_completed

    def one(number: int) -> WordPage:
        if cancelled():
            return WordPage(number, [], 0.0, 0.0)
        document, plumber = _open_readers(source, password)
        try:
            page = document.load_page(number - 1)
            plan = plans.get(number)
            strategy = plan.strategy if plan is not None else "faithful"
            use_faithful = strategy in {"faithful", "broken_font"}
            if strategy in {"scanned", "editable", "table"}:
                use_faithful = False
            furniture: set[tuple[str, str]] = set()
            notes: list[str] = []
            capture = _capture_faithful if use_faithful else _capture_page
            record = capture(
                page,
                _plumber_at(plumber, number - 1),
                background_xrefs,
                furniture,
                _ColumnSchema(),
                number,
                notes,
                options,
            )
            counts["retries"].extend(notes)
            return record
        finally:
            _close_readers(document, plumber)

    pages_by_number: dict[int, WordPage] = {}
    with ThreadPoolExecutor(max_workers=min(workers, len(selected_numbers))) as pool:
        futures = {pool.submit(one, number): number for number in selected_numbers}
        done = 0
        for future in as_completed(futures):
            number = futures[future]
            try:
                record = future.result()
            except Exception as exc:
                counts["skipped"].append(f"Page {number} was skipped. {exc}")
                record = WordPage(number, [], 0.0, 0.0)
            pages_by_number[number] = record
            if record.kind == "scan":
                counts["scans"] += 1
            else:
                counts["digital"] += 1
            done += 1
            progress(done, f"Read page {number}")
    return [pages_by_number[number] for number in selected_numbers if number in pages_by_number]


def _scaled_blocks(blocks: list, scale: float, offset_x: float, offset_y: float) -> list:
    output = []
    for block in blocks:
        clone = copy.copy(block)
        box = block.bbox
        clone.bbox = (
            box[0] * scale + offset_x,
            box[1] * scale + offset_y,
            box[2] * scale + offset_x,
            box[3] * scale + offset_y,
        )
        if isinstance(block, TextBlock):
            clone.size = max(6.0, block.size * scale)
        elif isinstance(block, ImageBlock):
            clone.width = block.width * scale
            clone.height = block.height * scale
        elif isinstance(block, TableBlock):
            clone.widths = [value * scale for value in block.widths]
        output.append(clone)
    return output


def _frame_for(record: WordPage, options: WordOptions):
    """Page size for Word. A4 is exact when the PDF page is already A4."""
    if not options.fit_a4:
        return record.blocks, record.width, record.height, (record.width, record.height)
    if abs(record.width - _A4_WIDTH) <= 4 and abs(record.height - _A4_HEIGHT) <= 4:
        return record.blocks, _A4_WIDTH, _A4_HEIGHT, (_A4_WIDTH, _A4_HEIGHT)
    scale = min(_A4_WIDTH / max(record.width, 1.0), _A4_HEIGHT / max(record.height, 1.0))
    offset_x = (_A4_WIDTH - record.width * scale) / 2.0
    offset_y = (_A4_HEIGHT - record.height * scale) / 2.0
    blocks = _scaled_blocks(record.blocks, scale, offset_x, offset_y)
    return blocks, _A4_WIDTH, _A4_HEIGHT, (_A4_WIDTH, _A4_HEIGHT)


def write_word_build(build: WordBuild, docx_path: str | Path) -> dict:
    """Write the current pages. Unchanged pages are the blocks stored on the build."""
    from core.html_layout import MERGE_PARAGRAPHS

    output = Path(docx_path).expanduser()
    if output.suffix.lower() != ".docx":
        output = output.with_suffix(".docx")
    output.parent.mkdir(parents=True, exist_ok=True)
    word = Document()
    options = build.options
    font_token = _FONT.set(str(getattr(options, "font", "") or ""))
    try:
        return _write_word_build(build, output, word, options)
    finally:
        _FONT.reset(font_token)


def _set_default_font(word, family: str) -> None:
    """The Normal style and the document defaults use `family` for every script."""
    if not family:
        return
    for properties in (
        word.styles["Normal"].element.get_or_add_rPr(),
        word.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
        if word.styles.element.find(qn("w:docDefaults")) is not None else None,
    ):
        if properties is None:
            continue
        fonts = properties.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            properties.insert(0, fonts)
        for slot in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            if fonts.get(qn(slot)) is not None:
                del fonts.attrib[qn(slot)]
        for slot in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            fonts.set(qn(slot), family)


def _write_word_build(build: WordBuild, output: Path, word, options) -> dict:
    from core.html_layout import MERGE_PARAGRAPHS

    try:
        _set_default_font(word, (_FONT.get() or "").strip())
    except Exception:
        pass

    direction = options.direction if options.direction in {"auto", "rtl", "ltr"} else "rtl"
    counts = {
        "pages": 0,
        "paragraphs": 0,
        "headings": 0,
        "tables": 0,
        "digital": sum(1 for page in build.pages if page.kind != "scan"),
        "scans": sum(1 for page in build.pages if page.kind == "scan"),
        "retries": list(build.counts.get("retries") or []),
        "skipped": list(build.counts.get("skipped") or []),
        "page_plans": list(build.counts.get("page_plans") or []),
        "decision_log": list(build.counts.get("decision_log") or []),
        "missing_fonts": list(build.counts.get("missing_fonts") or []),
    }
    token = MERGE_PARAGRAPHS.set(bool(options.merge_paragraphs))
    direction_token = _DIRECTION.set(direction)
    try:
        # A faithful page sets the direction of each line itself, so the
        # document follows the letters it holds even when the caller left the
        # direction at its Arabic default.
        any_faithful = any(getattr(page, "faithful", None) is not None for page in build.pages)
        follows_letters = direction == "auto" or (any_faithful and direction == "rtl")
        if direction == "ltr" or (follows_letters and build.chrome.majority == "ltr"):
            _mark_document_ltr(word)
            section_bidi = False
        else:
            _mark_document_rtl(word)
            section_bidi = True
        started = False
        signature = None
        for record in build.pages:
            payload = getattr(record, "faithful", None)
            if payload is not None:
                from core import faithful

                scale, offset_x, offset_y, page_size = faithful.frame_for(
                    record.width, record.height, bool(options.fit_a4)
                )
                # Body page labels become a live PAGE field with one format.
                label_format = getattr(payload, "page_label_format", "") or ""
                # Only a page whose own number line was lifted gets a footer
                # field; otherwise the footer would add a "N | Page" line the
                # PDF never printed.
                use_page_field = bool(options.lock_footer and label_format)
                label_place = getattr(payload, "page_label_place", "") or "footer"
                footer_key = (label_format, getattr(payload, "page_label_side", "") or build.chrome.footer_number_side, label_place)
                sign = (round(page_size[0], 1), round(page_size[1], 1), "faithful", footer_key if use_page_field else ())
                if not started or sign != signature:
                    _start_section(
                        word,
                        None,
                        not started,
                        build.chrome,
                        True,
                        record.number - 1,
                        show_header=False,
                        show_footer=use_page_field,
                        page_size=page_size,
                        section_bidi=section_bidi,
                    )
                    if use_page_field:
                        _apply_faithful_page_footer(
                            word.sections[-1],
                            build.chrome,
                            label_format or "{PAGE} | Page",
                            getattr(payload, "page_label_side", "") or "left",
                            section_bidi=section_bidi,
                            place=label_place,
                            start=getattr(payload, "page_label_number", None),
                        )
                    started = True
                else:
                    _add_page_break(word)
                signature = sign
                faithful.write_page(word, payload, scale, offset_x, offset_y, counts)
                counts["pages"] += 1
                continue
            blocks, frame_w, frame_h, page_size = _frame_for(record, options)
            cover = bool(record.cover or _is_cover_sheet(record.blocks))
            header_on = bool(options.lock_header and record.show_header and not cover)
            footer_on = bool(options.lock_footer and record.show_footer and not cover)
            sign = (round(frame_w, 1), round(frame_h, 1), cover, header_on, footer_on)
            if not started or sign != signature:
                _start_section(
                    word,
                    None,
                    not started,
                    build.chrome,
                    cover,
                    record.number - 1,
                    show_header=header_on,
                    show_footer=footer_on,
                    lock_header=True,
                    lock_footer=True,
                    page_size=page_size,
                    section_bidi=section_bidi,
                )
                started = True
            else:
                _add_page_break(word)
            signature = sign
            if getattr(build.options, "flow", True) and not getattr(record, "frames", False):
                from core.flow_layout import write_flow_page

                # The Word header already carries the running header's art.
                header_top = 0.0
                if header_on and build.chrome.header_png and build.chrome.header_band > 0:
                    header_top = float(build.chrome.header_top) + float(build.chrome.header_band)
                write_flow_page(word, blocks, frame_w, frame_h, counts, header_bottom=header_top)
            else:
                _write_page(word, blocks, frame_w, frame_h, counts)
            counts["pages"] += 1
        _collapse_section_breaks(word)
        word.save(output)
    finally:
        _DIRECTION.reset(direction_token)
        MERGE_PARAGRAPHS.reset(token)
    return counts


def rescan_word_page(
    pdf_path: str | Path,
    build: WordBuild,
    page_number: int,
    mode: str,
    password: str = "",
) -> WordPage:
    """
    Re-read one page.

    `rtl` rebuilds the paragraphs in reading order and drops overlapping copies.
    `images` replaces figures only. `tables` replaces grids only.
    `full` reads the page again with the same layout choices.
    """
    record = _find_word_page(build, page_number)
    source = Path(pdf_path).expanduser()
    document_pdf, plumber = _open_readers(source, password)
    try:
        if page_number < 1 or page_number > document_pdf.page_count:
            raise ConversionError(f"Page {page_number} is not in this PDF.")
        background = _running_backgrounds(document_pdf)
        page = document_pdf.load_page(page_number - 1)
        plumber_page = _plumber_at(plumber, page_number - 1)
        furniture = _furniture_for([page])
        schema = _ColumnSchema()
        notes = build.counts.setdefault("retries", [])
        frozen = record.cover or _is_cover_sheet(record.blocks)
        if mode not in {"full", "images", "tables", "rtl"}:
            raise ConversionError(f"Unknown page mode: {mode}")
        if _is_faithful(build.options):
            # A faithful page is one unit: its picture and its text are read
            # together, so every re-scan reads the whole page again. `rtl`
            # also drops glyphs that repeat a neighbour.
            fresh = _capture_faithful(
                page,
                plumber_page,
                background,
                furniture,
                schema,
                page_number,
                notes,
                build.options,
                "rtl" if mode == "rtl" else "",
            )
            record.undo = _remember_page(record)
            record.blocks = fresh.blocks
            record.kind = fresh.kind
            record.cover = fresh.cover
            record.width = fresh.width
            record.height = fresh.height
            record.faithful = fresh.faithful
        elif mode == "full":
            fresh = _capture_page(
                page, plumber_page, background, furniture, schema, page_number, notes, build.options
            )
            snapshot = _remember_page(record)
            record.blocks = fresh.blocks
            record.kind = fresh.kind
            record.cover = fresh.cover
            record.width = fresh.width
            record.height = fresh.height
            if fresh.cover:
                record.show_header = False
                record.show_footer = False
            record.undo = snapshot
        elif mode == "images":
            fresh_blocks = _page_picture(page) if frozen else _refresh_images(page, record.blocks, background)
            record.undo = _remember_page(record)
            record.blocks = fresh_blocks
        elif mode == "tables":
            if frozen:
                raise ConversionError("This page is kept as a cover picture.")
            fresh_blocks = _refresh_tables(page, plumber_page, record.blocks, build.options.tables_as_images)
            record.undo = _remember_page(record)
            record.blocks = fresh_blocks
        elif mode == "rtl":
            if frozen:
                raise ConversionError("This page is kept as a cover picture.")
            fresh = _capture_page(
                page,
                plumber_page,
                background,
                furniture,
                schema,
                page_number,
                notes,
                build.options,
                "rtl",
            )
            if not any(isinstance(block, TextBlock) and block.text.strip() for block in fresh.blocks):
                raise ConversionError("That page did not produce any text.")
            record.undo = _remember_page(record)
            record.blocks = fresh.blocks
            record.kind = fresh.kind
            record.cover = False
        else:
            raise ConversionError(f"Unknown page mode: {mode}")
    finally:
        _close_readers(document_pdf, plumber)
    return record


def undo_word_page(build: WordBuild, page_number: int) -> WordPage:
    """Put back the blocks this page had before its latest re-scan."""
    record = _find_word_page(build, page_number)
    if not record.undo:
        raise ConversionError("This page has no re-scan to undo.")
    saved = record.undo
    record.blocks = saved["blocks"]
    record.show_header = saved["show_header"]
    record.show_footer = saved["show_footer"]
    record.cover = saved["cover"]
    record.kind = saved["kind"]
    record.faithful = saved.get("faithful")
    record.undo = None
    return record


def set_word_page_chrome(build: WordBuild, page_number: int, show_header: bool, show_footer: bool) -> WordPage:
    """Hide or show the running header and footer on one page."""
    record = _find_word_page(build, page_number)
    show_header = bool(show_header)
    show_footer = bool(show_footer)
    if record.show_header == show_header and record.show_footer == show_footer:
        return record
    record.undo = _remember_page(record)
    record.show_header = show_header
    record.show_footer = show_footer
    return record


def convert_arabic_pdf_to_docx(
    pdf_path: str | Path,
    docx_path: str | Path,
    password: str = "",
    options: WordOptions | None = None,
) -> dict:
    """
    Convert one Arabic PDF to a .docx file.

    Each PDF page is written into one Word page. A new section is opened when
    the sheet size changes, or when one page needs a different header or footer.
    Pages that share a section are separated by a single page break. Nothing in
    one page is joined to the text or the table of another page.
    """
    build = read_word_build(pdf_path, password, options)
    counts = write_word_build(build, docx_path)
    try:
        from core.docx_validate import validate_docx

        counts["docx_validation"] = validate_docx(docx_path)
    except Exception as exc:
        counts["docx_validation"] = {"ok": False, "errors": [str(exc)], "warnings": []}
    return counts


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convert an Arabic PDF to a right-to-left Word document."
    )
    parser.add_argument("pdf", help="Path to the PDF")
    parser.add_argument("docx", nargs="?", help="Path to the .docx file to write")
    parser.add_argument("--password", default="", help="Password, if the PDF asks for one")
    parser.add_argument(
        "--mode",
        default="faithful",
        choices=("faithful", "editable"),
        help="faithful: the Word page looks like the PDF page (default). editable: flowing paragraphs and tables",
    )
    parser.add_argument(
        "--fit-a4",
        action="store_true",
        help="Scale every page onto A4. Off by default: the Word page keeps the PDF page size",
    )
    parser.add_argument("--pages", default="", help="Pages to convert, such as 1-5 or 3,4,7. Empty means all")
    args = parser.parse_args(argv)
    source = Path(args.pdf)
    target = Path(args.docx) if args.docx else source.with_suffix(".docx")
    try:
        options = WordOptions(mode=args.mode, fit_a4=bool(args.fit_a4), page_spec=args.pages or "")
        counts = convert_arabic_pdf_to_docx(source, target, args.password, options)
    except ConversionError as exc:
        print(exc, file=sys.stderr)
        return 1
    print(
        f"Wrote {target}  pages={counts['pages']}  "
        f"paragraphs={counts['paragraphs']}  headings={counts['headings']}  "
        f"tables={counts['tables']}  digital={counts['digital']}  scans={counts['scans']}  "
        f"retries={len(counts['retries'])}"
    )
    for note in counts["retries"]:
        print(note, file=sys.stderr)
    for note in counts["skipped"]:
        print(note, file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
