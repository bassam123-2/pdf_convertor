# -*- coding: utf-8 -*-
"""
Faithful mode: the Word page looks and prints like the PDF page.

The page is written in two layers.

The background is everything that is not text: fills, lines, shapes,
gradients, and pictures. It is drawn once from a copy of the page whose text
has been removed, and placed as one page-size picture behind the text. The
text is taken out of the copy, not painted over, so no letter is drawn twice.

The text is real Unicode text. Each line is one text box at the position it
had on the PDF page, with the PDF's font family, size, weight, slant, and
color. Word does not set type exactly as the PDF did, so each word and each
space is given the width it had on the PDF page. A line therefore ends where
it ended on the PDF and does not run into its neighbour.

Arabic goes through the same reader as the flowing mode: letters are turned
back into logical order and lam-alef stays one unit. A line that reads from
the right is a right-to-left paragraph pinned to its right edge.

Text that cannot be written as text stays in the picture: text set at an
angle, outlined text, and letters from a font whose Unicode map is wrong when
the picture reader cannot read them either. Nothing is dropped.
"""

from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass, field
from io import BytesIO
from xml.sax.saxutils import escape

import fitz
from docx.oxml import parse_xml
from docx.shared import Emu

from core import font_catalog
from core.arabic_fixer import (
    _BIDI_MARKS,
    _Atom,
    _fold_doachashmee_forms,
    _is_arabic,
    _logical_pieces,
    _logical_text,
)

_A4_SHORT = 595.28
_A4_LONG = 841.89

# Word puts the baseline of an exact-height line four fifths of the way down
# the line, whatever the font. Measured in Word and in LibreOffice.
_BASELINE_SHARE = 0.8

# The longer side of the background picture, in pixels, and the limits on the
# zoom that gives. About 190 dots per inch on an A4 page.
_PICTURE_SIDE = 2200.0
_PICTURE_ZOOM = (1.5, 3.0)

_NAMESPACES = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)
_SHAPE_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


@dataclass
class Chunk:
    """A word or a space on one line, in one font and color."""

    text: str
    font: str
    flags: int
    size: float
    bold: bool
    italic: bool
    color: str
    rtl: bool
    space: bool
    # Width on the PDF page in points. None when the glyphs could not be matched.
    width: float | None
    # How far the glyphs sit above the line's baseline, in points.
    rise: float = 0.0


@dataclass
class Line:
    """One text box: a line, or the part of a line inside one column."""

    x0: float
    y0: float
    x1: float
    y1: float
    baseline: float
    size: float
    rtl: bool
    text: str
    chunks: list[Chunk] = field(default_factory=list)


@dataclass
class FaithfulLink:
    """A clickable region: URL, mailto, or internal page jump."""

    bbox: tuple[float, float, float, float]
    kind: str  # "uri" | "mailto" | "goto"
    target: str


@dataclass
class FaithfulPage:
    """One PDF page as a background picture and its text lines."""

    width: float
    height: float
    lines: list[Line] = field(default_factory=list)
    background: bytes = b""
    # Glyphs left in the picture because they could not be written as text.
    kept_in_picture: int = 0
    # Glyphs whose font map is wrong and that were read from the picture.
    read_from_picture: int = 0
    notes: list[str] = field(default_factory=list)
    # Native tables rebuilt from vector grids or OCR'd raster pictures.
    tables: list = field(default_factory=list)
    # Hyperlinks preserved from the PDF.
    links: list[FaithfulLink] = field(default_factory=list)
    # Fonts named on this page that are not installed locally.
    missing_fonts: list[str] = field(default_factory=list)
    # Page-number chrome stripped from body lines (written as a PAGE field).
    page_label_format: str = ""
    page_label_side: str = ""
    # "header" or "footer": where the PDF printed the page label.
    page_label_place: str = ""
    # The number the PDF printed on this page, when it printed one.
    page_label_number: int | None = None
    # Upright rules (x, y0, y1) used to clamp text boxes inside columns.
    column_rules: list[tuple[float, float, float]] = field(default_factory=list)


def frame_for(width: float, height: float, fit_a4: bool) -> tuple[float, float, float, tuple[float, float]]:
    """
    Scale, left offset, top offset, and the Word page size.

    Without `fit_a4` the Word page is the PDF page. With it, the page is A4
    in the PDF page's own orientation and the content is scaled as a whole.
    """
    if not fit_a4:
        return 1.0, 0.0, 0.0, (width, height)
    sheet = (_A4_LONG, _A4_SHORT) if width > height else (_A4_SHORT, _A4_LONG)
    if abs(width - sheet[0]) <= 4 and abs(height - sheet[1]) <= 4:
        return 1.0, 0.0, 0.0, sheet
    scale = min(sheet[0] / max(width, 1.0), sheet[1] / max(height, 1.0))
    return scale, (sheet[0] - width * scale) / 2.0, (sheet[1] - height * scale) / 2.0, sheet


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------


def _fill_regions(page: fitz.Page) -> list[tuple[float, float, float, float, float]]:
    """Filled shapes and pictures: the panels a block of text can sit on."""
    regions = []
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        drawings = []
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None or drawing.get("fill") is None:
            continue
        width = float(rect.x1) - float(rect.x0)
        height = float(rect.y1) - float(rect.y0)
        if width < 4 or height < 4:
            continue
        regions.append((float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1), width * height))
    try:
        images = page.get_image_info()
    except Exception:
        images = []
    for image in images or []:
        box = image.get("bbox")
        if not box:
            continue
        width = float(box[2]) - float(box[0])
        height = float(box[3]) - float(box[1])
        if width >= 4 and height >= 4:
            regions.append((float(box[0]), float(box[1]), float(box[2]), float(box[3]), width * height))
    return regions


def _vertical_rules(page: fitz.Page) -> list[tuple[float, float, float]]:
    """Upright lines on the page, such as the rules between table columns."""
    rules: list[tuple[float, float, float]] = []
    try:
        from core.table_rebuild import visible_drawings
        drawings = visible_drawings(page)
    except Exception:
        return rules
    for drawing in drawings:
        rect = drawing.get("rect")
        stroked = drawing.get("color") is not None
        if rect is not None and drawing.get("fill") is not None:
            if float(rect.x1) - float(rect.x0) <= 2.5 and float(rect.y1) - float(rect.y0) >= 3:
                rules.append(((float(rect.x0) + float(rect.x1)) / 2.0, float(rect.y0), float(rect.y1)))
        if not stroked:
            continue
        for item in drawing.get("items") or []:
            kind = item[0]
            if kind == "l":
                first, second = item[1], item[2]
                if abs(float(first.x) - float(second.x)) <= 0.75 and abs(float(first.y) - float(second.y)) >= 3:
                    rules.append((float(first.x), min(float(first.y), float(second.y)), max(float(first.y), float(second.y))))
            elif kind in {"re", "qu"}:
                box = item[1] if kind == "re" else item[1].rect
                if float(box.y1) - float(box.y0) >= 3:
                    rules.append((float(box.x0), float(box.y0), float(box.y1)))
                    rules.append((float(box.x1), float(box.y0), float(box.y1)))
    return rules


def _region_at(regions: list[tuple], x: float, y: float) -> int:
    """The smallest filled shape under this point, or -1 for bare paper."""
    best = -1
    area = float("inf")
    for index, (left, top, right, bottom, size) in enumerate(regions):
        if left <= x <= right and top <= y <= bottom and size < area:
            best = index
            area = size
    return best


def _center(atom: _Atom) -> tuple[float, float]:
    return (atom.x0 + atom.x1) / 2.0, (atom.y0 + atom.y1) / 2.0


def _split_band(band: list[_Atom], regions: list[tuple], rules: list[tuple], gap_ems: float = 2.0) -> list[list[_Atom]]:
    """
    Cut one baseline into the pieces that belong to different blocks.

    Glyphs on the same baseline are one line only when they are neighbours.
    The line is cut at a gap of two ems, at a smaller gap when the two sides
    sit on different filled shapes, and where an upright rule passes between
    them. A sidebar and the paragraph beside it share a baseline and nothing
    else, so they become two boxes.
    """
    ordered = sorted(band, key=lambda atom: (atom.x0, atom.origin_x))
    parts: list[list[_Atom]] = []
    current: list[_Atom] = []
    waiting: list[_Atom] = []
    edge = 0.0
    last: _Atom | None = None
    for atom in ordered:
        if not atom.text.strip():
            waiting.append(atom)
            continue
        cut = False
        if last is not None:
            gap = atom.x0 - edge
            size = max(last.size, atom.size, 1.0)
            if gap >= gap_ems * size:
                cut = True
            elif gap >= 0.5 * size and _region_at(regions, *_center(last)) != _region_at(regions, *_center(atom)):
                cut = True
            elif gap > 0.5:
                middle = (atom.y0 + atom.y1) / 2.0
                cut = any(
                    edge - 0.5 <= x <= atom.x0 + 0.5 and top <= middle <= bottom for x, top, bottom in rules
                )
        if cut:
            parts.append(current)
            current = []
            edge = atom.x1
        else:
            current.extend(waiting)
            edge = atom.x1 if last is None else max(edge, atom.x1)
        waiting = []
        current.append(atom)
        last = atom
    if current:
        parts.append(current)
    return parts


class _PictureWord(_Atom):
    """A word read from the page picture, standing where untrusted glyphs stood."""


def _unmapped(char: str) -> bool:
    """A character that stands for 'no Unicode value': the font map has a hole here."""
    code = ord(char)
    if code == 0xFFFD or 0xE000 <= code <= 0xF8FF or 0xF0000 <= code <= 0x10FFFF:
        return True
    return code < 0x20 and char not in "\t\n\r"


def _overfull_arabic_text(text: str, box: float, size: float) -> bool:
    """
    True when an Arabic string stores more letters than its box can hold.

    Broken ToUnicode maps often emit duplicated/mangled letter runs
    (e.g. تعامتتعاملات). Fitting that string into the original width makes
    Word draw overlapping glyphs that look like fake bold.
    """
    letters = [char for char in text if _is_arabic(char)]
    if len(letters) < 5:
        return False
    box = max(0.5, box)
    compact = "".join(letters)
    # Concatenated twin in visual or logical order (glyphs may be stored either way).
    for candidate in (compact, compact[::-1]):
        for span in range(3, min(7, len(candidate) // 2 + 1)):
            prefix = candidate[:span]
            if not prefix:
                continue
            # A twin repeats right after itself (one stray letter allowed).
            # A prefix that only comes back later, as لات in التعاملات, is
            # how ordinary words are spelt.
            idx = candidate.find(prefix, span)
            if idx >= span and idx - span <= 1:
                return True
    # Extreme density only — Arabic is narrow; avoid flagging normal words.
    expected = size * 0.42 * len(letters)
    return len(letters) >= 8 and expected > box * 2.15


def _overfull_arabic(atom: _Atom) -> bool:
    """True when a single atom's Arabic letters cannot fit its own box."""
    return _overfull_arabic_text(atom.text, atom.x1 - atom.x0, atom.size)


def _overfull_token_ids(atoms: list[_Atom]) -> set[int]:
    """
    Mark every glyph in an Arabic *word* whose assembled ToUnicode is bad.

    Glyphs are often one letter each; mangling shows once they join. Space
    glyphs (and gaps ≥ 0.35 em) end a word so a whole line is not one token.
    """
    from core.docx_converter import _cluster_baselines

    marked: set[int] = set()
    for band in _cluster_baselines(atoms):
        ordered = sorted(band, key=lambda atom: (atom.x0, atom.origin_x))
        token: list[_Atom] = []
        edge = 0.0

        def flush() -> None:
            nonlocal token
            letters = [atom for atom in token if atom.text.strip()]
            token = []
            if not letters:
                return
            text = "".join(atom.text for atom in letters)
            box = max(atom.x1 for atom in letters) - min(atom.x0 for atom in letters)
            size = max(atom.size for atom in letters)
            logical = ""
            try:
                logical = _logical_text(letters) or ""
            except Exception:
                logical = text
            damaged = any(_unmapped(char) for atom in letters for char in atom.text)
            # A word with an unmapped hole glues into garbage once that hole is
            # dropped — mark the whole word so OCR can replace it in one piece.
            if damaged or _overfull_arabic_text(text, box, size) or _overfull_arabic_text(logical, box, size):
                for atom in letters:
                    marked.add(id(atom))

        for atom in ordered:
            blank = not atom.text.strip()
            if blank:
                flush()
                continue
            if not token:
                token = [atom]
                edge = atom.x1
                continue
            gap = atom.x0 - edge
            if gap > max(atom.size, token[-1].size) * 0.35:
                flush()
                token = [atom]
            else:
                token.append(atom)
            edge = max(edge, atom.x1)
        flush()
    return marked


# Letters and signs that never sit inside an Arabic word: Latin letters,
# Latin-1 signs and stray symbols a broken font map leaves behind.
_FOREIGN_IN_ARABIC = re.compile(r"[A-Za-z\u00a1-\u00bf\u00c0-\u024f\u2022\u2122\\\[\]{}|^~`@#$&*_<>]")


def _mixed_token_ids(atoms: list[_Atom]) -> set[int]:
    """
    Glyphs of Arabic words that carry a foreign letter or sign ("KŠيوما", "nي").

    A font map that is only partly repaired leaves such words: most letters
    read well, one or two stay Latin. The word is read again from the page
    picture. An English word, a number or an Arabic word with ordinary
    punctuation is not touched.
    """
    from core.docx_converter import _cluster_baselines

    found: set[int] = set()
    for band in _cluster_baselines(atoms):
        ordered = sorted(band, key=lambda atom: (atom.x0, atom.origin_x))
        token: list[_Atom] = []
        for atom in ordered + [None]:
            gap = atom is not None and token and atom.x0 - token[-1].x1 > max(1.5, atom.size * 0.25)
            if atom is not None and atom.text.strip() and not gap:
                token.append(atom)
                continue
            if token:
                text = "".join(item.text for item in token)
                arabic_fonts = {item.font for item in token if any(_is_arabic(char) for char in item.text)}
                # The foreign letter must come from the Arabic letters' own
                # font: a bullet from another font beside a word is a bullet.
                foreign = [item for item in token if _FOREIGN_IN_ARABIC.search(item.text) and item.font in arabic_fonts]
                if arabic_fonts and foreign and len(text) <= 24:
                    found.update(id(item) for item in token)
            token = [atom] if atom is not None and atom.text.strip() else []
    return found


def _suspect_atoms(page: fitz.Page, atoms: list[_Atom]) -> set[int]:
    """
    Glyphs whose letters cannot be trusted, by font.

    A glyph with no Unicode value is always one. A font is one when a tenth
    of its letters have no Unicode value. On a page whose Arabic words carry
    Latin letters inside them, the fonts that write those words are too: the
    glyph was drawn as an Arabic letter and stored as a Latin one. Other
    fonts on the same page keep their text. Overfull Arabic tokens (more
    letters than the box allows) are also suspects.
    """
    from core.docx_converter import _cluster_baselines, _encoding_broken

    suspects: set[int] = set()
    totals: dict[str, int] = {}
    holes: dict[str, int] = {}
    for atom in atoms:
        text = atom.text.strip()
        if not text:
            continue
        totals[atom.font] = totals.get(atom.font, 0) + len(text)
        missing = sum(1 for char in text if _unmapped(char))
        if missing:
            holes[atom.font] = holes.get(atom.font, 0) + missing
            suspects.add(id(atom))
        if _overfull_arabic(atom):
            suspects.add(id(atom))
    suspects |= _overfull_token_ids(atoms)
    suspects |= _mixed_token_ids(atoms)
    bad_fonts = {font for font, count in holes.items() if count >= 3 and count >= totals.get(font, 0) * 0.1}
    try:
        broken = _encoding_broken(page, atoms)
    except Exception:
        broken = False
    if broken:
        mixed: dict[str, int] = {}
        arabic_fonts: set[str] = set()
        for band in _cluster_baselines(atoms):
            ordered = sorted(band, key=lambda atom: (atom.x0, atom.origin_x))
            token: list[_Atom] = []
            for atom in ordered + [None]:
                if atom is not None and atom.text.strip():
                    token.append(atom)
                    continue
                arabic = {item.font for item in token if any(_is_arabic(char) for char in item.text)}
                latin = {item.font for item in token if any(char.isascii() and char.isalpha() for char in item.text)}
                arabic_fonts |= arabic
                for font in arabic & latin:
                    mixed[font] = mixed.get(font, 0) + 1
                token = []
        guilty = {font for font, count in mixed.items() if count >= 2}
        bad_fonts |= guilty or arabic_fonts
    if bad_fonts:
        for atom in atoms:
            if atom.font in bad_fonts and atom.text.strip():
                suspects.add(id(atom))
    return suspects


def _dedupe_fake_bold(atoms: list[_Atom]) -> list[_Atom]:
    """
    Keep one copy when the PDF drew the same word twice with a small offset.

    Fake bold is often stroke+fill or a second draw shifted by a fraction of
    a point. The survivor is marked bold. General: no file-specific rules.

    Glyphs are looked up in a grid of small cells, so each glyph is compared
    only with the glyphs near it. Comparing every pair took seconds on a
    dense page.
    """
    ordered = sorted(atoms, key=lambda atom: (round(atom.origin_y, 1), atom.x0, -((atom.x1 - atom.x0))))
    kept: list[_Atom] = []
    cell = 8.0
    grid: dict[tuple[int, int], list[_Atom]] = {}
    for atom in ordered:
        text = atom.text.strip()
        if not text:
            kept.append(atom)
            continue
        cx = (atom.x0 + atom.x1) / 2.0
        gx, gy = int(cx // cell), int(atom.origin_y // cell)
        duplicate = False
        for nx in (gx - 1, gx, gx + 1):
            for ny in (gy - 1, gy, gy + 1):
                for other in grid.get((nx, ny), ()):
                    if other.text.strip() != text:
                        continue
                    if abs(atom.origin_y - other.origin_y) > max(1.5, atom.size * 0.25):
                        continue
                    # An overprint copy is shifted by a fraction of its own glyph.
                    # The next letter of a double letter (the second l of "all")
                    # is a whole glyph further on, however narrow the glyph is.
                    dx = abs(cx - ((other.x0 + other.x1) / 2.0))
                    glyph = max(0.5, min(atom.x1 - atom.x0, other.x1 - other.x0))
                    if dx > min(max(2.5, atom.size * 0.35), glyph * 0.5):
                        continue
                    other.bold = True
                    other.flags = int(other.flags) | 16
                    duplicate = True
                    break
                if duplicate:
                    break
            if duplicate:
                break
        if not duplicate:
            kept.append(atom)
            grid.setdefault((gx, gy), []).append(atom)
    return kept


def _word_around(ordered: list[_Atom], zone: list[_Atom]) -> list[_Atom]:
    """Glyphs of the visual word that contains this zone (stops at spaces)."""
    letters = [atom for atom in zone if atom.text.strip()]
    if not letters:
        return list(zone)
    index = {id(atom): i for i, atom in enumerate(ordered)}
    left_i = min(index[id(atom)] for atom in letters if id(atom) in index)
    right_i = max(index[id(atom)] for atom in letters if id(atom) in index)
    while left_i > 0:
        prev = ordered[left_i - 1]
        cur = ordered[left_i]
        if not prev.text.strip():
            break
        if cur.x0 - prev.x1 > max(prev.size, cur.size) * 0.35:
            break
        left_i -= 1
    while right_i + 1 < len(ordered):
        cur = ordered[right_i]
        nxt = ordered[right_i + 1]
        if not nxt.text.strip():
            break
        if nxt.x0 - cur.x1 > max(nxt.size, cur.size) * 0.35:
            break
        right_i += 1
    return [atom for atom in ordered[left_i : right_i + 1] if atom.text.strip() or atom in zone]


def _spacer_atom(host: _Atom, x0: float, x1: float, baseline: float) -> _Atom:
    """A space that keeps neighbours from gluing across a removed hole."""
    return _Atom(
        " ",
        x0,
        host.y0,
        max(x0 + 0.5, x1),
        host.y1,
        x0,
        baseline,
        host.level,
        host.size,
        host.bold,
        host.color,
        host.font,
        host.italic,
        host.flags,
        host.space,
    )


def _ocr_words_for(page: fitz.Page, atoms: list[_Atom], suspects: set[int]) -> list[_Atom]:
    """
    Words read from the page picture where the suspect glyphs are.

    A few suspect words are read from strips around their lines, which is
    quicker and reads small type larger. When suspects cover much of the
    page, the whole page is read once.
    """
    from core.docx_converter import _ocr_word_atoms

    boxes = []
    for atom in atoms:
        if id(atom) in suspects:
            boxes.append(fitz.Rect(page.rect.x0, atom.y0 - 3, page.rect.x1, atom.y1 + 3))
    if not boxes:
        return []
    boxes.sort(key=lambda box: box.y0)
    strips: list[fitz.Rect] = []
    for box in boxes:
        if strips and box.y0 <= strips[-1].y1 + 4:
            strips[-1] |= box
        else:
            strips.append(fitz.Rect(box))
    covered = sum(strip.height for strip in strips)
    if covered > page.rect.height * 0.35 or len(strips) > 12:
        return _ocr_word_atoms(page)
    words: list[_Atom] = []
    for strip in strips:
        words.extend(_ocr_word_atoms(page, strip))
    return words


def _replace_unreadable(
    bands: list[list[_Atom]],
    suspects: set[int],
    words: list[_Atom],
    keep_unread: bool = False,
) -> tuple[list[list[_Atom]], list[_Atom], list[_Atom], int]:
    """
    Swap the untrusted glyphs for the words the picture reader found there.

    OCR words replace the stretch (never sit beside the old glyphs). A hole
    with no OCR hit stays in the picture and a spacer is left so neighbours
    do not glue into one mangled token. Whole-word expansion runs only when
    most of the word is already suspect (overfull / heavy damage).

    Returns the bands, the glyphs to remove from the picture, the glyphs that
    stay in the picture, and how many glyphs were read from the picture.
    """
    taken: set[int] = set()
    result: list[list[_Atom]] = []
    replaced: list[_Atom] = []
    kept: list[_Atom] = []
    read = 0
    for band in bands:
        ordered = sorted(band, key=lambda atom: (atom.x0, atom.origin_x))
        zones: list[list[_Atom]] = []
        current: list[_Atom] = []
        for atom in ordered:
            if id(atom) in suspects:
                current.append(atom)
            elif not atom.text.strip() and current:
                current.append(atom)
            elif current:
                zones.append(current)
                current = []
        if current:
            zones.append(current)
        if not zones:
            result.append(band)
            continue
        resolved: list[list[_Atom]] = []
        for zone in zones:
            letters = [atom for atom in zone if atom.text.strip()]
            if not letters:
                continue
            word = _word_around(ordered, zone)
            word_letters = [atom for atom in word if atom.text.strip()]
            suspect_n = sum(1 for atom in word_letters if id(atom) in suspects)
            # Expand to the whole word only when damage dominates that word.
            if word_letters and suspect_n >= max(2, int(len(word_letters) * 0.4 + 0.5)):
                resolved.append(word)
            else:
                resolved.append(list(zone))
        inside = {id(atom) for zone in resolved for atom in zone}
        line = [atom for atom in band if id(atom) not in inside]
        for zone in resolved:
            letters = [atom for atom in zone if atom.text.strip()]
            if not letters:
                continue
            left = min(atom.x0 for atom in letters)
            right = max(atom.x1 for atom in letters)
            top = min(atom.y0 for atom in letters)
            bottom = max(atom.y1 for atom in letters)
            fonts: dict[str, int] = {}
            for atom in letters:
                fonts[atom.font] = fonts.get(atom.font, 0) + 1
            host = next(atom for atom in letters if atom.font == max(fonts, key=fonts.get))
            baseline = sorted(atom.origin_y for atom in letters)[len(letters) // 2]
            slack = host.size * 0.45
            found = [
                word
                for word in words
                if id(word) not in taken
                and left - 1.5 <= (word.x0 + word.x1) / 2.0 <= right + 1.5
                and top - slack <= (word.y0 + word.y1) / 2.0 <= bottom + slack
            ]
            if not found and keep_unread:
                # No page picture behind the text: the PDF's own letters are
                # better than a hole.
                line.extend(zone)
                continue
            if not found:
                kept.extend(letters)
                # Spacer stops اخقي + الأخلاقية from becoming one token.
                line.append(_spacer_atom(host, left, right, baseline))
                continue
            unique: list[_Atom] = []
            for word in sorted(found, key=lambda item: (item.x0, item.y0)):
                taken.add(id(word))
                twin = False
                for prior in unique:
                    if prior.text.strip() != word.text.strip():
                        continue
                    if abs(((prior.x0 + prior.x1) / 2.0) - ((word.x0 + word.x1) / 2.0)) <= 2.5:
                        twin = True
                        break
                if twin:
                    continue
                unique.append(word)
            for word in unique:
                line.append(
                    _PictureWord(
                        word.text,
                        word.x0,
                        top,
                        word.x1,
                        bottom,
                        word.x0,
                        baseline,
                        word.level,
                        host.size,
                        host.bold,
                        host.color,
                        host.font,
                        host.italic,
                        host.flags,
                        host.space,
                    )
                )
            replaced.extend(letters)
            read += len(letters)
        if line:
            result.append(line)
    return result, replaced, kept, read


def _styled_chars(pieces: list[tuple[_Atom, str]]) -> list[tuple[str, _Atom]]:
    """
    The line's characters as Word should store them, each with its glyph.

    This is `plain_arabic` done one character at a time, so every character
    keeps the font and color of the glyph it came from.
    """
    raw = "".join(part for _atom, part in pieces)
    folded = _fold_doachashmee_forms(raw)
    spread: list[tuple[str, _Atom]] = []
    position = 0
    for atom, part in pieces:
        for _char in part:
            char = folded[position]
            position += 1
            for item in unicodedata.normalize("NFKC", char):
                if item in _BIDI_MARKS or item == "\u00ad":
                    continue
                if item in "\xa0\t":
                    item = " "
                spread.append((item, atom))
    cleaned: list[tuple[str, _Atom]] = []
    for char, atom in spread:
        if char.isspace():
            if not cleaned or cleaned[-1][0] == " ":
                continue
            char = " "
        cleaned.append((char, atom))
    while cleaned and cleaned[-1][0] == " ":
        cleaned.pop()
    return cleaned


def _reads_right_to_left(text: str, level_odd: bool) -> bool:
    """Whether Word should treat this word or mark as right-to-left text."""
    strong_right = False
    strong_left = False
    for char in text:
        kind = unicodedata.bidirectional(char)
        if kind in {"R", "AL", "AN"}:
            strong_right = True
        elif kind in {"L", "EN"}:
            strong_left = True
    if strong_right:
        return True
    if strong_left:
        return False
    return level_odd


def _slot(atom: _Atom, letters: list[_Atom]) -> float | None:
    """The room a space takes on the PDF line: the gap between the words on either side."""
    middle = (atom.x0 + atom.x1) / 2.0
    left = [other.x1 for other in letters if other is not atom and (other.x0 + other.x1) / 2.0 < middle]
    right = [other.x0 for other in letters if other is not atom and (other.x0 + other.x1) / 2.0 > middle]
    if not atom.text.strip():
        if left and right:
            return max(0.0, min(right) - max(left))
        return max(0.0, atom.x1 - atom.x0) or None
    # A word from the picture reader carries the space that follows it.
    if atom.level % 2 == 1:
        return max(0.0, atom.x0 - max(left)) if left else None
    return max(0.0, min(right) - atom.x1) if right else None


def _style_of(atom: _Atom, baseline: float) -> tuple:
    rise = baseline - atom.origin_y
    if abs(rise) < 0.75:
        rise = 0.0
    return (atom.font, atom.flags, round(atom.size, 2), atom.bold, atom.italic, atom.color, round(rise * 2) / 2.0)


def _line_is_arabic(text: str) -> bool:
    """True when the visible letters are mostly Arabic (RTL box required)."""
    arabic = sum(1 for char in text if _is_arabic(char))
    latin = sum(1 for char in text if char.isascii() and char.isalpha())
    return arabic > 0 and arabic >= latin


def _build_line(atoms: list[_Atom], page_reads_left: bool) -> Line | None:
    """One text box from the glyphs of one line piece."""
    if not any(atom.text.strip() for atom in atoms):
        return None
    turned_line = any(atom.level % 2 == 1 for atom in atoms)
    # An English page can still hold a line that is all Arabic (a bilingual
    # title or column). That line reads right to left whatever the page does.
    arabic_letters = sum(1 for atom in atoms for char in atom.text if _is_arabic(char))
    latin_letters = sum(1 for atom in atoms for char in atom.text if char.isascii() and char.isalpha())
    left_base = turned_line and page_reads_left and latin_letters >= arabic_letters
    pieces, turned = _logical_pieces(atoms, ltr_base=left_base)
    letters: list[_Atom] = []
    seen: set[int] = set()
    for atom, _part in pieces:
        if atom.text.strip() and id(atom) not in seen:
            seen.add(id(atom))
            letters.append(atom)
    if not letters:
        return None
    baseline = sorted(atom.origin_y for atom in letters)[len(letters) // 2]
    chars = _styled_chars(pieces)
    text = "".join(char for char, _atom in chars)
    if not left_base:
        trusted = _logical_text(atoms)
        if not trusted:
            return None
        if trusted != text:
            # A repair changed the line in a way the glyphs cannot follow.
            # The trusted string is written whole, in the line's main style.
            chars = []
            text = trusted
    if not text:
        return None
    # On a mostly-English page, Arabic column lines still need an RTL box even
    # when the PDF bidi levels did not mark the run as "turned".
    rtl = bool(turned or _line_is_arabic(text))
    line = Line(
        x0=min(atom.x0 for atom in letters),
        y0=min(atom.y0 for atom in letters),
        x1=max(atom.x1 for atom in letters),
        y1=max(atom.y1 for atom in letters),
        baseline=baseline,
        size=max(atom.size for atom in letters),
        rtl=rtl,
        text=text,
    )
    if not chars:
        weights: dict[tuple, float] = {}
        hosts: dict[tuple, _Atom] = {}
        for atom in letters:
            key = _style_of(atom, baseline)[:6]
            weights[key] = weights.get(key, 0.0) + (atom.x1 - atom.x0)
            hosts.setdefault(key, atom)
        host = hosts[max(weights, key=weights.get)]
        odd = host.level % 2 == 1
        for part in re.findall(r" +|[^ ]+", text):
            blank = not part.strip()
            line.chunks.append(
                Chunk(
                    text=" " if blank else part,
                    font=host.font,
                    flags=host.flags,
                    size=host.size,
                    bold=host.bold,
                    italic=host.italic,
                    color=host.color,
                    rtl=_reads_right_to_left(part, odd or rtl),
                    space=blank,
                    width=None,
                )
            )
        return line
    group: list[tuple[str, _Atom]] = []
    group_key: tuple | None = None

    def close() -> None:
        if not group:
            return
        part = "".join(char for char, _atom in group)
        blank = group[0][0] == " "
        first = group[0][1]
        members: list[_Atom] = []
        marked: set[int] = set()
        for _char, atom in group:
            if id(atom) not in marked:
                marked.add(id(atom))
                members.append(atom)
        if blank:
            width = _slot(first, letters)
        else:
            width = max(atom.x1 for atom in members) - min(atom.x0 for atom in members)
        style = _style_of(first, baseline)
        odd = sum(1 for atom in members if atom.level % 2 == 1) * 2 > len(members)
        line.chunks.append(
            Chunk(
                text=part,
                font=first.font,
                flags=first.flags,
                size=first.size,
                bold=first.bold,
                italic=first.italic,
                color=first.color,
                rtl=_reads_right_to_left(part, odd or rtl),
                space=blank,
                width=width,
                rise=style[6],
            )
        )

    for char, atom in chars:
        key = (char == " ", _style_of(atom, baseline))
        if key != group_key:
            close()
            group = []
            group_key = key
        group.append((char, atom))
    close()
    return line


def _page_reads_left(atoms: list[_Atom]) -> bool:
    """True when the page has more Latin letters than Arabic ones."""
    arabic = 0
    latin = 0
    for atom in atoms:
        for char in atom.text:
            if _is_arabic(char):
                arabic += 1
            elif char.isascii() and char.isalpha():
                latin += 1
    return latin > arabic


def _other_text(page: fitz.Page) -> bool:
    """True when the page paints text this reader does not write: angled or outlined."""
    from core.docx_converter import _upright

    try:
        traces = page.get_texttrace() or []
    except Exception:
        return False
    from core.flow_layout import _filled_spots, _outline_copy

    filled = _filled_spots(traces)
    for trace in traces:
        kind = int(trace.get("type") or 0)
        if kind == 3:
            continue
        if kind == 1 and _outline_copy(trace, filled):
            continue  # an outline over the same filled text, not text of its own
        if kind != 0 or not _upright(trace):
            if any(chr(int(raw[0])).strip() for raw in trace.get("chars") or []):
                return True
    return False


def _removal_boxes(groups: list[list[_Atom]], kept: list[_Atom]) -> list[fitz.Rect]:
    """
    Boxes over the glyphs that are written as text, one per unbroken run.

    A run stops at a glyph that stays in the picture, so that glyph is never
    under a box.
    """
    boxes: list[fitz.Rect] = []
    for group in groups:
        letters = sorted((atom for atom in group if atom.text.strip()), key=lambda atom: atom.x0)
        run: list[_Atom] = []
        for atom in letters + [None]:
            joins = False
            if atom is not None and run:
                edge = max(item.x1 for item in run)
                joins = atom.x0 - edge <= max(atom.size, 2.0) and not any(
                    edge - 0.5 <= (other.x0 + other.x1) / 2.0 <= atom.x0 + 0.5
                    and other.y0 <= (atom.y0 + atom.y1) / 2.0 <= other.y1
                    for other in kept
                )
            if atom is not None and (not run or joins):
                run.append(atom)
                continue
            if run:
                top = min(item.y0 for item in run)
                bottom = max(item.y1 for item in run)
                trim = (bottom - top) * 0.2
                boxes.append(
                    fitz.Rect(
                        min(item.x0 for item in run) - 0.2,
                        top + trim,
                        max(item.x1 for item in run) + 0.2,
                        bottom - trim,
                    )
                )
            run = [atom] if atom is not None else []
    return boxes


def _visible_glyphs(page: fitz.Page) -> int:
    count = 0
    try:
        traces = page.get_texttrace() or []
    except Exception:
        return 0
    for trace in traces:
        if int(trace.get("type") or 0) == 3:
            continue
        count += sum(1 for raw in trace.get("chars") or [] if chr(int(raw[0])).strip())
    return count


def _encode(pixmap: fitz.Pixmap) -> bytes:
    """
    Compact picture bytes for the page background.

    Flat / low-colour graphics become a palette PNG. Photographic pages become
    JPEG when that is clearly smaller. Nothing is upscaled.
    """
    if pixmap.is_unicolor:
        try:
            if tuple(pixmap.pixel(0, 0))[:3] == (255, 255, 255):
                return b""
        except Exception:
            pass
    data = pixmap.tobytes("png")
    compact = _palette_png(data)
    if compact and len(compact) < len(data) * 0.85:
        data = compact
    if len(data) > 1_500_000:
        try:
            smaller = pixmap.tobytes("jpeg", jpg_quality=85)
        except Exception:
            smaller = b""
        if smaller and len(smaller) < len(data) * 0.6:
            return smaller
    return data


def _palette_png(data: bytes) -> bytes:
    """Quantize flat graphics (charts, ruled plates) to a small palette PNG."""
    try:
        from PIL import Image
    except Exception:
        return b""
    try:
        image = Image.open(BytesIO(data))
    except Exception:
        return b""
    if image.mode not in {"RGB", "RGBA", "L", "P"}:
        return b""
    # Photographic pages have many unique colours; skip those.
    sample = image.convert("RGB").resize(
        (min(120, image.width), min(120, image.height)),
        Image.Resampling.NEAREST,
    )
    colours = len({sample.getpixel((x, y)) for x in range(sample.width) for y in range(sample.height)})
    if colours > 48:
        return b""
    quantized = image.convert("RGB").quantize(colors=min(64, max(8, colours + 4)), method=Image.Quantize.MEDIANCUT)
    out = BytesIO()
    quantized.save(out, format="PNG", optimize=True)
    return out.getvalue()


def _zoom(page: fitz.Page) -> float:
    longest = max(float(page.rect.width), float(page.rect.height), 1.0)
    return min(_PICTURE_ZOOM[1], max(_PICTURE_ZOOM[0], _PICTURE_SIDE / longest))


def render_background(page: fitz.Page, boxes: list[fitz.Rect] | None, everything: bool) -> tuple[bytes, int]:
    """
    The page without the text that is written as text.

    With `everything`, all text is removed. Otherwise only the glyphs under
    `boxes` are. Graphics and pictures are never removed. The work is done on
    a copy, so the open PDF is not changed. Returns the picture and the
    number of glyphs still painted in it.
    """
    scratch = fitz.open()
    try:
        scratch.insert_pdf(page.parent, from_page=page.number, to_page=page.number)
        copy = scratch[0]
        if copy.rotation:
            copy.remove_rotation()
        if everything:
            copy.add_redact_annot(copy.rect, fill=False, cross_out=False)
        else:
            for box in boxes or []:
                if box.width > 0 and box.height > 0:
                    copy.add_redact_annot(box, fill=False, cross_out=False)
        if everything or boxes:
            copy.apply_redactions(
                images=fitz.PDF_REDACT_IMAGE_NONE,
                graphics=fitz.PDF_REDACT_LINE_ART_NONE,
                text=fitz.PDF_REDACT_TEXT_REMOVE,
            )
        left = _visible_glyphs(copy)
        zoom = _zoom(copy)
        pixmap = copy.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        return _encode(pixmap), left
    finally:
        scratch.close()


def picture_page(page: fitz.Page, note: str = "") -> FaithfulPage:
    """The whole page as one picture, text included. Nothing on it is editable."""
    payload = FaithfulPage(float(page.rect.width), float(page.rect.height))
    zoom = _zoom(page)
    pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    payload.background = _encode(pixmap)
    payload.kept_in_picture = _visible_glyphs(page)
    if note:
        payload.notes.append(note)
    return payload


@dataclass
class PageReading:
    """A page's glyphs once: doubles removed, untrusted words read from the picture."""

    atoms: list
    bands: list
    replaced: list
    kept: list
    suspects: set
    read: int = 0
    removed: int = 0

    @property
    def letters(self) -> list:
        """Every glyph the readers should use, the picture's words included."""
        return [atom for band in self.bands for atom in band]


_READINGS: dict[tuple, PageReading] = {}


def _words_on_dark_boxes(page: fitz.Page, words: list) -> list:
    """
    Light words on dark bars (a title banner, a header cell) that the page
    reader missed: each such bar is cut out, turned light-on-dark into
    dark-on-light, and read on its own.
    """
    import tempfile

    from core.arabic_fixer import plain_arabic
    from core.docx_converter import _drawing_paints
    from core.image_tables import _tesseract_text, tessdata_for

    _folder, langs = tessdata_for()
    if not langs:
        return []
    found: list = []
    boxes = []
    for x0, y0, x1, y1, colour, area in _drawing_paints(page):
        try:
            red, green, blue = int(colour[0:2], 16), int(colour[2:4], 16), int(colour[4:6], 16)
        except ValueError:
            continue
        if red * 0.3 + green * 0.59 + blue * 0.11 >= 150 or area < 300 or y1 - y0 > 80:
            continue
        rect = fitz.Rect(x0, y0, x1, y1)
        if any(rect.contains(fitz.Point((w.x0 + w.x1) / 2, (w.y0 + w.y1) / 2)) for w in words):
            continue
        if any(rect.intersects(other) and (rect & other).get_area() > rect.get_area() * 0.5 for other in boxes):
            continue
        boxes.append(rect)
    if len(boxes) > 30:
        return []
    with tempfile.TemporaryDirectory(prefix="pdf_tool_bar_") as work:
        for index, rect in enumerate(boxes):
            try:
                pixmap = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=rect, alpha=False)
                pixmap.invert_irect(pixmap.irect)
                path = f"{work}/b{index}.png"
                pixmap.save(path)
                text = " ".join(_tesseract_text(path, langs).split())
            except Exception:
                text = ""
            text = " ".join(token for token in plain_arabic(text).split() if re.search(r"\w", token))
            arabic_letters = sum(1 for char in text if _is_arabic(char))
            latin_words = re.findall(r"[A-Za-z]{4,}", text)
            if arabic_letters < 2 and not any(re.search(r"[aeiouAEIOU]", word) for word in latin_words):
                continue  # pieces of a logo
            size = max(7.0, min(24.0, rect.height * 0.55))
            arabic = any(_is_arabic(char) for char in text)
            found.append(_Atom(f"{text} ", rect.x0 + 2, rect.y0 + 2, rect.x1 - 2, rect.y1 - 2, rect.x0 + 2, rect.y1 - rect.height * 0.3, 1 if arabic else 0, size, True, "FFFFFF"))
    return found


def page_reading(page: fitz.Page, keep_unread: bool = True, from_picture: bool = False) -> PageReading:
    """
    The page's trusted glyphs, shared by the line reader and the table readers.

    Reading the picture is slow, so one page is read once; a table cell then
    gets the same repaired words as the paragraph around it.
    """
    from core.docx_converter import _atoms_on_page, _cluster_baselines, _separate_overlapping_atoms

    key = (id(page.parent), getattr(page.parent, "name", ""), page.number, keep_unread, from_picture)
    cached = _READINGS.get(key)
    if cached is not None:
        return cached
    if from_picture:
        # No text layer (a scan, or "Print to PDF" letters drawn as shapes):
        # the words are read from the page picture, whole.
        from core.docx_converter import _ocr_word_atoms

        try:
            words = _ocr_word_atoms(page)
        except Exception:
            words = []
        try:
            words += _words_on_dark_boxes(page, words)
        except Exception:
            pass
        reading = PageReading(words, _cluster_baselines(words), [], [], set(), len(words), 0)
        _READINGS[key] = reading
        return reading
    atoms = _atoms_on_page(page, upright_only=True)
    before = sum(1 for atom in atoms if atom.text.strip())
    atoms = _separate_overlapping_atoms(_dedupe_fake_bold(atoms))
    removed = before - sum(1 for atom in atoms if atom.text.strip())
    suspects = _suspect_atoms(page, atoms) if atoms else set()
    bands = _cluster_baselines(atoms)
    replaced: list = []
    kept: list = []
    read = 0
    if suspects:
        try:
            words = _ocr_words_for(page, atoms, suspects)
        except Exception:
            words = []
        bands, replaced, kept, read = _replace_unreadable(bands, suspects, words, keep_unread=keep_unread)
    reading = PageReading(atoms, bands, replaced, kept, suspects, read, removed)
    if len(_READINGS) > 16:
        _READINGS.clear()
    _READINGS[key] = reading
    return reading


_JOINED: dict[str, float] = {}


def _joined_ratio(family: str) -> float:
    """
    How much narrower Arabic is joined than letter by letter.

    Joined letters drop their tails, so a word is about a quarter narrower
    than its letters' own widths added up (0.76 measured for Times New
    Roman). Laying out a sample to measure each family was dropped: on
    Windows, MuPDF's layout of an installed font file could end the process.
    """
    return _JOINED.get(family, 0.75)


def _sizes_from_widths(lines: list) -> None:
    """
    Type sizes for lines read from a picture, from how wide each line is.

    A word's box says little about its size: tall letters and dots change
    it from word to word. The line's printed width does not: it is the
    text's width in the output font times the size. Lines of one paragraph
    (close together, sizes alike) share their middle size, so a paragraph
    is not set in three sizes.
    """
    try:
        from core.docx_converter import output_fonts
        from core.flow_layout import _advances

        family = output_fonts()[1] or "Times New Roman"
        advances = _advances(family) or _advances("Times New Roman")
    except Exception:
        advances = None
    if not advances:
        return
    default = advances.get("\u0628", 0.5)
    joined = _joined_ratio(family)
    # The letters' height is steady from line to line but says little about
    # the size in the output font; the width of a long line says that, but
    # a word the reader missed makes it jump. One scale for the whole page,
    # taken from its long lines, turns heights into sizes.
    ratios = []
    for line in lines:
        text = (line.text or "").strip()
        if len(text) < 25:
            continue
        em = sum(advances.get(char, default) * (joined if _is_arabic(char) else 1.0) for char in text)
        base = float(line.size or 10)
        if em > 0 and base > 0:
            ratios.append((line.x1 - line.x0) / em / base)
    scale = sorted(ratios)[len(ratios) // 2] if ratios else 1.0
    scale = max(0.85, min(1.35, scale))
    estimates: dict[int, float] = {index: float(line.size or 10) * scale for index, line in enumerate(lines)}
    order = sorted(range(len(lines)), key=lambda i: lines[i].y0)
    groups: list[list[int]] = []
    for index in order:
        line = lines[index]
        size = estimates.get(index, float(line.size or 10))
        if groups:
            last = lines[groups[-1][-1]]
            last_size = estimates.get(groups[-1][-1], float(last.size or 10))
            close = 0 <= line.y0 - last.y0 <= 2.4 * max(size, last_size)
            alike = abs(size - last_size) <= 0.25 * max(size, last_size)
            overlap = min(line.x1, last.x1) - max(line.x0, last.x0) > 0
            if close and alike and overlap:
                groups[-1].append(index)
                continue
        groups.append([index])
    for group in groups:
        sizes = sorted(estimates[i] for i in group if i in estimates)
        if not sizes:
            continue
        middle = round(sizes[len(sizes) // 2] * 2) / 2
        for i in group:
            lines[i].size = middle
    # One body size for the page: a line read a little short (a word the
    # reader missed) must not come out in its own size. The body size is the
    # size most letters have; lines near it take it.
    weights: dict[float, int] = {}
    for line in lines:
        weights[float(line.size or 10)] = weights.get(float(line.size or 10), 0) + len(line.text or "")
    if not weights:
        return
    body = max(weights, key=weights.get)
    near = sum(count for size, count in weights.items() if body * 0.8 <= size <= body * 1.25)
    body = round(sum(size * count for size, count in weights.items() if body * 0.8 <= size <= body * 1.25) / max(1, near) * 2) / 2
    for line in lines:
        size = float(line.size or 10)
        long_line = len(line.text or "") >= 45
        if body * 0.8 <= size <= body * 1.25 or (long_line and size > body):
            # A long line is body text: a heading is short. Its letters'
            # boxes read tall when the reader took the line's whole height.
            line.size = body
        elif size > body * 1.25 and len(line.text or "") <= 90:
            # The picture says nothing about weight; a line clearly larger
            # than the body text is a heading, and headings are bold.
            for chunk in line.chunks or []:
                chunk.bold = True


def _picture_page_furniture(payload, height: float) -> None:
    """
    Drop the small lines a browser prints at the sheet's edges (the page
    title and time at the top, the address and page count at the bottom).
    """
    kept = []
    for line in payload.lines:
        small = float(line.size or 10) <= 7.5
        if small and (line.y1 <= height * 0.04 or line.y0 >= height * 0.96):
            continue
        kept.append(line)
    payload.lines = kept


def read_page(page: fitz.Page, repair: str = "", background: bool = True, from_picture: bool = False) -> FaithfulPage:
    """
    Read one PDF page into a background picture and text lines.

    `repair="rtl"` also drops a glyph that sits inside another glyph and
    repeats its letters, as the flowing mode's repair pass does.
    """
    payload = FaithfulPage(float(page.rect.width), float(page.rect.height))
    number = page.number + 1
    reading = page_reading(page, keep_unread=not background, from_picture=from_picture and not background)
    atoms = reading.atoms
    if reading.removed:
        payload.notes.append(
            f"Page {number}: removed {reading.removed} overlapping duplicate glyph(s) (fake bold / overprint)."
        )
    bands = reading.bands
    replaced = reading.replaced
    kept = reading.kept
    if reading.suspects:
        payload.read_from_picture = reading.read
        payload.kept_in_picture = len(kept)
        if payload.read_from_picture:
            payload.notes.append(
                f"Page {number}: {payload.read_from_picture} glyphs come from a font whose Unicode map is wrong. "
                "Their words were read from the page picture. The rest of the page is the PDF's own text."
            )
        if payload.kept_in_picture:
            payload.notes.append(
                f"Page {number}: {payload.kept_in_picture} glyphs have no usable Unicode value and the picture "
                "reader could not read them. They stay in the page picture."
            )
    # Letters drawn as shapes are filled paths: read as cell fills and rules
    # they would cut every line into pieces.
    regions = [] if from_picture else _fill_regions(page)
    rules = [] if from_picture else _vertical_rules(page)
    payload.column_rules = list(rules)
    reads_left = _page_reads_left(atoms)
    written: list[list[_Atom]] = []
    for band in bands:
        # Words read from the picture leave a hole where the reader missed
        # one; only a column gutter, far wider, cuts such a line.
        for part in _split_band(band, regions, rules, 5.0 if from_picture else 2.0):
            try:
                line = _build_line(part, reads_left)
            except Exception:
                line = None
            if line is None:
                continue
            payload.lines.append(line)
            written.append(part)
    payload.lines.sort(key=lambda line: (round(line.baseline, 1), line.x0))
    if from_picture:
        _picture_page_furniture(payload, float(page.rect.height))
        _sizes_from_widths(payload.lines)
    other = _other_text(page)
    if other:
        payload.notes.append(
            f"Page {number}: text set at an angle or drawn as an outline stays in the page picture."
        )
    if not background:
        # The flowing layout writes no page picture; skip drawing one.
        _strip_page_labels(payload)
        payload.links = _read_links(page)
        payload.missing_fonts = _missing_fonts(page)
        _enrich_tables_and_charts(page, payload)
        return payload
    try:
        if other or kept:
            # Some text stays in the picture, so only the written glyphs are removed.
            groups = [[atom for atom in part if not isinstance(atom, _PictureWord)] for part in written]
            digital = [atom for group in groups for atom in group if atom.text.strip()]
            if replaced:
                groups.append(replaced)
            total = _visible_glyphs(page)
            payload.background, left = render_background(page, _removal_boxes(groups, kept), everything=False)
            # Removing clearly fewer glyphs than were written would print text twice.
            if total - left < (len(digital) + len(replaced)) * 0.9:
                return picture_page(
                    page,
                    f"Page {number}: the text could not be separated from the graphics. "
                    "The page is kept as a picture.",
                )
            payload.kept_in_picture = left
        else:
            payload.background, left = render_background(page, None, everything=True)
            if left:
                # The text could not be taken out of the page. Writing it again
                # as text would print it twice, so the page stays one picture.
                return picture_page(
                    page,
                    f"Page {number}: the text could not be separated from the graphics. "
                    "The page is kept as a picture.",
                )
    except Exception as exc:
        return picture_page(
            page, f"Page {number}: the background could not be drawn ({exc}). The page is kept as a picture."
        )
    _strip_page_labels(payload)
    payload.links = _read_links(page)
    payload.missing_fonts = _missing_fonts(page)
    _enrich_tables_and_charts(page, payload)
    return payload


_PAGE_LABEL = re.compile(
    r"^\s*[\d٠-٩]{1,4}\s*[|\u066d]?\s*p\s*a\s*g\s*e\s*$"
    r"|^\s*(?:page|صفحة|الصفحة)\s*[\d٠-٩]{1,4}\s*$"
    r"|^\s*[\d٠-٩]{1,4}\s*(?:page|صفحة|الصفحة)\s*$",
    re.IGNORECASE,
)


def _strip_page_labels(payload: FaithfulPage) -> None:
    """
    Lift body lines that are only a page number / 'N | Page' into chrome.

    Those lines become a live PAGE field in the Word footer/header so every
    page uses the same format and direction.
    """
    kept: list[Line] = []
    for line in payload.lines:
        compact = re.sub(r"\s+", " ", (line.text or "").strip())
        spaced = re.sub(r"(?<=[A-Za-z])\s+(?=[A-Za-z])", "", compact)
        # A label at the foot of the sheet moves into the Word footer, one at
        # the head into the Word header. A label in the body stays text.
        at_foot = float(line.y0) >= float(payload.height) * 0.8
        at_head = float(line.y1) <= float(payload.height) * 0.2
        if (at_foot or at_head) and (_PAGE_LABEL.match(compact) or _PAGE_LABEL.match(spaced)):
            digits = re.search(r"[\d٠-٩]{1,4}", compact)
            if digits and payload.page_label_number is None:
                payload.page_label_number = int(unicodedata.normalize("NFKC", digits.group(0)).translate(
                    str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")))
            if not payload.page_label_place:
                payload.page_label_place = "header" if at_head else "footer"
            if not payload.page_label_format:
                if "|" in compact or "|" in spaced:
                    payload.page_label_format = "{PAGE} | Page"
                elif re.search(r"page|صفحة", compact, re.I):
                    payload.page_label_format = (
                        "{PAGE} Page" if re.match(r"^\s*[\d٠-٩]", compact) else "Page {PAGE}"
                    )
                else:
                    payload.page_label_format = "{PAGE}"
                mid = payload.width / 2.0
                payload.page_label_side = (
                    "left" if line.x0 < mid * 0.7 else "right" if line.x1 > mid * 1.3 else "center"
                )
            continue
        kept.append(line)
    payload.lines = kept


def _read_links(page: fitz.Page) -> list[FaithfulLink]:
    links: list[FaithfulLink] = []
    try:
        raw = page.get_links() or []
    except Exception:
        return links
    for item in raw:
        box = item.get("from")
        if box is None:
            continue
        bbox = (float(box.x0), float(box.y0), float(box.x1), float(box.y1))
        uri = (item.get("uri") or "").strip()
        if uri.lower().startswith("mailto:"):
            links.append(FaithfulLink(bbox, "mailto", uri))
        elif uri:
            links.append(FaithfulLink(bbox, "uri", uri))
        elif item.get("page") is not None:
            links.append(FaithfulLink(bbox, "goto", str(int(item["page"]) + 1)))
    return links


def _missing_fonts(page: fitz.Page) -> list[str]:
    missing: list[str] = []
    seen: set[str] = set()
    try:
        fonts = page.get_fonts(full=True) or []
    except Exception:
        return missing
    for entry in fonts:
        name = ""
        if len(entry) >= 4:
            name = str(entry[3] or (entry[4] if len(entry) > 4 else "") or "")
        if not name or name in seen:
            continue
        seen.add(name)
        choice = font_catalog.choose_font(name, 0)
        if not choice.exact and choice.source in {
            "name only",
            "sans class",
            "serif class",
            "mono class",
        }:
            missing.append(name)
    return missing


def _enrich_tables_and_charts(page: fitz.Page, payload: FaithfulPage) -> None:
    """Rebuild tables, keep charts as one picture, and drop stray chart labels."""
    from core import table_rebuild

    try:
        regions = table_rebuild.classify_regions(page)
    except Exception as exc:
        payload.notes.append(f"Table/chart classifier failed: {exc}")
        return
    charts = table_rebuild.chart_boxes(regions)
    if charts:
        before = len(payload.lines)
        payload.lines = table_rebuild.drop_lines_in_boxes(payload.lines, charts)
        dropped = before - len(payload.lines)
        if dropped:
            payload.notes.append(
                f"Dropped {dropped} text label(s) that belong inside a chart/figure region."
            )
        for region in regions:
            if region.kind == "chart":
                payload.notes.append(
                    f"Chart/figure at {tuple(round(v, 1) for v in region.bbox)}: {region.reason}"
                )
    try:
        tables = table_rebuild.rebuild_tables(page, regions)
    except Exception as exc:
        payload.notes.append(f"Table rebuild failed: {exc}")
        return
    if not tables:
        return
    payload.tables = tables
    # Paint out regions we will cover: usable native grids, and unusable grids
    # that carry a fallback picture (so text is not drawn twice).
    boxes = []
    for table in tables:
        if not table.rows:
            continue
        if getattr(table, "usable", True):
            boxes.append(table.bbox)
        elif getattr(table, "fallback_png", b""):
            boxes.append(table.bbox)
    if boxes:
        payload.lines = table_rebuild.drop_lines_in_boxes(payload.lines, boxes)
        zoom = _zoom(page)
        payload.background = table_rebuild.paint_out_regions(
            payload.background, payload.width, payload.height, boxes, zoom
        )
    for table in tables:
        payload.notes.extend(table.notes)
        for row, col, conf in table.low_confidence_cells:
            payload.notes.append(f"Low-confidence table cell r{row + 1}c{col + 1} (score {conf:.0f})")


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------


def _emu(points: float) -> int:
    """
    Convert points to EMUs (English Metric Units) at full precision.

    1 pt = 12700 EMU. We only round to the nearest EMU integer — never to a
    whole point — so placed boxes keep sub-point accuracy.
    """
    return int(round(float(points) * 12700.0))


def _arabic_family() -> str:
    """The family for Arabic when the PDF's own font is not installed."""
    from core.docx_converter import _ARABIC_FONT

    for name in (_ARABIC_FONT, "Arial", "Times New Roman", "Tahoma"):
        found = font_catalog.installed_name(name)
        if found:
            return found
    return _ARABIC_FONT


_RESOLVED: dict[tuple, tuple] = {}


def _resolve(chunk: Chunk) -> tuple:
    """Latin family, complex-script family, bold, italic, and the two faces."""
    from core.docx_converter import _FONT

    fixed = (_FONT.get() or "").strip()
    key = (chunk.font, chunk.flags, chunk.bold, chunk.italic, fixed)
    if key in _RESOLVED:
        return _RESOLVED[key]
    flags = chunk.flags | (16 if chunk.bold else 0) | (2 if chunk.italic else 0)
    if fixed:
        # One family for the whole file; weight and slant still come from the PDF.
        # The faithful layout measures every line in the face it writes; when
        # the chosen family is not on this machine it cannot, and a stand-in
        # face may lack Arabic, so the PDF's own family is used instead.
        bold = bool(chunk.bold or flags & 16)
        italic = bool(chunk.italic or flags & 2)
        face = font_catalog.find_face(fixed, bold, italic)
        if face is not None:
            _RESOLVED[key] = (fixed, fixed, bold, italic, face, face)
            return _RESOLVED[key]
    choice = font_catalog.choose_font(chunk.font, flags)
    latin_face = font_catalog.find_face(choice.family, choice.bold, choice.italic)
    complex_family = choice.family
    if not (choice.exact and latin_face is not None and font_catalog.covers(latin_face, "\u0628")):
        complex_family = _arabic_family()
    complex_face = font_catalog.find_face(complex_family, choice.bold, choice.italic)
    _RESOLVED[key] = (choice.family, complex_family, choice.bold, choice.italic, latin_face, complex_face)
    return _RESOLVED[key]


def _joined(text: str) -> bool:
    """True when the letters join, so their widths cannot be added up one by one."""
    return any(_is_arabic(char) for char in text)


def _natural_width(chunk: Chunk, size: float) -> float | None:
    """Width Word gives this chunk before any correction, in points."""
    _latin, _complex, _bold, _italic, latin_face, complex_face = _resolve(chunk)
    face = complex_face if chunk.rtl else latin_face
    if face is None:
        return None
    if _joined(chunk.text):
        value = font_catalog.shaped_advance(face, chunk.text)
    else:
        if any(not char.isspace() and not font_catalog.covers(face, char) for char in chunk.text):
            return None
        value = font_catalog.text_advance(face, chunk.text)
    return None if value is None else value * size


def _word_size(size: float, scale: float) -> float:
    """Word stores sizes in half points."""
    return max(1.0, round(size * scale * 2.0) / 2.0)


def _measured_line_width(line: Line, scale: float) -> float | None:
    """Sum of natural chunk widths after fitting scale, when fonts are known."""
    total = 0.0
    known = 0
    for chunk in line.chunks:
        size = _word_size(chunk.size, scale)
        natural = _natural_width(chunk, size)
        if natural is None:
            if chunk.width is not None:
                total += chunk.width * scale
                known += 1
            continue
        total += natural
        known += 1
    if known == 0:
        return None
    return total


def _fit(line: Line, scale: float) -> list[tuple[int, int]]:
    """
    Character spacing (twentieths of a point) and width scale (percent) per chunk.

    Each word and each space is brought to the width it has on the PDF line.
    Words in a joined script are scaled, because spacing would break the
    joins. Other words and all spaces get character spacing. What rounding
    leaves over is carried into the next chunk, so the error does not add up
    along the line.
    """
    sizes = [_word_size(chunk.size, scale) for chunk in line.chunks]
    naturals = [_natural_width(chunk, size) for chunk, size in zip(line.chunks, sizes)]
    targets = [None if chunk.width is None else chunk.width * scale for chunk in line.chunks]
    if any(target is None for target in targets):
        # Without glyph positions, words keep their natural width and the
        # spaces share what is left of the line.
        for index, chunk in enumerate(line.chunks):
            if targets[index] is None and not chunk.space:
                targets[index] = naturals[index]
        free = [index for index, chunk in enumerate(line.chunks) if chunk.space and targets[index] is None]
        known = [target for index, target in enumerate(targets) if index not in free]
        if free and all(target is not None for target in known):
            share = ((line.x1 - line.x0) * scale - sum(known)) / len(free)
            for index in free:
                targets[index] = max(sizes[index] * 0.1, share)
        else:
            for index in free:
                targets[index] = naturals[index]
    fitted: list[tuple[int, int]] = []
    drawn = 0.0
    wanted = 0.0
    for index, chunk in enumerate(line.chunks):
        natural = naturals[index]
        target = targets[index]
        if natural is None or target is None:
            step = target if target is not None else (natural or 0.0)
            drawn += step
            wanted += step
            fitted.append((0, 100))
            continue
        wanted += target
        room = max(wanted - drawn, 0.0)
        count = max(1, len(chunk.text))
        if chunk.space:
            twips = int(round((room - natural) / count * 20.0))
            twips = max(-int(natural / count * 0.85 * 20.0), min(twips, 31000))
            drawn += natural + count * twips / 20.0
            fitted.append((twips, 100))
        elif _joined(chunk.text):
            ratio = room / natural if natural > 0 else 1.0
            # Do not crush Arabic below ~88%: Word often ignores deep w:w on
            # complex script, and the glyphs then spill past the column border.
            percent = 100 if abs(ratio - 1.0) <= 0.015 else int(round(min(2.0, max(0.88, ratio)) * 100))
            drawn += natural * percent / 100.0
            fitted.append((0, percent))
        else:
            twips = int(round((room - natural) / count * 20.0))
            limit = int(sizes[index] * 0.25 * 20.0)
            twips = max(-limit, min(twips, limit * 2))
            drawn += natural + count * twips / 20.0
            fitted.append((twips, 100))
    return fitted


_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def _run_xml(chunk: Chunk, text: str, size: float, spacing: int, percent: int) -> str:
    latin, complex_family, bold, italic, _latin_face, _complex_face = _resolve(chunk)
    half = int(round(size * 2))
    parts = [
        "<w:r><w:rPr>",
        f'<w:rFonts w:ascii="{escape(latin)}" w:hAnsi="{escape(latin)}" '
        f'w:cs="{escape(complex_family)}" w:eastAsia="{escape(latin)}"/>',
    ]
    if bold:
        parts.append("<w:b/><w:bCs/>")
    else:
        parts.append('<w:b w:val="0"/><w:bCs w:val="0"/>')
    if italic:
        parts.append("<w:i/><w:iCs/>")
    parts.append("<w:noProof/>")
    if len(chunk.color) == 6:
        parts.append(f'<w:color w:val="{chunk.color.upper()}"/>')
    if spacing:
        parts.append(f'<w:spacing w:val="{spacing}"/>')
    if percent != 100:
        parts.append(f'<w:w w:val="{percent}"/>')
    parts.append('<w:kern w:val="0"/>')
    if chunk.rise:
        parts.append(f'<w:position w:val="{int(round(chunk.rise * 2))}"/>')
    parts.append(f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>')
    if chunk.rtl:
        parts.append('<w:rtl/><w:lang w:val="ar-SA" w:bidi="ar-SA"/>')
    else:
        parts.append('<w:lang w:val="en-US"/>')
    cleaned = _XML_BAD.sub("", text)
    parts.append(f'</w:rPr><w:t xml:space="preserve">{escape(cleaned)}</w:t></w:r>')
    return "".join(parts)


def _line_height(line: Line, scale: float) -> float:
    """
    Height of the exact line, in points.

    The baseline sits four fifths down an exact line. The line is made tall
    enough for the tallest letter above the baseline and the deepest below
    it in the faces Word will use, so no glyph is clipped.
    """
    above = 0.0
    below = 0.0
    for chunk in line.chunks:
        if chunk.space:
            continue
        size = _word_size(chunk.size, scale)
        _latin, _complex, _bold, _italic, latin_face, complex_face = _resolve(chunk)
        face = complex_face if chunk.rtl else latin_face
        if face is not None:
            ascent, descent = face.ascent, face.descent
        elif chunk.rtl:
            ascent, descent = 1.1, 0.55
        else:
            ascent, descent = 0.95, 0.3
        rise = chunk.rise * scale
        above = max(above, ascent * size + rise)
        below = max(below, descent * size - rise)
    if above <= 0 and below <= 0:
        above = line.size * scale
        below = line.size * scale * 0.3
    return max(above / _BASELINE_SHARE, below / (1.0 - _BASELINE_SHARE), 4.0) * 1.03


def _column_bounds(
    line: Line,
    rules: list[tuple[float, float, float]],
    page_w: float,
) -> tuple[float, float]:
    """
    Nearest upright rules left and right of this line's midpoint.

    Used to keep a text box inside its table cell / column frame. Falls back
    to the page edges when no rule crosses the line's baseline.
    """
    mid_x = (line.x0 + line.x1) / 2.0
    mid_y = (line.y0 + line.y1) / 2.0
    left = 0.0
    right = page_w
    for x, top, bottom in rules:
        if mid_y < top - 2.0 or mid_y > bottom + 2.0:
            continue
        if x < mid_x - 0.5:
            left = max(left, x)
        elif x > mid_x + 0.5:
            right = min(right, x)
    return left, right


def _line_box(
    line: Line,
    scale: float,
    off_x: float,
    off_y: float,
    page_w: float,
    page_h: float,
    rules: list[tuple[float, float, float]] | None = None,
) -> tuple[float, float, float, float, list[str]]:
    """
    Page-anchored left/top/width/height for one text box, with clamps.

    RTL boxes pin to the PDF line's RIGHT edge and keep the PDF line width.
    Expanding past a column rule or the page is refused: the box is clamped
    and a flag is returned so the caller can report it.
    """
    height = math.ceil(_line_height(line, scale) * 20.0) / 20.0
    top = line.baseline * scale + off_y - _BASELINE_SHARE * height
    pdf_left = line.x0 * scale + off_x
    pdf_right = line.x1 * scale + off_x
    pdf_width = max(1.0, pdf_right - pdf_left)
    measured = _measured_line_width(line, scale)
    # A small slack for metric mismatch — never a second "pad past the rule".
    slack = max(1.0, line.size * scale * 0.12)
    needed = measured if measured is not None else pdf_width
    flags: list[str] = []

    col_left = off_x
    col_right = page_w * scale + off_x
    if rules:
        raw_l, raw_r = _column_bounds(line, rules, page_w)
        col_left = max(col_left, raw_l * scale + off_x)
        col_right = min(col_right, raw_r * scale + off_x)

    page_left = off_x
    page_right = page_w * scale + off_x
    page_top = off_y
    page_bottom = page_h * scale + off_y

    if line.rtl:
        # Arabic starts at the right edge of the PDF line. Do not push that
        # edge further right (that is what spilled past column borders).
        right = min(pdf_right, col_right - 0.5)
        # Prefer the original line width; open leftward only within the column
        # when Word's face is a little wider — never past the cell.
        want = min(max(pdf_width, needed * 1.01), max(pdf_width, right - col_left - 0.5))
        width = max(1.0, min(want + slack * 0.25, right - col_left - 0.25))
        left = right - width
        if left < col_left + 0.25:
            left = col_left + 0.25
            width = max(1.0, right - left)
            flags.append("clamped_to_cell")
        if abs(right - pdf_right) > 0.75 or abs(width - pdf_width) > slack + 1.0:
            if "clamped_to_cell" not in flags:
                flags.append("width_pinned")
    else:
        left = max(pdf_left, col_left + 0.25)
        want = min(max(pdf_width, needed * 1.01 + slack * 0.25), max(1.0, col_right - left - 0.25))
        width = max(1.0, want)
        if left + width > col_right - 0.25:
            width = max(1.0, col_right - 0.25 - left)
            flags.append("clamped_to_cell")

    # Page clamp — never place a box outside the sheet.
    if left < page_left:
        width -= page_left - left
        left = page_left
        flags.append("clamped_to_page")
    if left + width > page_right:
        width = max(1.0, page_right - left)
        flags.append("clamped_to_page")
    if top < page_top:
        top = page_top
        flags.append("clamped_to_page")
    if top + height > page_bottom:
        height = max(1.0, page_bottom - top)
        flags.append("clamped_to_page")

    # Detect a word/line that would still sit past its cell after clamping.
    if rules and (left < col_left - 0.5 or left + width > col_right + 0.5):
        flags.append("outside_cell")
    return left, top, max(1.0, width), max(1.0, height), flags


def _line_xml(
    line: Line,
    scale: float,
    off_x: float,
    off_y: float,
    ident: int,
    order: int,
    page_w: float = 0.0,
    page_h: float = 0.0,
    rules: list[tuple[float, float, float]] | None = None,
) -> tuple[str, list[str]]:
    """
    One visual line → one text box; style changes are runs inside that box.

    Returns the DrawingML fragment and any clamp/placement flags.
    """
    fitted = _fit(line, scale)
    runs: list[str] = []
    pending: tuple | None = None
    for chunk, (spacing, percent) in zip(line.chunks, fitted):
        size = _word_size(chunk.size, scale)
        key = (
            chunk.font,
            chunk.flags,
            chunk.bold,
            chunk.italic,
            chunk.color,
            chunk.rtl,
            chunk.rise,
            size,
            spacing,
            percent,
        )
        if pending is not None and pending[0] == key:
            pending = (key, pending[1], pending[2] + chunk.text, size, spacing, percent)
            continue
        if pending is not None:
            runs.append(_run_xml(pending[1], pending[2], pending[3], pending[4], pending[5]))
        pending = (key, chunk, chunk.text, size, spacing, percent)
    if pending is not None:
        runs.append(_run_xml(pending[1], pending[2], pending[3], pending[4], pending[5]))

    if page_w <= 0:
        page_w = max(line.x1 + 36.0, 595.0)
    if page_h <= 0:
        page_h = max(line.y1 + 36.0, 842.0)
    left, top, width, height, flags = _line_box(line, scale, off_x, off_y, page_w, page_h, rules)
    wrap = "none"
    align = "right" if line.rtl else "left"
    twips = int(round(height * 20))
    paragraph = (
        "<w:p><w:pPr>"
        '<w:widowControl w:val="0"/>'
        f'<w:bidi w:val="{1 if line.rtl else 0}"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="{twips}" w:lineRule="exact"/>'
        '<w:ind w:left="0" w:right="0" w:firstLine="0"/>'
        f'<w:jc w:val="{align}"/>'
        f"</w:pPr>{''.join(runs)}</w:p>"
    )
    xml = (
        f"<w:r {_NAMESPACES}><w:rPr><w:noProof/></w:rPr><w:drawing>"
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="{1000 + order}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{_emu(left)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{_emu(top)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{_emu(width)}" cy="{_emu(height)}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{ident}" name="Text {ident}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="{_SHAPE_URI}">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{_emu(width)}" cy="{_emu(height)}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
        f"<wps:txbx><w:txbxContent>{paragraph}</w:txbxContent></wps:txbx>"
        f'<wps:bodyPr rot="0" vert="horz" wrap="{wrap}" lIns="0" tIns="0" rIns="0" bIns="0" '
        'anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )
    return xml, flags


def write_page(
    document,
    payload: FaithfulPage,
    scale: float,
    off_x: float,
    off_y: float,
    counts: dict,
) -> None:
    """
    Put one page into the Word document.

    The background picture is anchored to the page at the top-left corner,
    behind the text. The text boxes follow it, each anchored to the page at
    its own PDF position. They all hang on one paragraph that has no height.
    """
    from core.html_layout import _float_drawing, _next_id, _pin_host

    host = document.add_paragraph()
    _pin_host(host)
    if payload.background:
        run = host.add_run()
        run.add_picture(BytesIO(payload.background), width=Emu(_emu(payload.width * scale)))
        _float_drawing(run, off_x, off_y, payload.width * scale, payload.height * scale, behind=True)
        for anchor in run._r.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor"):
            # Keep the plate behind text and locked so a click lands on editable
            # text boxes / tables instead of selecting the whole page as a picture.
            anchor.set("relativeHeight", "1")
            anchor.set("behindDoc", "1")
            anchor.set("locked", "1")
        for properties in run._r.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"):
            properties.set("name", "Page background")
            properties.set("descr", "Non-editable page graphics")
    rules = list(getattr(payload, "column_rules", None) or [])
    clamped = 0
    outside = 0
    for order, line in enumerate(payload.lines):
        ident = _next_id()
        flags: list[str] = []
        try:
            xml, flags = _line_xml(
                line, scale, off_x, off_y, ident, order, payload.width, payload.height, rules
            )
            element = parse_xml(xml)
        except Exception:
            # The line is still written, whole and in one style, so no text is lost.
            host_chunk = next((chunk for chunk in line.chunks if not chunk.space), None)
            plain = Line(line.x0, line.y0, line.x1, line.y1, line.baseline, line.size, line.rtl, line.text)
            plain.chunks = [
                Chunk(
                    text=line.text,
                    font=host_chunk.font if host_chunk else "",
                    flags=host_chunk.flags if host_chunk else 0,
                    size=line.size,
                    bold=bool(host_chunk and host_chunk.bold),
                    italic=bool(host_chunk and host_chunk.italic),
                    color=host_chunk.color if host_chunk else "",
                    rtl=line.rtl,
                    space=False,
                    width=None,
                )
            ]
            try:
                xml, flags = _line_xml(
                    plain, scale, off_x, off_y, ident, order, payload.width, payload.height, rules
                )
                element = parse_xml(xml)
            except Exception as exc:
                counts.setdefault("skipped", []).append(f"A line could not be written: {line.text[:40]} ({exc})")
                continue
        if any(flag.startswith("clamped") for flag in flags):
            clamped += 1
        if "outside_cell" in flags:
            outside += 1
            sample = (line.text or "").strip()[:48]
            counts.setdefault("skipped", []).append(
                f"Text box would leave its column/cell; clamped: {sample}"
            )
        host._p.append(element)
        counts["paragraphs"] = counts.get("paragraphs", 0) + 1
    if clamped:
        payload.notes.append(f"{clamped} text box(es) were clamped to the page or column edge.")
    if outside:
        payload.notes.append(f"{outside} text box(es) would have extended past a column/cell border.")
    for link in payload.links or []:
        try:
            _write_link(host, link, scale, off_x, off_y)
        except Exception as exc:
            counts.setdefault("skipped", []).append(f"A hyperlink could not be written: {exc}")
    _pin_host(host)
    for table in payload.tables or []:
        usable = bool(getattr(table, "rows", None)) and getattr(table, "usable", True)
        if not usable:
            # Weak rebuild: float the original region picture (even when a page
            # background exists — the hole was painted out above).
            if getattr(table, "fallback_png", b""):
                _write_fallback_picture(document, table, scale, off_x, off_y, counts)
            counts.setdefault("table_reports", []).append(
                {
                    "usable": False,
                    "rows": len(getattr(table, "rows", None) or []),
                    "low_confidence": len(getattr(table, "low_confidence_cells", None) or []),
                    "notes": list(getattr(table, "notes", None) or []),
                }
            )
            continue
        try:
            _write_floating_table(document, table, scale, off_x, off_y, counts)
        except Exception as exc:
            counts.setdefault("skipped", []).append(f"A table could not be written: {exc}")
            if getattr(table, "fallback_png", b""):
                _write_fallback_picture(document, table, scale, off_x, off_y, counts)


def _write_link(host, link: FaithfulLink, scale: float, off_x: float, off_y: float) -> None:
    """
    Preserve a PDF hyperlink as a Word external hyperlink on a tiny text run
    anchored at the link's page position. Internal page jumps are recorded as
    a tooltip pointing at the target page number.
    """
    from core.html_layout import _next_id
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    left = link.bbox[0] * scale + off_x
    top = link.bbox[1] * scale + off_y
    width = max(4.0, (link.bbox[2] - link.bbox[0]) * scale)
    height = max(4.0, (link.bbox[3] - link.bbox[1]) * scale)
    ident = _next_id()
    if link.kind == "goto":
        label = f"Page {link.target}"
        target = ""
    else:
        label = link.target
        target = link.target
    tip = escape(label[:120])
    # The PDF already prints the link's words under this box. The box only
    # needs to be clickable, so it holds blank space as wide as the link,
    # not a second, visible copy of the address or of "Page N".
    filler = "\u00a0" * max(1, int(width / 2.5))
    # Build a text-box run, then attach a real relationship hyperlink when URI.
    xml = (
        f"<w:r {_NAMESPACES}><w:rPr><w:noProof/></w:rPr><w:drawing>"
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="{200 + ident}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{_emu(left)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{_emu(top)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{_emu(width)}" cy="{_emu(height)}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{ident}" name="Link {ident}" descr="{tip}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="{_SHAPE_URI}">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{_emu(width)}" cy="{_emu(height)}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
        "<wps:txbx><w:txbxContent><w:p><w:pPr>"
        '<w:spacing w:before="0" w:after="0" w:line="40" w:lineRule="exact"/>'
        f"</w:pPr><w:r><w:rPr><w:noProof/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/>"
        f"</w:rPr><w:t xml:space=\"preserve\">{filler}</w:t></w:r></w:p></w:txbxContent></wps:txbx>"
        '<wps:bodyPr rot="0" vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" '
        'anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )
    element = parse_xml(xml)
    if target and (target.lower().startswith("http") or target.lower().startswith("mailto:")):
        # Wrap the run in a hyperlink that points at the relationship target.
        try:
            part = host.part
            rel = part.relate_to(
                target,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                is_external=True,
            )
            wrapper = OxmlElement("w:hyperlink")
            wrapper.set(qn("r:id"), rel)
            wrapper.set(qn("w:tooltip"), label[:120])
            wrapper.append(element)
            host._p.append(wrapper)
            return
        except Exception:
            pass
    host._p.append(element)


def _write_fallback_picture(document, table, scale: float, off_x: float, off_y: float, counts: dict) -> None:
    from core.html_layout import _float_drawing, _pin_host

    host = document.add_paragraph()
    _pin_host(host)
    run = host.add_run()
    width = (table.bbox[2] - table.bbox[0]) * scale
    height = (table.bbox[3] - table.bbox[1]) * scale
    run.add_picture(BytesIO(table.fallback_png), width=Emu(_emu(width)))
    _float_drawing(run, table.bbox[0] * scale + off_x, table.bbox[1] * scale + off_y, width, height, behind=False)
    counts["tables"] = counts.get("tables", 0) + 1


_TBL_PR_ORDER = (
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
    "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
    "tblLook", "tblCaption", "tblDescription",
)


def _order_tbl_pr(props) -> None:
    """Put w:tblPr children in schema order; Word may refuse a file otherwise."""
    rank = {name: index for index, name in enumerate(_TBL_PR_ORDER)}
    children = list(props)
    children.sort(key=lambda node: rank.get(node.tag.rsplit("}", 1)[-1], len(rank)))
    for node in children:
        props.remove(node)
        props.append(node)


def _write_floating_table(document, table, scale: float, off_x: float, off_y: float, counts: dict) -> None:
    """
    Native w:tbl anchored to the page at the PDF table position (faithful mode).
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import RGBColor

    rows = table.rows
    cols = max(len(row) for row in rows)
    rows = [list(row) + [""] * (cols - len(row)) for row in rows]
    word_table = document.add_table(rows=len(rows), cols=cols)
    word_table.autofit = False
    word_table.allow_autofit = False
    try:
        word_table.style = "Table Grid"
    except KeyError:
        pass

    widths = list(table.widths) if table.widths and len(table.widths) == cols else [(table.bbox[2] - table.bbox[0]) / cols] * cols
    widths = [max(12.0, w * scale) for w in widths]
    column_dxa = [max(200, int(w * 20)) for w in widths]
    total_dxa = sum(column_dxa)

    props = word_table._tbl.tblPr
    # Floating position relative to the page.
    tblp = OxmlElement("w:tblpPr")
    tblp.set(qn("w:horzAnchor"), "page")
    tblp.set(qn("w:vertAnchor"), "page")
    tblp.set(qn("w:tblpX"), str(int(round((table.bbox[0] * scale + off_x) * 20))))
    tblp.set(qn("w:tblpY"), str(int(round((table.bbox[1] * scale + off_y) * 20))))
    # No tblpXSpec/tblpYSpec: a spec overrides tblpX/tblpY and sends the
    # table to the page edge.
    existing_p = props.find(qn("w:tblpPr"))
    if existing_p is not None:
        props.remove(existing_p)
    props.append(tblp)
    overlap = OxmlElement("w:tblOverlap")
    overlap.set(qn("w:val"), "never")
    old_overlap = props.find(qn("w:tblOverlap"))
    if old_overlap is not None:
        props.remove(old_overlap)
    props.append(overlap)

    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:w"), str(total_dxa))
    tbl_w.set(qn("w:type"), "dxa")
    old_w = props.find(qn("w:tblW"))
    if old_w is not None:
        props.remove(old_w)
    props.append(tbl_w)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    old_layout = props.find(qn("w:tblLayout"))
    if old_layout is not None:
        props.remove(old_layout)
    props.append(layout)
    # Cells are stored in printed order, left to right. A right-to-left
    # table (bidiVisual) would mirror them: the Arabic label column of a
    # definitions table would land on the left. Say left to right outright.
    bidi = OxmlElement("w:bidiVisual")
    bidi.set(qn("w:val"), "0")
    props.append(bidi)

    _order_tbl_pr(props)

    grid = word_table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for node, dxa in zip(grid.findall(qn("w:gridCol")), column_dxa):
            node.set(qn("w:w"), str(dxa))

    heights = list(table.heights) if table.heights and len(table.heights) >= len(rows) else []
    for ri, row in enumerate(rows):
        tr = word_table.rows[ri]._tr
        tr_pr = tr.get_or_add_trPr()
        if getattr(table, "header_repeat", False) and ri == 0:
            header = OxmlElement("w:tblHeader")
            tr_pr.append(header)
        if ri < len(heights):
            h = OxmlElement("w:trHeight")
            h.set(qn("w:val"), str(max(80, int(heights[ri] * scale * 20))))
            h.set(qn("w:hRule"), "atLeast")
            old_h = tr_pr.find(qn("w:trHeight"))
            if old_h is not None:
                tr_pr.remove(old_h)
            tr_pr.append(h)
        for ci, text in enumerate(row):
            cell = word_table.cell(ri, ci)
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = OxmlElement("w:tcW")
            tc_w.set(qn("w:w"), str(column_dxa[ci]))
            tc_w.set(qn("w:type"), "dxa")
            old_tc_w = tc_pr.find(qn("w:tcW"))
            if old_tc_w is not None:
                tc_pr.remove(old_tc_w)
            tc_pr.append(tc_w)
            fill = ""
            if ri < len(table.fills) and ci < len(table.fills[ri]):
                fill = table.fills[ri][ci] or ""
            if fill:
                shade = OxmlElement("w:shd")
                shade.set(qn("w:val"), "clear")
                shade.set(qn("w:color"), "auto")
                shade.set(qn("w:fill"), fill.upper())
                old_shade = tc_pr.find(qn("w:shd"))
                if old_shade is not None:
                    tc_pr.remove(old_shade)
                tc_pr.append(shade)
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_before = 0
            paragraph.paragraph_format.space_after = 0
            conf = 100.0
            if ri < len(getattr(table, "confidences", []) or []) and ci < len(table.confidences[ri]):
                conf = table.confidences[ri][ci]
            header_white = False
            if fill:
                try:
                    r = int(fill[0:2], 16)
                    g = int(fill[2:4], 16)
                    b = int(fill[4:6], 16)
                except ValueError:
                    r = g = b = 255
                header_white = r * 0.3 + g * 0.59 + b * 0.11 < 150
            # Keep multi-line cells (stacked dates) as soft breaks, not one glued run.
            parts = (text or "").split("\n") or [""]
            from core.docx_converter import _arabic_count, _set_paragraph_ltr, _set_paragraph_rtl, _set_run_font

            arabic_cell = 0 < _arabic_count(text or "") >= len(re.findall(r"[A-Za-z]", text or ""))
            if text and arabic_cell:
                _set_paragraph_rtl(paragraph, text)
            elif text:
                _set_paragraph_ltr(paragraph)
            size = max(6.5, min(16.0, float(getattr(table, "text_size", 0) or 9.0) * scale))
            # A wider output font takes the narrower PDF font's room.
            from core.flow_layout import _width_factor

            size *= _width_factor(text or "", getattr(table, "text_font", "") or "")
            for index, part in enumerate(parts):
                if index:
                    paragraph.add_run().add_break()
                run = paragraph.add_run(part)
                _set_run_font(run, size, bool(header_white), part)
                if header_white:
                    run.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                if conf and conf < 55:
                    try:
                        from docx.enum.text import WD_COLOR_INDEX

                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    except Exception:
                        pass
    for ri, row_spans in enumerate(getattr(table, "spans", None) or []):
        if ri >= len(rows):
            break
        for first, last in row_spans:
            if last > first and last < cols:
                try:
                    word_table.cell(ri, first).merge(word_table.cell(ri, last))
                except Exception:
                    pass
    counts["tables"] = counts.get("tables", 0) + 1
