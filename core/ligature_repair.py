# -*- coding: utf-8 -*-
"""
Give lam-alef ligatures back their letters when the PDF's character map lost them.

Word draws لا، لأ، لإ and لآ as one ligature glyph. Some PDF writers
(Ghostscript among them) store a wrong Unicode value for that glyph, so a
reader gets a Greek or Canadian letter in the middle of an Arabic word:
"التعامϼت" or "التعامᕙت" instead of "التعاملات". A converter that trusts the
map either prints the wrong letter or drops the whole line as unreadable.

Two repairs, both independent of any one document:

1. Shifted values. The ligatures live at U+FEF5..U+FEFC. A writer that
   loses the high byte of the code stores them at U+03F5..U+03FC, the same
   low byte. That shift is exact, so the letter is recovered with certainty.

2. Any other non-Arabic letter wedged between Arabic letters of one word is
   a ligature whose map is wrong. Which of the four it is comes from the
   document itself: each spelling is tried, and the one that forms a word
   the document also prints correctly elsewhere wins. With no evidence the
   plain لا is used, because it is by far the most common.
"""

from __future__ import annotations

import re
import unicodedata
from collections import Counter

_ARABIC = re.compile(r"[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]")
_LIGATURES = {
    0xFEF5: "لآ", 0xFEF6: "لآ", 0xFEF7: "لأ", 0xFEF8: "لأ",
    0xFEF9: "لإ", 0xFEFA: "لإ", 0xFEFB: "لا", 0xFEFC: "لا",
}
_CANDIDATES = ("لا", "لأ", "لإ", "لآ")
# Letters a broken map has been seen to produce in place of a ligature.
_FOREIGN = re.compile(r"[Ͱ-Ͽ᐀-ᙿ-�]")


def truncated_ligature(char: str) -> str | None:
    """The ligature a ten-bit truncated code point stands for, if any."""
    code = ord(char)
    if 0x03F5 <= code <= 0x03FC:
        return _LIGATURES.get(0xFE00 | (code & 0xFF))
    return None


def _is_arabic(text: str) -> bool:
    return bool(text) and bool(_ARABIC.search(text))


def _in_arabic_word(atoms: list, index: int, gap: float = 2.5) -> bool:
    """True when an Arabic letter touches this glyph on either side, on its line."""
    atom = atoms[index]
    for step in (-1, 1):
        other_index = index + step
        if not 0 <= other_index < len(atoms):
            continue
        other = atoms[other_index]
        if not other.text.strip():
            continue
        if abs(other.origin_y - atom.origin_y) > max(1.5, atom.size * 0.3):
            continue
        near = other.x0 - gap <= atom.x1 and atom.x0 <= other.x1 + gap
        if near and _is_arabic(other.text):
            return True
    return False


def _word_spans(atoms: list) -> list[tuple[int, int]]:
    """Index ranges of runs of touching non-space glyphs on one baseline."""
    spans = []
    start = None
    for index, atom in enumerate(atoms):
        solid = bool(atom.text.strip())
        if start is not None:
            prev = atoms[index - 1]
            joined = (
                solid
                and abs(atom.origin_y - prev.origin_y) <= max(1.5, atom.size * 0.3)
                and atom.x0 - 2.5 <= prev.x1
                and prev.x0 <= atom.x1 + 2.5
            )
            if not joined:
                spans.append((start, index))
                start = None
        if solid and start is None:
            start = index
    if start is not None:
        spans.append((start, len(atoms)))
    return spans


def _key(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    return re.sub(r"[^ء-ي]", "", text)


def vocabulary(texts) -> Counter:
    """Arabic words a document prints with a sound map, in both orders."""
    words: Counter = Counter()
    for text in texts:
        text = "".join(truncated_ligature(ch) or ch for ch in text)
        for raw in text.split():
            if _FOREIGN.search(raw):
                continue
            word = _key(raw)
            if len(word) >= 3:
                words[word] += 1
                words[word[::-1]] += 1
    return words


def repair_atoms(atoms: list, known: Counter | None = None) -> int:
    """
    Rewrite broken ligature glyphs in place. Returns how many were repaired.

    `atoms` are glyphs in painting order with text, x0, x1, origin_y and
    size. `known` is the document's vocabulary (see vocabulary()).
    """
    repaired = 0
    pending: list[int] = []
    for index, atom in enumerate(atoms):
        if len(atom.text) != 1:
            continue
        exact = truncated_ligature(atom.text)
        if exact and _in_arabic_word(atoms, index):
            atom.text = exact
            repaired += 1
        elif _FOREIGN.match(atom.text) and _in_arabic_word(atoms, index):
            pending.append(index)
    if not pending:
        return repaired
    spans = _word_spans(atoms)
    for index in pending:
        span = next(((s, e) for s, e in spans if s <= index < e), (index, index + 1))
        best, best_score = _CANDIDATES[0], -1
        for candidate in _CANDIDATES:
            score = 0
            # The glyphs may be stored in visual order, so try the ligature
            # both ways round inside the word.
            for form in (candidate, candidate[::-1]):
                word = "".join(form if i == index else atoms[i].text for i in range(*span))
                score = max(score, (known or Counter()).get(_key(word), 0))
            if score > best_score:
                best, best_score = candidate, score
        atoms[index].text = best
        repaired += 1
    return repaired


_DOCUMENT_WORDS: dict[tuple, Counter] = {}


def document_vocabulary(page) -> Counter:
    """The vocabulary of the document a fitz page belongs to, read once."""
    document = page.parent
    key = (id(document), getattr(document, "name", ""), document.page_count)
    if key not in _DOCUMENT_WORDS:
        texts = []
        for other in document:
            try:
                texts.append(other.get_text("text"))
            except Exception:
                continue
        if len(_DOCUMENT_WORDS) > 16:
            _DOCUMENT_WORDS.clear()
        _DOCUMENT_WORDS[key] = vocabulary(texts)
    return _DOCUMENT_WORDS[key]
