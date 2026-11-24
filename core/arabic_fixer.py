# -*- coding: utf-8 -*-
"""Arabic letters for the Word engine.

Presentation forms are folded to ordinary Unicode. A lam-alef ligature stays
one unit. Word applies the paragraph direction later, so this module does not
reverse a line a second time.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

_BIDI_MARKS = (
    "\u200e",  # left-to-right mark
    "\u200f",  # right-to-left mark
    "\u202a",  # left-to-right embedding
    "\u202b",  # right-to-left embedding
    "\u202c",  # pop directional formatting
    "\u202d",  # left-to-right override
    "\u202e",  # right-to-left override
    "\u2066",  # left-to-right isolate
    "\u2067",  # right-to-left isolate
    "\u2068",  # first-strong isolate
    "\u2069",  # pop directional isolate
)


def _is_arabic(char: str) -> bool:
    """Arabic letters, including the presentation forms some PDFs still store."""
    code = ord(char)
    return (
        0x0600 <= code <= 0x06FF
        or 0x0750 <= code <= 0x077F
        or 0x08A0 <= code <= 0x08FF
        or 0xFB50 <= code <= 0xFDFF
        or 0xFE70 <= code <= 0xFEFF
    )


# Letters that mark Urdu or Sindhi. A real heh doachashmee belongs with them.
# A presentation form of that letter inside ordinary Arabic is a bad font map.
_URDU_BASE = frozenset(
    chr(code)
    for code in (
        0x0679,
        0x067E,
        0x0686,
        0x0688,
        0x0691,
        0x06A9,
        0x06AF,
        0x06BA,
        0x06BE,
        0x06C1,
        0x06C2,
        0x06C3,
        0x06D2,
        0x06D3,
    )
)


def _fold_doachashmee_forms(text: str) -> str:
    """
    Map a presentation form of heh doachashmee to ordinary heh.

    NFKC turns U+FBAB into U+06BE and then stops. That is the right letter
    in Urdu. In an Arabic run the font stored the ordinary heh under the
    doachashmee presentation form, which is visible because the rest of the
    run has no Urdu letters. The base letter U+06BE is left as it is.
    """
    if any(char in _URDU_BASE for char in text):
        return text
    folded = []
    for char in text:
        if char != "\u06be" and unicodedata.normalize("NFKC", char) == "\u06be":
            folded.append("\u0647")
        else:
            folded.append(char)
    return "".join(folded)


def plain_arabic(text: str) -> str:
    """
    Return the string Word should store.

    NFKC folds Arabic presentation forms (isolated, initial, medial, final)
    back to the standard letters. The character order is left unchanged.
    There is no reversing step: Word's own right-to-left layout is the only
    reorder, so the letters are not flipped twice.
    """
    if not text:
        return ""
    cleaned = _fold_doachashmee_forms(text)
    cleaned = unicodedata.normalize("NFKC", cleaned)
    for mark in _BIDI_MARKS:
        cleaned = cleaned.replace(mark, "")
    cleaned = cleaned.replace("\xa0", " ").replace("\u00ad", "")
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    return cleaned.strip()


def _arabic_count(text: str) -> int:
    return sum(1 for char in text if _is_arabic(char))


@dataclass
class _Atom:
    """One letter, or a lam-alef ligature that shares a single origin."""

    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    origin_x: float
    origin_y: float
    level: int
    size: float
    bold: bool
    color: str = ""
    # The PDF's own font name and flags for this glyph, and that font's space
    # width in points. Empty for words that came from the picture reader.
    font: str = ""
    italic: bool = False
    flags: int = 0
    space: float = 0.0


def _keep_visual_order(atoms: list[_Atom]) -> str:
    """Join the line from the left. Used when the PDF already stored logical text."""
    ordered = sorted(atoms, key=lambda atom: (atom.x0, atom.origin_x))
    return "".join(atom.text for atom in ordered)


def _rtl_from_visual(atoms: list[_Atom]) -> str:
    """
    Turn one visual line into logical Unicode.

    Glyphs are painted from the left. An Arabic run is reversed so the word
    starts with its first letter, and a lam-alef unit is reversed as a whole
    so لا stays لا. An English run painted on the same line, such as an email
    address, stays left to right. The runs are then read from the right, which
    is where an Arabic sentence starts.
    """
    ordered = sorted(atoms, key=lambda atom: (atom.x0, atom.origin_x))
    runs: list[tuple[bool, list[_Atom]]] = []
    for atom in ordered:
        odd = atom.level % 2 == 1
        if runs and runs[-1][0] == odd:
            runs[-1][1].append(atom)
        else:
            runs.append((odd, [atom]))
    pieces: list[tuple[bool, str]] = []
    for odd, group in runs:
        sequence = list(reversed(group)) if odd else group
        text = "".join(atom.text for atom in sequence)
        if not odd:
            text = _unmirror_latin(text)
        pieces.append((odd, text))
    groups: list[tuple[bool, list[str]]] = []
    for odd, text in pieces:
        if not odd and groups and groups[-1][0] is False:
            groups[-1][1].append(text)
        else:
            groups.append((odd, [text]))
    groups.reverse()
    return "".join("".join(parts) for _odd, parts in groups)


def _clone_atom(atom: _Atom) -> _Atom:
    return _Atom(
        atom.text,
        atom.x0,
        atom.y0,
        atom.x1,
        atom.y1,
        atom.origin_x,
        atom.origin_y,
        atom.level,
        atom.size,
        atom.bold,
        getattr(atom, "color", ""),
        getattr(atom, "font", ""),
        getattr(atom, "italic", False),
        getattr(atom, "flags", 0),
        getattr(atom, "space", 0.0),
    )


def _is_mark(text: str) -> bool:
    """A vowel or tanween drawn on top of a letter, not a letter of its own."""
    return bool(text) and all(unicodedata.category(char) == "Mn" for char in text)


# Parentheses in an Arabic line are often painted mirrored: ")" on the left
# of an English phrase and "(" on its right. Logical text opens on the left.
_CLOSING_TO_OPEN = {")": "(", "]": "[", "}": "{"}


def _drop_covered_spaces(atoms: list[_Atom]) -> list[_Atom]:
    """
    Drop a space that is painted on top of a letter.

    Those spaces are duplicates. Leaving them in splits للإنقاذ into
    للإنقا ذ and البحث into البح ث.
    """
    letters = [atom for atom in atoms if atom.text.strip()]
    kept = []
    for atom in atoms:
        if atom.text.strip():
            kept.append(atom)
            continue
        if atom.x1 <= atom.x0:
            continue
        covered = False
        for letter in letters:
            center = (atom.x0 + atom.x1) / 2.0
            if letter.x0 - 0.4 <= center <= letter.x1 + 0.4:
                covered = True
                break
        if not covered:
            kept.append(atom)
    return kept


def _bind_marks(atoms: list[_Atom]) -> list[_Atom]:
    """Put each vowel on the letter it is drawn on, after that letter."""
    marks = [atom for atom in atoms if _is_mark(atom.text)]
    others = [atom for atom in atoms if not _is_mark(atom.text)]
    bases = [atom for atom in others if atom.text.strip()]
    for mark in marks:
        if not bases:
            others.append(mark)
            continue
        anchor = mark.origin_x if mark.x1 <= mark.x0 else (mark.x0 + mark.x1) / 2.0
        covering = [base for base in bases if base.x0 - 1.0 <= anchor <= base.x1 + 1.0]
        host = min(
            covering or bases,
            key=lambda base: abs(((base.x0 + base.x1) / 2.0) - anchor),
        )
        host.text += mark.text
    return others


def _adopt_latin_brackets(atoms: list[_Atom]) -> list[_Atom]:
    """A closing bracket just left of an English word belongs to that word."""
    ordered = sorted(atoms, key=lambda atom: (atom.x0, atom.origin_x))
    for index, atom in enumerate(ordered):
        if atom.text not in _CLOSING_TO_OPEN:
            continue
        nxt = None
        for later in ordered[index + 1 :]:
            if later.text.strip():
                nxt = later
                break
        if nxt is None:
            continue
        latin = any(char.isascii() and char.isalnum() for char in nxt.text)
        if latin and nxt.level % 2 == 0:
            atom.level = nxt.level
    return ordered


def _unmirror_latin(text: str) -> str:
    """Turn a visually mirrored English phrase ")Alkal, 2021(" back into "(Alkal, 2021)"."""
    if len(text) >= 2 and text[0] in _CLOSING_TO_OPEN and text[-1] == _CLOSING_TO_OPEN[text[0]]:
        return text[-1] + text[1:-1] + text[0]
    return text


def _drop_phantom_alef(atoms: list[_Atom]) -> list[_Atom]:
    """
    Drop a zero-width alef sitting on a lam-alef ligature.

    The ligature already contains its alef. The extra alef has no width, and
    leaving it turns أولاً into أولاا and أهلاً into أهلاا.
    """
    hosts = [
        atom
        for atom in atoms
        if "ل" in atom.text and "ا" in atom.text and (atom.x1 - atom.x0) > 0.8
    ]
    if not hosts:
        return atoms
    kept = []
    for atom in atoms:
        width = atom.x1 - atom.x0
        if (
            atom.text == "ا"
            and width <= 0.8
            and any(host.x0 - 1.5 <= atom.origin_x <= host.x1 + 1.5 for host in hosts)
        ):
            continue
        kept.append(atom)
    return kept


# A comma or hyphen between two numbers is neutral. An odd bidi level on that
# mark splits 1,500 into 500,1 and 2021-2022 into 2022-2021.
_LTR_SEPARATORS = set(",-–—−/:;.+%")
_QUOTES = set("\"'«»“”")


def _neighbor_atom(ordered: list[_Atom], index: int, step: int) -> _Atom | None:
    cursor = index + step
    while 0 <= cursor < len(ordered):
        if ordered[cursor].text.strip():
            return ordered[cursor]
        cursor += step
    return None


def _is_ltr_token(atom: _Atom) -> bool:
    """Digits or English, not an Arabic word that happens to contain a digit."""
    text = atom.text.strip()
    if not text or any(_is_arabic(char) for char in text):
        return False
    return any(char.isascii() and char.isalnum() for char in text)


def _single_arabic_letter(atom: _Atom) -> bool:
    letters = [char for char in atom.text if _is_arabic(char)]
    return len(letters) == 1 and len(atom.text.strip()) <= 2


def _adopt_ltr_separators(atoms: list[_Atom]) -> list[_Atom]:
    """Keep a comma, hyphen, or slash with the numbers on either side of it."""
    ordered = sorted(atoms, key=lambda atom: (atom.x0, atom.origin_x))
    for index, atom in enumerate(ordered):
        text = atom.text.strip()
        if len(text) != 1 or text not in _LTR_SEPARATORS:
            continue
        previous = _neighbor_atom(ordered, index, -1)
        nxt = _neighbor_atom(ordered, index, 1)
        if previous is None or nxt is None:
            continue
        if _is_ltr_token(previous) and _is_ltr_token(nxt):
            atom.level = 0 if previous.level % 2 else previous.level
    return ordered


def _bind_quotes(atoms: list[_Atom]) -> list[_Atom]:
    """A quote wrapped around one Arabic letter stays on that letter's side."""
    ordered = sorted(atoms, key=lambda atom: (atom.x0, atom.origin_x))
    for index, atom in enumerate(ordered):
        if atom.text.strip() not in _QUOTES:
            continue
        host = None
        for candidate in (
            _neighbor_atom(ordered, index, 1),
            _neighbor_atom(ordered, index, -1),
        ):
            if candidate is not None and _single_arabic_letter(candidate):
                host = candidate
                break
        if host is not None:
            atom.level = host.level
    return ordered


def _prepare_atoms(atoms: list[_Atom]) -> list[_Atom]:
    """Copy the glyphs and apply the repairs that must happen before reordering."""
    prepared = _drop_covered_spaces([_clone_atom(atom) for atom in atoms])
    prepared = _drop_phantom_alef(prepared)
    prepared = _bind_marks(prepared)
    prepared = _adopt_latin_brackets(prepared)
    prepared = _adopt_ltr_separators(prepared)
    return _bind_quotes(prepared)


def _logical_text(atoms: list[_Atom]) -> str:
    """
    Read one line the way the PDF marked it.

    An odd bidi level means the glyph was painted from the left and has to be
    turned back into logical Arabic. A lexical score is not used here: a
    reversed word such as للاخ starts with لل and would otherwise beat the
    real reading واستخدموا معامل ارتباط بيرسون من خلال فحص.
    """
    if not atoms:
        return ""
    atoms = _prepare_atoms(atoms)
    kept = plain_arabic(_keep_visual_order(atoms))
    if any(atom.level % 2 == 1 for atom in atoms):
        restored = plain_arabic(_rtl_from_visual(atoms))
        if restored:
            return restored
    return kept


def _logical_pieces(atoms: list[_Atom], ltr_base: bool = False) -> tuple[list[tuple[_Atom, str]], bool]:
    """
    The glyphs of one line in the order `_logical_text` reads them.

    Each glyph comes with the text it adds, so a caller can keep the glyph's
    font and color next to its letters. The second value is True when the
    line was turned from visual order, which is when it reads from the right.
    Joining the pieces gives the string `_logical_text` cleans up. A caller
    that needs the styles compares the two and falls back to the plain string
    when they differ.

    With `ltr_base`, the line is an English line that holds Arabic words.
    Each Arabic word is turned back into logical order and the line still
    reads from the left, so the sentence is not stored back to front.
    """
    if not atoms:
        return [], False
    prepared = _prepare_atoms(atoms)
    ordered = sorted(prepared, key=lambda atom: (atom.x0, atom.origin_x))
    if not any(atom.level % 2 == 1 for atom in prepared):
        return [(atom, atom.text) for atom in ordered], False
    runs: list[tuple[bool, list[_Atom]]] = []
    for atom in ordered:
        odd = atom.level % 2 == 1
        if runs and runs[-1][0] == odd:
            runs[-1][1].append(atom)
        else:
            runs.append((odd, [atom]))
    if ltr_base:
        flat: list[tuple[_Atom, str]] = []
        for odd, group in runs:
            sequence = list(reversed(group)) if odd else list(group)
            flat.extend((atom, atom.text) for atom in sequence)
        return flat, False
    pieces: list[tuple[bool, list[tuple[_Atom, str]]]] = []
    for odd, group in runs:
        sequence = list(reversed(group)) if odd else list(group)
        items = [(atom, atom.text) for atom in sequence]
        if not odd and len(items) >= 2:
            text = "".join(part for _atom, part in items)
            first, last = items[0], items[-1]
            mirrored = (
                len(text) >= 2
                and text[0] in _CLOSING_TO_OPEN
                and text[-1] == _CLOSING_TO_OPEN[text[0]]
            )
            # ")Alkal, 2021(" is stored mirrored. The two brackets trade places.
            if mirrored and len(first[1]) == 1 and len(last[1]) == 1:
                items[0] = (first[0], last[1])
                items[-1] = (last[0], first[1])
        pieces.append((odd, items))
    groups: list[tuple[bool, list[list[tuple[_Atom, str]]]]] = []
    for odd, items in pieces:
        if not odd and groups and groups[-1][0] is False:
            groups[-1][1].append(items)
        else:
            groups.append((odd, [items]))
    groups.reverse()
    return [item for _odd, parts in groups for items in parts for item in items], True


