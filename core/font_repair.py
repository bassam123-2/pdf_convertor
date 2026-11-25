# -*- coding: utf-8 -*-
"""
Read the true letters of a PDF font whose Unicode map is wrong, by shape.

Some PDF writers (macOS Quartz among them) embed a subset of an Arabic font
and store a wrong Unicode value for many of its glyphs: a reader gets
"المناPضة" or "ب|ن" for المناقضة and بين. The glyph drawings themselves are
untouched. A subset glyph is an exact copy of the glyph in the full font, so
when the full font is installed (Sakkal Majalla, Arial, Times New Roman,
Traditional Arabic ... ship with Windows), each embedded glyph can be found
in it by its outline, and the full font says which letter it is:

    embedded glyph 26  --outline-->  installed glyph "lam.medi"
                       --cmap/GSUB-->  ل

No OCR, no guessing, and the text keeps its exact position, size, weight and
colour, because it is still the PDF's own text layer.

A font is only rewritten when its stored values disagree with the shapes:
a healthy font is left exactly as it was.
"""

from __future__ import annotations

import io
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import fitz

_ARABIC = re.compile(r"[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]")
# Share of a font's glyphs whose stored letter must differ from its shape
# before the font is treated as broken.
BROKEN_SHARE = 0.2


# --------------------------------------------------------------------------
# Glyph outlines


def _outline_keys(font) -> dict[str, tuple]:
    """Glyph name -> a key made of its advance and its drawn outline."""
    from fontTools.pens.recordingPen import DecomposingRecordingPen

    glyphs = font.getGlyphSet()
    metrics = font["hmtx"].metrics if "hmtx" in font else {}
    units = float(font["head"].unitsPerEm or 1000)
    keys: dict[str, tuple] = {}
    for name in font.getGlyphOrder():
        try:
            pen = DecomposingRecordingPen(glyphs)
            glyphs[name].draw(pen)
        except Exception:
            continue
        if not pen.value:
            continue  # an empty glyph (space) says nothing about its letter
        scale = 1000.0 / units
        ops = tuple(
            (op, tuple((round(x * scale), round(y * scale)) for x, y in points))
            for op, points in pen.value
        )
        advance = round(metrics.get(name, (0, 0))[0] * scale)
        keys[name] = (advance, ops)
    return keys


def _glyph_text(font) -> dict[str, str]:
    """
    Glyph name -> the letters it draws, from the cmap and the GSUB features.

    A positional form (beh.init) is reached from its letter by a single
    substitution, a ligature (lam_alef) from its letters by a ligature
    substitution. Following those rules backwards gives every form its text.
    """
    text: dict[str, str] = {}
    try:
        cmap = font.getBestCmap() or {}
    except Exception:
        cmap = {}
    for code, name in sorted(cmap.items()):
        char = chr(code)
        # Presentation forms fold to their letter; a base letter wins.
        folded = _fold(char)
        if name not in text or (len(text[name]) == 1 and not _ARABIC.match(text[name])):
            text[name] = folded
        elif "\uFB50" <= text[name] <= "\uFEFF":
            text[name] = folded
    if "GSUB" not in font:
        return text
    lookups = []
    try:
        for lookup in font["GSUB"].table.LookupList.Lookup:
            for sub in lookup.SubTable:
                while getattr(sub, "LookupType", None) == 7 or type(sub).__name__ == "ExtensionSubst":
                    sub = sub.ExtSubTable
                lookups.append(sub)
    except Exception:
        return text
    for _round in range(6):
        changed = False
        for sub in lookups:
            kind = type(sub).__name__
            if kind == "SingleSubst":
                for source, target in (sub.mapping or {}).items():
                    if source in text and target not in text:
                        text[target] = text[source]
                        changed = True
            elif kind == "AlternateSubst":
                for source, targets in (getattr(sub, "alternates", None) or {}).items():
                    for target in targets:
                        if source in text and target not in text:
                            text[target] = text[source]
                            changed = True
            elif kind == "LigatureSubst":
                for first, ligatures in (sub.ligatures or {}).items():
                    for ligature in ligatures:
                        parts = [first, *ligature.Component]
                        if ligature.LigGlyph not in text and all(part in text for part in parts):
                            text[ligature.LigGlyph] = "".join(text[part] for part in parts)
                            changed = True
        if not changed:
            break
    return text


def _fold(char: str) -> str:
    import unicodedata

    if "\uFB50" <= char <= "\uFDFF" or "\uFE70" <= char <= "\uFEFF":
        folded = unicodedata.normalize("NFKC", char)
        return folded or char
    return char


def _core_arabic(text: str) -> bool:
    return all("\u0621" <= char <= "\u064A" for char in text)


def _settle(options: set[str]) -> str | None:
    """
    The letters for a shape that several characters share, or None.

    One option is the answer. When a font draws a letter and its Persian or
    Urdu twin alike (ي and ی, ك and ک, ه and ە), the letter of the Arabic
    alphabet is taken. Any other clash is left alone: the shape cannot tell.
    """
    if len(options) == 1:
        return next(iter(options))
    core = [option for option in options if _core_arabic(option)]
    if len(core) == 1:
        return core[0]
    return None


# --------------------------------------------------------------------------
# Installed fonts


@dataclass
class _Reference:
    path: str
    index: int
    keys: dict[tuple, str] = field(default_factory=dict)  # outline key -> text


_REFERENCES: dict[tuple[str, int], _Reference | None] = {}
_FAMILY_FILES: dict[str, list[tuple[str, int]]] | None = None


def _font_dirs() -> list[Path]:
    dirs = []
    if sys.platform.startswith("win"):
        dirs.append(Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts")
        local = os.environ.get("LOCALAPPDATA", "")
        if local:
            dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    elif sys.platform == "darwin":
        dirs += [Path("/System/Library/Fonts"), Path("/Library/Fonts"), Path.home() / "Library" / "Fonts"]
    else:
        dirs += [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"), Path.home() / ".fonts",
                 Path.home() / ".local" / "share" / "fonts"]
    extra = os.environ.get("PDF_TOOL_FONT_DIRS", "")
    dirs += [Path(part) for part in extra.split(os.pathsep) if part]
    return [d for d in dirs if d.is_dir()]


def _key_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def _family_files() -> dict[str, list[tuple[str, int]]]:
    """Compact family/PostScript name -> installed font files (path, index)."""
    global _FAMILY_FILES
    if _FAMILY_FILES is not None:
        return _FAMILY_FILES
    from fontTools.ttLib import TTCollection, TTFont

    found: dict[str, list[tuple[str, int]]] = {}
    for folder in _font_dirs():
        for path in folder.rglob("*"):
            if path.suffix.lower() not in {".ttf", ".otf", ".ttc", ".otc"}:
                continue
            try:
                if path.suffix.lower() in {".ttc", ".otc"}:
                    fonts = list(TTCollection(str(path), lazy=True).fonts)
                else:
                    fonts = [TTFont(str(path), lazy=True, fontNumber=0)]
            except Exception:
                continue
            for index, font in enumerate(fonts):
                try:
                    names = font["name"]
                    labels = {names.getDebugName(1) or "", names.getDebugName(4) or "",
                              names.getDebugName(6) or "", names.getDebugName(16) or ""}
                except Exception:
                    continue
                for label in labels:
                    if label:
                        found.setdefault(_key_name(label), []).append((str(path), index))
    _FAMILY_FILES = found
    return found


def _candidates(pdf_name: str) -> list[tuple[str, int]]:
    """Installed files that may hold the full font of an embedded subset."""
    base = re.sub(r"^[A-Z]{6}\+", "", pdf_name or "")
    files = _family_files()
    keys = [_key_name(base)]
    family = re.split(r"[-,]", base, maxsplit=1)[0]
    keys.append(_key_name(re.sub(r"(PSMT|MT|PS)$", "", family)))
    keys.append(_key_name(re.sub(r"(?<=[a-z])(?=[A-Z])", " ", re.sub(r"(PSMT|MT|PS)$", "", family))))
    seen, out = set(), []
    # Exact PostScript name first (it carries the weight), then the family.
    for key in keys:
        for item in files.get(key, []):
            if item not in seen:
                seen.add(item)
                out.append(item)
    return out


def _reference(path: str, index: int) -> _Reference | None:
    marker = (path, index)
    if marker in _REFERENCES:
        return _REFERENCES[marker]
    from fontTools.ttLib import TTFont

    try:
        font = TTFont(path, fontNumber=index)
        keys = _outline_keys(font)
        texts = _glyph_text(font)
    except Exception:
        _REFERENCES[marker] = None
        return None
    reference = _Reference(path, index)
    choices: dict[tuple, set[str]] = {}
    for name, key in keys.items():
        letters = texts.get(name)
        if letters:
            choices.setdefault(key, set()).add(letters)
    for key, options in choices.items():
        chosen = _settle(options)
        if chosen:
            reference.keys[key] = chosen
    _REFERENCES[marker] = reference
    return reference


# --------------------------------------------------------------------------
# Embedded fonts


@dataclass
class EmbeddedMap:
    """Letters of one embedded font, by glyph id, read from shapes."""

    name: str
    letters: dict[int, str]
    advances: dict[int, float]
    units: float
    matched: int
    reference: str


_EMBEDDED: dict[tuple, EmbeddedMap | None] = {}


def embedded_map(document: fitz.Document, xref: int, pdf_name: str) -> EmbeddedMap | None:
    """Glyph id -> letters for one embedded TrueType/OpenType font, or None."""
    marker = (id(document), getattr(document, "name", ""), xref)
    if marker in _EMBEDDED:
        return _EMBEDDED[marker]
    _EMBEDDED[marker] = None
    try:
        _name, ext, _kind, data = document.extract_font(xref)
    except Exception:
        return None
    if not data or ext not in {"ttf", "otf", "cff", "ttc"}:
        return None
    from fontTools.ttLib import TTFont

    try:
        font = TTFont(io.BytesIO(data))
        keys = _outline_keys(font)
        order = font.getGlyphOrder()
        units = float(font["head"].unitsPerEm or 1000)
        metrics = font["hmtx"].metrics if "hmtx" in font else {}
    except Exception:
        return None
    best: EmbeddedMap | None = None
    for path, index in _candidates(pdf_name):
        reference = _reference(path, index)
        if reference is None:
            continue
        letters: dict[int, str] = {}
        for gid, name in enumerate(order):
            key = keys.get(name)
            if key is not None and key in reference.keys:
                letters[gid] = reference.keys[key]
            elif key is None and gid and float(metrics.get(name, (0, 0))[0]) > units * 0.1:
                # Nothing drawn but the pen moves on: a space.
                letters[gid] = " "
        drawn = sum(1 for name in order if name in keys)
        if drawn and len(letters) >= max(3, drawn * 0.6) and (best is None or len(letters) > best.matched):
            advances = {gid: float(metrics.get(name, (0, 0))[0]) for gid, name in enumerate(order)}
            best = EmbeddedMap(pdf_name, letters, advances, units, len(letters), path)
    _EMBEDDED[marker] = best
    return best


# --------------------------------------------------------------------------
# Pages


def _base(name: str) -> str:
    return re.sub(r"^[A-Z]{6}\+", "", name or "")


_PAGE_MAPS: dict[tuple, dict[str, list[EmbeddedMap]]] = {}


def page_maps(page: fitz.Page) -> dict[str, list[EmbeddedMap]]:
    """Compact font name (see _key_name) -> shape maps of its subsets on this page."""
    marker = (id(page.parent), getattr(page.parent, "name", ""), page.number)
    if marker in _PAGE_MAPS:
        return _PAGE_MAPS[marker]
    maps: dict[str, list[EmbeddedMap]] = {}
    try:
        fonts = page.get_fonts(full=True)
    except Exception:
        fonts = []
    for entry in fonts:
        xref, ext, _kind, basefont = int(entry[0]), entry[1], entry[2], str(entry[3] or "")
        if ext not in {"ttf", "otf", "cff"}:
            continue
        found = embedded_map(page.parent, xref, basefont)
        if found is not None:
            maps.setdefault(_key_name(_base(basefont)), []).append(found)
    if len(_PAGE_MAPS) > 64:
        _PAGE_MAPS.clear()
    _PAGE_MAPS[marker] = maps
    return maps


def _pick(maps: list[EmbeddedMap], trace: dict) -> EmbeddedMap | None:
    """
    The subset a trace was drawn with.

    A page may hold several subsets of one font under the same name, each
    with its own glyph numbers. The glyph widths the trace drew are compared
    with each subset's advances; the subset that agrees wins.
    """
    if len(maps) == 1:
        return maps[0]
    size = float(trace.get("size") or 0) or 1.0
    best, best_error = None, None
    for candidate in maps:
        error = 0.0
        counted = 0
        for raw in trace.get("chars") or []:
            gid = int(raw[1])
            if gid < 0:
                continue
            box = raw[3]
            drawn = float(box[2]) - float(box[0])
            expected = candidate.advances.get(gid)
            if expected is None:
                error += size
                counted += 1
                continue
            error += abs(drawn - expected * size / candidate.units)
            counted += 1
        if counted and (best_error is None or error < best_error):
            best, best_error = candidate, error
    return best


_GARBAGE = re.compile(r"[\x00-\x1f\ufffd\ue000-\uf8ff\u0370-\u03ff\u1400-\u167f\u0180-\u024f]")


def _garbage(char: str) -> bool:
    """A stored value that is plainly not the glyph's letter."""
    return bool(_GARBAGE.match(char))


def _agreement(candidate: EmbeddedMap, trace: dict) -> tuple[int, int]:
    """How many glyphs of a trace store the letter their shape draws, of those known."""
    agree = known = 0
    for raw in trace.get("chars") or []:
        gid = int(raw[1])
        if gid < 0:
            continue
        letters = candidate.letters.get(gid)
        if not letters or letters == " ":
            continue
        known += 1
        stored = _fold(chr(int(raw[0])))
        if stored == letters or stored == letters[:1]:
            agree += 1
    return agree, known


def _choose(maps: list[EmbeddedMap], trace: dict) -> tuple[EmbeddedMap | None, bool]:
    """
    The subset a trace was drawn with, and whether its stored values are sound.

    A page can hold many fonts of one name, each numbering its glyphs its own
    way; reading a trace with the wrong one turns "The" into "Rhe". So:

    * A subset whose shapes agree with most stored letters is the one, and
      the font is healthy: only plainly broken glyphs are replaced.
    * When none agrees, the font map is broken. A single subset is taken; of
      several, the one whose glyph widths match the trace, and only when they
      match closely. Otherwise nothing is changed.
    """
    scored = []
    for candidate in maps:
        agree, known = _agreement(candidate, trace)
        scored.append((agree / known if known else 0.0, known, candidate))
    sound = [item for item in scored if item[1] >= 2 and item[0] >= 0.6]
    if sound:
        return max(sound, key=lambda item: (item[0], item[1]))[2], True
    if not any(item[1] for item in scored):
        return None, True
    if len(maps) == 1:
        return maps[0], False
    fitting = [candidate for candidate in maps if _widths_match(candidate, trace)]
    if len(fitting) != 1:
        return None, True  # no subset, or more than one, fits: leave the trace alone
    return fitting[0], False


def _widths_match(candidate: EmbeddedMap, trace: dict) -> bool:
    """True when the trace's glyph widths are this subset's advances, glyph for glyph."""
    size = float(trace.get("size") or 0) or 1.0
    checked = 0
    for raw in trace.get("chars") or []:
        gid = int(raw[1])
        if gid < 0:
            continue
        expected = candidate.advances.get(gid)
        if expected is None:
            return False
        drawn = float(raw[3][2]) - float(raw[3][0])
        if abs(drawn - expected * size / candidate.units) > max(0.3, size * 0.03):
            return False
        checked += 1
    return checked >= 3


def _stored_disagrees(maps: list[EmbeddedMap], traces: list[dict]) -> bool:
    """True when the PDF's own values for these glyphs often differ from their shapes."""
    total = wrong = 0
    for trace in traces:
        chosen = _pick(maps, trace)
        if chosen is None:
            continue
        for raw in trace.get("chars") or []:
            gid = int(raw[1])
            letters = chosen.letters.get(gid)
            if not letters:
                continue
            stored = _fold(chr(int(raw[0])))
            total += 1
            if stored != letters[:1] and stored != letters:
                wrong += 1
    return total >= 5 and wrong >= total * BROKEN_SHARE


def _next_completes(raw, trace: dict, letters: str) -> bool:
    """True when the PDF stored a ligature as its first letter plus a no-glyph second letter."""
    chars = list(trace.get("chars") or [])
    try:
        index = chars.index(raw)
    except ValueError:
        return False
    following = chars[index + 1] if index + 1 < len(chars) else None
    return bool(following and int(following[1]) < 0 and _fold(chr(int(following[0]))) == letters[1:])


def repaired_traces(page: fitz.Page, traces: list[dict]) -> list[dict]:
    """
    The page's text traces with the letters of broken fonts read from shape.

    Each repaired glyph keeps its position and size. Its text becomes the
    letter its shape draws; a lam-alef glyph becomes the two letters and the
    glyph that follows it on the same spot is dropped, as the PDF readers
    store a ligature. Arabic glyphs get an odd bidi level so the line is put
    back in reading order.
    """
    try:
        maps = page_maps(page)
    except Exception:
        return traces
    if not maps:
        return traces
    by_font: dict[str, list[dict]] = {}
    for trace in traces:
        key = _key_name(_base(str(trace.get("font") or "")))
        if int(trace.get("type") or 0) == 0 and key in maps:
            by_font.setdefault(key, []).append(trace)
    # Every font is checked glyph by glyph: a healthy font keeps all its
    # values, and a font with a few broken glyphs (a lam-alef stored as
    # U+FFFD) has just those glyphs read from their shapes.
    broken = set(by_font)
    if not broken:
        return traces
    fixed: list[dict] = []
    for trace in traces:
        key = _key_name(_base(str(trace.get("font") or "")))
        if key not in broken or int(trace.get("type") or 0) != 0:
            fixed.append(trace)
            continue
        chosen, healthy = _choose(maps[key], trace)
        if chosen is None:
            fixed.append(trace)
            continue
        chars = []
        arabic = False
        changed = False
        last_letters = ""
        for raw in trace.get("chars") or []:
            gid = int(raw[1])
            if gid < 0:
                # The second letter of a ligature the PDF split itself; when
                # the shape already gave both letters it is not added again.
                if len(last_letters) > 1:
                    continue
                chars.append(tuple(raw))
                continue
            letters = chosen.letters.get(gid)
            stored = _fold(chr(int(raw[0])))
            # In a healthy font only a glyph with no real letter stored (U+FFFD,
            # a control code, a private-use or foreign-script code) is replaced.
            if healthy and not _garbage(chr(int(raw[0]))):
                letters = None
            if not letters or stored == letters or (len(letters) > 1 and stored == letters[0] and _next_completes(raw, trace, letters)):
                # The PDF's own value agrees with the shape: keep it as it is.
                last_letters = "" if not letters or len(letters) == 1 else ""
                chars.append(tuple(raw))
                continue
            last_letters = letters
            changed = True
            if _ARABIC.search(letters):
                arabic = True
            # A ligature drawn as two glyphs puts the second, with no width, on
            # the first one's origin. Stored as a no-glyph letter it joins the
            # first, so the pair is never turned around.
            box = raw[3]
            joined = (
                chars
                and abs(float(raw[2][0]) - float(chars[-1][2][0])) <= 0.8
                and abs(float(box[2]) - float(box[0])) <= 0.5
            )
            chars.append((ord(letters[0]), -1 if joined else gid, raw[2], raw[3]))
            for extra in letters[1:]:
                chars.append((ord(extra), -1, raw[2], raw[3]))
        if not changed:
            fixed.append(trace)
            continue
        copy = dict(trace)
        copy["chars"] = tuple(chars)
        if arabic and int(copy.get("bidi_lvl") or 0) % 2 == 0:
            copy["bidi_lvl"] = 1
        copy["_shape_repaired"] = True
        fixed.append(copy)
    return fixed
