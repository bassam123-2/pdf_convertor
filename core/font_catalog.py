# -*- coding: utf-8 -*-
"""
The fonts installed on this computer, and the one Word should use for a PDF font.

A PDF names its fonts the PostScript way: ABCDEF+Calibri-Bold, ArialMT,
TimesNewRomanPSMT. Word wants a family name: Calibri, Arial, Times New Roman.
The installed font files carry both names, so the match is read from the
files and is not a guess. A font that is not installed is mapped to the
family with the same metrics when one is known, and otherwise to a family of
the same class (serif, sans-serif, or fixed pitch).

Only the font files' name and metric tables are read. No font is changed.
"""

from __future__ import annotations

import os
import re
import struct
from dataclasses import dataclass
from pathlib import Path

# Families drawn with the same advance widths. The first installed name wins.
# These are the published metric-compatible sets, not names from one document.
_METRIC_SETS = (
    ("arial", ("Arial", "Liberation Sans", "Arimo", "Helvetica", "Nimbus Sans", "Nimbus Sans L")),
    ("helvetica", ("Arial", "Liberation Sans", "Arimo")),
    ("helveticaneue", ("Arial", "Liberation Sans", "Arimo")),
    ("nimbussan", ("Arial", "Liberation Sans", "Arimo")),
    ("liberationsans", ("Arial", "Arimo")),
    ("arimo", ("Arial", "Liberation Sans")),
    ("timesnewroman", ("Times New Roman", "Liberation Serif", "Tinos", "Nimbus Roman")),
    ("times", ("Times New Roman", "Liberation Serif", "Tinos")),
    ("nimbusrom", ("Times New Roman", "Liberation Serif", "Tinos")),
    ("liberationserif", ("Times New Roman", "Tinos")),
    ("tinos", ("Times New Roman", "Liberation Serif")),
    ("couriernew", ("Courier New", "Liberation Mono", "Cousine")),
    ("courier", ("Courier New", "Liberation Mono", "Cousine")),
    ("nimbusmon", ("Courier New", "Liberation Mono", "Cousine")),
    ("liberationmono", ("Courier New", "Cousine")),
    ("cousine", ("Courier New", "Liberation Mono")),
    ("calibri", ("Calibri", "Carlito")),
    ("carlito", ("Calibri",)),
    ("cambria", ("Cambria", "Caladea")),
    ("caladea", ("Cambria",)),
)

# Used when nothing better is known. One family per class of letterform.
_CLASS_FALLBACK = {"serif": "Times New Roman", "sans": "Arial", "mono": "Courier New"}

# Words at the end of a PostScript name that describe the style, not the family.
_STYLE_WORDS = (
    "bolditalic",
    "boldoblique",
    "semibold",
    "demibold",
    "extrabold",
    "ultrabold",
    "bold",
    "italic",
    "oblique",
    "regular",
    "roman",
    "normal",
    "medium",
    "book",
    "light",
    "black",
    "heavy",
    "condensed",
    "narrow",
    "psmt",
    "mt",
    "ps",
)


@dataclass(frozen=True)
class Face:
    """One installed font face and the numbers Word lays lines out with."""

    family: str
    bold: bool
    italic: bool
    path: str
    index: int
    postscript: str
    units: int
    hhea_ascent: int
    hhea_descent: int
    hhea_gap: int
    typo_ascent: int
    typo_descent: int
    typo_gap: int
    win_ascent: int
    win_descent: int
    use_typo: bool

    @property
    def ascent(self) -> float:
        """Height above the baseline, in ems, as Word measures a line."""
        return self._line()[0]

    @property
    def descent(self) -> float:
        """Depth below the baseline, in ems, as Word measures a line."""
        return self._line()[1]

    def _line(self) -> tuple[float, float]:
        units = float(self.units or 1000)
        if self.use_typo and (self.typo_ascent or self.typo_descent):
            return self.typo_ascent / units, abs(self.typo_descent) / units
        if self.win_ascent or self.win_descent:
            return self.win_ascent / units, self.win_descent / units
        return self.hhea_ascent / units, abs(self.hhea_descent) / units


@dataclass(frozen=True)
class FontChoice:
    """The Word family for one PDF font, and how sure the match is."""

    family: str
    bold: bool
    italic: bool
    exact: bool
    source: str


def _key(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def strip_subset(name: str) -> str:
    """ABCDEF+Calibri-Bold is the font Calibri-Bold, embedded as a subset."""
    return re.sub(r"^[A-Z]{6}\+", "", (name or "").strip())


def _tables(handle, offset: int) -> dict[str, tuple[int, int]]:
    handle.seek(offset)
    head = handle.read(12)
    if len(head) < 12:
        return {}
    count = struct.unpack(">H", head[4:6])[0]
    records = handle.read(16 * count)
    found = {}
    for position in range(0, len(records) - 15, 16):
        tag = records[position : position + 4].decode("latin-1")
        table_offset, length = struct.unpack(">II", records[position + 8 : position + 16])
        found[tag] = (table_offset, length)
    return found


def _table(handle, tables: dict, tag: str, limit: int = 1 << 20) -> bytes:
    if tag not in tables:
        return b""
    offset, length = tables[tag]
    handle.seek(offset)
    return handle.read(min(length, limit))


def _names(data: bytes) -> dict[int, str]:
    """Family, style, and PostScript names. English Windows names are preferred."""
    if len(data) < 6:
        return {}
    count, strings = struct.unpack(">HH", data[2:6])
    best: dict[int, tuple[int, str]] = {}
    for index in range(count):
        start = 6 + index * 12
        if start + 12 > len(data):
            break
        platform, encoding, language, name_id, length, offset = struct.unpack(">HHHHHH", data[start : start + 12])
        if name_id not in (1, 2, 4, 6, 16, 17):
            continue
        raw = data[strings + offset : strings + offset + length]
        try:
            if platform == 3 or platform == 0:
                text = raw.decode("utf-16-be")
            else:
                text = raw.decode("mac_roman")
        except UnicodeDecodeError:
            continue
        text = text.strip("\x00 ").strip()
        if not text:
            continue
        rank = 3 if (platform == 3 and language == 0x409) else 2 if platform == 3 else 1
        if name_id not in best or rank > best[name_id][0]:
            best[name_id] = (rank, text)
    return {name_id: text for name_id, (_rank, text) in best.items()}


def _read_faces(path: Path) -> list[Face]:
    faces: list[Face] = []
    try:
        with open(path, "rb") as handle:
            tag = handle.read(4)
            offsets = [0]
            if tag == b"ttcf":
                handle.read(4)
                count = struct.unpack(">I", handle.read(4))[0]
                offsets = list(struct.unpack(f">{count}I", handle.read(4 * count)))
            for index, offset in enumerate(offsets):
                tables = _tables(handle, offset)
                names = _names(_table(handle, tables, "name"))
                family = names.get(1) or names.get(16)
                if not family:
                    continue
                head = _table(handle, tables, "head", 64)
                hhea = _table(handle, tables, "hhea", 64)
                os2 = _table(handle, tables, "OS/2", 128)
                units = struct.unpack(">H", head[18:20])[0] if len(head) >= 20 else 1000
                mac_style = struct.unpack(">H", head[44:46])[0] if len(head) >= 46 else 0
                ascent, descent, gap = struct.unpack(">hhh", hhea[4:10]) if len(hhea) >= 10 else (0, 0, 0)
                selection = typo_ascent = typo_descent = typo_gap = win_ascent = win_descent = 0
                if len(os2) >= 78:
                    selection = struct.unpack(">H", os2[62:64])[0]
                    typo_ascent, typo_descent, typo_gap = struct.unpack(">hhh", os2[68:74])
                    win_ascent, win_descent = struct.unpack(">HH", os2[74:78])
                style = (names.get(2) or "").lower()
                bold = bool(mac_style & 1) or bool(selection & 0x20) or "bold" in style
                italic = bool(mac_style & 2) or bool(selection & 0x01) or "italic" in style or "oblique" in style
                faces.append(
                    Face(
                        family=family,
                        bold=bold,
                        italic=italic,
                        path=str(path),
                        index=index,
                        postscript=names.get(6, ""),
                        units=units or 1000,
                        hhea_ascent=ascent,
                        hhea_descent=descent,
                        hhea_gap=gap,
                        typo_ascent=typo_ascent,
                        typo_descent=typo_descent,
                        typo_gap=typo_gap,
                        win_ascent=win_ascent,
                        win_descent=win_descent,
                        use_typo=bool(selection & 0x80),
                    )
                )
    except (OSError, struct.error, ValueError):
        return faces
    return faces


def _font_files() -> list[Path]:
    """Font files Windows has registered for all users and for this user."""
    files: list[Path] = []
    seen: set[str] = set()
    try:
        import winreg
    except ImportError:
        return files
    system_folder = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    user_folder = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for root, folder in ((winreg.HKEY_LOCAL_MACHINE, system_folder), (winreg.HKEY_CURRENT_USER, user_folder)):
        try:
            with winreg.OpenKey(root, key_path) as key:
                for index in range(winreg.QueryInfoKey(key)[1]):
                    _name, value, _kind = winreg.EnumValue(key, index)
                    path = Path(str(value))
                    if not path.is_absolute():
                        path = folder / path
                    if path.suffix.lower() not in {".ttf", ".ttc", ".otf", ".otc"}:
                        continue
                    marker = str(path).lower()
                    if marker in seen or not path.is_file():
                        continue
                    seen.add(marker)
                    files.append(path)
        except OSError:
            continue
    return files


class _Catalog:
    def __init__(self) -> None:
        self.by_family: dict[str, list[Face]] = {}
        self.by_postscript: dict[str, Face] = {}
        self.names: dict[str, str] = {}
        for path in _font_files():
            for face in _read_faces(path):
                key = _key(face.family)
                self.by_family.setdefault(key, []).append(face)
                self.names.setdefault(key, face.family)
                if face.postscript:
                    self.by_postscript.setdefault(_key(face.postscript), face)


_CATALOG: _Catalog | None = None


def _catalog() -> _Catalog:
    global _CATALOG
    if _CATALOG is None:
        _CATALOG = _Catalog()
    return _CATALOG


def reset_catalog() -> None:
    """Forget the installed fonts. Tests use this after pointing at other folders."""
    global _CATALOG
    _CATALOG = None
    _ADVANCES.clear()
    _PIL_FACES.clear()


def is_installed(family: str) -> bool:
    return _key(family) in _catalog().by_family


def installed_name(family: str) -> str:
    """The family as its font file spells it, or nothing when it is not installed."""
    return _catalog().names.get(_key(family), "")


def find_face(family: str, bold: bool = False, italic: bool = False) -> Face | None:
    """The installed face Word draws for this family and style."""
    faces = _catalog().by_family.get(_key(family))
    if not faces:
        return None
    return min(faces, key=lambda face: (face.bold != bold) * 2 + (face.italic != italic))


def _split_style(name: str) -> tuple[str, bool, bool]:
    """Family key, bold, italic from a PostScript font name."""
    base = strip_subset(name)
    lowered = base.lower()
    bold = bool(re.search(r"bold|black|heavy|semibold|demi", lowered))
    italic = bool(re.search(r"italic|oblique", lowered))
    family = re.split(r"[-,]", base, maxsplit=1)[0]
    key = _key(family)
    changed = True
    while changed:
        changed = False
        for word in _STYLE_WORDS:
            if key.endswith(word) and len(key) > len(word) + 2:
                key = key[: -len(word)]
                changed = True
    return key, bold, italic


def _spaced(name: str) -> str:
    """TimesNewRoman is written Times New Roman when no installed name is known."""
    base = re.split(r"[-,]", strip_subset(name), maxsplit=1)[0]
    base = re.sub(r"(PSMT|MT|PS)$", "", base)
    return re.sub(r"(?<=[a-z])(?=[A-Z])", " ", base).strip() or "Arial"


def choose_font(pdf_name: str, flags: int = 0) -> FontChoice:
    """
    The Word family for a PDF font.

    `flags` are the PDF font's own: 2 italic, 4 serif, 8 fixed pitch, 16 bold.
    The same family is used when it is installed. Otherwise a family with the
    same metrics, and otherwise one of the same class.
    """
    catalog = _catalog()
    flag_bold = bool(flags & 16)
    flag_italic = bool(flags & 2)
    base = strip_subset(pdf_name)
    face = catalog.by_postscript.get(_key(base))
    if face is not None:
        return FontChoice(face.family, face.bold or flag_bold, face.italic or flag_italic, True, "postscript name")
    key, name_bold, name_italic = _split_style(base)
    bold = flag_bold or name_bold
    italic = flag_italic or name_italic
    if key in catalog.by_family:
        return FontChoice(catalog.names[key], bold, italic, True, "family name")
    whole = _key(re.split(r"[-,]", base, maxsplit=1)[0])
    if whole in catalog.by_family:
        return FontChoice(catalog.names[whole], bold, italic, True, "family name")
    # The longest alias first, so Times New Roman is not taken for Times.
    for alias, families in sorted(_METRIC_SETS, key=lambda item: -len(item[0])):
        if key.startswith(alias) or whole.startswith(alias):
            for family in families:
                if _key(family) in catalog.by_family:
                    return FontChoice(catalog.names[_key(family)], bold, italic, False, "same metrics")
    letter_class = "mono" if flags & 8 else "serif" if flags & 4 else "sans"
    fallback = _CLASS_FALLBACK[letter_class]
    if catalog.by_family and _key(fallback) not in catalog.by_family:
        fallback = next(iter(catalog.names.values()))
    if not catalog.by_family:
        # The installed fonts could not be read. Keep the PDF's own family name:
        # Word substitutes on its side when the family is missing there.
        return FontChoice(_spaced(base), bold, italic, False, "name only")
    return FontChoice(fallback, bold, italic, False, f"{letter_class} class")


_ADVANCES: dict[tuple[str, int], dict[str, float]] = {}
_PIL_FACES: dict[tuple[str, int], object] = {}
_MEASURE_PX = 1000


def _pil_face(face: Face):
    key = (face.path, face.index)
    if key not in _PIL_FACES:
        try:
            from PIL import ImageFont

            _PIL_FACES[key] = ImageFont.truetype(
                face.path, _MEASURE_PX, index=face.index, layout_engine=ImageFont.Layout.BASIC
            )
        except Exception:
            _PIL_FACES[key] = None
    return _PIL_FACES[key]


def char_advance(face: Face, char: str) -> float | None:
    """Advance width of one character in ems, with no kerning, as Word sets text."""
    table = _ADVANCES.setdefault((face.path, face.index), {})
    if char in table:
        return table[char]
    pil = _pil_face(face)
    if pil is None:
        return None
    try:
        value = float(pil.getlength(char)) / _MEASURE_PX
    except Exception:
        return None
    table[char] = value
    return value


def has_glyph(face: Face, char: str) -> bool:
    """True when the face draws this character itself."""
    pil = _pil_face(face)
    if pil is None:
        return False
    try:
        return bool(pil.getmask(char).getbbox()) or char.isspace()
    except Exception:
        return False


def text_advance(face: Face, text: str) -> float | None:
    """Width of unshaped text in ems: the sum of the letters' own advances."""
    total = 0.0
    for char in text:
        value = char_advance(face, char)
        if value is None:
            return None
        total += value
    return total


_MISSING: dict[tuple[str, int], object] = {}
_COVERAGE: dict[tuple[str, int, str], bool] = {}


def covers(face: Face, char: str) -> bool:
    """
    True when the face has its own glyph for this character.

    A face draws the same "missing glyph" box for every character it lacks,
    so a character whose picture equals that box is not in the face.
    """
    key = (face.path, face.index, char)
    if key in _COVERAGE:
        return _COVERAGE[key]
    answer = False
    try:
        from PIL import ImageFont

        def picture(font, letter: str) -> tuple:
            mask = font.getmask(letter)
            return (mask.size, mask.getbbox(), tuple(mask.histogram()), font.getlength(letter))

        slot = (face.path, face.index)
        if slot not in _MISSING:
            small = ImageFont.truetype(face.path, 48, index=face.index, layout_engine=ImageFont.Layout.BASIC)
            _MISSING[slot] = (small, picture(small, "\uffff"))
        small, blank = _MISSING[slot]
        answer = char.isspace() or picture(small, char) != blank
    except Exception:
        answer = False
    _COVERAGE[key] = answer
    return answer


class _Shaper:
    """Windows' own text shaping, used to measure joined Arabic the way Word sets it."""

    _UNITS = 1000

    def __init__(self) -> None:
        import ctypes
        from ctypes import wintypes

        self.ctypes = ctypes
        self.gdi = ctypes.WinDLL("gdi32", use_last_error=True)
        self.usp = ctypes.WinDLL("usp10", use_last_error=True)
        self.gdi.CreateCompatibleDC.restype = ctypes.c_void_p
        self.gdi.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
        self.gdi.CreateFontW.restype = ctypes.c_void_p
        self.gdi.CreateFontW.argtypes = [ctypes.c_int] * 5 + [wintypes.DWORD] * 8 + [wintypes.LPCWSTR]
        self.gdi.SelectObject.restype = ctypes.c_void_p
        self.gdi.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        self.gdi.GetTextFaceW.argtypes = [ctypes.c_void_p, ctypes.c_int, wintypes.LPWSTR]
        self.usp.ScriptStringAnalyse.restype = ctypes.c_long
        self.usp.ScriptStringAnalyse.argtypes = [
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            wintypes.DWORD,
            ctypes.c_int,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p),
        ]
        self.usp.ScriptString_pSize.restype = ctypes.POINTER(wintypes.SIZE)
        self.usp.ScriptString_pSize.argtypes = [ctypes.c_void_p]
        self.usp.ScriptStringFree.argtypes = [ctypes.POINTER(ctypes.c_void_p)]
        self.dc = self.gdi.CreateCompatibleDC(None)
        if not self.dc:
            raise OSError("no device context")
        self.fonts: dict[tuple[str, bool, bool], int | None] = {}
        self.widths: dict[tuple[str, bool, bool, str], float | None] = {}

    def _font(self, family: str, bold: bool, italic: bool):
        key = (family, bold, italic)
        if key not in self.fonts:
            handle = self.gdi.CreateFontW(
                -self._UNITS, 0, 0, 0, 700 if bold else 400, 1 if italic else 0, 0, 0, 1, 0, 0, 0, 0, family
            )
            self.fonts[key] = handle or None
        return self.fonts[key]

    def width(self, family: str, bold: bool, italic: bool, text: str) -> float | None:
        key = (family, bold, italic, text)
        if key in self.widths:
            return self.widths[key]
        ctypes = self.ctypes
        value = None
        font = self._font(family, bold, italic)
        if font and text:
            self.gdi.SelectObject(self.dc, font)
            analysis = ctypes.c_void_p()
            buffer = ctypes.create_unicode_buffer(text)
            # SSA_GLYPHS | SSA_RTL: shape the letters and read them from the right.
            status = self.usp.ScriptStringAnalyse(
                self.dc,
                buffer,
                len(text),
                len(text) * 2 + 16,
                -1,
                0x80 | 0x100,
                0,
                None,
                None,
                None,
                None,
                None,
                ctypes.byref(analysis),
            )
            if status == 0 and analysis:
                size = self.usp.ScriptString_pSize(analysis)
                if size:
                    value = float(size.contents.cx) / self._UNITS
                self.usp.ScriptStringFree(ctypes.byref(analysis))
        self.widths[key] = value
        return value


_SHAPER: _Shaper | None | bool = None


def shaped_advance(face: Face, text: str) -> float | None:
    """
    Width in ems of joined text, such as an Arabic word, in this face.

    Arabic letters change shape and width when they join, so adding up the
    letters' own widths is wrong. Windows shapes the word the way Word does.
    The answer is None where that service is missing.
    """
    global _SHAPER
    if _SHAPER is None:
        try:
            _SHAPER = _Shaper()
        except Exception:
            _SHAPER = False
    if not _SHAPER:
        return None
    try:
        return _SHAPER.width(face.family, face.bold, face.italic, text)
    except Exception:
        return None
