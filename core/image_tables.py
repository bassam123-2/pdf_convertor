# -*- coding: utf-8 -*-
"""
Read a table that is only a picture (a screenshot pasted into the document).

The PDF holds no letters for such a table, only pixels. The region is drawn
again at high resolution (small print reads far better when it is large),
img2table finds the ruled grid in the pixels, and Tesseract reads each cell.
The result is the table's rows, cell boxes and column widths in PDF points,
ready to be written as an editable Word table.

Nothing runs when img2table or Tesseract is missing; the caller then keeps
the picture as before.
"""

from __future__ import annotations

import os
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import fitz

# Drawing resolution for the picture of the table.
DPI = 300
# Words read with less confidence than this are dropped by img2table.
MIN_CONFIDENCE = 40


@dataclass
class ImageTable:
    rows: list[list[str]]
    widths: list[float]
    heights: list[float]
    bbox: tuple[float, float, float, float]
    notes: list[str] = field(default_factory=list)


def _languages(folder: Path) -> set[str]:
    return {path.stem for path in folder.glob("*.traineddata")}


def _quietly(factory, *args, **kwargs):
    """
    Build img2table's reader without its banner.

    It runs "tesseract --version" to check the program is there, and that
    prints the version and library list into the terminal on every start.
    """
    import sys

    try:
        sys.stdout.flush()
        sys.stderr.flush()
        saved = os.dup(1), os.dup(2)
        null = os.open(os.devnull, os.O_WRONLY)
        os.dup2(null, 1)
        os.dup2(null, 2)
    except Exception:
        return factory(*args, **kwargs)
    try:
        return factory(*args, **kwargs)
    finally:
        os.dup2(saved[0], 1)
        os.dup2(saved[1], 2)
        for handle in (null, *saved):
            os.close(handle)


def tessdata_for(wanted: tuple[str, ...] = ("eng", "ara")) -> tuple[str | None, str]:
    """
    A tessdata folder that holds every wanted language, and the "+"-joined list.

    An installed Tesseract often has English only, while the Arabic model
    ships with this tool. When no single folder has both, the two models are
    copied once into a folder of the tool's own.
    """
    from engine.tesseract_bin import ara_tessdata_dir, find_exe

    folders: list[Path] = []
    override = os.environ.get("PDF_TOOL_TESSDATA", "").strip()
    if override:
        folders.append(Path(override))
    exe = find_exe()
    if exe:
        folders.append(Path(exe).resolve().parent / "tessdata")
    for prefix in ("/usr/share/tesseract-ocr/5/tessdata", "/usr/share/tesseract-ocr/4.00/tessdata", "/usr/share/tessdata"):
        folders.append(Path(prefix))
    ara = ara_tessdata_dir()
    if ara is not None:
        folders.append(ara)
    folders = [folder for folder in folders if folder.is_dir()]
    for folder in folders:
        have = _languages(folder)
        if all(lang in have for lang in wanted):
            return str(folder), "+".join(wanted)
    # Gather each model from wherever it is.
    found: dict[str, Path] = {}
    for folder in folders:
        for lang in wanted:
            path = folder / f"{lang}.traineddata"
            if lang not in found and path.is_file():
                found[lang] = path
    if not found:
        return None, ""
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".cache") / "pdf_tool" / "tessdata"
    try:
        base.mkdir(parents=True, exist_ok=True)
        for lang, path in found.items():
            target = base / path.name
            if not target.is_file() or target.stat().st_size != path.stat().st_size:
                shutil.copyfile(path, target)
        for extra in ("osd.traineddata",):
            for folder in folders:
                if (folder / extra).is_file() and not (base / extra).is_file():
                    shutil.copyfile(folder / extra, base / extra)
        # Tesseract reads its output settings (hocr, tsv ...) from the same
        # folder; without them it writes nothing.
        for sub in ("configs", "tessconfigs"):
            if (base / sub).is_dir():
                continue
            for folder in folders:
                if (folder / sub).is_dir():
                    shutil.copytree(folder / sub, base / sub)
                    break
    except OSError:
        return None, ""
    langs = [lang for lang in wanted if lang in found]
    return str(base), "+".join(langs)


_OCR = None
_READERS: dict[str, object] = {}


def _reader(langs: str):
    """An img2table Tesseract reader for one language set, made once."""
    if langs in _READERS:
        return _READERS[langs]
    base = _ocr()
    if base is None:
        return None
    try:
        from img2table.ocr import TesseractOCR

        tessdata, _all = tessdata_for()
        _READERS[langs] = _quietly(TesseractOCR, n_threads=1, lang=langs, tessdata_dir=tessdata)
    except Exception:
        _READERS[langs] = None
    return _READERS[langs]


def _ocr():
    """One img2table Tesseract reader for the process, or None."""
    global _OCR
    if _OCR is not None:
        return _OCR or None
    _OCR = False
    try:
        from img2table.ocr import TesseractOCR

        from engine.tesseract_bin import find_exe
    except Exception:
        return None
    exe = find_exe()
    if not exe:
        return None
    folder = str(Path(exe).resolve().parent)
    if folder not in os.environ.get("PATH", "").split(os.pathsep):
        os.environ["PATH"] = folder + os.pathsep + os.environ.get("PATH", "")
    tessdata, langs = tessdata_for()
    if not langs:
        return None
    try:
        _OCR = _quietly(TesseractOCR, n_threads=1, lang=langs, tessdata_dir=tessdata)
    except Exception:
        _OCR = False
        return None
    return _OCR


def available() -> bool:
    return _ocr() is not None


def read_table(page: fitz.Page, bbox: tuple[float, float, float, float]) -> ImageTable | None:
    """The table inside one picture region of a page, or None."""
    reader = _ocr()
    if reader is None:
        return None
    try:
        from img2table.document import Image as TableImage
    except Exception:
        return None
    rect = fitz.Rect(bbox) & page.rect
    if rect.is_empty or rect.width < 40 or rect.height < 30:
        return None
    try:
        pixmap = page.get_pixmap(dpi=DPI, clip=rect, alpha=False)
    except Exception:
        return None
    with tempfile.TemporaryDirectory(prefix="pdf_tool_tbl_") as work:
        path = os.path.join(work, "table.png")
        pixmap.save(path)
        tables = _extract(TableImage, path, reader)
        # Reading Arabic and English at once makes Tesseract guess Arabic
        # digits for small Latin letters (a1 -> 281). A table with no Arabic
        # in it is read again with the English model alone, and the other
        # way round.
        script = _script(tables)
        single = {"latin": "eng", "arabic": "ara"}.get(script)
        if single and single in (getattr(reader, "lang", "") or "").split("+") and getattr(reader, "lang", "") != single:
            again = _reader(single)
            if again is not None:
                tables = _extract(TableImage, path, again) or tables
        if not tables:
            # img2table's own reading can drop an Arabic table entirely. Its
            # grid is found without OCR, and each cell is read on its own.
            tables = _extract(TableImage, path, None)
            if tables:
                _read_cells_one_by_one(path, tables)
    if not tables:
        return None
    # The largest grid is the table; it must fill most of the picture.
    table = max(tables, key=lambda t: (t.bbox.x2 - t.bbox.x1) * (t.bbox.y2 - t.bbox.y1))
    scale = 72.0 / DPI
    covered = ((table.bbox.x2 - table.bbox.x1) * (table.bbox.y2 - table.bbox.y1)) * scale * scale
    if covered < rect.width * rect.height * 0.4:
        return None
    rows_cells = [table.content[key] for key in sorted(table.content)]
    if len(rows_cells) < 2:
        return None
    cols = max(len(row) for row in rows_cells)
    if cols < 2:
        return None
    rows: list[list[str]] = []
    for row in rows_cells:
        texts = []
        for cell in row:
            value = cell.value if cell.value is not None else ""
            texts.append(str(value).strip())
        texts += [""] * (cols - len(texts))
        rows.append(texts)
    # A merged cell repeats in every column it spans; keep it once.
    for row_cells, texts in zip(rows_cells, rows):
        seen: set[tuple] = set()
        for index, cell in enumerate(row_cells):
            key = (cell.bbox.x1, cell.bbox.y1, cell.bbox.x2, cell.bbox.y2)
            if key in seen:
                texts[index] = ""
            seen.add(key)
    if sum(1 for row in rows for text in row if text) < 4:
        return None
    first = rows_cells[0]
    lefts = sorted({cell.bbox.x1 for row in rows_cells for cell in row})
    edges = _edges([cell.bbox.x1 for cell in first] + [first[-1].bbox.x2]) if len(first) == cols else None
    if edges is None or len(edges) != cols + 1:
        edges = _edges(lefts + [table.bbox.x2])
    if len(edges) != cols + 1:
        span = (table.bbox.x2 - table.bbox.x1) / cols
        edges = [table.bbox.x1 + span * index for index in range(cols + 1)]
    widths = [(edges[i + 1] - edges[i]) * scale for i in range(cols)]
    heights = [max(cell.bbox.y2 for cell in row) * scale - min(cell.bbox.y1 for cell in row) * scale for row in rows_cells]
    box = (
        rect.x0 + table.bbox.x1 * scale,
        rect.y0 + table.bbox.y1 * scale,
        rect.x0 + table.bbox.x2 * scale,
        rect.y0 + table.bbox.y2 * scale,
    )
    return ImageTable(rows, widths, heights, box, [f"table read from its picture at {DPI} dpi (img2table + Tesseract)"])


def _tesseract_text(image_path: str, langs: str) -> str:
    import subprocess

    from engine.tesseract_bin import find_exe

    exe = find_exe()
    tessdata, _all = tessdata_for()
    if not exe or not tessdata:
        return ""
    env = dict(os.environ, TESSDATA_PREFIX=tessdata)
    try:
        done = subprocess.run(
            [exe, image_path, "stdout", "-l", langs, "--psm", "6"],
            capture_output=True, text=True, encoding="utf-8", env=env, timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return ""
    return (done.stdout or "").strip()


def _read_cells_one_by_one(path: str, tables: list) -> None:
    """Read every cell of img2table's grids with Tesseract, in the table's own script."""
    import re

    from PIL import Image

    _tessdata, available_langs = tessdata_for()
    have = set(available_langs.split("+"))
    picture = Image.open(path)
    with tempfile.TemporaryDirectory(prefix="pdf_tool_cell_") as work:
        sample = _tesseract_text(path, "+".join(lang for lang in ("ara", "eng") if lang in have))
        arabic = len(re.findall(r"[\u0621-\u064A]", sample))
        latin = len(re.findall(r"[A-Za-z]", sample))
        langs = "ara" if arabic > latin and "ara" in have else ("eng" if latin > arabic * 3 and "eng" in have else available_langs)
        done: dict[tuple, str] = {}
        for table in tables:
            for row in table.content.values():
                for cell in row:
                    key = (cell.bbox.x1, cell.bbox.y1, cell.bbox.x2, cell.bbox.y2)
                    if key not in done:
                        pad = 4
                        crop = picture.crop((max(0, key[0] + pad), max(0, key[1] + pad), key[2] - pad, key[3] - pad))
                        if crop.width < 8 or crop.height < 8:
                            done[key] = ""
                        else:
                            target = os.path.join(work, f"c{len(done)}.png")
                            crop.save(target)
                            done[key] = _tesseract_text(target, langs)
                    try:
                        cell.value = done[key]
                    except Exception:
                        object.__setattr__(cell, "value", done[key])


def _extract(TableImage, path: str, reader) -> list:
    try:
        return TableImage(path).extract_tables(
            ocr=reader,
            implicit_rows=False,
            borderless_tables=False,
            min_confidence=MIN_CONFIDENCE,
        )
    except Exception:
        return []


def _script(tables: list) -> str:
    """latin, arabic or mixed, from the letters a first reading found."""
    import re

    text = " ".join(str(cell.value or "") for table in tables for row in table.content.values() for cell in row)
    arabic = len(re.findall(r"[\u0621-\u064A]", text))
    latin = len(re.findall(r"[A-Za-z]", text))
    if arabic <= max(2, latin * 0.05):
        return "latin"
    if latin <= max(2, arabic * 0.05):
        return "arabic"
    return "mixed"


def _edges(values: list[float]) -> list[float]:
    out: list[float] = []
    for value in sorted(values):
        if not out or value - out[-1] > 4:
            out.append(float(value))
    return out


def tables_on_page(page: fitz.Page, dpi: int = 110) -> list[tuple[float, float, float, float]]:
    """
    Boxes (PDF points) of ruled tables seen in a picture of the whole page.

    For pages whose text and rules are only drawings or pixels (a "Print to
    PDF" page, a scan), where no text layer tells a table finder anything.
    """
    try:
        from img2table.document import Image as TableImage
    except Exception:
        return []
    if _ocr() is None:
        return []
    try:
        pixmap = page.get_pixmap(dpi=dpi, alpha=False)
    except Exception:
        return []
    scale = 72.0 / dpi
    with tempfile.TemporaryDirectory(prefix="pdf_tool_page_") as work:
        path = os.path.join(work, "page.png")
        pixmap.save(path)
        found = _extract(TableImage, path, None)
    boxes = []
    for table in found:
        rows = list(table.content.values())
        if len(rows) < 2 or max((len(row) for row in rows), default=0) < 2:
            continue
        box = (table.bbox.x1 * scale, table.bbox.y1 * scale, table.bbox.x2 * scale, table.bbox.y2 * scale)
        if (box[2] - box[0]) < 60 or (box[3] - box[1]) < 30:
            continue
        boxes.append(box)
    return boxes
