# -*- coding: utf-8 -*-
"""Grayscale page renders, cached, so a review click only slices an array."""

from __future__ import annotations

import hashlib
import io
import threading
from pathlib import Path

import fitz
import numpy as np
from PIL import Image, ImageDraw

PAGE_DPI = 200
_SCALE = PAGE_DPI / 72.0
_BANKS: dict[str, "PageBank"] = {}
_PNG: dict[tuple, bytes] = {}
_LOCK = threading.Lock()


def pdf_key(pdf_bytes: bytes) -> str:
    return hashlib.sha256(pdf_bytes or b"").hexdigest()


def _cache_dir(key: str) -> Path:
    try:
        from engine.paths import page_cache_dir

        folder = page_cache_dir() / key
    except Exception:
        folder = Path(__file__).resolve().parent / ".page_cache" / key
    folder.mkdir(parents=True, exist_ok=True)
    return folder


class PageBank:
    """One PDF, rendered once at 200 DPI. Clicks never open it again."""

    def __init__(self, pdf_bytes: bytes) -> None:
        self.key = pdf_key(pdf_bytes)
        self.pdf = pdf_bytes
        self.pages: dict[int, np.ndarray] = {}
        self.size: dict[int, tuple[float, float]] = {}
        self.ready = False
        self.running = False
        self.error = ""
        self.render_ms = 0.0

    def start(self) -> None:
        if self.ready or self.running:
            return
        self.running = True
        threading.Thread(target=self._render, daemon=True).start()

    def _render(self) -> None:
        import time

        started = time.perf_counter()
        try:
            if not self._load_disk():
                self._render_pdf()
                self._save_disk()
            self.ready = True
        except Exception as exc:
            self.error = str(exc)
        finally:
            self.running = False
            self.render_ms = (time.perf_counter() - started) * 1000.0

    def _load_disk(self) -> bool:
        folder = _cache_dir(self.key)
        meta = folder / "sizes.txt"
        if not meta.is_file():
            return False
        pages: dict[int, np.ndarray] = {}
        size: dict[int, tuple[float, float]] = {}
        for line in meta.read_text(encoding="utf-8").splitlines():
            page_s, width_s, height_s = line.split()
            page = int(page_s)
            path = folder / f"p{page}.png"
            if not path.is_file():
                return False
            image = Image.open(path).convert("L")
            try:
                pages[page] = np.asarray(image).copy()
            finally:
                image.close()
            size[page] = (float(width_s), float(height_s))
        if not pages:
            return False
        self.pages = pages
        self.size = size
        return True

    def _save_disk(self) -> None:
        folder = _cache_dir(self.key)
        lines = []
        for page, array in self.pages.items():
            Image.fromarray(array, "L").save(folder / f"p{page}.png")
            width, height = self.size[page]
            lines.append(f"{page} {width} {height}")
        (folder / "sizes.txt").write_text("\n".join(lines), encoding="utf-8")

    def _render_pdf(self) -> None:
        document = fitz.open(stream=self.pdf, filetype="pdf")
        try:
            matrix = fitz.Matrix(_SCALE, _SCALE)
            for index in range(document.page_count):
                page = document.load_page(index)
                pixmap = page.get_pixmap(matrix=matrix, colorspace=fitz.csGRAY, alpha=False)
                array = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width).copy()
                self.pages[index + 1] = array
                self.size[index + 1] = (float(page.rect.width), float(page.rect.height))
        finally:
            document.close()


def ensure_page_bank(pdf_bytes: bytes) -> PageBank | None:
    """Start the one-time render if this file is not already in memory."""
    if not pdf_bytes:
        return None
    key = pdf_key(pdf_bytes)
    with _LOCK:
        bank = _BANKS.get(key)
        if bank is None:
            bank = PageBank(pdf_bytes)
            _BANKS[key] = bank
            bank.start()
        elif not bank.ready and not bank.running and not bank.pages:
            bank.start()
    return bank


def field_window(field: str, spans: dict, page_width: float) -> tuple[float, float] | None:
    """One cell's x range. A date is the glyphs beside the colour, not the whole left margin."""
    own = spans.get(field) if spans else None
    color = spans.get("color") if spans else None
    make = spans.get("make") if spans else None
    # Stored colour and make spans swallow the year and the admin number.
    # The card box is the word itself.
    if field == "color" and color:
        return (float(color[0]) - 2.0, float(color[0]) + 52.0)
    if field == "make" and make:
        return (float(make[0]) - 4.0, min(float(make[0]) + 68.0, float(make[1])))
    if field == "date" and color:
        return (max(4.0, float(color[0]) - 140.0), float(color[0]) + 14.0)
    if own and field not in {"date", "plate", "year", "admin", "color", "make"}:
        return (float(own[0]), float(own[1]))
    if field == "year" and color and make:
        return (float(color[1]) + 2.0, max(float(color[1]) + 28.0, float(make[0]) - 2.0))
    if field == "plate":
        if make and float(make[1]) < page_width - 24:
            return (min(page_width - 16.0, float(make[1]) + 48.0), min(page_width - 4.0, float(make[1]) + 130.0))
        return (max(4.0, page_width - 110.0), page_width - 6.0)
    if field == "admin" and make:
        return (float(make[1]) + 4.0, min(page_width - 8.0, float(make[1]) + 48.0))
    if own:
        return (float(own[0]), float(own[1]))
    return None


def tighten_window(array: np.ndarray, top: float, bottom: float, guess: tuple[float, float]) -> tuple[float, float]:
    """Snap a guessed cell onto the ink blob nearest that guess. Gaps separate the columns."""
    y0 = max(0, int(top * _SCALE))
    y1 = min(array.shape[0], int(bottom * _SCALE))
    if y1 - y0 < 4:
        return guess
    ink = array[y0:y1] < 150
    left = max(0.0, guess[0] - 36.0)
    right = min(array.shape[1] / _SCALE, guess[1] + 36.0)
    blobs: list[tuple[float, float]] = []
    start: float | None = None
    end = left
    x = left
    while x < right:
        x0 = int(x * _SCALE)
        x1 = min(ink.shape[1], int((x + 2.0) * _SCALE))
        density = float(ink[:, x0:x1].mean()) if x1 > x0 else 0.0
        if density > 0.04:
            if start is None:
                start = x
            end = x + 2.0
        elif start is not None:
            blobs.append((start, end))
            start = None
        x += 2.0
    if start is not None:
        blobs.append((start, end))
    merged: list[list[float]] = []
    for blob in blobs:
        if merged and blob[0] - merged[-1][1] <= 14.0 and blob[1] - merged[-1][0] <= 48.0:
            merged[-1][1] = blob[1]
        else:
            merged.append([blob[0], blob[1]])
    wide = [tuple(blob) for blob in merged if blob[1] - blob[0] >= 14.0]
    if not wide:
        return guess
    center = (guess[0] + guess[1]) / 2.0
    chosen = min(wide, key=lambda blob: abs((blob[0] + blob[1]) / 2.0 - center))
    if chosen[0] < guess[0] - 16.0:
        right_side = [blob for blob in wide if blob[0] >= guess[0] - 8.0]
        if right_side:
            chosen = min(right_side, key=lambda blob: abs((blob[0] + blob[1]) / 2.0 - center))
        else:
            return guess
    return (chosen[0] - 1.5, chosen[1] + 1.5)


def row_band(baselines: dict | None, page: int, y_value: float) -> tuple[float, float] | None:
    """This plate's row, from the midpoint to the plates above and below.

    Neighbours are the other stored baselines on the page, ordered by height.
    A header or separator in the row list is not a neighbour.
    """
    if not baselines or y_value <= 0:
        return None
    heights = sorted(
        {
            round(float(value), 2)
            for key, value in baselines.items()
            if str(key).startswith(f"{int(page)}|") and float(value) > 0
        }
    )
    if not heights:
        return None
    nearest = min(heights, key=lambda item: abs(item - float(y_value)))
    if abs(nearest - float(y_value)) > 6:
        return None
    index = heights.index(nearest)
    gaps = [right - left for left, right in zip(heights, heights[1:]) if 8.0 <= right - left <= 28.0]
    pitch = sorted(gaps)[len(gaps) // 2] if gaps else 13.0
    above = heights[index - 1] if index else None
    below = heights[index + 1] if index + 1 < len(heights) else None
    if above is not None and not 8.0 <= nearest - above <= 28.0:
        above = None
    if below is not None and not 8.0 <= below - nearest <= 28.0:
        below = None
    # The plate word fills its own row. Stop before the midpoint, so the box
    # cannot take the next plate's letters.
    rise = (nearest - above) * 0.46 if above is not None else pitch * 0.48
    drop = (below - nearest) * 0.40 if below is not None else pitch * 0.40
    top = nearest - min(rise, pitch * 0.48)
    bottom = nearest + min(drop, pitch * 0.40)
    if bottom - top < 7:
        top, bottom = nearest - 4.0, nearest + 3.0
    return (top, bottom)


def field_band(field: str, y_value: float) -> tuple[float, float]:
    """Where that column's ink sits relative to the plate baseline."""
    if field == "date":
        return (y_value - 9.2, y_value + 0.8)
    if field == "color":
        return (y_value - 1.0, y_value + 12.0)
    if field == "make":
        return (y_value - 1.2, y_value + 6.8)
    return (y_value - 4.0, y_value + 5.0)


def _pitch(baselines: dict, page: int, y_value: float) -> float:
    values = sorted(
        float(value)
        for key, value in (baselines or {}).items()
        if str(key).startswith(f"{int(page)}|") and float(value) > 0
    )
    gaps = [right - left for left, right in zip(values, values[1:]) if 8.0 <= right - left <= 28.0]
    if gaps:
        return sorted(gaps)[len(gaps) // 2]
    return 13.0


def card_png(
    bank: PageBank,
    page: int,
    y_value: float,
    field: str,
    spans: dict,
    baselines: dict | None = None,
) -> bytes:
    """
    Plate column and the target cell, one row above and below dimmed.

    The red box is that cell on the centre row. The source is the cached
    grayscale render, then scaled up so the letters stay readable.
    """
    if bank is None or not bank.ready or page not in bank.pages or y_value <= 0:
        return b""
    cached = _PNG.get((bank.key, int(page), round(float(y_value), 2), field))
    if cached is not None:
        return cached
    array = bank.pages[int(page)]
    page_width, page_height = bank.size.get(int(page), (array.shape[1] / _SCALE, array.shape[0] / _SCALE))
    plate = field_window("plate", spans or {}, page_width)
    cell = field_window(field, spans or {}, page_width) or plate
    owned = row_band(baselines, page, y_value) or field_band(field, y_value)
    band_top, band_bottom = owned
    if cell is not None and field in {"color", "make"}:
        cell = tighten_window(array, band_top, band_bottom, cell)
    if plate is None and cell is None:
        return b""
    if plate is None:
        plate = cell
    if cell is None:
        cell = plate
    pitch = _pitch(baselines or {}, page, y_value)
    top = max(0.0, band_top - pitch)
    bottom = min(page_height, band_bottom + pitch)
    left = max(0.0, min(plate[0], cell[0]) - 4.0)
    right = min(page_width, max(plate[1], cell[1]) + 4.0)
    if right - left < 24:
        right = min(page_width, left + 24)
    y0 = max(0, int(top * _SCALE))
    y1 = min(array.shape[0], int(bottom * _SCALE))
    x0 = max(0, int(left * _SCALE))
    x1 = min(array.shape[1], int(right * _SCALE))
    if y1 - y0 < 8 or x1 - x0 < 8:
        return b""
    chip = np.array(array[y0:y1, x0:x1])
    center0 = int(max(0.0, band_top - top) * _SCALE)
    center1 = int(min(bottom - top, band_bottom - top) * _SCALE)
    center0 = max(0, min(chip.shape[0] - 1, center0))
    center1 = max(center0 + 4, min(chip.shape[0], center1))
    rgb = np.repeat(chip[:, :, None], 3, axis=2)
    dim = (rgb.astype("float32") * 0.42 + 255.0 * 0.58).astype("uint8")
    if center0 > 0:
        rgb[:center0] = dim[:center0]
    if center1 < rgb.shape[0]:
        rgb[center1:] = dim[center1:]
    image = Image.fromarray(rgb, "RGB")
    grow = 3.0
    image = image.resize((max(1, int(image.width * grow)), max(1, int(image.height * grow))), Image.Resampling.LANCZOS)
    if image.width > 2000:
        grow = 2000.0 / max(1, (x1 - x0))
        image = image.resize((2000, max(1, int((y1 - y0) * grow))), Image.Resampling.LANCZOS)
    draw = ImageDraw.Draw(image)
    box_left = (max(cell[0], left) - left) * _SCALE * grow
    box_right = (min(cell[1], right) - left) * _SCALE * grow
    box_top = center0 * grow
    box_bottom = center1 * grow
    draw.rectangle(
        [box_left, box_top, box_right, box_bottom],
        outline=(176, 32, 32),
        width=max(3, int(round(grow))),
    )
    output = io.BytesIO()
    image.save(output, format="PNG")
    payload = output.getvalue()
    _PNG[(bank.key, int(page), round(float(y_value), 2), field)] = payload
    if len(_PNG) > 48:
        _PNG.pop(next(iter(_PNG)))
    return payload


def describe_pages(pdf_bytes: bytes) -> list[dict]:
    """Each page is a text layer or a scan. This does not run OCR."""
    from core.excel_converter import _text_char_count, text_layer_is_garbage

    if not pdf_bytes:
        return []
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    found = []
    try:
        for index in range(document.page_count):
            page = document.load_page(index)
            if text_layer_is_garbage(page):
                kind, weight = "نص غير مقروء", 4
            elif _text_char_count(page) >= 20:
                kind, weight = "طبقة نص", 1
            else:
                kind, weight = "مسح", 3
            found.append({"page": index + 1, "kind": kind, "weight": weight})
    finally:
        document.close()
    return found


def difficulty_score(pages: list[dict], columns_missing: bool) -> int:
    score = sum(int(item.get("weight") or 0) for item in pages)
    if columns_missing:
        score += 5
    return score


def cell_png(
    bank: PageBank,
    page: int,
    y_value: float,
    field: str,
    spans: dict,
    baselines: dict | None = None,
    scale: int = 3,
) -> bytes:
    """The target cell only, grayscale, scaled up. Vision uses 3. The cluster sheet uses 4."""
    if bank is None or not bank.ready or page not in bank.pages or y_value <= 0:
        return b""
    array = bank.pages[int(page)]
    page_width, page_height = bank.size.get(int(page), (array.shape[1] / _SCALE, array.shape[0] / _SCALE))
    window = field_window(field, spans or {}, page_width)
    if window is None:
        return b""
    band_top, band_bottom = row_band(baselines, page, y_value) or field_band(field, y_value)
    if field in {"color", "make"}:
        window = tighten_window(array, band_top, band_bottom, window)
    top = max(0.0, band_top)
    bottom = min(page_height, band_bottom)
    y0 = max(0, int(top * _SCALE))
    y1 = min(array.shape[0], int(bottom * _SCALE))
    x0 = max(0, int(window[0] * _SCALE))
    x1 = min(array.shape[1], int(window[1] * _SCALE))
    if y1 - y0 < 4 or x1 - x0 < 4:
        return b""
    chip = array[y0:y1, x0:x1]
    image = Image.fromarray(chip, "L").convert("RGB")
    grow = max(1, int(scale))
    image = image.resize((image.width * grow, image.height * grow), Image.Resampling.LANCZOS)
    output = io.BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


def cell_ink(bank: PageBank, page: int, y_value: float, field: str, spans: dict, baselines: dict | None = None) -> np.ndarray | None:
    """A small ink map for the learned library, cut from the cached page."""
    if bank is None or not bank.ready or page not in bank.pages:
        return None
    array = bank.pages[int(page)]
    page_width = bank.size.get(int(page), (array.shape[1] / _SCALE, 0))[0]
    window = field_window(field, spans or {}, page_width)
    if window is None:
        return None
    band_top, band_bottom = row_band(baselines, page, y_value) or field_band(field, y_value)
    if field in {"color", "make"}:
        window = tighten_window(array, band_top, band_bottom, window)
    top = max(0, int(band_top * _SCALE))
    bottom = min(array.shape[0], int(band_bottom * _SCALE))
    left = max(0, int(window[0] * _SCALE))
    right = min(array.shape[1], int(window[1] * _SCALE))
    if bottom - top < 4 or right - left < 4:
        return None
    chip = array[top:bottom, left:right]
    ink = chip < 160
    if float(ink.mean()) < 0.02:
        return np.zeros((16, 32), dtype="float32")
    ys = np.linspace(0, chip.shape[0] - 1, 16).astype(int)
    xs = np.linspace(0, chip.shape[1] - 1, 32).astype(int)
    return (chip[ys][:, xs] < 160).astype("float32")
