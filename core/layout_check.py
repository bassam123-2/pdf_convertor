# -*- coding: utf-8 -*-
"""Compare a reconstructed page with the PDF page it came from.

The check is structural. It asks where the ink sits (top, middle, bottom,
left, center, right), not whether the letters match. A reading that piles
every line into the top of the sheet fails even when the words are complete.
"""

from __future__ import annotations

import fitz

# 0 is the same structure. Above this, the page is read again.
DISCREPANCY_LIMIT = 0.34
_ROWS = 8
_COLS = 3


def page_is_ruled(page: fitz.Page) -> bool:
    """True when the page is crossed by several long horizontal rules."""
    try:
        drawings = page.get_drawings()
    except Exception:
        return False
    rules = 0
    width = float(page.rect.width) or 1.0
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is not None:
            if float(rect.width) >= width * 0.4 and float(rect.height) <= 2.5:
                rules += 1
        for item in drawing.get("items") or []:
            if not item or item[0] != "l" or len(item) < 3:
                continue
            start, end = item[1], item[2]
            try:
                if abs(float(start.y) - float(end.y)) <= 1.5 and abs(float(start.x) - float(end.x)) >= width * 0.35:
                    rules += 1
            except (AttributeError, TypeError, ValueError):
                continue
    return rules >= 4


def _distribution(weights: list[list[float]]) -> list[list[float]]:
    total = sum(value for row in weights for value in row)
    if total <= 0:
        return [[0.0 for _ in row] for row in weights]
    return [[value / total for value in row] for row in weights]


def _empty(rows: int, cols: int) -> list[list[float]]:
    return [[0.0 for _ in range(cols)] for _ in range(rows)]


def _density_from_pixmap(pixmap: fitz.Pixmap) -> list[list[float]]:
    """
    Which bands contain ink.

    A band is on or off. A solid reconstructed box and a line of glyphs then
    agree whenever they sit in the same part of the page.
    """
    grid = _empty(_ROWS, _COLS)
    width = pixmap.width or 1
    height = pixmap.height or 1
    samples = pixmap.samples
    stride = pixmap.stride or width
    for y in range(height):
        page_y = y / height
        if page_y < 0.08 or page_y > 0.92:
            continue
        row_index = min(_ROWS - 1, int((page_y - 0.08) / 0.84 * _ROWS))
        start = y * stride
        for x in range(width):
            if samples[start + x] >= 242:
                continue
            column = min(_COLS - 1, int(x / width * _COLS))
            grid[row_index][column] = 1.0
    return grid


def page_density(page: fitz.Page) -> list[list[float]]:
    """Where the PDF's ink sits, from a small grayscale render of the page."""
    try:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(0.25, 0.25), colorspace=fitz.csGRAY, alpha=False)
    except Exception:
        return _distribution(_empty(_ROWS, _COLS))
    return _density_from_pixmap(pixmap)


def render_layout(page: fitz.Page, boxes: list[tuple[float, float, float, float]]) -> list[list[float]]:
    """
    Render the reconstructed boxes to a temporary page image.

    The image is not written to disk. It is the layout Word will produce,
    so it can be compared with the PDF render band by band.
    """
    blank = _distribution(_empty(_ROWS, _COLS))
    if not boxes:
        return blank
    scratch = fitz.open()
    try:
        rendered = scratch.new_page(width=float(page.rect.width), height=float(page.rect.height))
        for box in boxes:
            rect = fitz.Rect(*box)
            if rect.is_empty or rect.is_infinite:
                continue
            rendered.draw_rect(rect, color=None, fill=(0, 0, 0))
        pixmap = rendered.get_pixmap(matrix=fitz.Matrix(0.25, 0.25), colorspace=fitz.csGRAY, alpha=False)
    except Exception:
        return blank
    finally:
        scratch.close()
    return _density_from_pixmap(pixmap)


def layout_discrepancy(
    page: fitz.Page,
    boxes: list[tuple[float, float, float, float]],
    kept_table: bool,
) -> float:
    """
    How far the reconstruction sits from the PDF page.

    0 means the ink and the boxes occupy the same parts of the sheet.
    A ruled page that came back with no table is treated as a failed layout.
    """
    source = page_density(page)
    rebuilt = render_layout(page, boxes)
    distance = 0.0
    for source_row, rebuilt_row in zip(source, rebuilt):
        distance += sum(abs(left - right) for left, right in zip(source_row, rebuilt_row))
    score = distance / max(1, _ROWS * _COLS)
    if page_is_ruled(page) and not kept_table:
        score = min(1.0, score + 0.5)
    return score
