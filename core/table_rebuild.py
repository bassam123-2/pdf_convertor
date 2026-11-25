# -*- coding: utf-8 -*-
"""
Classify PDF tables and charts, and rebuild raster tables as native Word grids.

A table is either:
  - vector/text: ruled or aligned text spans inside a bbox (editable source)
  - raster: an embedded picture of a table (must be OCR'd cell by cell)

A chart/figure is a picture region whose labels must stay inside one rendered
picture; those labels must never become free-floating text atoms.
"""

from __future__ import annotations

import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

import fitz

# Cells below this OCR score (0–100) are flagged; text is still kept.
_LOW_CONFIDENCE = 55.0


@dataclass
class CellHit:
    text: str
    confidence: float
    low_confidence: bool = False


@dataclass
class ImageRegion:
    """One embedded picture (or ruled region) on a page, classified."""

    kind: str  # "raster_table" | "vector_table" | "chart" | "other"
    bbox: tuple[float, float, float, float]
    reason: str
    xref: int | None = None
    image_bytes: bytes = b""


@dataclass
class RebuiltTable:
    """A native table ready for faithful Word output."""

    kind: str  # "vector" | "raster"
    bbox: tuple[float, float, float, float]
    rows: list[list[str]] = field(default_factory=list)
    fills: list[list[str]] = field(default_factory=list)
    widths: list[float] = field(default_factory=list)
    heights: list[float] = field(default_factory=list)
    confidences: list[list[float]] = field(default_factory=list)
    rules: bool = True
    rtl: bool = False
    header_repeat: bool = False
    # Original picture kept as a silent fallback (never shown unless rebuild fails).
    fallback_png: bytes = b""
    notes: list[str] = field(default_factory=list)
    low_confidence_cells: list[tuple[int, int, float]] = field(default_factory=list)
    # When False, Word keeps the picture visible and the native grid is withheld
    # (still reported) so a weak OCR pass never silently replaces a clear plate.
    usable: bool = True


def table_columns_too_narrow(table: RebuiltTable) -> bool:
    """
    True when a column is too skinny for its longest cell.

    That is the crushed-date / vertical-letter failure mode on revision sheets:
    Word wraps one character per line inside a 10–20 pt column.
    """
    rows = table.rows or []
    widths = table.widths or []
    if not rows or not widths:
        return False
    cols = min(len(widths), max(len(row) for row in rows))
    for ci in range(cols):
        width = float(widths[ci] or 0.0)
        longest = 0
        for row in rows:
            if ci >= len(row):
                continue
            # Multi-line cells (stacked dates) are judged by the longest line,
            # not the joined length — wrapping is intentional there.
            cell = (row[ci] or "").strip()
            if not cell:
                continue
            line_len = max(len(line.strip()) for line in cell.splitlines()) or len(cell)
            longest = max(longest, line_len)
        if longest <= 0:
            continue
        # ~4.5 pt per Latin glyph at 10–11 pt; under ~3 pt/char text stacks vertically.
        if width > 0 and width < max(28.0, longest * 3.0) and longest >= 6:
            return True
        if width < 22.0 and longest >= 4:
            return True
    return False


def table_is_usable(table: RebuiltTable) -> bool:
    """Reject OCR grids that are mostly empty or mostly low-confidence."""
    rows = table.rows or []
    if not rows:
        return False
    cols = max(len(row) for row in rows)
    # A geometric fallback that shards prose into dozens of columns is not a table.
    if cols >= 8 and len(rows) >= 12:
        return False
    if table_columns_too_narrow(table):
        return False
    if table.kind == "vector":
        # Real data tables have short header-like cells; prose frames do not.
        long_cells = sum(1 for row in rows for cell in row if len((cell or "").strip()) > 80)
        if long_cells >= max(3, len(rows) // 2) and cols <= 3:
            return False
        return True
    filled = [(ri, ci, text) for ri, row in enumerate(rows) for ci, text in enumerate(row) if (text or "").strip()]
    if len(filled) < 4:
        return False
    low = len(table.low_confidence_cells)
    if low / max(len(filled), 1) > 0.45:
        return False
    # A header row that is entirely empty usually means the grid was mis-cut.
    if rows and not any((cell or "").strip() for cell in rows[0]):
        return False
    return True


def clip_region_png(page: fitz.Page, bbox: tuple[float, float, float, float], zoom: float = 2.0) -> bytes:
    """PNG of one page rectangle — used when a rebuilt table is not usable."""
    rect = fitz.Rect(*bbox)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=rect, alpha=False)
    return pix.tobytes("png")


def _is_prose_column_frame(page: fitz.Page, table) -> bool:
    """
    True when find_tables matched a bilingual/two-column text frame, not a data grid.

    Typical hit: 1–2 rows × 2 columns, each cell a long multi-line paragraph,
    covering most of the page. Rebuilding that as w:tbl clips Arabic past the
    column border.
    """
    try:
        rows = int(getattr(table, "row_count", 0) or 0)
        cols = int(getattr(table, "col_count", 0) or 0)
    except Exception:
        rows = cols = 0
    bbox = tuple(float(v) for v in table.bbox)
    page_area = max(1.0, float(page.rect.width) * float(page.rect.height))
    area = max(1.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
    # Two-column page frames: few rows, two columns, large area.
    if cols <= 2 and rows <= 3 and area >= page_area * 0.35:
        return True
    try:
        extracted = table.extract() or []
    except Exception:
        extracted = []
    if not extracted:
        return cols <= 2 and area >= page_area * 0.35
    cells = [((cell or "").strip()) for row in extracted for cell in (row or [])]
    if not cells:
        return False
    multiline = sum(1 for cell in cells if cell.count("\n") >= 3 or len(cell) > 180)
    # Most cells are long prose → column frame / callout, not a data table.
    if multiline >= max(1, len([c for c in cells if c]) - 0) and cols <= 3 and rows <= 4:
        return True
    if multiline >= 2 and area >= page_area * 0.4:
        return True
    return False


def classify_regions(page: fitz.Page) -> list[ImageRegion]:
    """
    Decide what each notable region on the page is.

    Uses page content only: text spans inside a bbox vs an embedded image,
    grid structure in the pixels, and nearby captions. No file-name rules.
    """
    regions: list[ImageRegion] = []
    page_area = max(1.0, float(page.rect.width) * float(page.rect.height))
    text = page.get_text("text") or ""
    atoms = _text_boxes(page)

    # Vector / text tables from PyMuPDF.
    try:
        found = page.find_tables(vertical_strategy="lines", horizontal_strategy="lines")
        tables = list(getattr(found, "tables", None) or [])
    except Exception:
        tables = []
    for table in tables:
        bbox = tuple(float(v) for v in table.bbox)
        inside = sum(1 for box in atoms if _center_inside(box, bbox))
        if inside < 4 or (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]) <= page_area * 0.02:
            continue
        if _is_prose_column_frame(page, table):
            # Leave as flowing/faithful text; do not rebuild as w:tbl.
            continue
        regions.append(
            ImageRegion(
                "vector_table",
                bbox,
                f"ruled grid with {inside} text spans inside the bbox",
            )
        )

    # Embedded images: table picture vs chart vs other.
    try:
        infos = page.get_image_info(xrefs=True) or []
    except Exception:
        infos = []
    for info in infos:
        box = info.get("bbox")
        if not box:
            continue
        bbox = (float(box[0]), float(box[1]), float(box[2]), float(box[3]))
        area = max(1.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
        if area < page_area * 0.03:
            continue
        # Skip if already covered by a vector table.
        if any(r.kind == "vector_table" and _overlap_ratio(bbox, r.bbox) > 0.5 for r in regions):
            continue
        xref = int(info.get("xref") or 0) or None
        png = _extract_png(page, xref, bbox)
        spans_inside = sum(1 for box in atoms if _center_inside(box, bbox))
        caption = _nearby_caption(text, atoms, bbox)
        looks_table, table_why = _pixels_look_like_table(png)
        if spans_inside >= 6 and not looks_table:
            # Text sits on top of a figure; treat as chart/figure region.
            regions.append(
                ImageRegion("chart", bbox, f"picture with {spans_inside} overlapping text labels", xref, png)
            )
            continue
        if looks_table or (caption and re.search(r"\btable\b|جدول", caption, re.I)):
            reason = table_why if looks_table else f"caption near region: {caption[:40]!r}"
            regions.append(ImageRegion("raster_table", bbox, reason, xref, png))
            continue
        # Continuation plates: dark header and/or regular row rules on a mostly
        # light body, with almost no text layer of their own.
        header_only, header_why = _pixels_dark_header(png)
        ruled, ruled_why = _pixels_ruled_body(png)
        if spans_inside <= 2 and area >= page_area * 0.08 and (header_only or ruled):
            reason = header_why or ruled_why
            regions.append(
                ImageRegion(
                    "raster_table",
                    bbox,
                    reason + " (likely table plate / continuation)",
                    xref,
                    png,
                )
            )
            continue
        # Large image with almost no page text → chart/figure plate.
        if area >= page_area * 0.12 and spans_inside <= 2:
            regions.append(
                ImageRegion(
                    "chart",
                    bbox,
                    "large embedded picture with little or no text layer inside it",
                    xref,
                    png,
                )
            )
        else:
            regions.append(ImageRegion("other", bbox, "embedded picture", xref, png))
    return regions


def rebuild_tables(page: fitz.Page, regions: list[ImageRegion] | None = None) -> list[RebuiltTable]:
    """Build native tables for vector and raster regions on this page."""
    regions = regions if regions is not None else classify_regions(page)
    rebuilt: list[RebuiltTable] = []
    for region in regions:
        if region.kind == "vector_table":
            item = _rebuild_vector(page, region)
            if item is not None:
                rebuilt.append(item)
        elif region.kind == "raster_table":
            item = _rebuild_raster(page, region)
            if item is not None:
                rebuilt.append(item)
    return rebuilt


def chart_boxes(regions: list[ImageRegion]) -> list[tuple[float, float, float, float]]:
    return [r.bbox for r in regions if r.kind == "chart"]


def drop_lines_in_boxes(lines: list, boxes: list[tuple[float, float, float, float]]) -> list:
    """Remove text lines whose centers fall inside a chart/figure region."""
    if not boxes:
        return lines
    kept = []
    for line in lines:
        cx = (line.x0 + line.x1) / 2.0
        cy = (line.y0 + line.y1) / 2.0
        if any(box[0] <= cx <= box[2] and box[1] <= cy <= box[3] for box in boxes):
            continue
        kept.append(line)
    return kept


def paint_out_regions(png: bytes, page_width: float, page_height: float, boxes: list[tuple], zoom: float) -> bytes:
    """White out table regions in the background picture so native tables sit cleanly."""
    if not png or not boxes:
        return png
    try:
        from PIL import Image, ImageDraw
    except Exception:
        return png
    try:
        image = Image.open(BytesIO(png)).convert("RGB")
    except Exception:
        return png
    draw = ImageDraw.Draw(image)
    for box in boxes:
        left = max(0, int(box[0] * zoom) - 2)
        top = max(0, int(box[1] * zoom) - 2)
        right = min(image.width, int(box[2] * zoom) + 2)
        bottom = min(image.height, int(box[3] * zoom) + 2)
        if right > left and bottom > top:
            draw.rectangle([left, top, right, bottom], fill=(255, 255, 255))
    out = BytesIO()
    image.save(out, format="PNG", optimize=True)
    return out.getvalue()


def mark_continued_headers(tables_by_page: dict[int, list[RebuiltTable]]) -> None:
    """When consecutive pages share the same header row, ask Word to repeat it."""
    numbers = sorted(tables_by_page)
    for prev, curr in zip(numbers, numbers[1:]):
        if curr != prev + 1:
            continue
        for later in tables_by_page[curr]:
            for earlier in tables_by_page[prev]:
                if not earlier.rows or not later.rows:
                    continue
                if _row_key(earlier.rows[0]) == _row_key(later.rows[0]) and len(later.rows[0]) >= 2:
                    later.header_repeat = True


# --------------------------------------------------------------------------
# Internals
# --------------------------------------------------------------------------


def _row_key(row: list[str]) -> str:
    return "|".join(re.sub(r"\s+", " ", (cell or "").strip().lower()) for cell in row)


def _center_inside(box: tuple[float, float, float, float], outer: tuple[float, float, float, float]) -> bool:
    cx = (box[0] + box[2]) / 2.0
    cy = (box[1] + box[3]) / 2.0
    return outer[0] <= cx <= outer[2] and outer[1] <= cy <= outer[3]


def _overlap_ratio(a: tuple, b: tuple) -> float:
    x0 = max(a[0], b[0])
    y0 = max(a[1], b[1])
    x1 = min(a[2], b[2])
    y1 = min(a[3], b[3])
    if x1 <= x0 or y1 <= y0:
        return 0.0
    inter = (x1 - x0) * (y1 - y0)
    area = max(1.0, (a[2] - a[0]) * (a[3] - a[1]))
    return inter / area


def _text_boxes(page: fitz.Page) -> list[tuple[float, float, float, float]]:
    boxes = []
    try:
        for block in page.get_text("dict").get("blocks") or []:
            if block.get("type") != 0:
                continue
            for line in block.get("lines") or []:
                for span in line.get("spans") or []:
                    if (span.get("text") or "").strip():
                        boxes.append(tuple(float(v) for v in span["bbox"]))
    except Exception:
        pass
    return boxes


def _nearby_caption(text: str, atoms: list[tuple], bbox: tuple) -> str:
    """Caption sitting just above a region (e.g. 'Table 2.1 …')."""
    for box in atoms:
        if box[3] > bbox[1] or bbox[1] - box[3] > 36:
            continue
        if box[2] < bbox[0] - 20 or box[0] > bbox[2] + 20:
            continue
    # Fall back to raw page text search near the region is hard; use whole-page cues.
    match = re.search(r"(Table\s+\d+(?:\.\d+)?[^\n]{0,80}|جدول[^\n]{0,80})", text, re.I)
    return match.group(1).strip() if match else ""


def _extract_png(page: fitz.Page, xref: int | None, bbox: tuple) -> bytes:
    if xref:
        try:
            pix = fitz.Pixmap(page.parent, xref)
            if pix.n >= 5:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            return pix.tobytes("png")
        except Exception:
            pass
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), clip=fitz.Rect(*bbox), alpha=False)
        return pix.tobytes("png")
    except Exception:
        return b""


def _pixels_dark_header(png: bytes) -> tuple[bool, str]:
    """
    True when a solid dark header fill sits near the top of the picture.

    Raster tables often have white padding above the header, so the check
    scans the top fifth for a low-variance dark horizontal strip rather than
    averaging from y=0.
    """
    if not png:
        return False, ""
    try:
        from PIL import Image
        import numpy as np
    except Exception:
        return False, ""
    try:
        image = Image.open(BytesIO(png)).convert("RGB")
    except Exception:
        return False, ""
    width, height = image.size
    if width < 80 or height < 40:
        return False, ""
    arr = np.asarray(image)
    gray = arr.mean(axis=2)
    limit = max(20, height // 5)
    window = max(6, min(height // 12, 60))
    best = None
    for top in range(0, max(1, limit - window + 1), max(1, window // 2)):
        # Measure the middle of the strip so white page margins do not inflate std.
        mid_band = arr[top : top + window, width // 8 : 7 * width // 8, :]
        mean = mid_band.reshape(-1, 3).mean(axis=0)
        std = float(mid_band.reshape(-1, 3).std(axis=0).mean())
        luma = float(mean[0] * 0.30 + mean[1] * 0.59 + mean[2] * 0.11)
        mid = float(gray[top : top + window, width // 4 : 3 * width // 4].mean())
        if luma < 120 and mid < 120 and std < 70:
            best = (luma, std, top)
            break
    if best is None:
        return False, ""
    return True, f"dark header band (luma≈{best[0]:.0f} at y≈{best[2]})"

def _pixels_ruled_body(png: bytes) -> tuple[bool, str]:
    """
    True when the picture looks like a light table body with many row rules.

    Used for continuation plates that omit the dark header (or crop it off).
    Real charts usually fail the light-body + many-row-rules combination.
    """
    if not png:
        return False, ""
    try:
        from PIL import Image
        import numpy as np
    except Exception:
        return False, ""
    try:
        image = Image.open(BytesIO(png)).convert("RGB")
    except Exception:
        return False, ""
    width, height = image.size
    if width < 80 or height < 60:
        return False, ""
    arr = np.asarray(image)
    gray = arr.mean(axis=2)
    body = gray[height // 10 :, :]
    light = float((body > 200).mean())
    row_mean = gray.mean(axis=1)
    dips = [
        i
        for i in range(2, len(row_mean) - 2)
        if row_mean[i] < row_mean[i - 1] - 12 and row_mean[i] < row_mean[i + 1] - 6
    ]
    # Merge neighbouring dips into one rule.
    rules: list[int] = []
    for dip in dips:
        if not rules or dip - rules[-1] > 6:
            rules.append(dip)
    if light >= 0.55 and len(rules) >= 3:
        return True, f"light ruled body (light≈{light:.2f}, rules≈{len(rules)})"
    return False, ""


def _pixels_look_like_table(png: bytes) -> tuple[bool, str]:
    """Heuristic: dark header band and/or regular row rules on a light body."""
    if not png:
        return False, ""
    try:
        from PIL import Image
        import numpy as np
    except Exception:
        return False, ""
    try:
        image = Image.open(BytesIO(png)).convert("RGB")
    except Exception:
        return False, ""
    width, height = image.size
    if width < 80 or height < 60:
        return False, ""
    arr = np.asarray(image)
    dark, dark_why = _pixels_dark_header(png)
    ruled, ruled_why = _pixels_ruled_body(png)
    gray = arr.mean(axis=2)
    row_mean = gray.mean(axis=1)
    dips = 0
    for i in range(2, len(row_mean) - 2):
        if row_mean[i] < row_mean[i - 1] - 18 and row_mean[i] < row_mean[i + 1] - 8:
            dips += 1
    col_mean = gray.mean(axis=0)
    gutters = 0
    for i in range(2, len(col_mean) - 2):
        if col_mean[i] > col_mean[i - 1] + 12 and col_mean[i] > col_mean[i + 1] + 12:
            gutters += 1
    if dark and dips >= 3 and gutters >= 2:
        return True, f"{dark_why} and grid-like separators (rows≈{dips}, cols≈{gutters})"
    if dips >= 6 and gutters >= 3:
        return True, f"regular grid separators (rows≈{dips}, cols≈{gutters})"
    if dark and dips >= 2 and gutters >= 2:
        return True, f"{dark_why} with row/column separators (rows≈{dips}, cols≈{gutters})"
    if dark and ruled:
        return True, f"{dark_why}; {ruled_why}"
    if ruled and dips >= 4:
        return True, ruled_why
    return False, ""


def _rebuild_vector(page: fitz.Page, region: ImageRegion) -> RebuiltTable | None:
    try:
        found = page.find_tables(vertical_strategy="lines", horizontal_strategy="lines")
        tables = list(getattr(found, "tables", None) or [])
    except Exception:
        tables = []
    best = None
    best_overlap = 0.0
    for table in tables:
        bbox = tuple(float(v) for v in table.bbox)
        ratio = _overlap_ratio(bbox, region.bbox)
        if ratio > best_overlap:
            best_overlap = ratio
            best = table
    if best is not None and best_overlap >= 0.4 and _is_prose_column_frame(page, best):
        return None
    if best is None or best_overlap < 0.4:
        return _geometric_vector_fallback(page, region)
    try:
        raw = best.extract()
    except Exception:
        return _geometric_vector_fallback(page, region)
    rows = [[(cell or "").strip() for cell in row] for row in (raw or [])]
    rows = [row for row in rows if any(cell for cell in row)]
    if len(rows) < 2:
        return _geometric_vector_fallback(page, region)
    # One fat row of prose per column is a frame, not a data table.
    if len(rows) <= 2 and max(len(row) for row in rows) <= 3:
        longish = sum(1 for row in rows for cell in row if len(cell) > 120 or cell.count("\n") >= 2)
        if longish >= 1:
            return None
    cols = max(len(row) for row in rows)
    rows = [row + [""] * (cols - len(row)) for row in rows]
    heights = _heights_from_finder(best, region.bbox, len(rows)) or (
        [(region.bbox[3] - region.bbox[1]) / max(len(rows), 1)] * len(rows)
    )
    # Double-line borders invent extra columns; merge by clustered left edges.
    remapped = _cluster_double_border_columns(rows, best, region.bbox)
    if remapped is not None:
        rows, widths = remapped
    else:
        widths = _widths_from_finder(best, region.bbox, cols) or _column_widths(region.bbox, cols)
        rows, widths, heights, _ = _collapse_gutter_columns(rows, widths, heights)
    cols = max((len(row) for row in rows), default=0)
    if cols < 2 or len(rows) < 2:
        return _geometric_vector_fallback(page, region)
    bbox, widths, heights = _snap_to_rules(page, region.bbox, widths, heights)
    rows = _reread_cells(page, bbox, rows, widths, heights)
    fills = _sample_fills(page, bbox, rows, widths, heights)
    rebuilt = RebuiltTable(
        kind="vector",
        bbox=bbox,
        rows=rows,
        fills=fills,
        widths=widths,
        heights=heights,
        confidences=[[100.0] * cols for _ in rows],
        rules=True,
        rtl=_page_rtl(page),
        notes=[f"vector table: {region.reason}"],
    )
    rebuilt.usable = table_is_usable(rebuilt)
    if not rebuilt.usable:
        try:
            rebuilt.fallback_png = clip_region_png(page, region.bbox)
            rebuilt.notes.append(
                "Table rebuild not usable (narrow columns or weak grid); "
                "original region kept as a picture."
            )
        except Exception:
            return None
    return rebuilt


def _reread_cells(page: fitz.Page, bbox: tuple, rows: list[list[str]], widths: list[float], heights: list[float]) -> list[list[str]]:
    """
    Read every cell again from the page's glyphs with the paragraph reader.

    The table finder's own text skips what the page reader repairs: Arabic
    presentation forms stay unfolded and in visual order, fake-bold copies
    stay doubled and lam-alef glyphs with a lost code stay foreign. The
    glyphs inside each cell box are read the way a paragraph is. A cell
    keeps the finder's text only when the reader finds nothing in it.
    """
    cols = max((len(row) for row in rows), default=0)
    if not rows or len(widths) != cols or len(heights) != len(rows):
        return rows
    try:
        from core.docx_converter import _atoms_in, _atoms_on_page, _cluster_baselines, _logical_text, _separate_overlapping_atoms
        from core.arabic_fixer import plain_arabic
        from core.faithful import _dedupe_fake_bold

        atoms = _separate_overlapping_atoms(_dedupe_fake_bold(_atoms_on_page(page, upright_only=True)))
    except Exception:
        return rows
    x_edges = [float(bbox[0])]
    for width in widths:
        x_edges.append(x_edges[-1] + width)
    y_edges = [float(bbox[1])]
    for height in heights:
        y_edges.append(y_edges[-1] + height)
    fresh = []
    for r, row in enumerate(rows):
        cells = []
        for c in range(cols):
            old = row[c] if c < len(row) else ""
            box = (x_edges[c], y_edges[r], x_edges[c + 1], y_edges[r + 1])
            # One line of the cell per printed line: the line breaks say how
            # wide the column must be, and Word keeps them.
            lines = [plain_arabic(_logical_text(band)).strip() for band in _cluster_baselines(_atoms_in(atoms, box))]
            text = "\n".join(line for line in lines if line)
            cells.append(text or old)
        fresh.append(cells)
    return fresh


def _ruling_lines(page: fitz.Page) -> tuple[list[tuple[float, float, float]], list[tuple[float, float, float]]]:
    """Vertical rules as (x, y0, y1) and horizontal rules as (y, x0, x1)."""
    verticals: list[tuple[float, float, float]] = []
    horizontals: list[tuple[float, float, float]] = []
    try:
        drawings = page.get_drawings()
    except Exception:
        return verticals, horizontals
    for drawing in drawings:
        for item in drawing.get("items") or []:
            if not item:
                continue
            if item[0] == "l" and len(item) >= 3:
                a, b = item[1], item[2]
                if abs(a.x - b.x) <= 1.0:
                    verticals.append(((a.x + b.x) / 2, min(a.y, b.y), max(a.y, b.y)))
                elif abs(a.y - b.y) <= 1.0:
                    horizontals.append(((a.y + b.y) / 2, min(a.x, b.x), max(a.x, b.x)))
            elif item[0] == "re" and len(item) >= 2:
                r = item[1]
                if r.width <= 2.0 and r.height > 4:
                    verticals.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
                elif r.height <= 2.0 and r.width > 4:
                    horizontals.append(((r.y0 + r.y1) / 2, r.x0, r.x1))
    return verticals, horizontals


def _snap_to_rules(page: fitz.Page, bbox: tuple, widths: list[float], heights: list[float]) -> tuple[tuple, list[float], list[float]]:
    """
    Move a table's outer edges and column edges onto the lines drawn for it.

    The table finder sometimes stops at the text inside the cells, so the
    last column comes out as wide as its longest word. When the page draws
    vertical rules across the table, the rules are the true column edges.
    """
    x0, y0, x1, y1 = (float(v) for v in bbox)
    height = max(1.0, y1 - y0)
    width = max(1.0, x1 - x0)
    verticals, horizontals = _ruling_lines(page)
    # Rules that run down most of the table's height, near its sides.
    xs = sorted(
        x for x, top, bottom in verticals
        if min(bottom, y1) - max(top, y0) >= height * 0.6 and x0 - width * 0.25 <= x <= x1 + width * 0.75
    )
    edges: list[float] = []
    for x in xs:
        if not edges or x - edges[-1] > 3.0:
            edges.append(x)
    if len(edges) != len(widths) + 1:
        return bbox, widths, heights
    # Every text column must still start inside its own rule pair.
    starts = [x0]
    for w in widths[:-1]:
        starts.append(starts[-1] + w)
    if any(not (edges[i] - 3.0 <= starts[i] <= edges[i + 1]) for i in range(len(widths))):
        return bbox, widths, heights
    new_x0, new_x1 = edges[0], edges[-1]
    ys = [y for y, left, right in horizontals if min(right, new_x1) - max(left, new_x0) >= (new_x1 - new_x0) * 0.6]
    new_y0 = min([y for y in ys if y <= y0 + 3.0] or [y0])
    new_y1 = max([y for y in ys if y >= y1 - 3.0] or [y1])
    new_widths = [edges[i + 1] - edges[i] for i in range(len(widths))]
    rows_y: list[float] = []
    for y in sorted(y for y in ys if new_y0 - 1.0 <= y <= new_y1 + 1.0):
        if not rows_y or y - rows_y[-1] > 3.0:
            rows_y.append(y)
    if len(rows_y) == len(heights) + 1:
        heights = [rows_y[i + 1] - rows_y[i] for i in range(len(heights))]
    else:
        # Spread the added height over the rows in proportion.
        total = sum(heights) or 1.0
        heights = [h * (new_y1 - new_y0) / total for h in heights]
    return (new_x0, new_y0, new_x1, new_y1), new_widths, heights


def _cluster_double_border_columns(
    rows: list[list[str]],
    table,
    bbox: tuple,
    gap: float = 12.0,
) -> tuple[list[list[str]], list[float]] | None:
    """
    Merge find_tables columns that belong to one real column.

    Double rules produce 2–3 left edges a few points apart. Text for the same
    logical column lands in different extract columns by row (header in the
    middle stroke, body in the outer). Clustering by x0 and joining each band
    restores Action|Prepared|Reviewed|Approved (and similar) grids.
    """
    if not rows:
        return None
    cols = max(len(row) for row in rows)
    if cols < 3:
        return None
    try:
        cells = [cell for cell in (getattr(table, "cells", None) or []) if cell]
    except Exception:
        cells = []
    if not cells:
        return None
    raw_lefts = sorted({round(float(cell[0]), 2) for cell in cells})
    if len(raw_lefts) < 2:
        return None
    # One left edge per extract column when counts match; otherwise spread.
    if len(raw_lefts) == cols:
        col_x0 = raw_lefts
    elif len(raw_lefts) > cols:
        col_x0 = raw_lefts[:cols]
    else:
        # Fewer unique edges than extract cols — cannot map reliably.
        return None
    bands: list[list[int]] = [[0]]
    for index in range(1, cols):
        if col_x0[index] - col_x0[bands[-1][-1]] > gap:
            bands.append([index])
        else:
            bands[-1].append(index)
    # Fold a trailing hairline gutter into the previous band.
    span_right = float(bbox[2])
    if len(bands) >= 2:
        last_left = col_x0[bands[-1][0]]
        if span_right - last_left < 14.0:
            empty_last = not any(
                ((row[ci] if ci < len(row) else "") or "").strip()
                for row in rows
                for ci in bands[-1]
            )
            if empty_last:
                bands[-2].extend(bands[-1])
                bands.pop()
    if len(bands) < 2 or len(bands) == cols:
        return None
    new_rows: list[list[str]] = []
    for row in rows:
        merged: list[str] = []
        for band in bands:
            parts = [((row[ci] if ci < len(row) else "") or "").strip() for ci in band]
            parts = [part for part in parts if part]
            if not parts:
                merged.append("")
            elif len(parts) == 1:
                merged.append(parts[0])
            else:
                # Same logical cell split across double-border columns.
                uniq: list[str] = []
                for part in parts:
                    if part not in uniq:
                        uniq.append(part)
                merged.append("\n".join(uniq) if len(uniq) > 1 else uniq[0])
        new_rows.append(merged)
    band_lefts = [col_x0[band[0]] for band in bands]
    new_widths: list[float] = []
    for index, left in enumerate(band_lefts):
        right = band_lefts[index + 1] if index + 1 < len(band_lefts) else span_right
        new_widths.append(max(16.0, right - left))
    total = sum(new_widths) or 1.0
    span = max(1.0, span_right - float(bbox[0]))
    if abs(total - span) > 1.0:
        factor = span / total
        new_widths = [width * factor for width in new_widths]
    return new_rows, new_widths


def _collapse_gutter_columns(
    rows: list[list[str]],
    widths: list[float],
    heights: list[float],
) -> tuple[list[list[str]], list[float], list[float], None]:
    """
    Drop columns that are empty in every row (true gutter leftovers).
    """
    if not rows or not widths:
        return rows, widths, heights, None
    cols = min(len(widths), max(len(row) for row in rows))
    keep: list[int] = []
    for ci in range(cols):
        texts = [(row[ci] if ci < len(row) else "").strip() for row in rows]
        if any(texts):
            keep.append(ci)
    if len(keep) < 2 or len(keep) == cols:
        return rows, widths, heights, None
    new_rows = [[(row[ci] if ci < len(row) else "") for ci in keep] for row in rows]
    new_widths: list[float] = []
    absorb = 0.0
    keep_set = set(keep)
    for ci in range(cols):
        width = float(widths[ci] if ci < len(widths) else 0.0)
        if ci in keep_set:
            new_widths.append(width + absorb)
            absorb = 0.0
        else:
            absorb += width
    if absorb and new_widths:
        new_widths[-1] += absorb
    span = sum(float(w) for w in widths[:cols]) or 1.0
    total = sum(new_widths) or 1.0
    if abs(total - span) > 1.0:
        factor = span / total
        new_widths = [w * factor for w in new_widths]
    return new_rows, new_widths, heights, None


def _widths_from_finder(table, bbox: tuple, cols: int) -> list[float] | None:
    """Prefer the detector's column edges over equal splits."""
    try:
        cells = [cell for cell in (getattr(table, "cells", None) or []) if cell]
    except Exception:
        cells = []
    if not cells:
        return None
    # Cluster left edges so double-border strokes become one column start.
    raw_lefts = sorted(float(c[0]) for c in cells)
    lefts: list[float] = []
    for x in raw_lefts:
        if not lefts or x - lefts[-1] > 10.0:
            lefts.append(x)
    if len(lefts) < 2:
        return None
    # Match detector column count when clustering reduced duplicates.
    if abs(len(lefts) - cols) > max(2, cols // 2) and cols >= 2:
        # Still usable if we got a sane band count.
        pass
    widths = []
    for i, left in enumerate(lefts):
        right = lefts[i + 1] if i + 1 < len(lefts) else float(bbox[2])
        if i == len(lefts) - 1:
            right = float(bbox[2])
        widths.append(max(8.0, right - left))
    total = sum(widths)
    span = max(1.0, float(bbox[2]) - float(bbox[0]))
    if abs(total - span) > span * 0.35:
        return None
    # If clustering disagrees with extract() column count, prefer clustered widths
    # only when lengths match; else equal-split later after gutter collapse.
    if len(widths) != cols:
        return None
    return widths


def _heights_from_finder(table, bbox: tuple, rows: int) -> list[float] | None:
    try:
        cells = [cell for cell in (getattr(table, "cells", None) or []) if cell]
    except Exception:
        cells = []
    if not cells or rows < 1:
        return None
    raw_tops = sorted(float(c[1]) for c in cells)
    tops: list[float] = []
    for y in raw_tops:
        if not tops or y - tops[-1] > 6.0:
            tops.append(y)
    if len(tops) < rows:
        return None
    heights = []
    for i in range(min(rows, len(tops))):
        top = tops[i]
        bottom = tops[i + 1] if i + 1 < len(tops) else float(bbox[3])
        if i == rows - 1:
            bottom = float(bbox[3])
        heights.append(max(8.0, bottom - top))
    while len(heights) < rows:
        heights.append(max(8.0, (float(bbox[3]) - float(bbox[1])) / rows))
    return heights[:rows]


def _geometric_vector_fallback(page: fitz.Page, region: ImageRegion) -> RebuiltTable | None:
    """Cluster text by column alignment when find_tables is empty."""
    words = []
    try:
        for block in page.get_text("dict").get("blocks") or []:
            if block.get("type") != 0:
                continue
            for line in block.get("lines") or []:
                for span in line.get("spans") or []:
                    text = (span.get("text") or "").strip()
                    if not text:
                        continue
                    box = tuple(float(v) for v in span["bbox"])
                    if not _center_inside(box, region.bbox):
                        continue
                    words.append((box, text))
    except Exception:
        return None
    if len(words) < 4:
        return None
    # Cluster x0 into columns. Require a stable gutter, not every indent.
    xs = sorted(box[0] for box, _ in words)
    col_starts = [xs[0]]
    for x in xs[1:]:
        if x - col_starts[-1] > 36:
            col_starts.append(x)
    # Two-column prose pages produce many false "columns" from ragged indents.
    if len(col_starts) < 2 or len(col_starts) > 6:
        return None
    # Cluster y into rows.
    ys = sorted((box[1] + box[3]) / 2.0 for box, _ in words)
    row_ys = [ys[0]]
    for y in ys[1:]:
        if y - row_ys[-1] > 10:
            row_ys.append(y)
    if len(row_ys) > 40:
        # A full page of flowing lines is not a table.
        return None
    grid = [["" for _ in col_starts] for _ in row_ys]
    for box, text in words:
        cy = (box[1] + box[3]) / 2.0
        ri = min(range(len(row_ys)), key=lambda i: abs(row_ys[i] - cy))
        ci = min(range(len(col_starts)), key=lambda i: abs(col_starts[i] - box[0]))
        grid[ri][ci] = (grid[ri][ci] + " " + text).strip() if grid[ri][ci] else text
    cols = len(col_starts)
    widths = _column_widths(region.bbox, cols)
    heights = [(region.bbox[3] - region.bbox[1]) / max(len(grid), 1)] * len(grid)
    rebuilt = RebuiltTable(
        kind="vector",
        bbox=region.bbox,
        rows=grid,
        fills=[["" for _ in range(cols)] for _ in grid],
        widths=widths,
        heights=heights,
        confidences=[[100.0] * cols for _ in grid],
        rules=False,
        rtl=_page_rtl(page),
        notes=["vector table via column-alignment fallback"],
    )
    rebuilt.usable = table_is_usable(rebuilt)
    if not rebuilt.usable:
        try:
            rebuilt.fallback_png = clip_region_png(page, region.bbox)
            rebuilt.notes.append(
                "Column-alignment table not usable; original region kept as a picture."
            )
        except Exception:
            return None
    return rebuilt


def _rebuild_raster(page: fitz.Page, region: ImageRegion) -> RebuiltTable | None:
    png = region.image_bytes or _extract_png(page, region.xref, region.bbox)
    if not png:
        return None
    # First choice: find the grid in the pixels and read each cell at high
    # resolution (img2table + Tesseract). The older word-cluster reading
    # below stays as the fallback.
    try:
        from core import image_tables

        found = image_tables.read_table(page, region.bbox)
    except Exception:
        found = None
    if found is not None:
        rebuilt = RebuiltTable(
            kind="raster",
            bbox=found.bbox,
            rows=found.rows,
            widths=found.widths,
            heights=found.heights,
            confidences=[[100.0] * len(found.rows[0]) for _ in found.rows],
            rules=True,
            rtl=_page_rtl(page),
            fallback_png=png,
            notes=[*found.notes, f"({region.reason})"],
        )
        rebuilt.usable = table_is_usable(rebuilt)
        if rebuilt.usable:
            return rebuilt
    words = _ocr_cells(png)
    if not words:
        return RebuiltTable(
            kind="raster",
            bbox=region.bbox,
            rows=[],
            fallback_png=png,
            notes=[f"raster table OCR returned no words; picture kept as fallback ({region.reason})"],
        )
    rows, fills, widths, heights, confidences, low = _words_to_grid(png, words, region.bbox)
    if len(rows) < 2:
        return RebuiltTable(
            kind="raster",
            bbox=region.bbox,
            rows=rows,
            fallback_png=png,
            notes=[f"raster table grid incomplete; picture kept as fallback ({region.reason})"],
            low_confidence_cells=low,
        )
    note = f"raster table OCR rebuild ({region.reason})"
    if low:
        note += f"; {len(low)} low-confidence cell(s)"
    rebuilt = RebuiltTable(
        kind="raster",
        bbox=region.bbox,
        rows=rows,
        fills=fills,
        widths=widths,
        heights=heights,
        confidences=confidences,
        rules=True,
        rtl=_page_rtl(page),
        fallback_png=png,
        notes=[note],
        low_confidence_cells=low,
    )
    rebuilt.usable = table_is_usable(rebuilt)
    if not rebuilt.usable:
        rebuilt.notes.append(
            "OCR grid not usable (empty header or too many low-confidence cells); "
            "original picture kept visible; native cells reported only."
        )
    return rebuilt


def _ocr_cells(png: bytes) -> list[tuple[float, float, float, float, str, float]]:
    """Word boxes in image pixel space with confidence."""
    # Prefer Tesseract eng+ara for mixed thesis tables; fall back to Windows eng/ar.
    words = _tesseract_tsv(png, "eng+ara") or _tesseract_tsv(png, "eng")
    if words:
        return words
    return _windows_ocr_words(png)


def _tesseract_tsv(png: bytes, language: str) -> list[tuple[float, float, float, float, str, float]]:
    try:
        from engine.tesseract_bin import find_exe, ara_tessdata_dir, legacy_tessdata_dir
    except Exception:
        find_exe = None  # type: ignore
    exe = None
    folder = None
    if find_exe is not None:
        exe = find_exe()
        try:
            folder = ara_tessdata_dir() if "ara" in language else legacy_tessdata_dir(__file__)
        except Exception:
            folder = None
    if not exe:
        return []
    # eng.traineddata may live next to ara or in a system tessdata.
    data_dirs = []
    if folder is not None:
        data_dirs.append(Path(folder))
    for candidate in (
        Path(r"C:\Program Files\Tesseract-OCR\tessdata"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tessdata"),
    ):
        if candidate.is_dir():
            data_dirs.append(candidate)
    if not data_dirs:
        return []
    picture = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    try:
        picture.write(png)
        picture.close()
        for data_dir in data_dirs:
            result = subprocess.run(
                [
                    exe,
                    picture.name,
                    "stdout",
                    "-l",
                    language,
                    "--tessdata-dir",
                    str(data_dir),
                    "--psm",
                    "6",
                    "tsv",
                ],
                capture_output=True,
                timeout=120,
                check=False,
            )
            text = result.stdout.decode("utf-8", errors="replace")
            words = []
            for line in text.splitlines()[1:]:
                parts = line.split("\t")
                if len(parts) < 12 or parts[0] != "5":
                    continue
                raw = parts[11].strip()
                if not raw:
                    continue
                try:
                    conf = float(parts[10])
                    left, top, width, height = (float(parts[6]), float(parts[7]), float(parts[8]), float(parts[9]))
                except ValueError:
                    continue
                if conf < 15 or width < 1 or height < 1:
                    continue
                words.append((left, top, left + width, top + height, raw, conf))
            if words:
                return words
    except (OSError, subprocess.TimeoutExpired):
        return []
    finally:
        Path(picture.name).unlink(missing_ok=True)
    return []


def _windows_ocr_words(png: bytes) -> list[tuple[float, float, float, float, str, float]]:
    try:
        from winrt.runtime import ApartmentType, init_apartment
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import BitmapAlphaMode, BitmapDecoder, BitmapPixelFormat
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    except Exception:
        return []
    try:
        init_apartment(ApartmentType.MTA)
    except Exception:
        pass
    engine = OcrEngine.try_create_from_language(Language("en-US")) or OcrEngine.try_create_from_language(
        Language("ar-SA")
    )
    if engine is None:
        return []
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(png)
    writer.store_async().get()
    writer.detach_stream()
    stream.seek(0)
    decoder = BitmapDecoder.create_async(stream).get()
    bitmap = decoder.get_software_bitmap_async(BitmapPixelFormat.BGRA8, BitmapAlphaMode.PREMULTIPLIED).get()
    result = engine.recognize_async(bitmap).get()
    words = []
    for line in result.lines:
        for item in line.words:
            text = (getattr(item, "text", "") or "").strip()
            if not text:
                continue
            box = item.bounding_rect
            conf = float(getattr(item, "confidence", 0) or 0)
            if 0 < conf <= 1:
                conf *= 100.0
            if conf <= 0:
                conf = 70.0
            words.append(
                (
                    float(box.x),
                    float(box.y),
                    float(box.x) + float(box.width),
                    float(box.y) + float(box.height),
                    text,
                    conf,
                )
            )
    return words


def _words_to_grid(
    png: bytes,
    words: list[tuple[float, float, float, float, str, float]],
    page_bbox: tuple[float, float, float, float],
) -> tuple:
    try:
        from PIL import Image
        import numpy as np
    except Exception:
        return [], [], [], [], [], []
    image = Image.open(BytesIO(png)).convert("RGB")
    width, height = image.size
    gray = np.asarray(image.convert("L"), dtype=float)
    # Row boundaries from horizontal dark lines; fall back to word clustering.
    row_cuts = _projection_cuts(gray.mean(axis=1), min_gap=max(8, height // 40))
    col_cuts = _projection_cuts(gray.mean(axis=0), min_gap=max(8, width // 30))
    if len(row_cuts) < 3:
        centers = sorted((w[1] + w[3]) / 2.0 for w in words)
        row_cuts = _cluster_cuts(centers, height)
    if len(col_cuts) < 3:
        centers = sorted((w[0] + w[2]) / 2.0 for w in words)
        col_cuts = _cluster_cuts(centers, width)
    row_cuts = _ensure_edges(row_cuts, height)
    col_cuts = _ensure_edges(col_cuts, width)
    n_rows = len(row_cuts) - 1
    n_cols = len(col_cuts) - 1
    if n_rows < 1 or n_cols < 1:
        return [], [], [], [], [], []
    cells: list[list[list[tuple[str, float]]]] = [[[] for _ in range(n_cols)] for _ in range(n_rows)]
    for x0, y0, x1, y1, text, conf in words:
        cx = (x0 + x1) / 2.0
        cy = (y0 + y1) / 2.0
        ri = max(0, min(n_rows - 1, next((i for i in range(n_rows) if row_cuts[i] <= cy < row_cuts[i + 1]), n_rows - 1)))
        ci = max(0, min(n_cols - 1, next((i for i in range(n_cols) if col_cuts[i] <= cx < col_cuts[i + 1]), n_cols - 1)))
        cells[ri][ci].append((text, conf))
    rows: list[list[str]] = []
    confidences: list[list[float]] = []
    low: list[tuple[int, int, float]] = []
    for ri, row in enumerate(cells):
        texts = []
        confs = []
        for ci, parts in enumerate(row):
            if not parts:
                texts.append("")
                confs.append(0.0)
                continue
            text = " ".join(p[0] for p in parts)
            conf = sum(p[1] for p in parts) / len(parts)
            texts.append(text)
            confs.append(conf)
            if conf < _LOW_CONFIDENCE:
                low.append((ri, ci, conf))
        rows.append(texts)
        confidences.append(confs)
    # Drop empty trailing columns/rows.
    while rows and all(not cell for cell in rows[-1]):
        rows.pop()
        confidences.pop()
    if not rows:
        return [], [], [], [], [], low
    # Column widths / row heights in PDF points.
    page_w = page_bbox[2] - page_bbox[0]
    page_h = page_bbox[3] - page_bbox[1]
    widths = [page_w * (col_cuts[i + 1] - col_cuts[i]) / max(width, 1) for i in range(n_cols)]
    heights = [page_h * (row_cuts[i + 1] - row_cuts[i]) / max(height, 1) for i in range(len(rows))]
    fills = _header_fills_from_pixels(image, row_cuts, col_cuts, len(rows), n_cols)
    return rows, fills, widths, heights, confidences, low


def _projection_cuts(profile, min_gap: int) -> list[float]:
    """Cut positions at dark-line dips in a 1-D projection."""
    import numpy as np

    profile = np.asarray(profile, dtype=float)
    if profile.size < 10:
        return []
    median = float(np.median(profile))
    threshold = median - max(12.0, (median - float(profile.min())) * 0.35)
    dips = []
    i = 1
    while i < len(profile) - 1:
        if profile[i] < threshold and profile[i] <= profile[i - 1] and profile[i] <= profile[i + 1]:
            dips.append(float(i))
            i += min_gap
        else:
            i += 1
    return dips


def _cluster_cuts(centers: list[float], size: int) -> list[float]:
    if not centers:
        return [0.0, float(size)]
    groups = [[centers[0]]]
    for value in centers[1:]:
        if value - groups[-1][-1] > max(10.0, size * 0.02):
            groups.append([value])
        else:
            groups[-1].append(value)
    mids = [sum(g) / len(g) for g in groups]
    cuts = [0.0]
    for left, right in zip(mids, mids[1:]):
        cuts.append((left + right) / 2.0)
    cuts.append(float(size))
    return cuts


def _ensure_edges(cuts: list[float], size: int) -> list[float]:
    values = [0.0] + [c for c in cuts if 2 < c < size - 2] + [float(size)]
    cleaned = [values[0]]
    for value in values[1:]:
        if value - cleaned[-1] >= 4:
            cleaned.append(value)
    if cleaned[-1] != float(size):
        cleaned.append(float(size))
    return cleaned


def _header_fills_from_pixels(image, row_cuts, col_cuts, n_rows, n_cols) -> list[list[str]]:
    from PIL import Image

    fills = [["" for _ in range(n_cols)] for _ in range(n_rows)]
    if n_rows < 1:
        return fills
    # Sample first row for a dark header.
    y0 = int(row_cuts[0]) + 2
    y1 = int(row_cuts[1]) - 2
    if y1 <= y0:
        return fills
    for ci in range(n_cols):
        x0 = int(col_cuts[ci]) + 2
        x1 = int(col_cuts[ci + 1]) - 2
        if x1 <= x0:
            continue
        crop = image.crop((x0, y0, x1, y1)).resize((1, 1), Image.Resampling.BOX)
        r, g, b = crop.getpixel((0, 0))[:3]
        if r * 0.3 + g * 0.59 + b * 0.11 < 150:
            fills[0][ci] = f"{r:02X}{g:02X}{b:02X}"
    return fills


def _column_widths(bbox: tuple, cols: int) -> list[float]:
    width = max(1.0, bbox[2] - bbox[0])
    return [width / cols] * cols


def _sample_fills(page, bbox, rows, widths, heights) -> list[list[str]]:
    fills = [["" for _ in row] for row in rows]
    # Cheap sample: top-left of each cell via a low-res pixmap.
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2), clip=fitz.Rect(*bbox), alpha=False)
    except Exception:
        return fills
    zoom = 1.2
    x_cursor = 0.0
    for ci, w in enumerate(widths):
        y_cursor = 0.0
        for ri, h in enumerate(heights):
            px = int((x_cursor + w * 0.1) * zoom)
            py = int((y_cursor + h * 0.35) * zoom)
            try:
                r, g, b = pix.pixel(min(px, pix.width - 1), min(py, pix.height - 1))[:3]
            except Exception:
                r = g = b = 255
            if r * 0.3 + g * 0.59 + b * 0.11 < 150:
                fills[ri][ci] = f"{r:02X}{g:02X}{b:02X}"
            y_cursor += h
        x_cursor += w
    return fills


def _page_rtl(page: fitz.Page) -> bool:
    text = page.get_text("text") or ""
    arabic = sum(1 for c in text if "\u0600" <= c <= "\u06FF")
    latin = sum(1 for c in text if c.isascii() and c.isalpha())
    return arabic > latin
