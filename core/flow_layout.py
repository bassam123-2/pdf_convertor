# -*- coding: utf-8 -*-
"""
Write one PDF page as ordinary, flowing Word content.

The frame writer (html_layout) pins every paragraph in a floating box at its
PDF position. That looks right, but nothing can be edited: a word added to a
line has nowhere to go, Word counts no words, and a box whose font is a
little wider than the PDF's spills over its neighbour.

Here the page is written the way a person would type it:

* Paragraphs follow each other in reading order, with the space above each
  one taken from the PDF.
* Where the page sets two columns side by side (an English column next to
  an Arabic one, say), the columns become the two cells of one borderless
  table row. Each column then wraps inside its own cell and the two can
  never run into each other.
* Tables are real Word tables and pictures are inline pictures.

Nothing here is specific to one document. The only inputs are the blocks a
page was read into and the size of the sheet.
"""

from __future__ import annotations

import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# A block wider than this share of the sheet belongs to no column.
_SPAN_SHARE = 0.62
# Narrowest empty strip between two columns, in points.
_MIN_GUTTER = 6.0


class _CellHost:
    """Lets the paragraph and table writers add content to a table cell."""

    def __init__(self, cell, document):
        self._cell = cell
        self.styles = document.styles
        self._first = True

    def add_paragraph(self, *args, **kwargs):
        if self._first and self._cell.paragraphs and not self._cell.paragraphs[0].text:
            self._first = False
            return self._cell.paragraphs[0]
        self._first = False
        return self._cell.add_paragraph(*args, **kwargs)

    def add_table(self, rows, cols):
        self._first = False
        return self._cell.add_table(rows, cols)


def _box(block) -> tuple[float, float, float, float]:
    return tuple(float(v) for v in block.bbox)


def _kind(block) -> str:
    name = type(block).__name__
    if name == "TableBlock":
        return "table"
    if name == "ImageBlock":
        return "image"
    return "text"


def find_gutter(blocks: list, page_width: float) -> float | None:
    """
    The x of the empty strip between two side-by-side columns, or None.

    Two text blocks that sit level with each other, one left and one right,
    leave a gap between them. A gutter is an x that lies in the gaps of
    several such pairs, in the middle part of the sheet. A title centred
    across both columns does not hide it: only level pairs are looked at.
    """
    texts = [b for b in blocks if _kind(b) == "text" and (_box(b)[2] - _box(b)[0]) < page_width * _SPAN_SHARE]
    gaps: list[tuple[float, float]] = []
    for i, a in enumerate(texts):
        ax0, ay0, ax1, ay1 = _box(a)
        for b in texts[i + 1:]:
            bx0, by0, bx1, by1 = _box(b)
            level = min(ay1, by1) - max(ay0, by0)
            if level < min(ay1 - ay0, by1 - by0) * 0.5:
                continue
            left, right = (ax1, bx0) if ax1 <= bx0 else (bx1, ax0) if bx1 <= ax0 else (None, None)
            if left is None or right - left < _MIN_GUTTER:
                continue
            gaps.append((left, right))
    if len(gaps) < 2:
        return None
    lo, hi = page_width * 0.2, page_width * 0.8
    best_x, best_votes = None, 0
    for x in range(int(lo), int(hi) + 1):
        votes = sum(1 for left, right in gaps if left <= x <= right)
        if votes > best_votes:
            best_x, best_votes = x, votes
    if best_x is None or best_votes < 2:
        return None
    # Centre the gutter in the strip every voting pair leaves open.
    shared = [(left, right) for left, right in gaps if left <= best_x <= right]
    left = max(l for l, _ in shared)
    right = min(r for _, r in shared)
    gutter = (left + right) / 2.0
    # Each side needs text of its own; a page of one column with a stray
    # side note is not two columns.
    left_side = [b for b in texts if _box(b)[2] <= gutter]
    right_side = [b for b in texts if _box(b)[0] >= gutter]
    if len(left_side) < 2 or len(right_side) < 2:
        return None
    return gutter


def _runs(blocks: list, gutter: float | None, page_width: float) -> list[tuple[str, list]]:
    """
    Cut the page into pieces: ("flow", blocks) or ("columns", [left, right]).

    A block that crosses the gutter ends the column run above it.
    """
    ordered = sorted(blocks, key=lambda b: (_box(b)[1], _box(b)[0]))
    if gutter is None:
        return [("flow", ordered)]
    pieces: list[tuple[str, list]] = []
    left: list = []
    right: list = []

    def flush() -> None:
        nonlocal left, right
        if left and right:
            pieces.append(("columns", [left, right]))
        elif left or right:
            pieces.append(("flow", left or right))
        left, right = [], []

    for block in ordered:
        x0, _y0, x1, _y1 = _box(block)
        if x1 <= gutter + 1.0:
            left.append(block)
        elif x0 >= gutter - 1.0:
            right.append(block)
        else:
            flush()
            pieces.append(("flow", [block]))
    flush()
    merged: list[tuple[str, list]] = []
    for kind, items in pieces:
        if merged and kind == "flow" and merged[-1][0] == "flow":
            merged[-1][1].extend(items)
        else:
            merged.append((kind, items))
    return merged


def _gaps(blocks: list, top: float | None = None) -> None:
    """Space above each block: its distance from the one before, in points."""
    previous_bottom = top
    for block in blocks:
        x0, y0, x1, y1 = _box(block)
        if previous_bottom is not None and hasattr(block, "gap_before"):
            size = float(getattr(block, "size", 11) or 11)
            # Ordinary line spacing is already in the paragraph; only extra room counts.
            block.gap_before = max(0.0, min(36.0, y0 - previous_bottom - size * 0.35))
        previous_bottom = y1


def _latin(text: str) -> bool:
    arabic = len(re.findall(r"[؀-ۿ]", text or ""))
    latin = len(re.findall(r"[A-Za-z]", text or ""))
    return latin > arabic


def write_paragraph(host, block, conv) -> None:
    """One text block as one paragraph, in the direction of its own letters."""
    text = conv._clean_flow_text(block.text or "")
    if not text or not re.search(r"[0-9A-Za-z؀-ۿ]", text):
        return
    paragraph = host.add_paragraph()
    paragraph.style = host.styles["Normal"]
    english = _latin(text)
    if english:
        conv._set_paragraph_ltr(paragraph)
        if block.align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif block.align == "right":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif block.align == "justify":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        conv._set_paragraph_rtl(paragraph, text)
        conv._set_block_align(paragraph, block.align)
    heading = block.kind == "heading" or block.role in {"title", "subtitle"}
    size = min(max(float(block.size or 11), 7.0), 28.0 if block.role else 18.0)
    from core.docx_converter import _FONT

    # A fixed family for the whole file wins over the PDF's own family.
    family = "" if _FONT.get() else (getattr(block, "font", "") or "")
    for part in _split_scripts(text):
        run = paragraph.add_run(part)
        conv._set_run_font(run, size, bool(block.bold or heading), part)
        if family:
            fonts = run._element.get_or_add_rPr().find(qn("w:rFonts"))
            if fonts is not None:
                for slot in ("w:ascii", "w:hAnsi", "w:cs"):
                    fonts.set(qn(slot), family)
        if getattr(block, "color", ""):
            try:
                from docx.shared import RGBColor

                run.font.color.rgb = RGBColor.from_string(block.color.lstrip("#").upper()[:6])
            except Exception:
                pass
    paragraph.paragraph_format.space_before = Pt(max(0.0, float(block.gap_before or 0.0)))
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    if block.fill:
        conv._shade_box(paragraph._p.get_or_add_pPr(), block.fill)


def _split_scripts(text: str) -> list[str]:
    """
    Latin stretches as runs of their own inside an Arabic line.

    Word picks the complex-script font and direction per run, so a run that
    mixes both would draw its Latin words with the Arabic settings. Only
    stretches with Latin letters are cut out; digits, spaces and punctuation
    stay with the Arabic around them so the line keeps its order.
    """
    if not re.search(r"[\u0600-\u06FF]", text) or not re.search(r"[A-Za-z]", text):
        return [text]
    pieces = re.split(r"([A-Za-z][A-Za-z0-9 .,'’&/:%-]*[A-Za-z0-9.%]|[A-Za-z])", text)
    return [piece for piece in pieces if piece] or [text]


def _write_blocks(host, blocks: list, width_pt: float, height_pt: float, document, conv, counts: dict) -> None:
    from docx.shared import Emu

    blocks = conv_merge(blocks)
    _gaps(blocks)
    for block in blocks:
        kind = _kind(block)
        if kind == "table":
            conv._add_table(host, block, int(width_pt * 12700), height_pt)
            counts["tables"] = counts.get("tables", 0) + 1
        elif kind == "image":
            if getattr(block, "data", None) and write_image(host, block, width_pt, height_pt):
                counts["images"] = counts.get("images", 0) + 1
        else:
            write_paragraph(host, block, conv)
            key = "headings" if block.kind == "heading" else "paragraphs"
            counts[key] = counts.get(key, 0) + 1


def write_image(host, block, width_pt: float, height_pt: float) -> bool:
    """
    One picture as an inline picture in its own paragraph.

    It keeps its PDF size, shrunk only to fit the column (or the page) it
    sits in, and is placed left, centred or right as it was in the PDF.
    Inline, it moves with the text around it and works the same inside a
    table cell as on the page.
    """
    from io import BytesIO

    from docx.shared import Emu

    width = float(block.width or (_box(block)[2] - _box(block)[0]))
    height = float(block.height or (_box(block)[3] - _box(block)[1]))
    if width <= 0 or height <= 0:
        return False
    scale = min(1.0, max(1.0, width_pt) / width, max(1.0, height_pt * 0.85) / height)
    width, height = width * scale, height * scale
    if width < 8 or height < 8:
        return False
    paragraph = host.add_paragraph()
    paragraph.style = host.styles["Normal"]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(max(0.0, float(getattr(block, "gap_before", 0.0) or 0.0)))
    paragraph.paragraph_format.space_after = Pt(0)
    try:
        paragraph.add_run().add_picture(BytesIO(block.data), width=Emu(int(width * 12700)))
    except Exception:
        paragraph._p.getparent().remove(paragraph._p)
        return False
    return True


def conv_merge(blocks: list) -> list:
    """Join consecutive lines of one paragraph (shared with the frame writer)."""
    from core.html_layout import merge_paragraphs

    return merge_paragraphs(blocks)


def _column_table(document, conv, widths_pt: list[float]):
    table = document.add_table(rows=1, cols=len(widths_pt))
    table.autofit = False
    conv._set_table_borders(table, show=False)
    properties = table._tbl.tblPr
    width = OxmlElement("w:tblW")
    width.set(qn("w:w"), str(int(sum(widths_pt) * 20)))
    width.set(qn("w:type"), "dxa")
    old = properties.find(qn("w:tblW"))
    if old is not None:
        properties.remove(old)
    properties.append(width)
    for old_layout in properties.findall(qn("w:tblLayout")):
        properties.remove(old_layout)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    properties.append(layout)
    # Cells are in printed order, left to right. Say so: in a right-to-left
    # section a table without this may be laid out mirrored.
    visual = OxmlElement("w:bidiVisual")
    visual.set(qn("w:val"), "0")
    properties.append(visual)
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for node, value in zip(grid.findall(qn("w:gridCol")), widths_pt):
            node.set(qn("w:w"), str(int(value * 20)))
    for cell, value in zip(table.rows[0].cells, widths_pt):
        conv._set_cell_width(cell, int(value * 20))
    conv._pad_table_cells(table, 80, 0)
    from core.faithful import _order_tbl_pr

    _order_tbl_pr(properties)
    return table


def write_flow_page(document, blocks: list, page_width: float, page_height: float, counts: dict) -> None:
    """Write one page's blocks as flowing paragraphs, column tables, tables and pictures."""
    from core import docx_converter as conv

    section = document.sections[-1]
    content_pt = (int(section.page_width) - int(section.left_margin) - int(section.right_margin)) / 12700.0
    content_pt = max(120.0, content_pt)
    height_pt = (int(section.page_height) - int(section.top_margin) - int(section.bottom_margin)) / 12700.0
    # The running footer is written by the section; leave its lines out here.
    body = [b for b in blocks if _kind(b) == "table" or _box(b)[1] < page_height * 0.92]
    gutter = find_gutter(body, page_width)
    for kind, items in _runs(body, gutter, page_width):
        if kind == "flow":
            _write_blocks(document, items, content_pt, height_pt, document, conv, counts)
            continue
        left, right = items
        left_edge = min(_box(b)[0] for b in left)
        right_edge = max(_box(b)[2] for b in right)
        span = max(1.0, right_edge - left_edge)
        share = min(0.8, max(0.2, (gutter - left_edge) / span))
        widths = [content_pt * share, content_pt * (1.0 - share)]
        table = _column_table(document, conv, widths)
        cells = table.rows[0].cells
        for cell, side, width in zip(cells, (left, right), widths):
            _write_blocks(_CellHost(cell, document), side, width - 8.0, height_pt, document, conv, counts)
        counts["column_tables"] = counts.get("column_tables", 0) + 1


def blocks_from_faithful(page, payload) -> list:
    """
    Paragraph, table and picture blocks from a faithful page reading.

    Each line becomes a one-line text block; the writer joins lines into
    paragraphs per column. Tables keep their rows and column widths. Every
    embedded picture smaller than the sheet is cut out at its PDF box.
    """
    import fitz

    from core.docx_converter import ImageBlock, TableBlock, TextBlock

    blocks: list = []
    table_boxes = []
    for table in getattr(payload, "tables", None) or []:
        rows = [list(row) for row in (getattr(table, "rows", None) or [])]
        if not rows:
            continue
        box = tuple(float(v) for v in table.bbox)
        table_boxes.append(box)
        if getattr(table, "fallback_png", None) and not getattr(table, "usable", True):
            blocks.append(ImageBlock(box, table.fallback_png, box[2] - box[0], box[3] - box[1], "center"))
            continue
        blocks.append(TableBlock(
            bbox=box,
            rows=rows,
            fills=[list(row) for row in (getattr(table, "fills", None) or [])],
            widths=list(getattr(table, "widths", None) or []),
            rules=bool(getattr(table, "rules", True)),
        ))

    def inside_table(line) -> bool:
        cx, cy = (line.x0 + line.x1) / 2.0, (line.y0 + line.y1) / 2.0
        return any(b[0] <= cx <= b[2] and b[1] <= cy <= b[3] for b in table_boxes)

    for line in payload.lines:
        if not (line.text or "").strip() or inside_table(line):
            continue
        chunks = [c for c in (line.chunks or []) if (c.text or "").strip()]
        weight = sum(len(c.text) for c in chunks) or 1
        bold = sum(len(c.text) for c in chunks if c.bold) >= weight * 0.6
        color = ""
        if chunks:
            painted = [c.color for c in chunks if c.color and c.color.lstrip("#").lower() not in {"000000", "000"}]
            if len(painted) >= len(chunks) * 0.6:
                color = painted[0]
        block = TextBlock(
            kind="heading" if bold and len(line.text) < 90 else "paragraph",
            bbox=(line.x0, line.y0, line.x1, line.y1),
            text=line.text,
            size=float(line.size or 11),
            bold=bold,
            align="right" if line.rtl else "left",
            color=color,
        )
        # The PDF's own family (Times New Roman, Calibri ...), so the Word
        # line takes about the room it took in the PDF.
        fonts: dict[str, int] = {}
        for chunk in chunks:
            if chunk.font:
                fonts[chunk.font] = fonts.get(chunk.font, 0) + len(chunk.text)
        if fonts:
            from core.font_catalog import _spaced

            block.font = _spaced(max(fonts, key=fonts.get))
        blocks.append(block)

    page_area = float(page.rect.width * page.rect.height) or 1.0
    try:
        infos = page.get_image_info() or []
    except Exception:
        infos = []
    for info in infos:
        box = info.get("bbox")
        if not box:
            continue
        rect = fitz.Rect(box) & page.rect
        if rect.is_empty or rect.width < 12 or rect.height < 12:
            continue
        if rect.get_area() / page_area > 0.85:
            continue  # a page-sized backdrop or scan, not a figure
        if any(fitz.Rect(t).contains(rect) for t in table_boxes):
            continue
        try:
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=rect, alpha=False)
            data = pixmap.tobytes("png")
        except Exception:
            continue
        blocks.append(ImageBlock((rect.x0, rect.y0, rect.x1, rect.y1), data, rect.width, rect.height, "center"))
    _set_alignment(blocks, float(page.rect.width))
    return blocks


def _set_alignment(blocks: list, page_width: float) -> None:
    """Centre a line that sits in the middle of its column, away from both edges."""
    gutter = find_gutter(blocks, page_width)
    texts = [b for b in blocks if _kind(b) == "text"]
    for block in texts:
        x0, _y0, x1, _y1 = _box(block)
        if gutter is not None and x1 <= gutter + 1:
            side = [b for b in texts if _box(b)[2] <= gutter + 1]
        elif gutter is not None and x0 >= gutter - 1:
            side = [b for b in texts if _box(b)[0] >= gutter - 1]
        else:
            side = texts
        left = min(_box(b)[0] for b in side)
        right = max(_box(b)[2] for b in side)
        width = max(1.0, right - left)
        if (x1 - x0) > width * 0.8:
            continue
        middle = (x0 + x1) / 2.0
        if abs(middle - (left + right) / 2.0) <= width * 0.06 and x0 - left > width * 0.08 and right - x1 > width * 0.08:
            block.align = "center"


def page_flows_well(page, payload) -> tuple[bool, str]:
    """
    Whether a faithful reading can be written as flowing text without loss.

    Flowing text has no page picture under it. A page whose glyphs could not
    all be read as text (a broken font map), whose text is drawn at an angle
    or as outlines, or that is mostly pictures and drawings, would lose what
    only the picture carried, or turn into a long run of loose pictures.
    Such a page keeps the positioned layout instead.
    """
    import fitz

    from core.faithful import _visible_glyphs

    glyphs = max(1, _visible_glyphs(page))
    kept = int(getattr(payload, "kept_in_picture", 0) or 0)
    if kept > max(3, glyphs * 0.02):
        return False, f"{kept} of {glyphs} glyphs could only be kept in the page picture"
    tokens = [t for line in payload.lines for t in (line.text or "").split() if re.search(r"[\u0600-\u06FF]", t)]
    mixed = sum(1 for t in tokens if re.search(r"[A-Za-z|•¢´¦…]", t))
    if len(tokens) >= 12 and mixed >= 5 and mixed / len(tokens) >= 0.04:
        return False, f"{mixed} of {len(tokens)} Arabic words carry Latin letters or symbols (broken font map)"
    if any("drawn as an outline" in note or "set at an angle" in note for note in payload.notes):
        return False, "text at an angle or drawn as outlines"
    if any(note.startswith("Chart/figure") and "overlapping text labels" in note for note in payload.notes):
        return False, "a figure carries text labels"
    area = float(page.rect.width * page.rect.height) or 1.0
    pictures = 0
    covered = 0.0
    try:
        infos = page.get_image_info() or []
    except Exception:
        infos = []
    for info in infos:
        rect = fitz.Rect(info.get("bbox") or (0, 0, 0, 0)) & page.rect
        if rect.is_empty or rect.width < 12 or rect.height < 12 or rect.get_area() / area > 0.85:
            continue
        pictures += 1
        covered += rect.get_area()
    if pictures > 6 or covered / area > 0.45:
        return False, f"{pictures} pictures cover {covered / area:.0%} of the page"
    return True, ""
