# -*- coding: utf-8 -*-
"""
Turn one PDF page into HTML, then into Word frames at the same coordinates.

Word's HTML importer drops absolute positioning, so the HTML is not handed to
Word as a web page. It is the box model: each element carries its PDF box, and
the bridge writes a page-relative frame for that box. Frames do not flow, so
a page cannot grow into the next one.
"""

from __future__ import annotations

import base64
import contextvars
import html
import re
from html.parser import HTMLParser
from io import BytesIO

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Pt
from lxml import etree

# Shared with the Word reader. When it is off, each visual line stays its own frame.
MERGE_PARAGRAPHS = contextvars.ContextVar("pdf_tool_merge_paragraphs", default=True)

_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_SHAPE_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_DRAWING_ID = 10


def _next_id() -> int:
    global _DRAWING_ID
    _DRAWING_ID += 1
    return _DRAWING_ID


def _pt(value: float) -> int:
    return max(0, int(round(float(value) * 12700)))


def _css_name(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def _image_mime(data: bytes) -> str:
    if data.startswith(b"\xff\xd8"):
        return "image/jpeg"
    if data.startswith(b"\x89PNG"):
        return "image/png"
    return "image/png"


def blocks_to_page_html(page_width: float, page_height: float, blocks: list) -> str:
    """
    One PDF page as a single absolutely positioned HTML page.

    Text, tables, and images keep the bounding box they had on the PDF.
    Arabic is marked right-to-left. A page break follows the container.
    """
    parts = [
        "<article class=\"pdf-page\"",
        f" data-page-width=\"{page_width:.2f}\"",
        f" data-page-height=\"{page_height:.2f}\"",
        " style=\"position:relative;overflow:hidden;direction:rtl;",
        f"width:{page_width:.2f}pt;height:{page_height:.2f}pt;page-break-after:always\">",
    ]
    ordered = sorted(blocks, key=lambda block: (0 if _kind(block) == "image" else 1 if _kind(block) == "table" else 2, _box(block)[1], _box(block)[0]))
    for block in ordered:
        piece = _block_html(block, page_width, page_height)
        if piece:
            parts.append(piece)
    parts.append("</article>")
    return "".join(parts)


def locked_size(source_width: float, source_height: float, layout_width: float) -> tuple[float, float]:
    """
    Scale to the layout width and keep the source ratio.

    The height follows the width. A header that is a wide, short band stays
    wide and short instead of collapsing to a thumbnail.
    """
    if layout_width <= 0:
        return 0.0, 0.0
    if source_width <= 0 or source_height <= 0:
        return layout_width, layout_width * 0.08
    height = layout_width * (source_height / source_width)
    return layout_width, height


def _object_type(block, page_height: float) -> str:
    """Header, footer, table, image, or text. One role per block."""
    kind = _kind(block)
    top, bottom = _box(block)[1], _box(block)[3]
    if kind == "table":
        return "table"
    if top <= page_height * 0.05 or (kind == "image" and bottom <= page_height * 0.14):
        return "header"
    # A corner picture starts above the footer and runs to the page edge.
    # Only its top decides that it is still a picture.
    if kind == "image":
        if top >= page_height * 0.88:
            return "footer"
        return "image"
    if top >= page_height * 0.88 or bottom >= page_height * 0.95:
        return "footer"
    return "text"


def _block_direction(text: str) -> str:
    """rtl or ltr for one string, using the conversion's language choice."""
    from core.docx_converter import text_direction

    return text_direction(text)


def _kind(block) -> str:
    name = type(block).__name__
    if name == "ImageBlock":
        return "image"
    if name == "TableBlock":
        return "table"
    return "text"


def _box(block) -> tuple[float, float, float, float]:
    raw = getattr(block, "bbox", (0, 0, 0, 0))
    return (float(raw[0]), float(raw[1]), float(raw[2]), float(raw[3]))


def _clip(box, page_width: float, page_height: float) -> tuple[float, float, float, float] | None:
    left = max(0.0, min(box[0], page_width - 4))
    top = max(0.0, min(box[1], page_height - 4))
    right = min(page_width, max(box[2], left + 4))
    bottom = min(page_height, max(box[3], top + 4))
    if right - left < 4 or bottom - top < 4:
        return None
    return left, top, right - left, bottom - top


def _block_html(block, page_width: float, page_height: float) -> str:
    clipped = _clip(_box(block), page_width, page_height)
    if clipped is None:
        return ""
    left, top, width, height = clipped
    obj = _object_type(block, page_height)
    if obj == "header" and _kind(block) == "image" and width > 0 and height > 0:
        width, height = locked_size(width, height, page_width)
        left = 0.0
    style = (
        f"position:absolute;left:{left:.2f}pt;top:{top:.2f}pt;"
        f"width:{'100%' if obj == 'header' else f'{width:.2f}pt'};"
        f"height:{height:.2f}pt;margin:0;padding:0;object-fit:contain;"
    )
    common = (
        f" data-object=\"{obj}\""
        f" data-x=\"{left:.2f}\" data-y=\"{top:.2f}\""
        f" data-w=\"{width:.2f}\" data-h=\"{height:.2f}\""
    )
    if _kind(block) == "image":
        payload = base64.b64encode(block.data).decode("ascii")
        mime = _image_mime(block.data)
        fit = html.escape(getattr(block, "fit", "") or "", quote=True)
        return (
            f"<img class=\"block\" data-kind=\"image\"{common} data-fit=\"{fit}\""
            f" style=\"{style}\" alt=\"\" src=\"data:{mime};base64,{payload}\"/>"
        )
    if _kind(block) == "table":
        return _table_html(block, common, style, width)
    text = _css_name(getattr(block, "text", "") or "")
    role = getattr(block, "role", "") or ""
    if role == "toc":
        text = re.sub(r"(?:\.\s*){2,}|[.\u00b7\u2022\u2024\u2026‧]{2,}", " ", text)
        text = _css_name(text).strip(" .\t")
    if not text or not re.search(r"[0-9A-Za-z\u0600-\u06FF]", text):
        return ""
    kind = "toc" if role == "toc" and getattr(block, "toc_page", "") else "heading" if getattr(block, "kind", "") == "heading" or role in {"title", "subtitle"} else "text"
    align = getattr(block, "align", "right") or "right"
    if align not in {"left", "center", "right"}:
        align = "right"
    size = float(getattr(block, "size", 12.0) or 12.0)
    bold = "1" if getattr(block, "bold", False) or role == "title" else "0"
    fill = re.sub(r"[^0-9A-Fa-f]", "", getattr(block, "fill", "") or "")[:6]
    ink = re.sub(r"[^0-9A-Fa-f]", "", getattr(block, "color", "") or "")[:6]
    toc = html.escape(getattr(block, "toc_page", "") or "", quote=True)
    css_align = "right" if align == "right" else "center" if align == "center" else "left"
    return (
        f"<p class=\"block\" data-kind=\"{kind}\"{common}"
        f" data-size=\"{size:.2f}\" data-bold=\"{bold}\" data-align=\"{align}\""
        f" data-fill=\"{fill}\" data-ink=\"{ink}\" data-toc=\"{toc}\""
        f" style=\"{style}direction:{'rtl' if _block_direction(text) == 'rtl' else 'ltr'};"
        f"text-align:{css_align};font-size:{size:.2f}pt\">"
        f"{html.escape(text)}</p>"
    )


def _table_html(block, common: str, style: str, width: float) -> str:
    rows = list(getattr(block, "rows", []) or [])
    if not rows:
        return ""
    columns = max(len(row) for row in rows)
    fills = list(getattr(block, "fills", []) or [])
    widths = list(getattr(block, "widths", []) or [])
    rules = "1" if getattr(block, "rules", True) else "0"
    border = "1px solid #8AA0B4" if rules == "1" else "0"
    table_direction = _block_direction(" ".join(str(cell) for row in rows for cell in row))
    parts = [
        f"<table class=\"block\" data-kind=\"table\"{common} data-rules=\"{rules}\"",
        f" style=\"{style}border-collapse:collapse;direction:{table_direction}\">",
    ]
    if widths and len(widths) == columns and sum(widths) > 0:
        parts.append("<colgroup>")
        for value in widths:
            parts.append(f"<col data-w=\"{float(value):.2f}\"/>")
        parts.append("</colgroup>")
    for row_index, row in enumerate(rows):
        parts.append("<tr>")
        padded = list(row) + [""] * (columns - len(row))
        for column, cell in enumerate(padded):
            fill = ""
            if row_index < len(fills) and column < len(fills[row_index]):
                fill = re.sub(r"[^0-9A-Fa-f]", "", fills[row_index][column] or "")[:6]
            paint = f"background:#{fill};" if fill else ""
            cell_direction = _block_direction(str(cell))
            parts.append(
                f"<td data-fill=\"{fill}\" style=\"border:{border};direction:{cell_direction};"
                f"vertical-align:middle;{paint}width:{width / columns:.2f}pt\">"
                f"{html.escape(_css_name(str(cell)))}</td>"
            )
        parts.append("</tr>")
    parts.append("</table>")
    return "".join(parts)


class _PageParser(HTMLParser):
    """Read the page container back into boxes. The bridge places those boxes."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.page_width = 0.0
        self.page_height = 0.0
        self.blocks: list[dict] = []
        self._mode = ""
        self._buf: list[str] = []
        self._item: dict | None = None
        self._table: dict | None = None
        self._fill = ""

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr = {key: value or "" for key, value in attrs}
        classes = attr.get("class", "")
        if tag == "article" and "pdf-page" in classes:
            self.page_width = float(attr.get("data-page-width") or 0)
            self.page_height = float(attr.get("data-page-height") or 0)
            return
        if tag == "img" and "block" in classes:
            self.blocks.append(self._media(attr))
            return
        if tag == "p" and "block" in classes:
            self._mode = "p"
            self._buf = []
            self._item = self._text_item(attr)
            return
        if tag == "table" and "block" in classes:
            self._table = self._text_item(attr)
            self._table["kind"] = "table"
            self._table["rules"] = attr.get("data-rules", "1") == "1"
            self._table["rows"] = []
            self._table["fills"] = []
            self._table["widths"] = []
            return
        if tag == "col" and self._table is not None and not self._mode:
            raw = attr.get("data-w", "")
            if raw:
                self._table["widths"].append(float(raw))
            return
        if tag == "tr" and self._table is not None:
            self._table["rows"].append([])
            self._table["fills"].append([])
            return
        if tag == "td" and self._table is not None:
            self._mode = "cell"
            self._buf = []
            self._fill = attr.get("data-fill", "")
            return

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.handle_starttag(tag, attrs)

    def handle_data(self, data: str) -> None:
        if self._mode in {"p", "cell"}:
            self._buf.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "p" and self._mode == "p" and self._item is not None:
            self._item["text"] = "".join(self._buf).strip()
            if self._item["text"]:
                self.blocks.append(self._item)
            self._mode = ""
            self._item = None
            return
        if tag == "td" and self._mode == "cell" and self._table is not None and self._table["rows"]:
            self._table["rows"][-1].append("".join(self._buf).strip())
            self._table["fills"][-1].append(self._fill)
            self._mode = ""
            return
        if tag == "table" and self._table is not None:
            if self._table["rows"]:
                self.blocks.append(self._table)
            self._table = None

    def _media(self, attr: dict[str, str]) -> dict:
        item = self._text_item(attr)
        item["kind"] = "image"
        item["src"] = attr.get("src", "")
        item["fit"] = attr.get("data-fit", "")
        return item

    def _text_item(self, attr: dict[str, str]) -> dict:
        return {
            "object": attr.get("data-object", "text"),
            "kind": attr.get("data-kind", "text"),
            "x": float(attr.get("data-x") or 0),
            "y": float(attr.get("data-y") or 0),
            "w": float(attr.get("data-w") or 0),
            "h": float(attr.get("data-h") or 0),
            "size": float(attr.get("data-size") or 12),
            "bold": attr.get("data-bold", "0") == "1",
            "align": attr.get("data-align", "right") or "right",
            "fill": attr.get("data-fill", ""),
            "ink": attr.get("data-ink", ""),
            "toc": attr.get("data-toc", ""),
            "text": "",
        }


def parse_page_html(source: str) -> tuple[float, float, list[dict]]:
    """The boxes declared by one page's HTML."""
    parser = _PageParser()
    parser.feed(source)
    parser.close()
    return parser.page_width, parser.page_height, parser.blocks


def _flow_together(previous, block) -> bool:
    """Same size, same edge, and ordinary line spacing: one paragraph, not two boxes."""
    if getattr(previous, "role", "") in {"toc", "title", "subtitle"} or getattr(block, "role", "") in {"toc", "title", "subtitle"}:
        return False
    if abs(float(getattr(previous, "size", 12)) - float(getattr(block, "size", 12))) > 1.6:
        return False
    if bool(getattr(previous, "bold", False)) != bool(getattr(block, "bold", False)):
        return False
    if _starts_list_item(getattr(block, "text", "") or ""):
        return False
    left = _box(previous)
    right = _box(block)
    if abs(left[2] - right[2]) > 24 and abs(left[0] - right[0]) > 24:
        return False
    gap = right[1] - left[3]
    size = max(float(getattr(previous, "size", 12)), 8.0)
    return -1.0 <= gap <= max(size * 0.85, 7.0)


_LIST_START = re.compile(
    r"^\s*(?:"
    r"[•●▪◦■□➢►✓\-–\uF0A7\uF0B7\uF0D8\uF0FC]"            # bullets, Wingdings bullets
    r"|\(?[\d٠-٩]{1,3}\s*[.)\-–](?=\s|[\u0600-\u06FF])"     # 1.  2)  ٣-
    r"|\.[\d٠-٩]{1,3}(?=\s*[\u0600-\u06FF])"                 # .1 in a right-to-left line
    r"|[\d٠-٩]{1,3}(?=[\u0600-\u06FF])"                        # 1أي
    r"|\(?[a-zA-Z\u0623-\u064A]\)\s"                            # a)  (b)  أ)
    r")"
)


def _starts_list_item(text: str) -> bool:
    """True when a line opens a new bullet or numbered item."""
    return bool(_LIST_START.match(text or ""))


def merge_paragraphs(blocks: list) -> list:
    """
    Collapse consecutive lines into one right-to-left paragraph.

    A real break is a change of size, alignment, or a gap wider than the line spacing.
    """
    if not MERGE_PARAGRAPHS.get():
        return list(blocks)
    ordered = sorted(blocks, key=lambda block: (_box(block)[1], _box(block)[0]))
    merged: list = []
    for block in ordered:
        if _kind(block) != "text":
            merged.append(block)
            continue
        if merged and _kind(merged[-1]) == "text" and _flow_together(merged[-1], block):
            previous = merged[-1]
            previous.text = f"{previous.text} {block.text}".strip()
            box = _box(previous)
            other = _box(block)
            previous.bbox = (min(box[0], other[0]), min(box[1], other[1]), max(box[2], other[2]), max(box[3], other[3]))
            continue
        merged.append(block)
    return merged


_FACE_PX = 32
_FACES: dict[str, object] = {}


def _installed_font(family: str):
    """The font file for the family Word will use, or None."""
    if family in _FACES:
        return _FACES[family]
    face = None
    try:
        import os
        import winreg
        from pathlib import Path

        from PIL import ImageFont

        fonts = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
            for index in range(winreg.QueryInfoKey(key)[1]):
                name, value, _kind = winreg.EnumValue(key, index)
                label = name.lower().split("(")[0].strip()
                if label != family.lower():
                    continue
                path = Path(value)
                if not path.is_absolute():
                    path = fonts / value
                if path.is_file():
                    face = ImageFont.truetype(str(path), _FACE_PX)
                    break
    except Exception:
        face = None
    _FACES[family] = face
    return face


def _output_faces(text: str) -> list:
    """The faces Word assigns: the complex-script face, the Latin face, or both."""
    from core.arabic_fixer import _is_arabic
    from core.docx_converter import _ARABIC_FONT, _LATIN_FONT

    faces = []
    if any(_is_arabic(char) for char in text):
        faces.append(_installed_font(_ARABIC_FONT))
    if any(char.isascii() and char.isalnum() for char in text):
        faces.append(_installed_font(_LATIN_FONT))
    return [face for face in faces if face]


def _text_width(text: str, size: float) -> float:
    """Width in the face Word will draw, measured per script."""
    from core.arabic_fixer import _is_arabic
    from core.docx_converter import _ARABIC_FONT, _LATIN_FONT

    width = 0.0
    piece = ""
    arabic = False
    for char in text + "\0":
        this = _is_arabic(char) if char != "\0" else None
        if char != "\0" and (not piece or this == arabic):
            piece += char
            arabic = bool(this)
            continue
        face = _installed_font(_ARABIC_FONT if arabic else _LATIN_FONT)
        if face is not None:
            width += float(face.getlength(piece)) * (size / _FACE_PX)
        else:
            width += sum((size * 0.52) if _is_arabic(item) else (size * 0.28 if item == " " else size * 0.50) for item in piece)
        piece = "" if char == "\0" else char
        arabic = bool(this) if char != "\0" else False
    return width


def _line_factor(text: str) -> float:
    """
    Ascent plus descent, in ems, of the face Word will use.

    The PDF box is the ink. A frame shorter than the face's line box clips
    the line. The factor comes from that face, not from one document.
    """
    faces = _output_faces(text)
    ratios = []
    for face in faces:
        ascent, descent = face.getmetrics()
        ratios.append((ascent + descent) / _FACE_PX)
    if ratios:
        return max(ratios)
    from core.arabic_fixer import _is_arabic

    if any(_is_arabic(char) for char in text):
        return 1.5
    return 1.15


def _fit_text_frames(items: list[dict], page_width: float, page_height: float) -> None:
    """
    Give each text frame the line box of the face Word will draw.

    The PDF box is the ink. Word clips a frame shorter than the face's
    ascent plus descent (noAutofit, insets 0). The frame grows into the
    empty gap around it. The type size comes down only when the line is
    wider than the frame and would wrap out of sight.
    """
    boxes = [
        (item["x"], item["y"], item["x"] + item["w"], item["y"] + item["h"], id(item))
        for item in items
    ]
    for item in items:
        if item.get("kind") not in {"text", "heading", "toc"}:
            continue
        text = item.get("text") or ""
        if not text:
            continue
        size = float(item["size"] or 12)
        one_line = item["h"] <= max(size * 1.9, 24.0)
        factor = _line_factor(text)
        needed_w = _text_width(text, size) + size * 0.4 if one_line else item["w"]
        if needed_w > item["w"] + 1.5:
            extra = needed_w - item["w"]
            align = item.get("align") or "right"
            left_extra = extra if align != "left" else 0.0
            right_extra = extra if align == "left" else 0.0
            if align == "center":
                left_extra = right_extra = extra / 2.0
            item["x"] = max(4.0, item["x"] - left_extra)
            item["w"] = min(page_width - 4.0 - item["x"], item["w"] + left_extra + right_extra)
        needed_h = size * factor if one_line else item["h"] + size * max(0.15, factor - 1.0)
        if needed_h > item["h"]:
            bottom = item["y"] + item["h"]
            limit = page_height - 4.0
            for left, top, right, _bottom, ident in boxes:
                if ident == id(item) or top <= bottom + 0.4:
                    continue
                if min(item["x"] + item["w"], right) - max(item["x"], left) < 8:
                    continue
                limit = min(limit, top - 1.0)
            item["h"] = min(needed_h, max(item["h"], limit - item["y"]))
        if one_line:
            width_now = _text_width(text, item["size"])
            if item["w"] > 8 and width_now > item["w"] * 0.98:
                item["size"] = max(6.5, float(item["size"]) * (item["w"] * 0.98) / width_now)


def write_page_html(document, blocks: list, page_width: float, page_height: float, counts: dict) -> None:
    """Build the page HTML and place every box on this Word page."""
    blocks = merge_paragraphs(blocks)
    source = blocks_to_page_html(page_width, page_height, blocks)
    _width, _height, items = parse_page_html(source)
    if not items:
        return
    _lock_toc_column(items)
    _fit_text_frames(items, page_width, page_height)
    host = document.add_paragraph()
    _pin_host(host)
    for item in items:
        if item.get("object") == "footer":
            continue
        if item["kind"] != "table" and item["y"] >= page_height * 0.88:
            continue
        if item["kind"] == "image":
            _place_image(host, item)
        elif item["kind"] == "table":
            _place_table(host, item)
            counts["tables"] += 1
        else:
            _place_text(host, item)
            if item["kind"] == "heading":
                counts["headings"] += 1
            else:
                counts["paragraphs"] += 1
    _pin_host(host)


def _lock_toc_column(items: list[dict]) -> None:
    """Give every contents line the same width so the page numbers share one edge."""
    rows = [item for item in items if item.get("kind") == "toc" and item.get("toc")]
    if not rows:
        return
    left = min(item["x"] for item in rows)
    right = max(item["x"] + item["w"] for item in rows)
    width = max(120.0, right - left)
    for item in rows:
        item["x"] = left
        item["w"] = width
        item["tab"] = width - 6.0


def _pin_host(paragraph) -> None:
    """The frames live on the page. The paragraph that holds them has no height."""
    properties = paragraph._p.get_or_add_pPr()
    spacing = properties.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        properties.append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "20")
    spacing.set(qn("w:lineRule"), "exact")


def _place_image(paragraph, item: dict) -> None:
    payload = item.get("src", "")
    marker = "base64,"
    if marker not in payload or item["w"] < 8 or item["h"] < 8:
        return
    try:
        data = base64.b64decode(payload.split(marker, 1)[1])
    except Exception:
        return
    if not data:
        return
    layout_w, layout_h = locked_size(item["w"], item["h"], item["w"])
    run = paragraph.add_run()
    run.add_picture(BytesIO(data), width=Emu(_pt(layout_w)))
    _float_drawing(run, item["x"], item["y"], layout_w, layout_h, behind=item.get("fit") != "page")


def _place_text(paragraph, item: dict) -> None:
    if item["w"] < 8 or item["h"] < 6 or not item.get("text"):
        return
    content = OxmlElement("w:txbxContent")
    content.append(_text_paragraph(item))
    _float_shape(paragraph, item, content, behind=False)


def _scratch():
    """One temporary document. Built paragraphs are moved out of it immediately."""
    global _SCRATCH
    if _SCRATCH is None:
        from docx import Document

        _SCRATCH = Document()
    return _SCRATCH


_SCRATCH = None


def _place_table(paragraph, item: dict) -> None:
    from core.docx_converter import (
        _definition_weights,
        _fit_table_height,
        _ink_on,
        _line_count,
        _lock_row,
        _pad_table_cells,
        _set_cell_width,
        _set_paragraph_ltr,
        _set_paragraph_rtl,
        _set_run_font,
        text_direction,
        _set_table_borders,
        _shade_box,
    )

    rows = item.get("rows") or []
    if not rows or item["w"] < 12 or item["h"] < 12:
        return
    columns = max(len(row) for row in rows)
    scratch = _scratch()
    table = scratch.add_table(rows=len(rows), cols=columns)
    table.autofit = False
    table.allow_autofit = False
    _set_table_borders(table, show=bool(item.get("rules")))
    defined = _definition_weights(rows, list(item.get("widths") or []), columns)
    if defined:
        weights = defined
    else:
        weights = list(item.get("widths") or [])
        if len(weights) != columns or sum(weights) <= 0:
            weights = [1.0] * columns
    total = sum(weights) or 1.0
    dxa = [max(80, int(item["w"] * 20 * weight / total)) for weight in weights]
    _fix_table_width(table, dxa)
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for node, width in zip(grid.findall(qn("w:gridCol")), dxa):
            node.set(qn("w:w"), str(width))
    fills = item.get("fills") or []
    for row_index, row in enumerate(rows):
        padded = list(row) + [""] * (columns - len(row))
        for column, text in enumerate(padded):
            cell = table.cell(row_index, column)
            _set_cell_width(cell, dxa[column])
            cell_paragraph = cell.paragraphs[0]
            rtl_cell = text_direction(text) == "rtl"
            if rtl_cell:
                _set_paragraph_rtl(cell_paragraph, text)
            else:
                _set_paragraph_ltr(cell_paragraph)
            cell_paragraph.paragraph_format.space_before = Pt(0)
            cell_paragraph.paragraph_format.space_after = Pt(0)
            cell_paragraph.paragraph_format.line_spacing = 1.0
            if text:
                run = cell_paragraph.add_run(text)
                _set_run_font(run, 10.0, row_index == 0 and item.get("rules"), text)
                if rtl_cell and re.fullmatch(r"[\d٠-٩]{1,4}", text):
                    cell_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    justification = cell_paragraph._p.get_or_add_pPr().find(qn("w:jc"))
                    if justification is not None:
                        justification.set(qn("w:val"), "left")
            fill = ""
            if row_index < len(fills) and column < len(fills[row_index]):
                fill = fills[row_index][column]
            if fill:
                _shade_box(cell._tc.get_or_add_tcPr(), fill)
                ink = _ink_on(fill)
                if ink is not None and text:
                    run.font.color.rgb = ink
        _lock_row(table.rows[row_index])
    _pad_table_cells(table, 100 if defined else 80, 60 if defined else 40)
    if defined:
        heights = []
        for row in rows:
            padded = list(row) + [""] * (columns - len(row))
            lines = 1
            for index, cell_text in enumerate(padded):
                column_width = item["w"] * (weights[index] / 100.0) - 8.0
                lines = max(lines, _line_count(str(cell_text), max(24.0, column_width), 10.0))
            heights.append(max(item["h"] / len(rows), lines * 13.5 + 6.0))
        cap = max(item["h"], item["h"] * 1.35)
        if sum(heights) > cap:
            scale = cap / sum(heights)
            heights = [height * scale for height in heights]
        item["h"] = sum(heights)
    else:
        heights = [max(10.0, item["h"] / len(rows))] * len(rows)
    _fit_table_height(table, heights)
    content = OxmlElement("w:txbxContent")
    table_element = table._tbl
    table_element.getparent().remove(table_element)
    content.append(table_element)
    _float_shape(paragraph, item, content, behind=True)


def _fix_table_width(table, dxa: list[int]) -> None:
    properties = table._tbl.tblPr
    table_width = OxmlElement("w:tblW")
    table_width.set(qn("w:w"), str(sum(dxa)))
    table_width.set(qn("w:type"), "dxa")
    existing = properties.find(qn("w:tblW"))
    if existing is not None:
        properties.remove(existing)
    properties.append(table_width)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    existing_layout = properties.find(qn("w:tblLayout"))
    if existing_layout is not None:
        properties.remove(existing_layout)
    properties.append(layout)


def _paint_ink(run, item: dict) -> None:
    """The PDF's ink. A light bar must not turn red heading text black."""
    ink = (item.get("ink") or "").strip()
    if len(ink) != 6:
        return
    from docx.shared import RGBColor

    run.font.color.rgb = RGBColor.from_string(ink)


def _text_paragraph(item: dict):
    from core.docx_converter import _set_paragraph_ltr, _set_paragraph_rtl, _set_run_font, text_direction

    scratch = _scratch()
    paragraph = scratch.add_paragraph()
    rtl_text = text_direction(item["text"]) == "rtl"
    if rtl_text:
        _set_paragraph_rtl(paragraph, item["text"])
    else:
        _set_paragraph_ltr(paragraph)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    if item["align"] == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        justification = paragraph._p.get_or_add_pPr().find(qn("w:jc"))
        if justification is not None:
            justification.set(qn("w:val"), "center")
    elif rtl_text and item["align"] == "left":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        justification = paragraph._p.get_or_add_pPr().find(qn("w:jc"))
        if justification is not None:
            justification.set(qn("w:val"), "right")
    elif not rtl_text and item["align"] == "right":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        justification = paragraph._p.get_or_add_pPr().find(qn("w:jc"))
        if justification is not None:
            justification.set(qn("w:val"), "right")
    if item.get("toc"):
        size = min(max(item["size"], 11.0), 14.0)
        title = paragraph.add_run(item["text"])
        _set_run_font(title, size, item["bold"], item["text"])
        _paint_ink(title, item)
        paragraph.add_run("\t")
        number = paragraph.add_run(item["toc"])
        _set_run_font(number, size, False, item["toc"])
        paragraph.paragraph_format.tab_stops.add_tab_stop(
            Pt(item.get("tab") or max(120.0, item["w"] - 6.0)),
            WD_TAB_ALIGNMENT.RIGHT,
            WD_TAB_LEADER.DOTS,
        )
    else:
        run = paragraph.add_run(item["text"])
        _set_run_font(run, min(max(item["size"], 6.5), 28.0), item["bold"], item["text"])
        _paint_ink(run, item)
    element = paragraph._p
    element.getparent().remove(element)
    return element


def _float_drawing(run, left: float, top: float, width: float, height: float, behind: bool) -> None:
    drawing = run._r.find(qn("w:drawing"))
    if drawing is None:
        return
    inline = drawing.find(qn("wp:inline"))
    if inline is None:
        return
    extent = inline.find(qn("wp:extent"))
    doc_pr = inline.find(qn("wp:docPr"))
    graphic = inline.find(qn("a:graphic"))
    if extent is None or graphic is None:
        return
    extent.set("cx", str(_pt(width)))
    extent.set("cy", str(_pt(height)))
    if doc_pr is not None:
        doc_pr.set("id", str(_next_id()))
    anchor = _anchor(left, top, width, height, behind)
    anchor.append(extent)
    effect = OxmlElement("wp:effectExtent")
    for edge in ("l", "t", "r", "b"):
        effect.set(edge, "0")
    anchor.append(effect)
    anchor.append(OxmlElement("wp:wrapNone"))
    if doc_pr is not None:
        anchor.append(doc_pr)
    locks = OxmlElement("wp:cNvGraphicFramePr")
    frame_locks = OxmlElement("a:graphicFrameLocks")
    frame_locks.set("noChangeAspect", "1")
    locks.append(frame_locks)
    anchor.append(locks)
    anchor.append(graphic)
    drawing.remove(inline)
    drawing.append(anchor)


def _float_shape(paragraph, item: dict, content, behind: bool) -> None:
    run = paragraph.add_run()
    drawing = OxmlElement("w:drawing")
    anchor = _anchor(item["x"], item["y"], item["w"], item["h"], behind)
    extent = OxmlElement("wp:extent")
    extent.set("cx", str(_pt(item["w"])))
    extent.set("cy", str(_pt(item["h"])))
    anchor.append(extent)
    effect = OxmlElement("wp:effectExtent")
    for edge in ("l", "t", "r", "b"):
        effect.set(edge, "0")
    anchor.append(effect)
    anchor.append(OxmlElement("wp:wrapNone"))
    doc_pr = OxmlElement("wp:docPr")
    doc_pr.set("id", str(_next_id()))
    doc_pr.set("name", f"Box {_DRAWING_ID}")
    anchor.append(doc_pr)
    locks = OxmlElement("wp:cNvGraphicFramePr")
    anchor.append(locks)
    graphic = OxmlElement("a:graphic")
    data = OxmlElement("a:graphicData")
    data.set("uri", _SHAPE_URI)
    shape = etree.Element(f"{{{_WPS}}}wsp", nsmap={"wps": _WPS})
    properties = etree.SubElement(shape, f"{{{_WPS}}}cNvSpPr")
    properties.set("txBox", "1")
    sp_pr = etree.SubElement(shape, f"{{{_WPS}}}spPr")
    xfrm = OxmlElement("a:xfrm")
    off = OxmlElement("a:off")
    off.set("x", "0")
    off.set("y", "0")
    ext = OxmlElement("a:ext")
    ext.set("cx", str(_pt(item["w"])))
    ext.set("cy", str(_pt(item["h"])))
    xfrm.append(off)
    xfrm.append(ext)
    sp_pr.append(xfrm)
    geom = OxmlElement("a:prstGeom")
    geom.set("prst", "rect")
    geom.append(OxmlElement("a:avLst"))
    sp_pr.append(geom)
    fill = (item.get("fill") or "").strip()
    if len(fill) == 6:
        solid = OxmlElement("a:solidFill")
        color = OxmlElement("a:srgbClr")
        color.set("val", fill.upper())
        solid.append(color)
        sp_pr.append(solid)
    else:
        sp_pr.append(OxmlElement("a:noFill"))
    line = OxmlElement("a:ln")
    line.append(OxmlElement("a:noFill"))
    sp_pr.append(line)
    text_box = etree.SubElement(shape, f"{{{_WPS}}}txbx")
    text_box.append(content)
    body = etree.SubElement(shape, f"{{{_WPS}}}bodyPr")
    body.set("wrap", "square")
    body.set("lIns", "0")
    body.set("tIns", "0")
    body.set("rIns", "0")
    body.set("bIns", "0")
    body.set("rtlCol", "1")
    body.set("anchor", "t")
    body.append(OxmlElement("a:noAutofit"))
    data.append(shape)
    graphic.append(data)
    anchor.append(graphic)
    drawing.append(anchor)
    run._r.append(drawing)


def _anchor(left: float, top: float, width: float, height: float, behind: bool):
    anchor = OxmlElement("wp:anchor")
    anchor.set("distT", "0")
    anchor.set("distB", "0")
    anchor.set("distL", "0")
    anchor.set("distR", "0")
    anchor.set("simplePos", "0")
    anchor.set("relativeHeight", str(200 + _DRAWING_ID))
    anchor.set("behindDoc", "1" if behind else "0")
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "1")
    anchor.set("allowOverlap", "1")
    simple = OxmlElement("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    anchor.append(simple)
    position_h = OxmlElement("wp:positionH")
    position_h.set("relativeFrom", "page")
    offset_h = OxmlElement("wp:posOffset")
    offset_h.text = str(_pt(left))
    position_h.append(offset_h)
    anchor.append(position_h)
    position_v = OxmlElement("wp:positionV")
    position_v.set("relativeFrom", "page")
    offset_v = OxmlElement("wp:posOffset")
    offset_v.text = str(_pt(top))
    position_v.append(offset_v)
    anchor.append(position_v)
    return anchor
