# -*- coding: utf-8 -*-
"""Faults found by the quality checks stay fixed. None of them is tied to one file."""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path

import fitz
import pytest

from core import flow_layout, ligature_repair
from core.docx_converter import TextBlock, WordOptions, convert_arabic_pdf_to_docx, read_word_build
from core.faithful import _dedupe_fake_bold, _overfull_arabic_text

ROOT = Path(__file__).resolve().parent.parent
FILES = ROOT / "tests" / "corpus" / "files"
RELATED = FILES / "related_party_bilingual.pdf"
BORDERED = FILES / "table_bordered.pdf"
FAKE_BOLD = FILES / "fake_bold_overprint.pdf"


@dataclass
class Glyph:
    text: str
    x0: float
    x1: float
    origin_y: float = 100.0
    size: float = 12.0
    y0: float = 90.0
    y1: float = 102.0
    origin_x: float = 0.0
    bold: bool = False
    flags: int = 0


# --- lam-alef ligatures with a wrong Unicode value -------------------------


def test_shifted_ligature_codes_map_back_exactly():
    assert [ligature_repair.truncated_ligature(chr(c)) for c in range(0x3F5, 0x3FD)] == [
        "لآ", "لآ", "لأ", "لأ", "لإ", "لإ", "لا", "لا",
    ]
    assert ligature_repair.truncated_ligature("a") is None


def test_foreign_letter_in_an_arabic_word_uses_the_documents_own_spelling():
    # "الأطراف" stored visually with a Canadian letter in place of لأ.
    letters = ["ف", "ا", "ر", "ط", "ᕓ", "ا"]
    glyphs = [Glyph(t, 100 - i * 6, 106 - i * 6) for i, t in enumerate(letters)]
    known = ligature_repair.vocabulary(["الأطراف ذوي العلاقة"])
    assert ligature_repair.repair_atoms(glyphs, known) == 1
    assert glyphs[4].text == "لأ"


def test_a_greek_letter_in_english_text_is_left_alone():
    glyphs = [Glyph(t, 10 + i * 6, 16 + i * 6) for i, t in enumerate("ϻab")]
    assert ligature_repair.repair_atoms(glyphs, None) == 0
    assert glyphs[0].text == "ϻ"


# --- doubled letters are not fake bold --------------------------------------


def test_double_letters_survive_fake_bold_removal():
    # "all": the two l glyphs are 2.7 pt apart, one glyph width.
    glyphs = [Glyph("a", 0, 5.3), Glyph("l", 5.3, 8.0), Glyph("l", 8.0, 10.7)]
    assert "".join(g.text for g in _dedupe_fake_bold(glyphs)) == "all"


def test_an_overprinted_copy_is_still_removed():
    glyphs = [Glyph("B", 0, 8), Glyph("B", 0.4, 8.4)]
    kept = _dedupe_fake_bold(glyphs)
    assert len(kept) == 1 and kept[0].bold


def test_ordinary_arabic_words_are_not_called_twins():
    assert not _overfull_arabic_text("التعاملات", 40, 12)
    assert not _overfull_arabic_text("تلاماعتلا", 40, 12)
    assert _overfull_arabic_text("تعامتتعاملات", 40, 12)


# --- flowing editable layout -------------------------------------------------


def _block(text, x0, y0, x1, y1):
    return TextBlock("paragraph", (x0, y0, x1, y1), text)


def test_two_side_by_side_columns_are_found():
    blocks = []
    for row in range(6):
        y = 100 + row * 20
        blocks.append(_block("English line", 40, y, 280, y + 12))
        blocks.append(_block("سطر عربي", 320, y, 560, y + 12))
    gutter = flow_layout.find_gutter(blocks, 600)
    assert gutter is not None and 280 <= gutter <= 320
    kinds = [kind for kind, _ in flow_layout._runs(blocks, gutter, 600)]
    assert kinds == ["columns"]


def test_one_column_page_has_no_gutter():
    blocks = [_block("line", 40, 100 + i * 20, 560, 112 + i * 20) for i in range(8)]
    assert flow_layout.find_gutter(blocks, 600) is None


def test_digits_stay_with_the_arabic_around_them():
    assert flow_layout._split_scripts("مقدمة .1") == ["مقدمة .1"]
    assert flow_layout._split_scripts("شركة Mezzan القابضة") == ["شركة ", "Mezzan", " القابضة"]


@pytest.mark.skipif(not RELATED.is_file(), reason="sample missing")
def test_editable_mode_writes_flowing_text_not_boxes(tmp_path):
    out = tmp_path / "related.docx"
    convert_arabic_pdf_to_docx(RELATED, out, options=WordOptions(mode="editable", page_spec="3"))
    with zipfile.ZipFile(out) as package:
        xml = package.read("word/document.xml").decode("utf-8")
    assert xml.count("<w:txbxContent") <= 2
    assert "Building a culture of professional conduct" in xml
    assert "التعاملات" in xml and "all employees" in xml


# --- pages and tables ----------------------------------------------------------


@pytest.mark.skipif(not RELATED.is_file(), reason="sample missing")
def test_cover_page_keeps_its_arabic_title():
    build = read_word_build(RELATED, options=WordOptions(mode="faithful", page_spec="1"))
    lines = [line.text for line in build.pages[0].faithful.lines]
    assert "سياسة التعاملات مع الأطراف ذوي" in lines
    assert "Related Parties Transactions" in lines


@pytest.mark.skipif(not FAKE_BOLD.is_file(), reason="sample missing")
def test_short_text_page_without_pictures_is_not_a_scan():
    build = read_word_build(FAKE_BOLD, options=WordOptions(mode="faithful"))
    assert build.pages[0].kind == "digital"
    assert build.pages[0].faithful is not None


@pytest.mark.skipif(not BORDERED.is_file(), reason="sample missing")
def test_floating_table_sits_on_its_drawn_grid(tmp_path):
    build = read_word_build(BORDERED, options=WordOptions(mode="faithful"))
    table = build.pages[0].faithful.tables[0]
    assert [round(v) for v in table.bbox] == [72, 110, 480, 290]
    assert [round(w) for w in table.widths] == [128, 140, 140]
    out = tmp_path / "bordered.docx"
    convert_arabic_pdf_to_docx(BORDERED, out, options=WordOptions(mode="faithful"))
    with zipfile.ZipFile(out) as package:
        xml = package.read("word/document.xml").decode("utf-8")
    assert "tblpXSpec" not in xml and "tblpYSpec" not in xml
    assert xml.index("<w:tblpPr") < xml.index("<w:tblW ")


def test_list_items_start_their_own_paragraph():
    from core.html_layout import merge_paragraphs

    lines = [
        _block("1. Interests regarding investment or", 40, 100, 280, 112),
        _block("ownership in a commercial activity", 40, 113, 270, 125),
        _block("2. A commercial activity or any facility", 40, 126, 280, 138),
    ]
    merged = merge_paragraphs(lines)
    assert [block.text[:2] for block in merged] == ["1.", "2."]


def test_side_by_side_columns_keep_printed_order(tmp_path):
    from docx import Document

    blocks = []
    for row in range(5):
        y = 100 + row * 20
        blocks.append(_block(f"English {row}", 40, y, 280, y + 12))
        blocks.append(_block(f"عربي {row}", 320, y, 560, y + 12))
    document = Document()
    flow_layout.write_flow_page(document, blocks, 600, 800, {})
    xml = document.element.xml
    assert '<w:bidiVisual w:val="0"/>' in xml
    assert xml.index("English 0") < xml.index("عربي 0")
    assert xml.count("<w:tblLayout") == 1


GUIDE = FILES / "competition_guide_ar.pdf"


@pytest.mark.skipif(not GUIDE.is_file(), reason="sample missing")
def test_broken_font_pages_stay_positioned_in_editable_mode():
    from core.font_repair import _candidates

    build = read_word_build(GUIDE, options=WordOptions(mode="editable", page_spec="2,9"))
    if _candidates("SakkalMajalla"):
        # The installed font reads the broken glyphs from their shapes: the
        # pages are sound text and may flow.
        texts = " ".join(getattr(b, "text", "") for page in build.pages for b in page.blocks)
        assert "المناهضة" in texts or any(page.frames for page in build.pages)
    else:
        # Both pages store Arabic letters as Latin ones: no flowing text for
        # them. They are read the way earlier versions read them (OCR text in
        # frames, or the faithful layout), not dropped into a page picture.
        assert all(page.frames for page in build.pages)


@pytest.mark.skipif(not RELATED.is_file(), reason="sample missing")
def test_clean_pages_still_flow():
    build = read_word_build(RELATED, options=WordOptions(mode="editable", page_spec="3"))
    assert build.pages[0].faithful is None and build.pages[0].blocks


def test_link_box_holds_no_visible_copy_of_the_address(tmp_path):
    links = FILES / "links_sample.pdf"
    if not links.is_file():
        pytest.skip("sample missing")
    out = tmp_path / "links.docx"
    convert_arabic_pdf_to_docx(links, out, options=WordOptions(mode="faithful"))
    with zipfile.ZipFile(out) as package:
        xml = package.read("word/document.xml").decode("utf-8")
    assert ">Page " not in xml


# --- invisible paint and coloured-cell tables -------------------------------


def _invisible_box_pdf(path):
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    # A black rectangle at opacity 0 under a heading, as Word leaves under
    # linked contents entries.
    page.draw_rect(fitz.Rect(60, 90, 400, 112), color=None, fill=(0, 0, 0), fill_opacity=0)
    page.insert_text((70, 106), "Contents entry under an invisible box", fontsize=12)
    for row in range(6):
        page.insert_text((70, 160 + row * 20), f"Body line {row} with ordinary words in it", fontsize=11)
    document.save(path)


def test_invisible_paint_is_not_a_black_bar(tmp_path):
    from core.docx_converter import _drawing_paints
    from core.table_rebuild import visible_drawings

    pdf = tmp_path / "invisible.pdf"
    _invisible_box_pdf(pdf)
    page = fitz.open(pdf)[0]
    assert page.get_drawings()  # the PDF does hold the rectangle
    assert visible_drawings(page) == []
    assert _drawing_paints(page) == []
    out = tmp_path / "invisible.docx"
    for mode in ("editable", "faithful"):
        convert_arabic_pdf_to_docx(pdf, out, options=WordOptions(mode=mode))
        with zipfile.ZipFile(out) as package:
            xml = package.read("word/document.xml").decode("utf-8")
        assert 'val="000000"' not in xml.replace('w:color w:val="000000"', "")


def _colour_table_pdf(path):
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    terms = ["Kingdom", "System", "Rules", "Guide", "Authority"]
    for row, term in enumerate(terms):
        top = 150 + row * 26
        page.draw_rect(fitz.Rect(440, top, 540, top + 26), color=None, fill=(0.16, 0.73, 0.69))
        page.insert_text((450, top + 17), term, fontsize=11, color=(1, 1, 1))
        page.insert_text((80, top + 17), f"The meaning of {term.lower()} in this guide", fontsize=11)
    document.save(path)


def test_coloured_cell_table_is_a_table_in_both_readers(tmp_path):
    import pdfplumber

    from core.docx_converter import TableBlock, extract_page

    pdf = tmp_path / "colour_table.pdf"
    _colour_table_pdf(pdf)
    page = fitz.open(pdf)[0]
    with pdfplumber.open(pdf) as plumber:
        blocks = extract_page(page, plumber.pages[0], set(), {}, "auto")
    tables = [b for b in blocks if isinstance(b, TableBlock)]
    assert len(tables) == 1 and len(tables[0].rows) == 5
    assert tables[0].rows[0][1] == "Kingdom"
    out = tmp_path / "colour_table.docx"
    convert_arabic_pdf_to_docx(pdf, out, options=WordOptions(mode="editable"))
    with zipfile.ZipFile(out) as package:
        xml = package.read("word/document.xml").decode("utf-8")
    assert "<w:tbl>" in xml and "Authority" in xml


EXTRACT = FILES / "guide_extract_7p.pdf"


@pytest.mark.skipif(not EXTRACT.is_file(), reason="sample missing")
def test_glyphs_drawn_as_shapes_are_not_colour_bars():
    from core.docx_converter import _drawing_paints

    document = fitz.open(EXTRACT)
    for page in document:
        assert "000000" not in {paint[4] for paint in _drawing_paints(page)}


def test_a_box_is_a_rectangle_not_a_letter():
    from core.table_rebuild import is_box

    rect = fitz.Rect(0, 0, 100, 20)
    assert is_box({"items": [("re", rect, 1)], "rect": rect})
    letter = [("l", fitz.Point(0, 0), fitz.Point(1, 1))] * 9 + [("c", 0, 0, 0, 0)] * 6
    assert not is_box({"items": letter, "rect": fitz.Rect(0, 0, 8, 10)})


SPLIT_GUIDE = FILES / "guide_split_7p.pdf"


@pytest.mark.skipif(not SPLIT_GUIDE.is_file(), reason="sample missing")
def test_label_cells_with_line_highlights_keep_their_rows():
    from core import table_rebuild

    page = fitz.open(SPLIT_GUIDE)[3]
    tables = table_rebuild.fill_tables(page)
    assert len(tables) == 1 and len(tables[0].rows) == 3


@pytest.mark.skipif(not SPLIT_GUIDE.is_file(), reason="sample missing")
def test_coloured_table_wins_over_a_line_grid_found_on_it(monkeypatch):
    # Newer PyMuPDF builds see the white gaps of a coloured table as a ruled
    # grid; that grid then failed to rebuild and the page lost its table.
    from core import table_rebuild

    class Grid:
        bbox = (52, 92, 546, 312)

    class Found:
        tables = [Grid()]

    page = fitz.open(SPLIT_GUIDE)[3]
    monkeypatch.setattr(fitz.Page, "find_tables", lambda self, *a, **k: Found())
    monkeypatch.setattr(table_rebuild, "_is_prose_column_frame", lambda page, table: False)
    kinds = [region.kind for region in table_rebuild.classify_regions(page) if "table" in region.kind]
    assert kinds == ["fill_table"]


# --- Word tables: thin border bars, shading boxes, pictures in cells ---------

LT13 = FILES / "lt13_programme_en.pdf"


def _word_style_table(path: Path) -> None:
    """A 3-column, 3-row table drawn the way Word exports it."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    xs = [50, 120, 260, 545]
    ys = [100, 120, 170, 220]
    grey = (0.8, 0.83, 0.88)
    for r in range(3):
        for c in range(3):
            cell = fitz.Rect(xs[c], ys[r], xs[c + 1], ys[r + 1])
            shade = (0.06, 0.09, 0.17) if r == 0 else ((0.89, 0.91, 0.94) if c == 0 else (1, 1, 1))
            page.draw_rect(cell, color=None, fill=shade)
            # Word shades each text line again inside the cell.
            page.draw_rect(fitz.Rect(cell.x0 + 4, cell.y0 + 4, cell.x1 - 4, cell.y0 + 15), color=None, fill=shade)
            ink = (1, 1, 1) if r == 0 else (0, 0, 0)
            page.insert_text((cell.x0 + 5, cell.y0 + 13), f"R{r}C{c} word", fontsize=8, color=ink)
    for y in ys:
        page.draw_rect(fitz.Rect(xs[0], y - 0.25, xs[-1], y + 0.25), color=None, fill=grey)
    for x in xs:
        page.draw_rect(fitz.Rect(x - 0.25, ys[0], x + 0.25, ys[-1]), color=None, fill=grey)
    doc.save(path)


def test_thin_border_bars_make_the_grid(tmp_path):
    from core.table_rebuild import rule_grids

    pdf = tmp_path / "word_table.pdf"
    _word_style_table(pdf)
    grids = rule_grids(fitz.open(pdf)[0])
    assert len(grids) == 1
    columns, rows_y, _cells = grids[0]
    assert len(columns) == 4 and len(rows_y) == 4


def test_word_table_keeps_three_columns_in_editable_mode(tmp_path):
    pdf = tmp_path / "word_table.pdf"
    _word_style_table(pdf)
    build = read_word_build(pdf, options=WordOptions(mode="editable", fit_a4=True, direction="auto"))
    tables = [b for page in build.pages for b in page.blocks if type(b).__name__ == "TableBlock"]
    assert len(tables) == 1
    assert [len(row) for row in tables[0].rows] == [3, 3, 3]
    assert tables[0].rows[1][2] == "R1C2 word"


def test_find_tables_never_runs_the_layout_step():
    from core import safe_tables

    seen = {}

    class Page:
        def find_tables(self, **kwargs):
            seen.update(kwargs)
            return []

    safe_tables.find_tables(Page(), vertical_strategy="lines")
    if safe_tables._HAS_LAYOUT:
        assert seen.get("use_layout") is False
    assert seen.get("vertical_strategy") == "lines"


def test_digits_inside_a_footer_line_stay(tmp_path):
    from core.docx_converter import _read_footer

    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((60, 820), "Venue x 30 November 2026 x Asia 2026", fontsize=8)
    page.insert_text((530, 820), "7", fontsize=8)
    title, _label, number, _gap = _read_footer(page)
    assert "30 November 2026" in title and title.endswith("2026")
    assert number == 7


def test_raised_mark_does_not_split_its_line():
    from core.docx_converter import _cluster_baselines
    from core.arabic_fixer import _Atom

    def atom(text, x, oy):
        return _Atom(text, x, oy - 10, x + 5, oy + 3, x, oy, 0, 15.0, False)

    line = [atom("ب", 10, 429.8), atom("ت", 20, 429.4), atom("ّ", 30, 425.3), atom("ث", 40, 429.8)]
    assert len(_cluster_baselines(line)) == 1


@pytest.mark.skipif(not LT13.exists(), reason="sample missing")
def test_word_programme_tables_and_photos():
    build = read_word_build(LT13, options=WordOptions(mode="editable", fit_a4=True, direction="auto", page_spec="2-3"))
    agenda, speakers = [[b for b in page.blocks if type(b).__name__ == "TableBlock"] for page in build.pages]
    assert len(agenda) == 1 and len(agenda[0].rows) == 9 and all(len(row) == 3 for row in agenda[0].rows)
    assert agenda[0].rows[0] == ["Time", "Session & Venue", "Session Description & Objectives"]
    assert len(speakers) == 1 and len(speakers[0].images) == 3


def test_cards_side_by_side_become_one_table_row(tmp_path):
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 60, 70), False)
    pix.clear_with(230)
    png = pix.tobytes("png")
    for k, x in enumerate((40, 220, 400)):
        page.insert_image(fitz.Rect(x, 200, x + 160, 400), stream=png)
        for line in range(3):
            page.insert_text((x + 20, 320 + line * 14), f"card{k} line{line} text", fontsize=9)
    pdf = tmp_path / "cards.pdf"
    doc.save(pdf)
    build = read_word_build(pdf, options=WordOptions(mode="editable", fit_a4=True, direction="auto"))
    tables = [b for page in build.pages for b in page.blocks if type(b).__name__ == "TableBlock"]
    assert len(tables) == 1 and len(tables[0].rows) == 1 and len(tables[0].rows[0]) == 3
    assert tables[0].rows[0][0].startswith("card0 line0")
