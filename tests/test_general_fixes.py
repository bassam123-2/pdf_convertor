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
    build = read_word_build(GUIDE, options=WordOptions(mode="editable", page_spec="2,9"))
    # Both pages store Arabic letters as Latin ones: no flowing text for them.
    # They are read the way earlier versions read them (OCR text in frames,
    # or the faithful layout), not dropped into a page picture by the flow.
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
