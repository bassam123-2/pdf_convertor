# -*- coding: utf-8 -*-
"""Text that must survive the trip from a PDF page into a Word frame."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import fitz
import pytest

from core.arabic_fixer import _Atom, _logical_text, plain_arabic
from core.docx_converter import (
    WordOptions,
    _layer_numbers,
    _read_footer,
    _restore_layer_numbers,
    _restore_quoted_letters,
    convert_arabic_pdf_to_docx,
    extract_page,
)

PDF = Path(
    r"c:\Users\jouan\.cursor\projects\c-Users-jouan-Desktop-PDF-TOOL\attachments"
    r"\f37a5719-9abf-4e51-91c7-bb5f2017e553\______________________________________________.pdf"
)
FONT = Path(r"C:\Windows\Fonts\trado.ttf")


def _atom(text: str, x: float, level: int, width: float | None = None) -> _Atom:
    span = width if width is not None else max(4.0, 7.0 * max(1, len(text)))
    return _Atom(text, x, 0.0, x + span, 12.0, x, 10.0, level, 14.0, False)


def _words(text: str) -> list[str]:
    cleaned = plain_arabic(text).replace("\u06be", "\u0647")
    cleaned = re.sub(r"([\u0600-\u06FF])([0-9A-Za-z])", r"\1 \2", cleaned)
    cleaned = re.sub(r"([0-9A-Za-z])([\u0600-\u06FF])", r"\1 \2", cleaned)
    return re.findall(r"[0-9A-Za-z\u0600-\u06FF]+", cleaned)


def _docx_text(xml: str) -> str:
    return " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


def word_delta(source: str, output: str) -> tuple[list[str], list[str]]:
    """Words present on the page but missing from the docx, and words that were added."""
    source_words = _words(source)
    output_words = _words(output)
    lost: list[str] = []
    spare = list(output_words)
    for word in source_words:
        if word in spare:
            spare.remove(word)
        else:
            lost.append(word)
    return lost, spare


def test_plain_arabic_folds_a_bad_heh_form_and_keeps_urdu():
    # U+FBAB is the presentation form this kind of font writes for ordinary heh.
    assert plain_arabic("ولائحت\ufbab") == "ولائحته"
    # U+06BE is a real letter in Urdu. It must not be rewritten.
    assert plain_arabic("ٹھی") == "ٹھی"
    assert "\u06be" in plain_arabic("ٹھ\u06be")


def test_quoted_single_letters_stay_in_logical_order():
    word = list("قحلم")
    atoms = []
    x = 10.0
    for char, level in (('"', 0), ("ب", 1), ('"', 0), (" ", 1)):
        atoms.append(_atom(char, x, level, 8))
        x += 10
    for char in word:
        atoms.append(_atom(char, x, 1, 8))
        x += 9
    assert _logical_text(atoms) == 'ملحق "ب"'

    atoms = []
    x = 10.0
    for char in ('"', "أ", '"', " "):
        atoms.append(_atom(char, x, 1, 8))
        x += 10
    for char in word:
        atoms.append(_atom(char, x, 1, 8))
        x += 9
    assert _logical_text(atoms) == 'ملحق "أ"'


def test_numbers_and_a_citation_keep_their_order():
    atoms = [_atom("1", 10, 0, 8), _atom(",", 18, 1, 4), _atom("500", 24, 0, 24)]
    atoms.append(_atom(" ", 50, 1, 4))
    x = 60.0
    for char in "قحلم":
        atoms.append(_atom(char, x, 1, 8))
        x += 9
    assert _logical_text(atoms) == "ملحق 1,500"

    atoms = [_atom("2021", 10, 0, 28), _atom("-", 40, 1, 6), _atom("2022", 48, 0, 28)]
    atoms.append(_atom(" ", 80, 1, 4))
    x = 90.0
    for char in "قحلم":
        atoms.append(_atom(char, x, 1, 8))
        x += 9
    assert _logical_text(atoms) == "ملحق 2021-2022"

    citation = list(")Smith, 2021(")
    atoms = [_atom(char, 10 + index * 7, 0, 6) for index, char in enumerate(citation)]
    atoms.append(_atom(" ", 130, 1, 4))
    x = 140.0
    for char in "قحلم":
        atoms.append(_atom(char, x, 1, 8))
        x += 9
    assert _logical_text(atoms) == "ملحق (Smith, 2021)"


def test_footer_on_a_real_page_keeps_spaces_heh_and_the_page_word():
    if not PDF.is_file():
        pytest.skip("sample PDF is not attached")
    document = fitz.open(PDF)
    try:
        title, label, number, _gap = _read_footer(document[4])
    finally:
        document.close()
    assert title == "دليل الامتثال لنظام المنافسة ولائحته التنفيذية"
    assert "\u06be" not in title
    assert label == "الصفحة"
    assert number == 5


def test_quoted_appendix_letters_survive_on_the_contents_page():
    if not PDF.is_file():
        pytest.skip("sample PDF is not attached")
    document = fitz.open(PDF)
    try:
        blocks = extract_page(document[1])
    finally:
        document.close()
    text = "\n".join(getattr(block, "text", "") for block in blocks)
    assert 'الملحق "أ"' in text
    assert 'الملحق "ب"' in text
    assert 'الملحق "ج"' in text


def test_page_five_keeps_the_hijri_dates_and_the_corner():
    if not PDF.is_file():
        pytest.skip("sample PDF is not attached")
    document = fitz.open(PDF)
    try:
        blocks = extract_page(document[4])
    finally:
        document.close()
    text = "\n".join(getattr(block, "text", "") for block in blocks)
    assert "وتاريخ 1440/06/29" in text
    assert "1441/01/25" in text
    images = [block for block in blocks if type(block).__name__ == "ImageBlock"]
    corners = [
        block
        for block in images
        if block.bbox[1] > 680 and (block.bbox[3] - block.bbox[1]) < 180
    ]
    assert corners, "the corner mark should stay a picture of its own"
    banners = [
        block
        for block in images
        if block.bbox[1] < 180 and (block.bbox[3] - block.bbox[1]) < 80
    ]
    assert banners, "the introduction banner should stay one picture"


def test_synthetic_page_words_match_the_docx(tmp_path: Path):
    if not FONT.is_file():
        pytest.skip("Traditional Arabic is not installed")
    source = tmp_path / "sample.pdf"
    target = tmp_path / "sample.docx"
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    page.insert_font(fontname="trad", fontfile=str(FONT))
    written = 'الملحق "ب" 2021-2022 (Smith, 2021) 1,500'
    page.insert_textbox(
        fitz.Rect(72, 120, 520, 180),
        written,
        fontname="trad",
        fontsize=16,
        align=fitz.TEXT_ALIGN_RIGHT,
    )
    page.insert_text(
        fitz.Point(360, 810),
        "دليل الامتثال ولائحته",
        fontname="trad",
        fontsize=9,
    )
    document.save(source)
    document.close()
    convert_arabic_pdf_to_docx(source, target, options=WordOptions(mode="editable", page_spec="1"))
    with zipfile.ZipFile(target) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
        footer = ""
        if "word/footer1.xml" in archive.namelist():
            footer = archive.read("word/footer1.xml").decode("utf-8")
    plain = _docx_text(xml)
    opened = fitz.open(source)
    try:
        body = opened[0].get_text("text").splitlines()[0]
    finally:
        opened.close()
    lost, added = word_delta(body, plain)
    assert lost == [], f"lost {lost} added {added}"
    assert "2021" in plain and "Smith" in plain
    assert "\u06be" not in footer


def test_contents_page_docx_keeps_the_quoted_letters(tmp_path: Path):
    if not PDF.is_file():
        pytest.skip("sample PDF is not attached")
    target = tmp_path / "page2.docx"
    convert_arabic_pdf_to_docx(PDF, target, options=WordOptions(mode="editable", page_spec="2"))
    with zipfile.ZipFile(target) as archive:
        xml = _docx_text(archive.read("word/document.xml").decode("utf-8"))
        footer = ""
        for name in archive.namelist():
            if name.startswith("word/footer"):
                footer += _docx_text(archive.read(name).decode("utf-8"))
    assert 'الملحق "أ"' in xml
    assert 'الملحق "ب"' in xml
    assert 'الملحق "ج"' in xml
    assert "دليل الامتثال لنظام المنافسة ولائحته التنفيذية" in footer
    assert "الصفحة" in footer
    assert "\u06be" not in footer


def _placed(text: str, x0: float, x1: float, level: int = 0) -> _Atom:
    return _Atom(text, x0, 0.0, x1, 12.0, x0, 10.0, level, 14.0, False)


def test_a_digit_glued_by_a_broken_map_is_not_part_of_the_number():
    atoms = [_placed("\u0640", 10, 16, 1)]
    cursor = 16.0
    for char in "61440/06/29":
        atoms.append(_placed(char, cursor, cursor + 6))
        cursor += 6
    tokens = [text for _y, text, _host in _layer_numbers(atoms)]
    assert tokens == ["1440/06/29"]
    year = [_placed("ا", 0, 8, 1)]
    cursor = 8.0
    for char in "2021":
        year.append(_placed(char, cursor, cursor + 6))
        cursor += 6
    assert _layer_numbers(year) == []


def test_a_page_with_no_text_layer_is_left_alone():
    document = fitz.open()
    page = document.new_page()
    page.draw_rect(fitz.Rect(40, 40, 220, 90), color=(0, 0, 0), fill=(0.9, 0.9, 0.9))
    words = [_Atom("Hello ", 48, 48, 120, 64, 48, 60, 0, 12, False)]
    restored = _restore_layer_numbers(page, _restore_quoted_letters(page, words))
    document.close()
    assert restored[0].text == "Hello "


def test_english_page_keeps_its_size_and_does_not_grow_an_arabic_footer(tmp_path: Path):
    source = tmp_path / "english.pdf"
    target = tmp_path / "english.docx"
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    page.insert_text(
        fitz.Point(72, 120),
        "The committee met in 2021-2022 and paid 1,500.",
        fontname="helv",
        fontsize=12,
    )
    page.insert_text(
        fitz.Point(72, 144),
        "A second line sits just below the first one.",
        fontname="helv",
        fontsize=12,
    )
    page.insert_text(fitz.Point(72, 800), "Guide note", fontname="helv", fontsize=9)
    page.insert_text(fitz.Point(400, 800), "Page 2", fontname="helv", fontsize=9)
    document.save(source)
    document.close()
    convert_arabic_pdf_to_docx(source, target, options=WordOptions(mode="editable", page_spec="1"))
    with zipfile.ZipFile(target) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
        footer = ""
        for name in archive.namelist():
            if name.startswith("word/footer") and name.endswith(".xml"):
                footer += _docx_text(archive.read(name).decode("utf-8"))
    plain = _docx_text(xml)
    assert "committee" in plain and "1,500" in plain and "2021-2022" in plain
    sizes = [int(value) for value in re.findall(r'w:sz w:val="(\d+)"', xml)]
    assert sizes and max(sizes) >= 22
    assert "الصفحة" not in footer
    assert "|" not in footer


def test_arabic_with_a_good_map_stays_digital():
    if not FONT.is_file():
        pytest.skip("Traditional Arabic is not installed")
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    page.insert_font(fontname="trad", fontfile=str(FONT))
    page.insert_textbox(
        fitz.Rect(72, 120, 520, 180),
        'الملحق "ب" يوضح القاعدة كما كتبت في النص الأصلي دون تغيير.',
        fontname="trad",
        fontsize=16,
        align=fitz.TEXT_ALIGN_RIGHT,
    )
    kind: dict[str, str] = {}
    blocks = extract_page(page, kind_out=kind)
    document.close()
    text = "\n".join(getattr(block, "text", "") for block in blocks)
    assert kind.get("kind") == "digital"
    assert '"ب"' in text
    assert "\u06be" not in text


def test_a_ruled_table_keeps_its_words():
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    page.insert_text(
        fitz.Point(72, 80),
        "The regional table lists every office and its staff count for the year.",
        fontname="helv",
        fontsize=12,
    )
    rows = [("North office", "12"), ("South office", "18"), ("Total offices", "30")]
    top = 120.0
    for index, (name, count) in enumerate(rows):
        page.insert_text(fitz.Point(80, top + 18 + index * 24), name, fontname="helv", fontsize=12)
        page.insert_text(fitz.Point(300, top + 18 + index * 24), count, fontname="helv", fontsize=12)
    for y in (top, top + 24, top + 48, top + 72):
        page.draw_line(fitz.Point(70, y), fitz.Point(400, y), width=0.6)
    for x in (70, 250, 400):
        page.draw_line(fitz.Point(x, top), fitz.Point(x, top + 72), width=0.6)
    blocks = extract_page(page)
    document.close()
    text = "\n".join(getattr(block, "text", "") or " ".join(cell for row in getattr(block, "rows", []) for cell in row) for block in blocks)
    assert "North" in text and "South" in text and "Total" in text
