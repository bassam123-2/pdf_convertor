# -*- coding: utf-8 -*-
"""Bilingual column pages must stay as text boxes, not false w:tbl frames."""

from __future__ import annotations

from pathlib import Path

import fitz
import pytest

from core import faithful, table_rebuild
from core.docx_converter import WordOptions, read_word_build

PDF = Path(
    r"C:\Users\jouan\.cursor\projects\c-Users-jouan-Desktop-PDF-TOOL\attachments"
    r"\da28dd0d-1afe-444e-be3d-e5ea56fa536e\2025-08-Related-Party-Transactions-AR_EN.pdf"
)
if not PDF.is_file():
    PDF = Path(r"C:\Users\jouan\Downloads\2025-08-Related-Party-Transactions-AR_EN.pdf")


@pytest.mark.skipif(not PDF.is_file(), reason="Related Party sample missing")
def test_bilingual_pages_are_not_false_tables():
    document = fitz.open(PDF)
    # Page 4 is the two-column conflict-of-interest list that was clipped.
    regions = table_rebuild.classify_regions(document[3])
    document.close()
    assert not any(region.kind == "vector_table" for region in regions)


@pytest.mark.skipif(not PDF.is_file(), reason="Related Party sample missing")
def test_arabic_column_lines_are_rtl_text_boxes():
    document = fitz.open(PDF)
    payload = faithful.read_page(document[3])
    document.close()
    right = [line for line in payload.lines if line.x0 > 280]
    assert len(right) >= 10
    assert sum(1 for line in right if line.rtl) >= len(right) * 0.7
    assert not payload.tables


@pytest.mark.skipif(not PDF.is_file(), reason="Related Party sample missing")
def test_real_tables_on_page_two_still_detected():
    document = fitz.open(PDF)
    regions = table_rebuild.classify_regions(document[1])
    document.close()
    assert any(region.kind == "vector_table" for region in regions)
