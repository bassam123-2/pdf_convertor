# -*- coding: utf-8 -*-
"""Round 4: tables, charts, footers, classifier for figure plates."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz
import pytest

from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx, read_word_build
from core.docx_validate import validate_docx
from core.page_strategy import classify_page
from core import table_rebuild

FILES = Path(__file__).resolve().parent / "corpus" / "files"
CHART = FILES / "chart_with_labels.pdf"
RASTER = FILES / "table_raster.pdf"
BORDERED = FILES / "table_bordered.pdf"
LINKS = FILES / "links_sample.pdf"


@pytest.mark.skipif(not CHART.is_file(), reason="chart sample missing")
def test_chart_page_stays_faithful_not_scanned():
    document = fitz.open(CHART)
    plan = classify_page(document[0], default_mode="faithful")
    document.close()
    assert plan.strategy == "faithful"
    assert "figure" in plan.reason or "picture" in plan.reason or "embedded" in plan.reason


@pytest.mark.skipif(not CHART.is_file(), reason="chart sample missing")
def test_chart_page_keeps_picture_without_stray_labels(tmp_path: Path):
    build = read_word_build(CHART, options=WordOptions(mode="faithful", fit_a4=False))
    assert build.pages
    payload = build.pages[0].faithful
    assert payload is not None
    assert payload.background, "chart page must keep a background picture"
    # Labels such as 90.66% must not appear as free text lines.
    joined = " ".join(line.text for line in payload.lines)
    assert "90.66" not in joined
    assert "99.62" not in joined
    notes = " ".join(payload.notes).lower()
    assert "chart" in notes or "figure" in notes


@pytest.mark.skipif(not RASTER.is_file(), reason="raster table sample missing")
def test_raster_table_classified_and_rebuilt():
    document = fitz.open(RASTER)
    page = document[0]
    regions = table_rebuild.classify_regions(page)
    kinds = {r.kind for r in regions}
    assert "raster_table" in kinds
    tables = table_rebuild.rebuild_tables(page, regions)
    document.close()
    # OCR may be weak on CI without tessdata; require either rows or a fallback picture.
    assert tables
    assert tables[0].rows or tables[0].fallback_png


@pytest.mark.skipif(not BORDERED.is_file(), reason="bordered table sample missing")
def test_vector_table_native_in_faithful(tmp_path: Path):
    build = read_word_build(BORDERED, options=WordOptions(mode="faithful", fit_a4=False))
    payload = build.pages[0].faithful
    assert payload is not None
    # Either a rebuilt native table or vector detection notes.
    if payload.tables:
        assert payload.tables[0].rows
        assert payload.tables[0].kind == "vector"
    else:
        # find_tables may miss synthetic grids; geometric fallback should still fire via regions.
        document = fitz.open(BORDERED)
        regions = table_rebuild.classify_regions(document[0])
        document.close()
        assert any(r.kind == "vector_table" for r in regions)


@pytest.mark.skipif(not CHART.is_file(), reason="chart sample missing")
def test_page_label_becomes_page_field(tmp_path: Path):
    out = tmp_path / "chart.docx"
    counts = convert_arabic_pdf_to_docx(CHART, out, options=WordOptions(mode="faithful", fit_a4=False))
    assert out.is_file()
    validation = counts.get("docx_validation") or validate_docx(out)
    assert validation["ok"], validation.get("errors")
    with zipfile.ZipFile(out) as archive:
        xml = archive.read("word/document.xml").decode("utf-8", errors="replace")
        # The label sits at the top of this sample, so its field is in the
        # header part; a label at the foot would be in a footer part.
        footers = [name for name in archive.namelist() if "footer" in name.lower() or "header" in name.lower()]
        blob = xml + "".join(archive.read(name).decode("utf-8", errors="replace") for name in footers)
    assert "PAGE" in blob or "w:fldChar" in blob


@pytest.mark.skipif(not LINKS.is_file(), reason="links sample missing")
def test_links_preserved():
    build = read_word_build(LINKS, options=WordOptions(mode="faithful", fit_a4=False))
    payload = build.pages[0].faithful
    assert payload is not None
    kinds = {link.kind for link in payload.links}
    assert "uri" in kinds
    assert "mailto" in kinds
    assert "goto" in kinds


def test_docx_validate_accepts_empty_word_doc(tmp_path: Path):
    from docx import Document

    path = tmp_path / "plain.docx"
    Document().save(path)
    report = validate_docx(path)
    assert report["ok"]


def test_is_page_label_handles_spaced_page_word():
    from core.docx_converter import _is_page_label

    assert _is_page_label("8 | Page")
    assert _is_page_label("8 | P a g e")
    assert _is_page_label("10 Page")
