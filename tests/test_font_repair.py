# -*- coding: utf-8 -*-
"""Glyphs whose stored letters are wrong are read from their shapes."""

from __future__ import annotations

from pathlib import Path

import fitz
import pytest

from core import font_repair
from core.docx_converter import _atoms_on_page, _cluster_baselines, _logical_text

FILES = Path(__file__).resolve().parent / "corpus" / "files"
BROKEN = FILES / "broken_cmap_arabic.pdf"        # every glyph stored as a Latin letter or symbol
SPLIT = FILES / "ligature_split_arabic.pdf"      # lam-alef stored as U+FFFD / control codes
EXPECTED = ["سياسة التعاملات مع الأطراف ذوي العلاقة", "المناقضة بين المنشآت لا يجوز"]

needs_font = pytest.mark.skipif(
    not font_repair._candidates("NotoNaskhArabic-Regular"),
    reason="Noto Naskh Arabic is not installed here",
)


def _lines(path: Path) -> list[str]:
    page = fitz.open(path)[0]
    return [_logical_text(band) for band in _cluster_baselines(_atoms_on_page(page))]


@needs_font
def test_a_fully_broken_map_is_read_from_shapes():
    assert "PQ" in fitz.open(BROKEN)[0].get_text() or "•" in fitz.open(BROKEN)[0].get_text()
    assert _lines(BROKEN) == EXPECTED


@needs_font
def test_a_few_broken_ligatures_are_read_from_shapes():
    assert _lines(SPLIT) == EXPECTED


def test_shared_shapes_prefer_the_arabic_alphabet():
    assert font_repair._settle({"ي", "ی"}) == "ي"
    assert font_repair._settle({"-", "−"}) is None
    assert font_repair._settle({"ب"}) == "ب"


def test_without_the_installed_font_nothing_changes():
    page = fitz.open(FILES / "related_party_bilingual.pdf")[2]
    traces = page.get_texttrace()
    # No installed copy of this PDF's fonts in a bare test run means no map,
    # and the traces come back untouched.
    if not any(font_repair._candidates(name) for name in {t["font"] for t in traces}):
        assert font_repair.repaired_traces(page, traces) == traces


def test_repair_can_be_switched_off(monkeypatch):
    monkeypatch.setenv("PDF_TOOL_SHAPE_REPAIR", "0")
    page = fitz.open(BROKEN)[0]
    text = "".join(atom.text for atom in _atoms_on_page(page))
    assert not any("؀" <= char <= "ۿ" for char in text)
