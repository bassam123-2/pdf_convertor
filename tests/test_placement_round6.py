# -*- coding: utf-8 -*-
"""Round 6: fake-bold dedupe, one RTL box with runs, clamp inside cells."""

from __future__ import annotations

from pathlib import Path

import fitz
import pytest
import runpy

from core import faithful
from core.faithful import _dedupe_fake_bold, _line_box
from core.docx_converter import _atoms_on_page

FILES = Path(__file__).resolve().parent / "corpus" / "files"
FAKE = FILES / "fake_bold_overprint.pdf"
RTL = FILES / "rtl_mixed_runs.pdf"
BILINGUAL = FILES / "bilingual_bordered_table.pdf"
BUILD = Path(__file__).resolve().parent / "corpus" / "_build_round6_samples.py"


@pytest.fixture(scope="module", autouse=True)
def _ensure_samples():
    if not FAKE.is_file() or not RTL.is_file() or not BILINGUAL.is_file():
        runpy.run_path(str(BUILD))


def test_fake_bold_overprint_emits_each_word_once():
    document = fitz.open(FAKE)
    atoms = _atoms_on_page(document[0], upright_only=True)
    document.close()
    raw_letters = [atom for atom in atoms if atom.text.strip()]
    # Overprint doubles the first phrase's letters.
    assert len(raw_letters) > 20
    cleaned = _dedupe_fake_bold(atoms)
    kept_letters = [atom for atom in cleaned if atom.text.strip()]
    assert len(kept_letters) < len(raw_letters)
    # Survivor of the overprint pair is marked bold.
    assert any(atom.bold for atom in cleaned if atom.text.strip() and atom.origin_y < 100)
    # "Once Only" was drawn once — eight letters, no overprint twin.
    once = [atom for atom in cleaned if abs(atom.origin_y - 120) < 1 and atom.text.strip()]
    assert len(once) == 8


def test_rtl_mixed_styles_stay_one_box():
    document = fitz.open(RTL)
    payload = faithful.read_page(document[0])
    rules = payload.column_rules
    document.close()
    arabic = [
        line
        for line in payload.lines
        if any(0x0600 <= ord(ch) <= 0x06FF for ch in line.text)
    ]
    assert arabic, "expected an Arabic line"
    assert len(arabic) == 1
    line = arabic[0]
    assert line.rtl
    # Style changes live as runs (chunks) inside the same box.
    assert len(line.chunks) >= 1
    left, _top, width, _height, flags = _line_box(
        line, 1.0, 0.0, 0.0, payload.width, payload.height, rules
    )
    assert left >= 28.0
    assert left + width <= 392.0
    assert "outside_cell" not in flags


def test_bilingual_words_stay_inside_cells():
    document = fitz.open(BILINGUAL)
    payload = faithful.read_page(document[0])
    rules = payload.column_rules
    mid = 300.0
    for line in payload.lines:
        left, _t, width, _h, flags = _line_box(
            line, 1.0, 0.0, 0.0, payload.width, payload.height, rules
        )
        center = (line.x0 + line.x1) / 2.0
        if center > mid:
            assert left >= mid - 1.0, (line.text, left, flags)
            assert left + width <= 556.0, (line.text, left, width, flags)
        else:
            assert left + width <= mid + 1.0, (line.text, left, width, flags)
        assert "outside_cell" not in flags
    hits = sum(line.text.count("الأخلاقية") for line in payload.lines)
    assert hits <= 1
    document.close()


def test_read_page_dedupes_fake_bold_end_to_end():
    document = fitz.open(FAKE)
    payload = faithful.read_page(document[0])
    document.close()
    joined = " ".join(line.text for line in payload.lines)
    assert joined.count("Boldish") == 1
