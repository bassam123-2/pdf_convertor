# -*- coding: utf-8 -*-
"""English word gaps and column gutters in the Word engine."""

from __future__ import annotations

from pathlib import Path

import fitz
import pytest

from core.docx_converter import (
    WordOptions,
    _Atom,
    _atoms_on_page,
    _cluster_baselines,
    _gap_spaces,
    _lines_on_page,
    _split_gutter,
    convert_arabic_pdf_to_docx,
)
from core.arabic_fixer import _logical_text

RESUME = Path(
    r"c:\Users\jouan\.cursor\projects\c-Users-jouan-Desktop-PDF-TOOL\attachments"
    r"\da28dd0d-1afe-444e-be3d-e5ea56fa536e\Hamza_Ahmad_-_Resume_-_Nov.2026.pdf"
)


def _english_atom(char: str, x0: float, size: float = 12.0, space: float = 3.0) -> _Atom:
    width = size * (0.55 if char != " " else 0.25)
    return _Atom(char, x0, 10.0, x0 + width, 22.0, x0, 20.0, 0, size, False, "000000", "Arial", False, 0, space)


def test_gap_spaces_insert_between_english_words():
    # "Audit" then a pen move, then "Manager" — no painted space glyph.
    atoms = []
    x = 10.0
    for char in "Audit":
        atoms.append(_english_atom(char, x))
        x = atoms[-1].x1
    x += 4.0  # pen jump ≈ one space
    for char in "Manager":
        atoms.append(_english_atom(char, x))
        x = atoms[-1].x1
    filled = atoms + _gap_spaces(atoms)
    assert "Audit Manager" == _logical_text(filled)


def test_split_gutter_cuts_a_narrow_sidebar():
    # Phone on the left, body on the right. The hole sits left of mid-page.
    left = [_english_atom(char, 20.0 + index * 7.0) for index, char in enumerate("+966")]
    right = [_english_atom(char, 280.0 + index * 7.0, size=13.0) for index, char in enumerate("body")]
    parts = _split_gutter(left + right, page_width=960.0)
    assert len(parts) == 2
    assert "".join(atom.text for atom in parts[0]).startswith("+966")
    assert "".join(atom.text for atom in parts[1]).startswith("body")


@pytest.mark.skipif(not RESUME.is_file(), reason="resume sample is not on this machine")
def test_resume_contact_stays_out_of_the_body_column():
    page = fitz.open(RESUME)[0]
    atoms = _atoms_on_page(page)
    lines = _lines_on_page(page, atoms, tables=[], figures=[])
    phone = [line for line in lines if "+966" in line.text]
    assert phone, "phone line missing"
    assert all("coordinating" not in line.text for line in phone)
    mail = [line for line in lines if "gmail.com" in line.text.lower()]
    assert mail
    assert all("remediation" not in line.text for line in mail)


@pytest.mark.skipif(not RESUME.is_file(), reason="resume sample is not on this machine")
def test_resume_faithful_keeps_page_size_and_spaced_english(tmp_path: Path):
    import re
    import zipfile

    out = tmp_path / "resume.docx"
    convert_arabic_pdf_to_docx(RESUME, out, options=WordOptions(mode="faithful", fit_a4=False))
    with zipfile.ZipFile(out) as archive:
        xml = archive.read("word/document.xml").decode("utf-8", errors="replace")
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
    assert 'w:w="19200"' in xml and 'w:h="10800"' in xml
    assert "Internal Audit Manager" in text
    assert "InternalAuditManager" not in text
    assert xml.count("w:txbxContent") > 20
    assert xml.count("a:blip") >= 1
    fonts = set(re.findall(r'w:ascii="([^"]+)"', xml))
    assert "Georgia" in fonts
