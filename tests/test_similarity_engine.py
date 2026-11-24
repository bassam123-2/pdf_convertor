# -*- coding: utf-8 -*-
"""Round 7: similarity metrics, no-scaling, project state, reconvert plumbing."""

from __future__ import annotations

import json
from pathlib import Path

import fitz
import numpy
import pytest
from PIL import Image

from core import similarity
from core.project_state import (
    ProjectState,
    load_project,
    remember_before_reconvert,
    save_project,
    undo_page_choice,
)

FILES = Path(__file__).resolve().parent / "corpus" / "files"
OUT = Path(__file__).resolve().parent / "corpus" / "_similarity_tmp"


def _write_rgb(path: Path, array: numpy.ndarray) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(array.astype(numpy.uint8), "RGB").save(path)
    return path


def test_ssim_identical_is_high():
    base = numpy.full((64, 64, 3), 220, dtype=numpy.uint8)
    base[10:30, 10:40] = (20, 20, 20)
    a = _write_rgb(OUT / "ident_a.png", base)
    b = _write_rgb(OUT / "ident_b.png", base.copy())
    metrics = similarity._visual_compare(a, b, OUT / "ident_diff.png")
    assert metrics.ssim >= 0.99
    assert metrics.score >= 99.0


def test_ssim_shifted_content_drops():
    base = numpy.full((80, 80, 3), 240, dtype=numpy.uint8)
    shifted = base.copy()
    base[20:40, 20:50] = (10, 10, 10)
    shifted[20:40, 28:58] = (10, 10, 10)  # 8 px ≈ shift
    a = _write_rgb(OUT / "shift_a.png", base)
    b = _write_rgb(OUT / "shift_b.png", shifted)
    metrics = similarity._visual_compare(a, b, OUT / "shift_diff.png")
    assert metrics.ssim < 0.99
    assert (OUT / "shift_diff.png").is_file()


def test_content_metrics_missing_and_garbled():
    assert similarity._garbled_arabic("كلمxة") is True
    assert similarity._garbled_arabic("كلمة") is False
    src = fitz.open()
    page = src.new_page(width=300, height=200)
    page.insert_text((40, 80), "Hello world ethics", fontsize=12)
    dst = fitz.open()
    other = dst.new_page(width=300, height=200)
    other.insert_text((40, 80), "Hello ethics", fontsize=12)  # missing "world"
    other.insert_text((40, 160), "EXTRA", fontsize=12)
    content = similarity._content_metrics(src[0], other)
    src.close()
    dst.close()
    assert content.missing >= 1
    assert content.extra >= 1
    assert content.score < 100.0


def test_placement_metrics_identical_and_shifted():
    doc = fitz.open()
    page = doc.new_page(width=400, height=200)
    page.insert_text((50, 100), "Alpha Beta", fontsize=14)
    # Clone via PDF bytes so we have two page objects.
    data = doc.tobytes()
    doc.close()
    a = fitz.open(stream=data, filetype="pdf")
    b = fitz.open(stream=data, filetype="pdf")
    identical = similarity._placement_metrics(a[0], b[0])
    assert identical.pct_within_1pt >= 99.0
    assert identical.mean_shift_pt <= 0.5
    # Shifted: draw the same words further right on a new page.
    shifted = fitz.open()
    sp = shifted.new_page(width=400, height=200)
    sp.insert_text((50 + 10, 100), "Alpha Beta", fontsize=14)  # ~10 pt shift
    moved = similarity._placement_metrics(a[0], shifted[0])
    a.close()
    b.close()
    shifted.close()
    assert moved.mean_shift_pt >= 5.0
    assert moved.pct_within_1pt < identical.pct_within_1pt


def test_thresholds_per_document_type():
    slides = similarity.thresholds_for("slides")
    text = similarity.thresholds_for("text_document")
    assert slides["weights"]["visual"] > text["weights"]["visual"]
    assert text["target_score"] >= slides["target_score"]


def test_project_state_undo_roundtrip(tmp_path: Path):
    path = tmp_path / "job.project.json"
    state = ProjectState(pdf="a.pdf", docx="a.docx", document_type="text_document")
    choice = state.page(3)
    choice.mode = "faithful"
    choice.score = 91.0
    remember_before_reconvert(choice)
    choice.mode = "editable"
    choice.score = 96.5
    save_project(state, path)
    loaded = load_project(path)
    assert loaded.page(3).mode == "editable"
    assert undo_page_choice(loaded.page(3)) is True
    assert loaded.page(3).mode == "faithful"
    assert loaded.page(3).score == 91.0
    assert undo_page_choice(loaded.page(3)) is False


def test_no_scaling_frame_matches_pdf_page():
    from core.faithful import frame_for

    scale, off_x, off_y, sheet = frame_for(612.0, 792.0, fit_a4=False)
    assert scale == 1.0
    assert abs(sheet[0] - 612.0) < 0.01
    assert abs(sheet[1] - 792.0) < 0.01
    assert abs(off_x) < 0.01 and abs(off_y) < 0.01
    # Landscape stays landscape.
    scale2, _, _, sheet2 = frame_for(960.0, 540.0, fit_a4=False)
    assert scale2 == 1.0
    assert sheet2[0] > sheet2[1]


def test_emu_preserves_subpoint_precision():
    from core.faithful import _emu

    a = _emu(100.0)
    b = _emu(100.04)
    assert b != a
    assert abs((b / 12700.0) - 100.04) < 0.001


def test_fit_uses_spacing_not_fontsize_change():
    """Character spacing / width scale adjusts to target width; size stays."""
    from core.faithful import Chunk, Line, _fit, _word_size

    line = Line(0, 0, 120, 14, 12, 12, False, "Hello world")
    line.chunks = [
        Chunk("Hello", "Helv", 0, 12, False, False, "000000", False, False, 40.0),
        Chunk(" ", "Helv", 0, 12, False, False, "000000", False, True, 4.0),
        Chunk("world", "Helv", 0, 12, False, False, "000000", False, False, 40.0),
    ]
    fitted = _fit(line, 1.0)
    assert _word_size(12, 1.0) == 12.0
    assert len(fitted) == 3
    # Each entry is (spacing_twips, width_percent) — size is not in the tuple.
    for spacing, percent in fitted:
        assert isinstance(spacing, int)
        assert percent >= 88


@pytest.mark.skipif(not (FILES / "table_bordered.pdf").is_file(), reason="corpus pdf missing")
def test_render_pdf_page_and_infer_type():
    pdf = FILES / "table_bordered.pdf"
    rendered = similarity.render_pdf_page(pdf, 1, scale=1.0)
    assert Path(rendered.image_path).is_file()
    assert rendered.width_pt > 0 and rendered.height_pt > 0
    assert rendered.renderer == "pymupdf"
    kind = similarity.infer_document_type(pdf)
    assert kind in {"text_document", "table_heavy", "slides", "scanned"}


def test_compare_page_on_corpus_sample(tmp_path: Path):
    from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx
    from core.visual_check import available_renderers

    if not available_renderers():
        pytest.skip("no Word/LibreOffice renderer")
    pdf = FILES / "table_bordered.pdf"
    if not pdf.is_file():
        pytest.skip("corpus missing")
    docx = tmp_path / "t.docx"
    convert_arabic_pdf_to_docx(pdf, docx, options=WordOptions(mode="faithful", fit_a4=False))
    result = similarity.compare_page(pdf, docx, 1, scale=1.0, out_dir=tmp_path / "cmp")
    assert 0.0 <= result.score <= 100.0
    assert result.size_matches
    assert result.visual.diff_image
    assert result.placement.words >= 1
