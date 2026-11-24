# -*- coding: utf-8 -*-
"""Page classifier picks a strategy from page properties."""

from __future__ import annotations

from pathlib import Path

import fitz
import pytest

from core.page_strategy import classify_document, classify_page

RESUME = Path(__file__).resolve().parent / "corpus" / "files" / "hamza_resume_ppt.pdf"
SCAN = Path(__file__).resolve().parent / "corpus" / "files" / "arabic_news_scan.pdf"


@pytest.mark.skipif(not RESUME.is_file(), reason="resume sample missing")
def test_resume_pages_are_faithful_slides():
    document = fitz.open(RESUME)
    plans = classify_document(document, default_mode="faithful")
    document.close()
    assert plans
    assert all(plan.strategy == "faithful" for plan in plans)
    assert plans[0].width > plans[0].height


@pytest.mark.skipif(not SCAN.is_file(), reason="scan sample missing")
def test_scan_pages_are_marked_scanned():
    document = fitz.open(SCAN)
    plan = classify_page(document[0], default_mode="faithful")
    document.close()
    assert plan.strategy == "scanned"


def test_override_wins():
    document = fitz.open()
    page = document.new_page(width=400, height=600)
    page.insert_text((72, 72), "Hello world " * 20)
    plans = classify_document(document, default_mode="faithful", overrides={1: "editable"})
    document.close()
    assert plans[0].strategy == "editable"
    assert "override" in plans[0].reason
