# -*- coding: utf-8 -*-
"""
Choose how each PDF page should be read, before conversion starts.

The choice is based on properties of the page (text layer, drawings, images,
aspect ratio), never on a file name or a page number. A caller can override
the choice per page through WordOptions.page_strategies.
"""

from __future__ import annotations

from core.safe_tables import find_tables as _safe_find_tables

from dataclasses import asdict, dataclass

import fitz


STRATEGIES = (
    "faithful",  # design / slide / mixed: background picture + text boxes
    "editable",  # flowing paragraphs and tables
    "scanned",  # OCR / picture page
    "table",  # ruled grid dominates
    "broken_font",  # text layer present but Unicode map is wrong
)


@dataclass(frozen=True)
class PagePlan:
    number: int
    strategy: str
    reason: str
    width: float
    height: float
    text_chars: int
    drawings: int
    images: int

    def as_dict(self) -> dict:
        return asdict(self)


def _text_chars(page: fitz.Page) -> int:
    try:
        return sum(len((span.get("text") or "").strip()) for block in page.get_text("dict").get("blocks") or [] if block.get("type") == 0 for line in block.get("lines") or [] for span in line.get("spans") or [])
    except Exception:
        return len((page.get_text("text") or "").strip())


def _drawing_count(page: fitz.Page) -> int:
    try:
        return len(page.get_drawings() or [])
    except Exception:
        return 0


def _image_count(page: fitz.Page) -> int:
    try:
        return len(page.get_images(full=True) or [])
    except Exception:
        return 0


def _looks_broken(page: fitz.Page) -> bool:
    """Private-use or replacement characters dominate the text layer."""
    try:
        text = page.get_text("text") or ""
    except Exception:
        return False
    letters = [char for char in text if char.isalpha() or ("\u0600" <= char <= "\u06FF")]
    if len(letters) < 20:
        return False
    bad = sum(1 for char in letters if ord(char) == 0xFFFD or 0xE000 <= ord(char) <= 0xF8FF)
    return bad / len(letters) >= 0.25


def _ruled_table(page: fitz.Page) -> bool:
    try:
        found = _safe_find_tables(page, vertical_strategy="lines", horizontal_strategy="lines")
    except Exception:
        return False
    tables = list(getattr(found, "tables", None) or [])
    return any(len(getattr(table, "rows", []) or []) >= 3 for table in tables)


def _image_coverage(page: fitz.Page) -> float:
    """Fraction of the page covered by the largest embedded image (0–1)."""
    try:
        infos = page.get_image_info() or []
    except Exception:
        return 0.0
    page_area = max(1.0, float(page.rect.width) * float(page.rect.height))
    best = 0.0
    for info in infos:
        box = info.get("bbox")
        if not box:
            continue
        area = max(0.0, (float(box[2]) - float(box[0])) * (float(box[3]) - float(box[1])))
        best = max(best, area / page_area)
    return best


def classify_page(page: fitz.Page, default_mode: str = "faithful") -> PagePlan:
    """Pick one strategy for this page and record why."""
    number = int(page.number) + 1
    width = float(page.rect.width)
    height = float(page.rect.height)
    chars = _text_chars(page)
    drawings = _drawing_count(page)
    images = _image_count(page)
    # Chart / raster-table plates: a short text layer (often just a page label)
    # plus embedded figures that do not cover the whole sheet. Faithful keeps
    # those pictures. A true scan has essentially no text layer and needs OCR.
    # Glyphs drawn as vector paths (no text operators), with or without a small
    # logo image: keep faithful/picture. OCR rebuilds crushed tables / displaced Arabic.
    if chars < 40 and drawings >= 20 and default_mode not in {"editable", "flow", "flowing"}:
        coverage = _image_coverage(page) if images else 0.0
        if coverage < 0.92:
            return PagePlan(
                number,
                "faithful",
                "dense vector drawings with no text layer (outlined glyphs); keep as page picture",
                width,
                height,
                chars,
                drawings,
                images,
            )
    if chars < 40 and images >= 1:
        coverage = _image_coverage(page)
        if (
            chars >= 5
            and coverage < 0.92
            and default_mode not in {"editable", "flow", "flowing"}
        ):
            return PagePlan(
                number,
                "faithful",
                "embedded figure(s) with a short text layer; keep pictures in faithful background (not OCR scan)",
                width,
                height,
                chars,
                drawings,
                images,
            )
        return PagePlan(
            number,
            "scanned",
            "little or no text layer; page looks like a scan or a full-page picture",
            width,
            height,
            chars,
            drawings,
            images,
        )
    if chars < 3 and drawings == 0 and images == 0:
        return PagePlan(number, "scanned", "little or no text layer; page looks like a scan or a picture", width, height, chars, drawings, images)
    # A short page of real text with no picture on it is not a scan: there is
    # nothing to read from pixels, and OCR would throw the text layer away.
    if _looks_broken(page):
        return PagePlan(number, "broken_font", "text layer has many unmapped characters", width, height, chars, drawings, images)
    landscape = width > height * 1.15
    dense_art = drawings >= 8 or (drawings >= 3 and images >= 1)
    if landscape and dense_art:
        return PagePlan(number, "faithful", "landscape page with many drawings or pictures (slide/design)", width, height, chars, drawings, images)
    if dense_art and chars < 1200:
        return PagePlan(number, "faithful", "vector fills and pictures dominate the page", width, height, chars, drawings, images)
    if _ruled_table(page) and default_mode == "editable":
        return PagePlan(number, "table", "ruled grid with several rows", width, height, chars, drawings, images)
    if default_mode in {"editable", "flow", "flowing"}:
        return PagePlan(number, "editable", "flowing text; caller asked for editable mode", width, height, chars, drawings, images)
    return PagePlan(number, "faithful", "digital text; faithful layout keeps positions", width, height, chars, drawings, images)


def classify_document(document: fitz.Document, default_mode: str = "faithful", overrides: dict[int, str] | None = None) -> list[PagePlan]:
    """One plan per page. `overrides` maps 1-based page numbers to a strategy name."""
    overrides = overrides or {}
    plans: list[PagePlan] = []
    for index in range(document.page_count):
        page = document.load_page(index)
        plan = classify_page(page, default_mode=default_mode)
        forced = overrides.get(plan.number) or overrides.get(str(plan.number))
        if forced:
            name = str(forced).strip().lower()
            if name not in STRATEGIES:
                raise ValueError(f"Unknown page strategy {forced!r}. Use one of {STRATEGIES}.")
            plan = PagePlan(plan.number, name, f"override: {forced}", plan.width, plan.height, plan.text_chars, plan.drawings, plan.images)
        plans.append(plan)
    return plans
