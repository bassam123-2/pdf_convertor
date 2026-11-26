# -*- coding: utf-8 -*-
"""
Check each converted page against its PDF page, before anything is written.

The words the PDF page prints (read glyph by glyph, as the quality checker
does) are compared with the words the converter kept for that page. A page
that lost many words, or gained many the PDF does not have, is read again
the other way; the reading that keeps more of the page's words wins. Words
the converter could not be sure of are listed so the user knows where to
look.

Arabic and English are counted the same way: a word is the letters between
spaces, with marks and tatweel taken away, and an Arabic word read backwards
still matches.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field, fields, is_dataclass

import fitz

_ARABIC = re.compile(r"[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]")
_LATIN = re.compile(r"[A-Za-zÀ-ɏ]")
# Running headers and footers are written once per section, not with the page.
_CHROME = 0.08


@dataclass
class PageScore:
    number: int
    source: int
    recall: float
    precision: float
    unsure: list[str] = field(default_factory=list)

    @property
    def weak(self) -> bool:
        return self.source >= 25 and (self.recall < 0.7 or self.precision < 0.6)

    def better_than(self, other: "PageScore") -> bool:
        return (self.recall + self.precision) > (other.recall + other.precision) + 0.05


def _bag(words: list[str]) -> Counter:
    return Counter(min(word, word[::-1]) if _ARABIC.search(word) else word for word in words)


def source_words(page: fitz.Page, boxes: list | None = None) -> list[str]:
    """The trusted words the PDF page prints in its body (and their boxes into `boxes`)."""
    from core.quality_check import _READABLE, _mixed_script, _trace_words, normalize_word

    height = float(page.rect.height)
    words = []
    seen: dict[str, list] = {}
    for text, rect in _trace_words(page):
        if rect.y1 < height * _CHROME or rect.y0 > height * (1 - _CHROME):
            continue
        word = normalize_word(text)
        if not word or not _READABLE.match(word) or _mixed_script(word):
            continue
        if any(abs(r.x0 - rect.x0) < 2.5 and abs(r.y0 - rect.y0) < 2.5 for r in seen.get(word, [])):
            continue
        seen.setdefault(word, []).append(rect)
        words.append(word)
        if boxes is not None:
            boxes.append(rect)
    return words


def _texts(value, out: list[str], depth: int = 0) -> None:
    """Every piece of text inside a page's blocks, tables and nested groups."""
    if depth > 6 or value is None:
        return
    if isinstance(value, str):
        out.append(value)
        return
    if isinstance(value, (bytes, bytearray, int, float, bool)):
        return
    if isinstance(value, (list, tuple)):
        for item in value:
            _texts(item, out, depth + 1)
        return
    if isinstance(value, dict):
        for item in value.values():
            _texts(item, out, depth + 1)
        return
    if is_dataclass(value):
        for spec in fields(value):
            if spec.name in {"text", "rows", "blocks", "children", "lines", "tables", "items", "cells"}:
                _texts(getattr(value, spec.name, None), out, depth + 1)


def _over_pictures(block, pictures: list, text_boxes: list | None = None) -> bool:
    """A table read from a picture: it lies on one, and the PDF prints no text there."""
    box = getattr(block, "bbox", None)
    if not box or type(block).__name__ != "TableBlock":
        return False
    rect = fitz.Rect(box)
    area = max(1.0, rect.width * rect.height)
    if not any((rect & picture).get_area() > area * 0.5 for picture in pictures):
        return False
    printed = sum(1 for word in text_boxes or [] if rect.contains(fitz.Point((word.x0 + word.x1) / 2, (word.y0 + word.y1) / 2)))
    return printed < 5


def output_words(record, pictures: list | None = None, text_boxes: list | None = None) -> list[str]:
    """
    The words a converted page keeps. A table read from a picture is left
    out: the PDF has no text for it to be held against.
    """
    from core.quality_check import normalize_word

    pieces: list[str] = []
    blocks = [block for block in (getattr(record, "blocks", None) or []) if not _over_pictures(block, pictures or [], text_boxes)]
    _texts(blocks, pieces)
    payload = getattr(record, "faithful", None)
    if payload is not None:
        _texts(getattr(payload, "lines", None), pieces)
        _texts(getattr(payload, "tables", None), pieces)
    words = []
    for piece in pieces:
        for raw in piece.split():
            word = normalize_word(raw)
            if word:
                words.append(word)
    return words


def score(page: fitz.Page, record, number: int) -> PageScore:
    text_boxes: list = []
    words = source_words(page, text_boxes)
    # Words the page keeps inside a picture (an infographic's labels) are
    # still on the page; they are not held against the text.
    kept_as_picture = [
        fitz.Rect(block.bbox) for block in (getattr(record, "blocks", None) or [])
        if type(block).__name__ == "ImageBlock" and getattr(block, "bbox", None)
        and getattr(block, "fit", "") != "page"
    ]
    if kept_as_picture:
        pairs = [
            (word, box) for word, box in zip(words, text_boxes)
            if not any(picture.contains(fitz.Point((box.x0 + box.x1) / 2, (box.y0 + box.y1) / 2)) for picture in kept_as_picture)
        ]
        words = [word for word, _box in pairs]
        text_boxes = [box for _word, box in pairs]
    source = _bag(words)
    try:
        pictures = [fitz.Rect(info["bbox"]) for info in page.get_image_info() if info.get("bbox")]
    except Exception:
        pictures = []
    out_words = output_words(record, pictures, text_boxes)
    result = _bag(out_words)
    common = sum((source & result).values())
    total_source = sum(source.values())
    total_result = sum(result.values())
    recall = common / total_source if total_source else 1.0
    precision = common / total_result if total_result else (1.0 if not total_source else 0.0)
    unsure = sorted({w for w in out_words if _ARABIC.search(w) and _LATIN.search(w)})
    # Words written more often than the page prints them: text read twice.
    # Words the PDF has no text for (a table read from a picture) are not
    # counted against the page.
    doubled = sum(max(0, result[word] - count) for word, count in source.items())
    precision = 1.0 - doubled / total_source if total_source else precision
    return PageScore(number, total_source, round(recall, 3), round(max(0.0, min(1.0, precision)), 3), unsure[:8])
