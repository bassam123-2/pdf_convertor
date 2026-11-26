# -*- coding: utf-8 -*-
"""
One way to call PyMuPDF's ``page.find_tables``.

Newer PyMuPDF releases gate the line finder with a layout analyser
(``use_layout=True`` by default) that can crash the whole process on some
Word-exported pages. The converter only needs the ruled-line finder, so the
layout step is turned off wherever the installed release offers it.
"""

from __future__ import annotations

import inspect


def _layout_switch() -> bool:
    try:
        import pymupdf.table as table

        return "use_layout" in inspect.signature(table.find_tables).parameters
    except Exception:
        return False


_HAS_LAYOUT = _layout_switch()


def find_tables(page, **kwargs):
    if _HAS_LAYOUT:
        kwargs.setdefault("use_layout", False)
    return page.find_tables(**kwargs)
