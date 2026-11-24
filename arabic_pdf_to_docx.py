# -*- coding: utf-8 -*-
"""Command-line entry for the Word engine. The code lives in core.docx_converter."""

from __future__ import annotations

from core.docx_converter import ConversionError, convert_arabic_pdf_to_docx, main

__all__ = ["ConversionError", "convert_arabic_pdf_to_docx", "main"]


def __getattr__(name: str):
    from core import docx_converter

    return getattr(docx_converter, name)


if __name__ == "__main__":
    raise SystemExit(main())
