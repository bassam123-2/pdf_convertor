# -*- coding: utf-8 -*-
"""Command-line entry for the Excel engine. The code lives in core.excel_converter."""

from __future__ import annotations

from core.excel_converter import ConversionError, ExcelWriteCount, convert_pdf_to_excel, main

__all__ = ["ConversionError", "ExcelWriteCount", "convert_pdf_to_excel", "main"]


def __getattr__(name: str):
    from core import excel_converter

    return getattr(excel_converter, name)


if __name__ == "__main__":
    raise SystemExit(main())
