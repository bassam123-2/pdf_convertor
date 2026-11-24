# -*- coding: utf-8 -*-
"""Temporary files for one conversion, removed as soon as the result is delivered."""

from __future__ import annotations

import gc
import os
import shutil
import tempfile
from contextlib import contextmanager
from pathlib import Path


@contextmanager
def conversion_workspace(prefix: str):
    """
    A private folder for the PDF and the output of one conversion.

    The folder, the uploaded PDF, and the generated file are deleted when the
    caller leaves this block, including when the conversion raises. The bytes
    returned to the user are already in memory by then.
    """
    parent = os.environ.get("PDF_TOOL_TEMP", "").strip() or None
    folder = Path(tempfile.mkdtemp(prefix=prefix, dir=parent))
    try:
        yield folder
    finally:
        shutil.rmtree(folder, ignore_errors=True)
        gc.collect()


def read_output(path: Path) -> bytes:
    """Read the finished file, then delete that file before the folder itself is removed."""
    data = path.read_bytes()
    try:
        path.unlink()
    except OSError:
        pass
    return data
