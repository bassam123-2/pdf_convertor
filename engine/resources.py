# -*- coding: utf-8 -*-
"""How many page workers to run. The reading of each page stays the same."""

from __future__ import annotations

import os


def reading_mode() -> str:
    """accurate is the default. fast uses fewer workers and smaller batches."""
    value = os.environ.get("PDF_TOOL_MODE", "accurate").strip().lower()
    return "fast" if value == "fast" else "accurate"


def worker_cap() -> int:
    return 4 if reading_mode() == "fast" else 6


def worker_count(jobs: int) -> int:
    """CPU count minus one, never below 1, never above the mode cap or the job count."""
    cpu = os.cpu_count() or 1
    available = max(1, cpu - 1)
    count = max(0, int(jobs))
    if count <= 0:
        return 1
    return max(1, min(worker_cap(), available, count))


def batch_size() -> int:
    """How many pages stay in memory together."""
    return 2 if reading_mode() == "fast" else 3
