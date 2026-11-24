# -*- coding: utf-8 -*-
"""Folders for the desktop app. The install directory is never used for data."""

from __future__ import annotations

import os
import sys
from pathlib import Path

from engine.version import APP_NAME


def _local_app_data() -> Path:
    configured = os.environ.get("LOCALAPPDATA", "").strip()
    if configured:
        return Path(configured)
    return Path.home() / "AppData" / "Local"


def app_home() -> Path | None:
    """
    The writable root, or None when this process is a source checkout.

    Source runs keep the project folders the tests already use. The packaged
    app, and any launch that sets PDF_TOOL_HOME, uses %LOCALAPPDATA%\\PDFTool.
    """
    override = os.environ.get("PDF_TOOL_HOME", "").strip()
    if override:
        return Path(override)
    if getattr(sys, "frozen", False):
        return _local_app_data() / APP_NAME
    return None


def using_app_data() -> bool:
    return app_home() is not None


def _ensure(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def _under(name: str, legacy: Path) -> Path:
    home = app_home()
    if home is None:
        return _ensure(legacy)
    return _ensure(home / name)


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def cache_dir() -> Path:
    override = os.environ.get("PDF_TOOL_CACHE", "").strip()
    if override:
        return _ensure(Path(override))
    return _under("cache", project_root() / ".extract_cache")


def learned_file() -> Path:
    override = os.environ.get("PDF_TOOL_LEARNED", "").strip()
    if override:
        return Path(override)
    home = app_home()
    if home is not None:
        return home / "learned" / "learned_corrections.json"
    return project_root() / "learned_corrections.json"


def crops_dir() -> Path:
    override = os.environ.get("PDF_TOOL_CROPS", "").strip()
    if override:
        return Path(override)
    return _under("crops", project_root() / "learned_crops")


def page_cache_dir() -> Path:
    override = os.environ.get("PDF_TOOL_PAGE_CACHE", "").strip()
    if override:
        return _ensure(Path(override))
    return _under("pages", project_root() / ".page_cache")


def log_dir() -> Path:
    override = os.environ.get("PDF_TOOL_LOG_DIR", "").strip()
    if override:
        return _ensure(Path(override))
    return _under("logs", project_root() / ".logs")


def temp_dir() -> Path:
    override = os.environ.get("PDF_TOOL_TEMP", "").strip()
    if override:
        return _ensure(Path(override))
    home = app_home()
    if home is not None:
        return _ensure(home / "temp")
    return _ensure(project_root() / ".tmp")


def checkpoint_dir() -> Path:
    override = os.environ.get("PDF_TOOL_CHECKPOINT", "").strip()
    if override:
        return _ensure(Path(override))
    home = app_home()
    if home is not None:
        return _ensure(home / "checkpoints")
    return _ensure(project_root() / ".extract_cache" / "partial")


def benchmark_file() -> Path:
    home = app_home()
    if home is not None:
        _ensure(home)
        return home / "benchmark.json"
    return project_root() / ".extract_cache" / "benchmark.json"


def vision_dir() -> Path:
    return _under("vision", project_root() / ".vision_cache")


def apply_temp_env() -> None:
    """Point the process temp folders at the app data temp directory."""
    if not using_app_data():
        return
    folder = temp_dir()
    os.environ["PDF_TOOL_TEMP"] = str(folder)
    os.environ["TEMP"] = str(folder)
    os.environ["TMP"] = str(folder)
    import tempfile

    tempfile.tempdir = str(folder)
