# -*- coding: utf-8 -*-
"""PyInstaller runtime hook. Runs before the window, including in page workers."""

from __future__ import annotations

import multiprocessing
import os
import sys


if getattr(sys, "frozen", False):
    local = os.environ.get("LOCALAPPDATA", "").strip()
    if local and not os.environ.get("PDF_TOOL_HOME", "").strip():
        os.environ["PDF_TOOL_HOME"] = os.path.join(local, "PDFTool")
    os.environ.setdefault("PDF_TOOL_OFFLINE", "1")
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    try:
        from engine.host import prepare_server

        prepare_server()
    except Exception:
        pass

multiprocessing.freeze_support()
