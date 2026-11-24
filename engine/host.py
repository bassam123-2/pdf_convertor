# -*- coding: utf-8 -*-
"""Process setup for the desktop window. Extraction stays in core."""

from __future__ import annotations

import os
from pathlib import Path

from engine.paths import apply_temp_env
from engine.tesseract_bin import find_exe
from engine.version import VERSION


def server_environment(base: dict[str, str] | None = None) -> dict[str, str]:
    """Environment for the local page. Data stays under %LOCALAPPDATA%\\PDFTool."""
    env = dict(base or os.environ)
    local = env.get("LOCALAPPDATA", "").strip() or str(Path.home() / "AppData" / "Local")
    home = Path(env.get("PDF_TOOL_HOME", "").strip() or str(Path(local) / "PDFTool"))
    folders = {
        "PDF_TOOL_HOME": home,
        "PDF_TOOL_CACHE": home / "cache",
        "PDF_TOOL_LOG_DIR": home / "logs",
        "PDF_TOOL_TEMP": home / "temp",
        "PDF_TOOL_CHECKPOINT": home / "checkpoints",
        "PDF_TOOL_PAGE_CACHE": home / "pages",
        "PDF_TOOL_CROPS": home / "crops",
    }
    for path in folders.values():
        path.mkdir(parents=True, exist_ok=True)
    (home / "learned").mkdir(parents=True, exist_ok=True)
    (home / "vision").mkdir(parents=True, exist_ok=True)
    for key, path in folders.items():
        env[key] = str(path)
    env["PDF_TOOL_LEARNED"] = str(home / "learned" / "learned_corrections.json")
    env["PDF_TOOL_VISION_LOG"] = str(home / "vision" / "vision_log.jsonl")
    env["TEMP"] = str(home / "temp")
    env["TMP"] = str(home / "temp")
    env["PDF_TOOL_OFFLINE"] = "1"
    env["PDF_TOOL_VERSION"] = VERSION
    env["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    env["STREAMLIT_SERVER_HEADLESS"] = "true"
    env["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    exe = find_exe()
    if exe:
        env["PDF_TOOL_TESSERACT"] = exe
    # This Tesseract build treats TESSDATA_PREFIX as the tessdata folder
    # itself. Pointing it at the folder that contains tesseract.exe makes
    # "-l eng" miss eng.traineddata, so the plate re-read returns nothing.
    if exe and "TESSDATA_PREFIX" not in env:
        data = Path(exe).resolve().parent / "tessdata"
        if (data / "eng.traineddata").is_file():
            env["TESSDATA_PREFIX"] = str(data)
    return env


def prepare_server() -> None:
    """Apply the packaged-app folders and the offline guard inside the page process."""
    if os.environ.get("PDF_TOOL_OFFLINE") == "1" or os.environ.get("PDF_TOOL_HOME", "").strip():
        apply_temp_env()
    if os.environ.get("PDF_TOOL_OFFLINE") == "1":
        from engine.offline import install

        install()
