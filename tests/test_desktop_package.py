# -*- coding: utf-8 -*-
"""Desktop host checks. These do not read the sample register."""

from __future__ import annotations

import os
import socket
from pathlib import Path

import pytest

from engine.checkpoint import clear, load_page, save_page
from engine.paths import cache_dir, learned_file
from engine.resources import batch_size, worker_count
from engine.tesseract_bin import ara_tessdata_dir
from engine.version import VERSION


def test_version_is_set():
    assert VERSION == "1.0.0"


def test_workers_follow_cpu_and_mode(monkeypatch):
    monkeypatch.setattr(os, "cpu_count", lambda: 8)
    monkeypatch.delenv("PDF_TOOL_MODE", raising=False)
    assert worker_count(20) == 6
    assert batch_size() == 3
    monkeypatch.setenv("PDF_TOOL_MODE", "fast")
    assert worker_count(20) == 4
    assert batch_size() == 2
    monkeypatch.setattr(os, "cpu_count", lambda: 2)
    monkeypatch.delenv("PDF_TOOL_MODE", raising=False)
    assert worker_count(20) == 1


def test_app_data_is_not_the_install_folder(tmp_path, monkeypatch):
    home = tmp_path / "PDFTool"
    monkeypatch.setenv("PDF_TOOL_HOME", str(home))
    monkeypatch.delenv("PDF_TOOL_LEARNED", raising=False)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "Local"))
    assert cache_dir() == home / "cache"
    assert "Program Files" not in str(learned_file())
    assert home in learned_file().parents


def test_checkpoint_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setenv("PDF_TOOL_HOME", str(tmp_path / "PDFTool"))
    save_page("stamp", 2, {"tables": [{"columns": ["a"]}], "error": ""})
    assert load_page("stamp", 2)["tables"][0]["columns"] == ["a"]
    assert load_page("stamp", 3) is None
    clear("stamp")
    assert load_page("stamp", 2) is None


def test_second_read_still_uses_the_core_tessdata_folder():
    from core.excel_converter import _tessdata_dir

    folder = _tessdata_dir()
    assert folder.parent.name == "core"


def test_arabic_model_is_in_the_repo():
    folder = ara_tessdata_dir()
    assert folder is not None
    assert (folder / "ara.traineddata").is_file()


def test_offline_guard_blocks_remote_and_allows_loopback():
    from engine.offline import install, uninstall

    install()
    try:
        with pytest.raises(OSError, match="offline"):
            socket.create_connection(("1.1.1.1", 80), timeout=2)
        probe = socket.socket()
        try:
            probe.connect(("127.0.0.1", 9))
        except OSError as exc:
            assert "offline" not in str(exc).lower()
        finally:
            probe.close()
    finally:
        uninstall()


def test_cloud_read_does_nothing_offline(monkeypatch):
    monkeypatch.setenv("PDF_TOOL_OFFLINE", "1")
    from core.excel_converter import vision_read_crop

    assert vision_read_crop(b"\x89PNG", "make", "secret") == ""


def test_tessdata_prefix_is_the_data_folder(tmp_path, monkeypatch):
    exe_dir = tmp_path / "tesseract"
    data = exe_dir / "tessdata"
    data.mkdir(parents=True)
    (data / "eng.traineddata").write_bytes(b"eng")
    exe = exe_dir / "tesseract.exe"
    exe.write_bytes(b"")
    monkeypatch.setattr("engine.host.find_exe", lambda: str(exe))
    from engine.host import server_environment

    env = server_environment(
        {
            "LOCALAPPDATA": str(tmp_path / "Local"),
            "PDF_TOOL_HOME": str(tmp_path / "PDFTool"),
        }
    )
    assert Path(env["TESSDATA_PREFIX"]) == data


def test_installer_script_keeps_the_same_version():
    text = Path("packaging/PDFTool.iss").read_text(encoding="utf-8")
    assert f'#define AppVersion "{VERSION}"' in text
    assert "محوّل PDF" in text
    assert "{localappdata}" in text
