# -*- coding: utf-8 -*-
"""A local timing sample, then an estimate for the PDF that is open."""

from __future__ import annotations

import json
import time
from pathlib import Path

from engine.paths import benchmark_file, using_app_data


def load() -> dict:
    path = benchmark_file()
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def record_run(pages: int, seconds: float) -> None:
    """Remember how long a finished conversion actually took. Failures are ignored."""
    if not using_app_data() or pages <= 0 or seconds <= 0:
        return
    data = load()
    data["seconds_per_page"] = round(float(seconds) / int(pages), 3)
    data["pages"] = int(pages)
    data["seconds"] = round(float(seconds), 3)
    _write(data)


def calibrate_if_needed() -> None:
    """Time one local Tesseract pass the first time the app data folder is used."""
    if not using_app_data():
        return
    data = load()
    if data.get("probe_seconds"):
        return
    seconds = _probe()
    if seconds <= 0:
        return
    data = load()
    data["probe_seconds"] = round(seconds, 3)
    _write(data)


def estimate_seconds(page_count: int) -> float | None:
    if page_count <= 0:
        return None
    data = load()
    per_page = data.get("seconds_per_page")
    if isinstance(per_page, (int, float)) and per_page > 0:
        return float(per_page) * int(page_count)
    probe = data.get("probe_seconds")
    if isinstance(probe, (int, float)) and probe > 0:
        # A tiny picture is faster than a full scanned page. The factor is only a first guess.
        return float(probe) * 8.0 * int(page_count)
    return None


def estimate_text(pdf_bytes: bytes) -> str:
    count = _page_count(pdf_bytes)
    if count <= 0:
        return ""
    seconds = estimate_seconds(count)
    if seconds is None:
        calibrate_if_needed()
        seconds = estimate_seconds(count)
    if seconds is None:
        return f"الملف فيه {count} صفحة. سيظهر الوقت المتبقي بعد الصفحة الأولى."
    return f"الوقت المتوقع لهذا الملف ({count} صفحة): حوالي {_clock(seconds)}."


def _clock(seconds: float) -> str:
    total = max(1, int(round(seconds)))
    minutes, rest = divmod(total, 60)
    if minutes <= 0:
        return f"{rest} ث"
    if rest < 15:
        return f"{minutes} د"
    return f"{minutes} د و {rest} ث"


def _write(data: dict) -> None:
    path = benchmark_file()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(".tmp")
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(path)
    except OSError:
        return


def _page_count(pdf_bytes: bytes) -> int:
    if not pdf_bytes:
        return 0
    try:
        import fitz
    except ImportError:
        return 0
    try:
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return 0
    try:
        return int(document.page_count)
    finally:
        document.close()


def _probe() -> float:
    from engine.tesseract_bin import ara_tessdata_dir, find_exe

    exe = find_exe()
    data = ara_tessdata_dir()
    if not exe or data is None:
        return 0.0
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return 0.0
    picture = Image.new("RGB", (280, 50), "white")
    ImageDraw.Draw(picture).text((8, 12), "12345", fill="black")
    folder = benchmark_file().parent
    folder.mkdir(parents=True, exist_ok=True)
    image_path = folder / "probe.png"
    try:
        picture.save(image_path)
        started = time.perf_counter()
        subprocess_run(exe, image_path, data)
        return time.perf_counter() - started
    except (OSError, TimeoutError):
        return 0.0
    finally:
        picture.close()
        image_path.unlink(missing_ok=True)


def subprocess_run(exe: str, image_path: Path, data: Path) -> None:
    import subprocess

    subprocess.run(
        [exe, str(image_path), "stdout", "-l", "eng", "--tessdata-dir", str(data), "--psm", "7"],
        capture_output=True,
        timeout=60,
        check=False,
    )
