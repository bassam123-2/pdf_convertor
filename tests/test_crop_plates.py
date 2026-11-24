"""The red box is on the same printed row as the plate named on the card."""

import io
import re
import subprocess
import tempfile
from pathlib import Path

import fitz
import numpy as np
import pytest
from PIL import Image

SAMPLE = Path(r"C:\Users\jouan\Downloads\كشف سيارات الافضل 2026.pdf")
PAGES = (2, 4, 5, 6, 8, 12)


def _digits(text: str) -> str:
    return re.sub(r"\D", "", text or "")


def _accepts(plate: str, readings: list[str]) -> bool:
    """The reading is this plate. One doubled stroke may be dropped; another plate is not."""
    for item in readings:
        if item == plate:
            return True
        if len(item) != len(plate) + 1:
            continue
        for index in range(1, len(item)):
            if item[index] == item[index - 1] and item[:index] + item[index + 1 :] == plate:
                return True
    return False


def _plate_chip(png: bytes) -> bytes:
    """The plate digits on the red-box row, tightened onto the ink."""
    image = Image.open(io.BytesIO(png)).convert("RGB")
    array = np.asarray(image)
    red = (array[:, :, 0] > 150) & (array[:, :, 1] < 90)
    rows = np.where(red.any(axis=1))[0]
    if len(rows) < 4:
        return b""
    top = max(0, int(rows[0]) - 8)
    bottom = min(image.height, int(rows[-1]) + 9)
    chip = image.crop((int(image.width * 0.70), top, image.width - 4, bottom)).convert("L")
    gray = np.asarray(chip)
    ys, xs = np.where(gray < 140)
    if len(xs) < 8:
        return b""
    tight = chip.crop(
        (
            max(0, int(xs.min()) - 8),
            max(0, int(ys.min()) - 8),
            min(chip.width, int(xs.max()) + 9),
            min(chip.height, int(ys.max()) + 9),
        )
    )
    tight = tight.resize((tight.width * 3, tight.height * 3), Image.Resampling.LANCZOS)
    buffer = io.BytesIO()
    tight.save(buffer, format="PNG")
    return buffer.getvalue()


def _readings(png: bytes, exe: str, recognize, plate: str) -> list[str]:
    chip = _plate_chip(png)
    if not chip:
        return []
    found: list[str] = []
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as handle:
        handle.write(chip)
        path = handle.name
    try:
        result = subprocess.run(
            [exe, path, "stdout", "--psm", "7", "-l", "eng", "-c", "tessedit_char_whitelist=0123456789"],
            capture_output=True,
            text=True,
            check=False,
        )
    finally:
        Path(path).unlink(missing_ok=True)
    tess = _digits(result.stdout)
    if tess:
        found.append(tess)
    if plate in found:
        return found
    recognized = recognize(chip, "en-US")
    if recognized is not None:
        windows = _digits(" ".join(line.text for line in recognized.lines))
        if windows:
            found.append(windows)
    return found


def test_card_box_reads_the_named_plate():
    if not SAMPLE.is_file():
        pytest.skip("sample PDF is not on this machine")
    from core.excel_converter import (
        ExtractedTable,
        _arabic_column_boxes,
        _page_anchor_words,
        _plate_token,
        _recognize_image,
        _store_visual_baselines,
        _tesseract_exe,
        _visual_plate_ys,
    )
    from page_images import PageBank, card_png

    exe = _tesseract_exe()
    if not exe:
        pytest.skip("Tesseract is not installed")

    pdf = SAMPLE.read_bytes()
    bank = PageBank(pdf)
    bank._render_pdf()
    bank.ready = True
    document = fitz.open(SAMPLE)
    mismatches: list[str] = []
    try:
        for page_number in PAGES:
            page = document.load_page(page_number - 1)
            anchors = _page_anchor_words(page)
            full: dict[str, float] = {}
            for word in anchors:
                token = _plate_token([word.text])
                if token and token not in full:
                    full[token] = float(word.center_y)
            table = ExtractedTable(["رقم اللوحة"], [[token] for token in full])
            table.baseline_at = {f"{page_number}|{token}": y for token, y in full.items()}
            if page_number == 5 and "11411" not in full:
                table.baseline_at["5|11411"] = sorted(full.values())[len(full) // 2]
            _visual, offset = _visual_plate_ys(page, anchors)
            _store_visual_baselines(table, page_number, page, anchors)
            if page_number == 2 and "11261" in full:
                assert table.baseline_at["2|11261"] < full["11261"] - 4
            if page_number == 5 and "5|11411" in table.baseline_at:
                planted = sorted(full.values())[len(full) // 2]
                assert abs(table.baseline_at["5|11411"] - (planted + offset)) < 0.2
                assert table.baseline_at["5|11411"] > 180
            boxes = _arabic_column_boxes(page, anchors)
            spans = {key: value for key, value in boxes.items() if value}
            for plate, y_value in full.items():
                stored = float(table.baseline_at[f"{page_number}|{plate}"])
                png = card_png(bank, page_number, stored, "color", spans, table.baseline_at)
                readings = _readings(png, exe, _recognize_image, plate) if png else []
                if not _accepts(plate, readings):
                    mismatches.append(f"page {page_number} plate {plate} read {readings}")
    finally:
        document.close()
    assert not mismatches, "\n".join(mismatches[:40])
