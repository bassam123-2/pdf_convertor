"""Image-only registers drawn from sample_expected.csv, with exact ground truth.

The page is skewed, noised, marked with a pink highlighter and a stamp, then
stored without a text layer so the reader has to look at the picture.
"""

from __future__ import annotations

import csv
from pathlib import Path

import fitz
import numpy as np
from PIL import Image, ImageDraw, ImageFont

GOLDEN = Path(__file__).parent / "golden" / "sample_expected.csv"
FONT = Path(r"C:\Windows\Fonts\tahoma.ttf")

MAKES = {
    "تويوتا": "Toyota",
    "شيفروليه": "Chevrolet",
    "إيسوزو": "Isuzu",
    "رينو": "Renault",
    "لكزس": "Lexus",
    "جي إم سي": "GMC",
    "جي ام سي": "GMC",
    "فورد": "Ford",
    "نيسان": "Nissan",
    "هوندا": "Honda",
    "هيونداي": "Hyundai",
}
COLORS = {
    "أبيض": "White",
    "أسود": "Black",
    "رمادي": "Gray",
    "فضي": "Silver",
    "أحمر": "Red",
    "أخضر": "Green",
    "صدفي": "Pearl",
    "كحلي": "Navy",
    "بني": "Brown",
    "رمادي غامق": "Dark gray",
}
PURPOSES = {
    "باص خاص": "Private bus",
    "خصوصي": "Private",
    "نقل خاص": "Private transport",
}

ARABIC_HEADERS = ["غرض الترخيص", "رقم اللوحة", "رقم إداري", "الصنع", "السنة", "اللون", "انتهاء التأمين"]
ENGLISH_HEADERS = ["Licence purpose", "Plate number", "Admin number", "Make", "Year", "Color", "Insurance expiry"]
FIELDS = ["purpose", "plate", "admin", "make", "year", "color", "insurance_expiry"]


def choose_rows(limit: int = 28) -> list[dict]:
    """Enough repeated colours for a comparison, taken from the expected file."""
    with GOLDEN.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    chosen: list[dict] = []
    seen: set[str] = set()

    def take(row: dict) -> None:
        plate = str(row.get("plate") or "")
        if plate and plate not in seen and len(chosen) < limit:
            seen.add(plate)
            chosen.append(dict(row))

    for row in rows:
        if str(row.get("color") or "") == "أبيض":
            take(row)
        if sum(1 for item in chosen if item.get("color") == "أبيض") >= 12:
            break
    for row in rows:
        if str(row.get("color") or "") == "صدفي":
            take(row)
        if sum(1 for item in chosen if item.get("color") == "صدفي") >= 10:
            break
    for row in rows:
        if str(row.get("color") or "") not in {"أبيض", "صدفي"}:
            take(row)
        if len(chosen) >= limit:
            break
    return chosen


def _translate(row: dict, english: bool) -> list[str]:
    purpose = str(row.get("purpose") or "")
    make = str(row.get("make") or "")
    color = str(row.get("color") or "")
    if english:
        purpose = PURPOSES.get(purpose, purpose)
        make = MAKES.get(make, make)
        color = COLORS.get(color, color)
    return [
        purpose,
        str(row.get("plate") or ""),
        str(row.get("admin") or ""),
        make,
        str(row.get("year") or ""),
        color,
        str(row.get("insurance_expiry") or ""),
    ]


def _escape(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


# Purpose, plate, admin, make, year, colour, date. Wide word columns stay apart.
_WIDTHS = (130, 78, 72, 110, 64, 90, 120)
_ROW_HEIGHT = 26
_ORIGIN = (28, 28)


def _cell_rect(column: int, row: int) -> fitz.Rect:
    x = _ORIGIN[0] + sum(_WIDTHS[:column])
    y = _ORIGIN[1] + row * _ROW_HEIGHT
    return fitz.Rect(x + 2, y + 2, x + _WIDTHS[column] - 2, y + _ROW_HEIGHT - 2)


def _draw_grid(page: fitz.Page, headers: list[str], body: list[list[str]], english: bool) -> tuple[fitz.Rect, fitz.Rect]:
    """One box per cell. Returns the colour cell that is highlighted, and the stamp box."""
    css = "@font-face {font-family: tahoma; src: url(tahoma.ttf);} div {font-family: tahoma; font-size: 13px;}"
    archive = fitz.Archive(str(FONT.parent))
    direction = "ltr" if english else "rtl"
    for row_index, row in enumerate([headers] + body):
        for column, cell in enumerate(row):
            box = _cell_rect(column, row_index)
            page.draw_rect(
                fitz.Rect(box.x0 - 2, box.y0 - 2, box.x1 + 2, box.y1 + 2),
                color=(0.2, 0.2, 0.2),
                width=0.7,
            )
            html = f'<div dir="{direction}">{_escape(cell)}</div>'
            page.insert_htmlbox(box, html, css=css, archive=archive)
    footer_y = _ORIGIN[1] + (len(body) + 2) * _ROW_HEIGHT
    page.insert_text((28, footer_y), "2026-04-06", fontname="helv", fontsize=14)
    highlight = _cell_rect(5, 3)
    stamp = fitz.Rect(page.rect.width - 130, footer_y - 10, page.rect.width - 20, footer_y + 90)
    return highlight, stamp


def _paint(pixmap: fitz.Pixmap, highlight: fitz.Rect, stamp: fitz.Rect, scale: float) -> Image.Image:
    """Skew, noise, a pink highlighter on one colour, and a stamp in the margin."""
    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    draw = ImageDraw.Draw(image, "RGBA")
    pink = (
        int(highlight.x0 * scale),
        int(highlight.y0 * scale),
        int(highlight.x1 * scale),
        int(highlight.y1 * scale),
    )
    draw.rectangle(pink, fill=(255, 105, 180, 110))
    stamp_box = (
        int(stamp.x0 * scale),
        int(stamp.y0 * scale),
        int(stamp.x1 * scale),
        int(stamp.y1 * scale),
    )
    draw.ellipse(stamp_box, outline=(180, 0, 0, 255), width=5)
    try:
        font = ImageFont.truetype(str(FONT), 28)
    except OSError:
        font = ImageFont.load_default()
    draw.text((stamp_box[0] + 18, stamp_box[1] + 28), "STAMP", fill=(180, 0, 0, 255), font=font)
    image = image.rotate(1.4, expand=True, fillcolor=(255, 255, 255), resample=Image.Resampling.BICUBIC)
    array = np.asarray(image).astype("int16")
    generator = np.random.default_rng(7)
    noise = generator.integers(-14, 15, size=array.shape, dtype="int16")
    mask = generator.random(array.shape[:2]) < 0.3
    array[mask] = np.clip(array[mask] + noise[mask], 0, 255)
    return Image.fromarray(array.astype("uint8"), "RGB")


def write_synthetic(language: str, pdf_path: Path, truth_path: Path, limit: int = 18) -> list[dict]:
    """Write one image PDF and the exact rows that were drawn."""
    english = language == "english"
    rows = choose_rows(limit)
    headers = ENGLISH_HEADERS if english else ARABIC_HEADERS
    body = [_translate(row, english) for row in rows]
    height = _ORIGIN[1] + (len(body) + 5) * _ROW_HEIGHT + 40
    width = _ORIGIN[0] + sum(_WIDTHS) + 40
    document = fitz.open()
    page = document.new_page(width=width, height=height)
    highlight, stamp = _draw_grid(page, headers, body, english)
    scale = 2.0
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    document.close()
    image = _paint(pixmap, highlight, stamp, scale)
    payload = fitz.open()
    out = payload.new_page(width=image.width, height=image.height)
    holder = Path(pdf_path).with_suffix(".png")
    holder.parent.mkdir(parents=True, exist_ok=True)
    image.save(holder, format="PNG")
    out.insert_image(out.rect, filename=str(holder))
    payload.save(pdf_path)
    payload.close()
    holder.unlink(missing_ok=True)
    truth_rows = []
    for row, cells in zip(rows, body):
        truth_rows.append(
            {
                "page": "1",
                "purpose": cells[0],
                "plate": cells[1],
                "admin": cells[2],
                "make": cells[3],
                "year": cells[4],
                "color": cells[5],
                "insurance_expiry": cells[6],
            }
        )
    with truth_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["page", "purpose", "plate", "admin", "make", "year", "color", "insurance_expiry"])
        writer.writeheader()
        writer.writerows(truth_rows)
    return truth_rows
