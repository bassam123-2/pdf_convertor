# -*- coding: utf-8 -*-
"""
Draw the app icon: packaging/icon.png (for the window) and packaging/icon.ico
(for the program file and the installer).

    python packaging/make_icon.py

A white page with a folded corner on a deep navy-to-teal square, with an
arrow turning it into a document. The colours are the interface's own.
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

HERE = Path(__file__).resolve().parent
NAVY = (20, 32, 51)
TEAL = (15, 110, 107)
GOLD = (214, 168, 76)


def _gradient(size: int) -> Image.Image:
    square = Image.new("RGB", (size, size))
    pixels = square.load()
    for y in range(size):
        for x in range(size):
            t = (x + y) / (2 * (size - 1))
            pixels[x, y] = tuple(int(NAVY[i] + (TEAL[i] - NAVY[i]) * t) for i in range(3))
    return square


def draw(size: int = 1024) -> Image.Image:
    s = size
    icon = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, s - 1, s - 1), radius=int(s * 0.22), fill=255)
    icon.paste(_gradient(s), (0, 0), mask)

    # Soft shadow under the page.
    shadow = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    left, top, right, bottom = int(s * 0.27), int(s * 0.17), int(s * 0.73), int(s * 0.83)
    fold = int(s * 0.14)
    page = [(left, top), (right - fold, top), (right, top + fold), (right, bottom), (left, bottom)]
    ImageDraw.Draw(shadow).polygon([(x + s * 0.015, y + s * 0.025) for x, y in page], fill=(0, 0, 0, 90))
    icon = Image.alpha_composite(icon, shadow.filter(ImageFilter.GaussianBlur(s * 0.02)))

    pen = ImageDraw.Draw(icon)
    pen.polygon(page, fill=(255, 255, 255, 255))
    pen.polygon([(right - fold, top), (right, top + fold), (right - fold, top + fold)], fill=(214, 224, 230, 255))

    # Text lines on the page, right-aligned as Arabic is.
    line = int(s * 0.035)
    for k, width in enumerate((0.30, 0.24, 0.30, 0.18)):
        y = int(s * (0.40 + k * 0.085))
        x1 = right - int(s * 0.07)
        pen.rounded_rectangle((x1 - int(s * width), y, x1, y + line), radius=line // 2, fill=(*TEAL, 255) if k == 0 else (150, 166, 180, 255))

    # A gold arrow: PDF in, Word out.
    cx, cy, r = int(s * 0.36), int(s * 0.72), int(s * 0.115)
    pen.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(*GOLD, 255))
    w = int(s * 0.03)
    pen.line((cx - r * 0.45, cy, cx + r * 0.40, cy), fill=(255, 255, 255, 255), width=w)
    pen.polygon(
        [(cx + r * 0.55, cy), (cx + r * 0.10, cy - r * 0.42), (cx + r * 0.10, cy + r * 0.42)],
        fill=(255, 255, 255, 255),
    )
    return icon


def main() -> None:
    big = draw(1024)
    big.resize((256, 256), Image.LANCZOS).save(HERE / "icon.png")
    big.save(HERE / "icon.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


if __name__ == "__main__":
    main()
