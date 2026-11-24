# -*- coding: utf-8 -*-
"""Build general synthetic corpus PDFs for tables/charts (no sample special-casing)."""
from pathlib import Path

import fitz
from PIL import Image, ImageDraw, ImageFont

out = Path(__file__).resolve().parent / "files"
out.mkdir(parents=True, exist_ok=True)

# 1) Vector/text table with borders
doc = fitz.open()
page = doc.new_page(width=595, height=842)
page.insert_text((72, 50), "1 | Page", fontsize=10)
page.insert_text((72, 80), "Bordered table sample", fontsize=14)
cols = [72, 200, 340, 480]
rows = [110, 140, 190, 240, 290]
headers = ["Name", "Value", "Note"]
data = [["Alpha", "10", "ok"], ["Beta", "20", "ok"], ["Gamma", "30", "gap"]]
page.draw_rect(fitz.Rect(cols[0], rows[0], cols[-1], rows[1]), color=(0.1, 0.2, 0.4), fill=(0.1, 0.2, 0.4))
for x in cols:
    page.draw_line(fitz.Point(x, rows[0]), fitz.Point(x, rows[-1]), color=(0, 0, 0), width=0.7)
for y in rows:
    page.draw_line(fitz.Point(cols[0], y), fitz.Point(cols[-1], y), color=(0, 0, 0), width=0.7)
for i, h in enumerate(headers):
    page.insert_text((cols[i] + 6, rows[0] + 18), h, fontsize=10, color=(1, 1, 1))
for r, row in enumerate(data):
    for c, cell in enumerate(row):
        page.insert_text((cols[c] + 6, rows[r + 1] + 18), cell, fontsize=10)
doc.save(out / "table_bordered.pdf")
doc.close()

# 2) Borderless aligned table
doc = fitz.open()
page = doc.new_page(width=595, height=842)
page.insert_text((72, 50), "2 | Page", fontsize=10)
page.insert_text((72, 80), "Borderless table sample", fontsize=14)
xs = [72, 220, 370]
ys = [120, 150, 180, 210]
cells = [
    ["ColA", "ColB", "ColC"],
    ["one", "two", "three"],
    ["four", "five", "six"],
    ["seven", "eight", "nine"],
]
for r, row in enumerate(cells):
    for c, text in enumerate(row):
        page.insert_text((xs[c], ys[r]), text, fontsize=11)
doc.save(out / "table_borderless.pdf")
doc.close()

# 3) Raster table image
img = Image.new("RGB", (900, 360), (255, 255, 255))
draw = ImageDraw.Draw(img)
draw.rectangle([0, 0, 900, 48], fill=(27, 54, 93))
cols_x = [0, 180, 360, 540, 720, 900]
rows_y = [0, 48, 120, 192, 264, 360]
headers = ["A", "B", "C", "D", "E"]
body = [
    ["a1", "b1", "c1", "d1", "e1"],
    ["a2", "b2", "c2", "d2", "e2"],
    ["a3", "b3", "c3", "d3", "e3"],
    ["a4", "b4", "c4", "d4", "e4"],
]
for x in cols_x:
    draw.line([(x, 0), (x, 360)], fill=(80, 80, 80), width=1)
for y in rows_y:
    draw.line([(0, y), (900, y)], fill=(80, 80, 80), width=1)
try:
    font = ImageFont.truetype("arial.ttf", 22)
except Exception:
    font = ImageFont.load_default()
for i, h in enumerate(headers):
    draw.text((cols_x[i] + 20, 12), h, fill=(255, 255, 255), font=font)
for r, row in enumerate(body):
    for c, t in enumerate(row):
        draw.text((cols_x[c] + 20, rows_y[r + 1] + 20), t, fill=(0, 0, 0), font=font)
png = out / "_raster_table.png"
img.save(png)
doc = fitz.open()
page = doc.new_page(width=595, height=842)
page.insert_text((72, 40), "3 | Page", fontsize=10)
page.insert_text((72, 70), "Table 1. Raster table sample", fontsize=12)
page.insert_image(fitz.Rect(72, 100, 523, 340), filename=str(png))
doc.save(out / "table_raster.pdf")
doc.close()

# 4) Chart with labels baked into the picture
chart = Image.new("RGB", (800, 500), (255, 255, 255))
cd = ImageDraw.Draw(chart)
cd.ellipse([200, 80, 600, 400], outline=(30, 90, 160), width=3)
cd.pieslice([200, 80, 600, 400], 0, 120, fill=(30, 90, 160))
cd.pieslice([200, 80, 600, 400], 120, 250, fill=(200, 80, 60))
cd.pieslice([200, 80, 600, 400], 250, 360, fill=(80, 160, 80))
cd.text((80, 40), "90.66%", fill=(0, 0, 0), font=font)
cd.text((620, 200), "99.62%", fill=(0, 0, 0), font=font)
cd.text((350, 420), "0.34", fill=(0, 0, 0), font=font)
chart_path = out / "_chart.png"
chart.save(chart_path)
doc = fitz.open()
page = doc.new_page(width=595, height=842)
page.insert_text((72, 40), "4 | Page", fontsize=10)
page.insert_image(fitz.Rect(72, 80, 523, 400), filename=str(chart_path))
page.insert_image(fitz.Rect(72, 420, 523, 700), filename=str(chart_path))
doc.save(out / "chart_with_labels.pdf")
doc.close()

# 5) Multi-page table
doc = fitz.open()
for n in (1, 2):
    page = doc.new_page(width=595, height=842)
    page.insert_text((72, 40), f"{n} | Page", fontsize=10)
    page.insert_text((72, 70), "Table 2. Multi-page", fontsize=12)
    cols = [72, 200, 340, 480]
    rows = [100, 130, 180, 230, 280]
    page.draw_rect(
        fitz.Rect(cols[0], rows[0], cols[-1], rows[1]), color=(0.1, 0.2, 0.4), fill=(0.1, 0.2, 0.4)
    )
    for x in cols:
        page.draw_line(fitz.Point(x, rows[0]), fitz.Point(x, rows[-1]), color=(0, 0, 0), width=0.7)
    for y in rows:
        page.draw_line(fitz.Point(cols[0], y), fitz.Point(cols[-1], y), color=(0, 0, 0), width=0.7)
    for i, h in enumerate(["H1", "H2", "H3"]):
        page.insert_text((cols[i] + 6, rows[0] + 18), h, fontsize=10, color=(1, 1, 1))
    for r in range(3):
        for c in range(3):
            page.insert_text((cols[c] + 6, rows[r + 1] + 18), f"p{n}r{r}c{c}", fontsize=10)
doc.save(out / "table_multipage.pdf")
doc.close()

# 6) Links
doc = fitz.open()
page = doc.new_page(width=595, height=842)
page2 = doc.new_page(width=595, height=842)
page = doc[0]
page2 = doc[1]
page.insert_text((72, 80), "Visit https://example.com for more.", fontsize=12)
page.insert_text((72, 120), "Email me at test@example.com", fontsize=12)
page.insert_text((72, 200), "Go to page 2", fontsize=12)
page2.insert_text((72, 80), "Target page", fontsize=14)
page.insert_link({"kind": fitz.LINK_URI, "from": fitz.Rect(120, 65, 280, 90), "uri": "https://example.com"})
page.insert_link(
    {"kind": fitz.LINK_URI, "from": fitz.Rect(160, 105, 320, 130), "uri": "mailto:test@example.com"}
)
page.insert_link({"kind": fitz.LINK_GOTO, "from": fitz.Rect(72, 185, 180, 210), "page": 1})
doc.save(out / "links_sample.pdf")
doc.close()

print("wrote corpus samples to", out)
