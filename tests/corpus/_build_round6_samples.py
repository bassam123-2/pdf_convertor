# -*- coding: utf-8 -*-
"""Synthetic PDFs for Round 6: fake bold, RTL mixed runs, bilingual bordered table."""
from pathlib import Path

import fitz

out = Path(__file__).resolve().parent / "files"
out.mkdir(parents=True, exist_ok=True)
ARIAL = r"C:\Windows\Fonts\arial.ttf"

# 1) Fake bold: same glyphs drawn twice with a small offset.
doc = fitz.open()
page = doc.new_page(width=400, height=200)
page.insert_text((40, 80), "Boldish Word", fontsize=18, fontname="helv")
page.insert_text((40.4, 80.15), "Boldish Word", fontsize=18, fontname="helv")
page.insert_text((40, 120), "Once Only", fontsize=14, fontname="helv")
doc.save(out / "fake_bold_overprint.pdf")
doc.close()

# 2) RTL line with mixed bold / regular runs inside one visual line + column frame.
# Place phrases adjacent so they share one baseline band (not split by a 2em gap).
doc = fitz.open()
page = doc.new_page(width=420, height=160)
page.draw_rect(fitz.Rect(30, 30, 390, 130), color=(0, 0, 0), width=0.8)
page.insert_font(fontname="ar", fontfile=ARIAL)
# One continuous visual line: regular word then bold overprint of the next word.
# Keep the horizontal gap under 2em so _split_band does not cut the line.
page.insert_text((210, 80), "عادي", fontsize=16, fontname="ar")
page.insert_text((160, 80), "عريض", fontsize=16, fontname="ar")
page.insert_text((160.35, 80.12), "عريض", fontsize=16, fontname="ar")
doc.save(out / "rtl_mixed_runs.pdf")
doc.close()

# 3) Bilingual side-by-side bordered two-column table (EN | AR).
doc = fitz.open()
page = doc.new_page(width=595, height=420)
outer = fitz.Rect(40, 50, 555, 360)
mid_x = 300.0
page.draw_rect(outer, color=(0, 0, 0), width=1.0)
page.draw_line(fitz.Point(mid_x, outer.y0), fitz.Point(mid_x, outer.y1), color=(0, 0, 0), width=1.0)
page.draw_line(fitz.Point(outer.x0, 90), fitz.Point(outer.x1, 90), color=(0, 0, 0), width=0.7)
page.insert_font(fontname="ar", fontfile=ARIAL)
page.insert_text((50, 75), "5. Board role", fontsize=12, fontname="helv")
page.insert_text((420, 75), "5. دور المجلس", fontsize=12, fontname="ar")
en = [
    "The Board shall oversee related-party",
    "transactions and report conflicts of",
    "interest to the shareholders directly.",
]
ar = [
    "يشرف مجلس الإدارة على معاملات",
    "الأطراف ذات العلاقة ويبلغ المساهمين",
    "بأي تعارض مصالح مباشرة.",
]
y = 120
for line in en:
    page.insert_text((50, y), line, fontsize=11, fontname="helv")
    y += 22
y = 120
for line in ar:
    page.insert_text((310, y), line, fontsize=11, fontname="ar")
    y += 22
page.insert_text((310, 190), "الأخلاقية بالأعمال", fontsize=11, fontname="ar")
page.insert_text((310.35, 190.12), "الأخلاقية بالأعمال", fontsize=11, fontname="ar")
doc.save(out / "bilingual_bordered_table.pdf")
doc.close()

print("wrote round-6 corpus PDFs into", out)
