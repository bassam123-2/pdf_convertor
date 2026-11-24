# -*- coding: utf-8 -*-
"""Quick probe: fake-bold pairs and RTL boxes after Round 6 geometry."""
from collections import Counter, defaultdict
from pathlib import Path

import fitz

from core import faithful
from core.docx_converter import _atoms_on_page

pdf = Path(r"C:\Users\jouan\Downloads\2025-08-Related-Party-Transactions-AR_EN.pdf")
doc = fitz.open(pdf)
page = doc[5]

atoms = _atoms_on_page(page, upright_only=True)
ar = [a for a in atoms if a.x0 > 300]
pairs = 0
for i, a in enumerate(ar):
    ta = a.text.strip()
    if len(ta) < 1:
        continue
    for b in ar[i + 1 : i + 20]:
        tb = b.text.strip()
        if ta != tb:
            continue
        if abs(((a.x0 + a.x1) / 2) - ((b.x0 + b.x1) / 2)) <= 3 and abs(a.origin_y - b.origin_y) <= 2:
            pairs += 1
print("raw identical near pairs", pairs)
cleaned = faithful._dedupe_fake_bold(atoms)
print(
    "after dedupe near pairs",
    sum(
        1
        for i, a in enumerate(cleaned)
        if a.text.strip()
        for b in cleaned[i + 1 : i + 20]
        if a.text.strip() == b.text.strip()
        and abs(((a.x0 + a.x1) / 2) - ((b.x0 + b.x1) / 2)) <= 3
        and abs(a.origin_y - b.origin_y) <= 2
    ),
)

payload = faithful.read_page(page)
over = []
for line in payload.lines:
    if not line.rtl:
        continue
    left, _t, width, _h, flags = faithful._line_box(
        line, 1.0, 0.0, 0.0, payload.width, payload.height, payload.column_rules
    )
    right = left + width
    if right > 561.5 or left < 305.5 or "outside_cell" in flags:
        over.append((round(left, 1), round(right, 1), flags, line.text[:40]))
print("rtl boxes outside column", len(over))
for row in over[:10]:
    print(" ", row)
doc.close()
