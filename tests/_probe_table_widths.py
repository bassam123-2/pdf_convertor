# -*- coding: utf-8 -*-
"""Inspect page-2 table column geometry on the Related Party PDF."""
from pathlib import Path

import fitz

from core import table_rebuild

pdf = Path(
    r"C:\Users\jouan\.cursor\projects\c-Users-jouan-Desktop-PDF-TOOL"
    r"\attachments\da28dd0d-1afe-444e-be3d-e5ea56fa536e"
    r"\2025-08-Related-Party-Transactions-AR_EN_1.pdf"
)
doc = fitz.open(pdf)
page = doc[1]
regions = table_rebuild.classify_regions(page)
print("regions", [(r.kind, tuple(round(v, 1) for v in r.bbox), r.reason[:60]) for r in regions])
tables = table_rebuild.rebuild_tables(page, regions)
for i, table in enumerate(tables):
    print(
        f"table {i} usable={table.usable} cols={len(table.widths)} "
        f"widths={[round(w, 1) for w in table.widths]} "
        f"fallback={len(table.fallback_png)} notes={table.notes}"
    )
    if table.rows:
        print("  header", table.rows[0][:6])
        print("  row1", [c[:40] for c in table.rows[1][:6]] if len(table.rows) > 1 else None)

found = page.find_tables(vertical_strategy="lines", horizontal_strategy="lines")
for ti, table in enumerate(list(getattr(found, "tables", None) or [])[:3]):
    print(f"finder {ti} rows={table.row_count} cols={table.col_count} bbox={tuple(round(v,1) for v in table.bbox)}")
    cells = [c for c in (table.cells or []) if c]
    print(f"  cells={len(cells)}")
    if cells:
        # unique x0 sorted
        xs = sorted({round(c[0], 1) for c in cells})
        print("  x0s", xs)
doc.close()
