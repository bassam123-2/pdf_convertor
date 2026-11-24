# -*- coding: utf-8 -*-
from pathlib import Path

import fitz

pdf = Path(
    r"C:\Users\jouan\.cursor\projects\c-Users-jouan-Desktop-PDF-TOOL"
    r"\attachments\da28dd0d-1afe-444e-be3d-e5ea56fa536e\1-3.pdf"
)
doc = fitz.open(pdf)
print("pages", doc.page_count, "size", doc[0].rect)
out = Path(r"C:\Users\jouan\Desktop\PDF_TOOL\tests\corpus\_similarity_tmp\src_1-3")
out.mkdir(parents=True, exist_ok=True)
for i in range(doc.page_count):
    page = doc[i]
    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
    pix.save(str(out / f"pdf-page-{i + 1}.png"))
    text = page.get_text() or ""
    print(f"page {i + 1}: {page.rect.width:.1f}x{page.rect.height:.1f} chars={len(text)}")
    try:
        tabs = list(page.find_tables().tables or [])
        print("  tables", len(tabs), [(t.row_count, t.col_count) for t in tabs[:3]])
    except Exception as exc:
        print("  tables err", exc)
doc.close()
print("wrote", out)
