# -*- coding: utf-8 -*-
from pathlib import Path

from core.visual_check import render_docx
import fitz

docx = Path(r"C:\Users\jouan\Desktop\PDF_TOOL\tests\corpus\_similarity_tmp\word_1-3.docx")
out = Path(r"C:\Users\jouan\Desktop\PDF_TOOL\tests\corpus\_similarity_tmp\word_1-3_pages")
out.mkdir(parents=True, exist_ok=True)
printed = out / "printed.pdf"
used = render_docx(docx, printed)
print("renderer", used)
doc = fitz.open(printed)
for i in range(doc.page_count):
    pix = doc[i].get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
    pix.save(str(out / f"word-page-{i + 1}.png"))
    print("wrote", out / f"word-page-{i + 1}.png")
doc.close()
