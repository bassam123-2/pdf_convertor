# -*- coding: utf-8 -*-
from pathlib import Path

import fitz

from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx, read_word_build
from core.page_strategy import classify_document

pdf = Path(
    r"C:\Users\jouan\.cursor\projects\c-Users-jouan-Desktop-PDF-TOOL"
    r"\attachments\da28dd0d-1afe-444e-be3d-e5ea56fa536e\1-3.pdf"
)
doc = fitz.open(pdf)
for i, page in enumerate(doc):
    imgs = page.get_images(full=True) or []
    drawings = len(page.get_drawings() or [])
    text = page.get_text("text") or ""
    words = page.get_text("words") or []
    print(
        f"p{i+1} images={len(imgs)} drawings={drawings} "
        f"text_len={len(text)} words={len(words)}"
    )
plans = classify_document(doc)
for plan in plans:
    print("plan", plan.as_dict())
doc.close()

out = Path(r"C:\Users\jouan\Desktop\PDF_TOOL\tests\corpus\_similarity_tmp\word_1-3.docx")
counts = convert_arabic_pdf_to_docx(pdf, out, options=WordOptions(mode="faithful", fit_a4=False))
print("convert", {k: counts[k] for k in ("pages", "paragraphs", "tables", "digital", "scans") if k in counts})
print("plans", counts.get("decision_log"))
print("notes", (counts.get("retries") or [])[:12])
print("wrote", out, out.stat().st_size)
