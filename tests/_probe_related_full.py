# -*- coding: utf-8 -*-
from pathlib import Path

import fitz

from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx
from core.page_strategy import classify_document
from core.visual_check import render_docx

pdf = Path(
    r"C:\Users\jouan\.cursor\projects\c-Users-jouan-Desktop-PDF-TOOL"
    r"\attachments\da28dd0d-1afe-444e-be3d-e5ea56fa536e"
    r"\2025-08-Related-Party-Transactions-AR_EN_1.pdf"
)
if not pdf.is_file():
    pdf = Path(r"C:\Users\jouan\Downloads\2025-08-Related-Party-Transactions-AR_EN.pdf")
print("pdf", pdf, pdf.is_file())
doc = fitz.open(pdf)
print("pages", doc.page_count)
for i in range(min(4, doc.page_count)):
    page = doc[i]
    text = page.get_text() or ""
    words = page.get_text("words") or []
    imgs = page.get_images() or []
    drawings = len(page.get_drawings() or [])
    print(
        f"p{i+1} chars={len(text)} words={len(words)} "
        f"images={len(imgs)} drawings={drawings}"
    )
for plan in classify_document(doc):
    print(f"plan {plan.number} {plan.strategy} chars={plan.text_chars} :: {plan.reason[:90]}")
doc.close()

out_dir = Path(r"C:\Users\jouan\Desktop\PDF_TOOL\tests\corpus\_similarity_tmp\related_full")
out_dir.mkdir(parents=True, exist_ok=True)
docx = out_dir / "related.docx"
counts = convert_arabic_pdf_to_docx(pdf, docx, options=WordOptions(mode="faithful", fit_a4=False))
print("convert pages", counts.get("pages"), "tables", counts.get("tables"), "scans", counts.get("scans"), "digital", counts.get("digital"))
for note in (counts.get("retries") or [])[:20]:
    print(" note", note)
print("docx", docx, docx.stat().st_size)

printed = out_dir / "printed.pdf"
print("renderer", render_docx(docx, printed))
printed_doc = fitz.open(printed)
src = fitz.open(pdf)
# Render first 3 pages side by side sources
for i in range(min(3, src.page_count, printed_doc.page_count)):
    src[i].get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False).save(str(out_dir / f"pdf-{i+1}.png"))
    printed_doc[i].get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False).save(str(out_dir / f"word-{i+1}.png"))
    print("wrote pair", i + 1)
src.close()
printed_doc.close()
