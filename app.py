# -*- coding: utf-8 -*-
"""
Arabic PDF workspace: Word and Excel in one local page.

The Word tab calls core.docx_converter.read_word_build, which keeps each page
so one page can be re-read without running the others again. The Excel tab calls
core.excel_converter.read_pdf_workbook, which keeps each row's page number.
The Excel tab can edit one page and re-scan that page, then writes a native
right-to-left sheet with a filled header, borders, and fitted columns.
Files are read and written on this computer.

Install once, from this folder:

    pip install -r requirements.txt

Start the app:

    streamlit run app.py

Streamlit prints a local address, usually http://localhost:8501
"""

from __future__ import annotations

import base64
import getpass
import hashlib
import html
import json
import os
import re
import sys
import threading
import time
from pathlib import Path

from excel_review import (
    audit_record,
    banner_facts,
    blank_text,
    classify_cells,
    clear_crop_library,
    clear_field_note,
    clip_acceptances,
    cluster_decision,
    cluster_review_items,
    crop_library_count,
    export_crop_library,
    load_review_session,
    progress_text,
    queue_status,
    read_print_date,
    remember_crop,
    save_review_session,
    suggestion_kind,
)
from page_images import card_png, cell_ink, cell_png, describe_pages, difficulty_score, ensure_page_bank
from io import BytesIO
from pathlib import Path

try:
    import streamlit as st
except ImportError:
    sys.stderr.write(
        "Install the libraries, then start the app:\n\n"
        "    pip install -r requirements.txt\n"
        "    streamlit run app.py\n"
    )
    raise SystemExit(1)

try:
    import pandas as pd
    from docx import Document

    from core.docx_converter import ConversionError as WordError
    from core.docx_converter import WordOptions
    from core.docx_converter import (
        read_word_build,
        rescan_word_page,
        set_word_page_chrome,
        undo_word_page,
        write_word_build,
    )
    from core.excel_converter import ConversionError as ExcelError
    from core.excel_converter import ExtractedTable
    from core.excel_converter import (
        delete_row_range,
        extract_page_tables,
        extract_row_band,
        fill_attributes_from_above,
        fill_blank_fields_from_above,
        merge_fragment_rows,
        read_pdf_workbook,
        load_cached_workbook,
        save_cached_workbook,
        validate_register,
        page_audit,
        normalize_edited_dates,
        repair_dates_in_range,
        replace_page_rows,
        replace_row_range,
        write_edited_excel,
        _coerce_excel_number,
        _export_column_kind,
        _nondata_kind,
        _protect_excel_date,
        _vehicle_field_at,
        suggest_cell,
        remember_correction,
        forget_correction,
        load_corrections,
        save_corrections,
        neighbourhood_make,
        crop_context_png,
        interpolate_row_y,
        vision_read_crop,
        partition_rows,
    )
    from utils.file_manager import conversion_workspace, read_output
except SystemExit as exc:
    _IMPORT_ERROR: BaseException | None = exc
except Exception as exc:
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None


_WORD_RESULT = "word_result"
_WORD_TOKEN = "word_token"
_EXCEL_RESULT = "excel_result"
_EXCEL_TOKEN = "excel_token"

_WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

_ERROR_RULES = (
    ("password", "هذا الملف يطلب كلمة مرور. أدخل الكلمة الصحيحة ثم أعد التحويل."),
    ("no pages", "ملف PDF لا يحتوي على صفحات."),
    ("not a pdf", "الملف المرفوع ليس ملف PDF."),
    ("is empty", "ملف PDF فارغ."),
    ("no table was found on that page", "لم يُعثر على جدول في هذه الصفحة."),
    ("no table was found", "لم يُعثر على جدول أو فاتورة. يحتاج الملف إلى نص مرتب في صفوف وأعمدة."),
    ("page list has no numbers", "نطاق الصفحات فارغ."),
    ("page list", "تعذّر قراءة نطاق الصفحات. مثال: 1-5 أو 3,4,7."),
    ("is not in this pdf", "رقم الصفحة خارج ملف PDF."),
    ("is not in this document", "هذه الصفحة ليست ضمن الصفحات التي حُوّلت."),
    ("cover picture", "هذه الصفحة محفوظة كغلاف مصور. ألغِ الحفاظ على الصفحة الأولى ثم أعد التحويل إذا أردت تعديل نصها."),
    ("no re-scan to undo", "لا يوجد تعديل على هذه الصفحة للتراجع عنه."),
    ("did not produce any text", "لم ينتج عن هذه الصفحة نص يمكن إصلاحه."),
    ("could not be opened", "تعذّر فتح ملف PDF. قد يكون الملف تالفاً."),
    ("may be damaged", "تعذّر فتح ملف PDF. قد يكون الملف تالفاً."),
    ("could not be read", "تعذّر قراءة الملف."),
    ("close the excel", "أغلق ملف إكسل إن كان مفتوحاً، ثم أعد المحاولة."),
)


_CSS = """
<style>
html, body, .stApp, [data-testid="stAppViewContainer"] {
  direction: rtl;
  color: #142033;
}
.stApp,
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stWidgetLabel"] p,
.stApp button,
.stApp input,
.stApp textarea {
  font-family: "IBM Plex Sans Arabic", "Segoe UI", Tahoma, sans-serif;
}
.stApp [data-testid="stMarkdownContainer"] p,
.stApp [data-testid="stMarkdownContainer"] h1,
.stApp [data-testid="stMarkdownContainer"] h2,
.stApp [data-testid="stMarkdownContainer"] h3,
.stApp [data-testid="stMarkdownContainer"] li,
.stApp [data-testid="stWidgetLabel"] p,
.hero, .hero h1, .hero p, .note, .note li {
  text-align: start !important;
}
.stApp button [data-testid="stMarkdownContainer"] p,
[data-testid="stTab"] [data-testid="stMarkdownContainer"] p {
  text-align: center !important;
}
[data-testid="stIconMaterial"] {
  font-family: "Material Symbols Rounded" !important;
  font-weight: 400 !important;
  font-style: normal !important;
  line-height: 1 !important;
  letter-spacing: normal !important;
  text-transform: none !important;
  white-space: nowrap !important;
  direction: ltr !important;
  width: 1.15rem !important;
  overflow: hidden !important;
}
[data-testid="stHeaderActionElements"] { display: none !important; }
.stApp { color-scheme: light; background: #F3F0E8; }
::selection { background: rgba(15, 110, 107, 0.18); }

[data-testid="stHeader"] {
  background: rgba(243, 240, 232, 0.92);
}
[data-testid="stDecoration"], footer, #MainMenu { display: none; }
[data-testid="stToolbar"] { visibility: hidden; height: 0; }

[data-testid="stMainBlockContainer"], .block-container {
  max-width: 1080px;
  padding-top: 1.5rem;
  padding-bottom: 3.2rem;
}

.hero {
  display: flex;
  align-items: stretch;
  justify-content: space-between;
  gap: 1.5rem;
  position: relative;
  overflow: hidden;
  color: #F7F4EE;
  border-radius: 24px;
  padding: 1.7rem 1.7rem 1.55rem;
  background:
    radial-gradient(420px 220px at 0% 0%, rgba(226, 196, 138, 0.22), transparent 70%),
    linear-gradient(155deg, #10243A 0%, #17324C 58%, #1A4148 100%);
  box-shadow: 0 18px 46px rgba(16, 36, 58, 0.16);
}
.hero-copy { position: relative; z-index: 1; min-width: 0; flex: 1 1 auto; }
.eyebrow, .hero-kicker {
  margin: 0;
  color: #E2C48A;
  font-size: 0.82rem;
  font-weight: 700;
}
.hero h1 {
  margin: 0.35rem 0 0.55rem !important;
  padding: 0 !important;
  font-family: "IBM Plex Sans Arabic", "Segoe UI", Tahoma, sans-serif;
  font-size: 2.15rem !important;
  line-height: 1.35 !important;
  font-weight: 700;
  letter-spacing: 0;
}
.hero .lead {
  margin: 0;
  max-width: 38rem;
  color: rgba(247, 244, 238, 0.82);
  font-size: 1.05rem;
  line-height: 1.8;
}
.pills {
  display: flex;
  flex-wrap: wrap;
  gap: 0.45rem;
  margin: 1.05rem 0 0;
  padding: 0;
  list-style: none;
}
.pills li {
  flex: 0 0 auto !important;
  width: max-content !important;
  max-width: none !important;
  min-width: max-content !important;
  white-space: nowrap !important;
  border: 1px solid rgba(247, 244, 238, 0.2);
  background: rgba(255, 255, 255, 0.05);
  border-radius: 999px;
  padding: 0.32rem 0.75rem;
  font-size: 0.86rem;
}
.hero-side {
  position: relative;
  z-index: 1;
  flex: 0 0 210px;
  display: flex;
  align-items: center;
}
.sheet {
  width: 100%;
  border-radius: 16px;
  padding: 0.95rem 0.95rem 0.8rem;
  background: rgba(255, 255, 255, 0.06);
  border: 1px solid rgba(255, 255, 255, 0.14);
  backdrop-filter: blur(6px);
}
.sheet b {
  display: block;
  margin-bottom: 0.55rem;
  color: #E2C48A;
  font-size: 0.78rem;
  font-weight: 700;
}
.sheet i {
  display: block;
  margin-top: 0.45rem;
  padding-top: 0.45rem;
  border-top: 1px solid rgba(247, 244, 238, 0.12);
  color: #F7F4EE;
  font-style: normal;
  font-size: 0.92rem;
  line-height: 1.55;
}

[data-testid="stTabs"] { margin-top: 1.1rem; }
[data-testid="stTabs"] [role="tablist"] {
  gap: 0.45rem;
  background: transparent;
  border: 0;
}
[data-testid="stTab"] {
  height: auto;
  border-radius: 12px;
  padding: 0.55rem 1rem;
  background: transparent;
  color: #4A5968;
  font-weight: 700;
}
[data-testid="stTab"] p { margin: 0; color: inherit; font-weight: 700; }
[data-testid="stTab"][aria-selected="true"] {
  background: #FFFFFF;
  color: #0F6E6B;
  box-shadow: 0 8px 22px rgba(20, 32, 51, 0.05);
}
[data-testid="stTab"] .react-aria-SelectionIndicator { display: none; }
[data-testid="stTabPanel"] {
  background: #FFFFFF;
  border-radius: 20px;
  padding: 1.2rem 1.2rem 1.05rem;
  box-shadow: 0 14px 36px rgba(20, 32, 51, 0.05);
}

.tool-head h2 {
  margin: 0.15rem 0 0.35rem !important;
  padding: 0 !important;
  font-size: 1.45rem !important;
  line-height: 1.4 !important;
  font-weight: 700;
  color: #142033;
}
.tool-head p { margin: 0; color: #4E5D6C; line-height: 1.75; }
.tool-head { margin-bottom: 0.95rem; }

.note {
  background: #F7F4EE;
  border: 1px solid rgba(20, 32, 51, 0.05);
  border-radius: 16px;
  padding: 0.95rem 0.95rem 0.8rem;
}
.note h3 {
  margin: 0 0 0.75rem !important;
  padding: 0 !important;
  font-size: 1rem !important;
  line-height: 1.5 !important;
  color: #142033;
}
.note ol { margin: 0; padding: 0; list-style: none; display: flex; flex-direction: column; gap: 0.72rem; }
.note li { display: grid; grid-template-columns: auto 1fr; gap: 0.65rem; align-items: start; }
.note .n {
  width: 1.55rem;
  height: 1.55rem;
  border-radius: 999px;
  background: #10243A;
  color: #F7F4EE;
  display: grid;
  place-items: center;
  font-size: 0.75rem;
  font-weight: 700;
}
.note strong { display: block; font-size: 0.94rem; color: #142033; }
.note span { display: block; margin-top: 0.1rem; color: #4E5D6C; font-size: 0.86rem; line-height: 1.6; }

.section-label {
  margin: 0.85rem 0 0.4rem;
  font-size: 0.95rem;
  font-weight: 700;
  color: #142033;
}
.preview {
  background: #FBF9F5;
  border: 1px solid rgba(20, 32, 51, 0.06);
  border-radius: 14px;
  padding: 0.85rem 1rem 0.45rem;
}
.preview h4 { margin: 0.15rem 0 0.35rem; font-size: 1.05rem; color: #142033; }
.preview p { margin: 0 0 0.55rem; color: #243246; line-height: 1.8; }
.file-meta { margin: 0.15rem 0 0.35rem; color: #4E5D6C; font-size: 0.9rem; }
.foot {
  margin-top: 0.9rem;
  color: #6A7784;
  font-size: 0.84rem;
  text-align: center;
}

[data-testid="stFileUploaderDropzone"] {
  background: #FBF9F5;
  border: 1.5px dashed #D5C7AE;
  border-radius: 16px;
  min-height: 5.5rem;
}
[data-testid="stFileUploaderDropzone"] button {
  min-width: 9rem;
  min-height: 2.5rem;
}
[data-testid="stFileUploaderDropzone"] button p {
  font-size: 0 !important;
  line-height: 0 !important;
  color: transparent !important;
}
[data-testid="stFileUploaderDropzone"] button p::after {
  content: "اختيار ملف";
  font-size: 0.95rem !important;
  line-height: 1.4 !important;
  font-weight: 700;
  color: #142033;
}
[data-testid="stFileUploaderDropzoneInstructions"] span {
  font-size: 0 !important;
  color: transparent !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] span::after {
  content: "ملف PDF واحد · حتى 200 م.ب";
  font-size: 0.9rem;
  line-height: 1.6;
  font-weight: 600;
  color: #3E4C5A;
}
[data-testid="stMetric"] {
  background: #F7F4EE;
  border: 1px solid rgba(20, 32, 51, 0.05);
  border-radius: 14px;
  padding: 0.7rem 0.8rem;
}
[data-testid="stAlert"], [data-testid="stExpander"] { border-radius: 14px; }
[data-testid="stTextInput"] input { direction: ltr; text-align: left; }
.stApp button[kind="primary"], [data-testid="stBaseButton-primary"] {
  background: #0F6E6B;
  border: 1px solid #0F6E6B;
  color: #FFFFFF;
  border-radius: 12px;
  font-weight: 700;
  min-height: 2.75rem;
}
.stApp button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover {
  background: #0C5C59;
  border-color: #0C5C59;
  color: #FFFFFF;
}
.stApp button:disabled { opacity: 0.5; }

@media (max-width: 820px) {
  .hero { flex-direction: column; padding: 1.3rem; }
  .hero h1 { font-size: 1.7rem; }
  .hero-side { flex-basis: auto; width: 100%; }
}
</style>
"""


_HERO = """
<section class="hero">
  <div class="hero-copy">
    <p class="hero-kicker">محوّل المستندات</p>
    <h1>PDF إلى وورد وإكسل</h1>
    <p class="lead">استخرج النصوص والفقرات إلى مستند وورد، أو الجداول وفواتير المبيعات إلى ملف إكسل. الكتابة العربية من اليمين إلى اليسار، والملف يبقى على هذا الجهاز.</p>
    <ul class="pills">
      <li>فقرات عربية من اليمين</li>
      <li>جداول بترويسة وحدود</li>
      <li>معالجة على هذا الجهاز</li>
    </ul>
  </div>
  <div class="hero-side">
    <div class="sheet">
      <b>ماذا تستلم</b>
      <i>مستند وورد بفقرات قابلة للتعديل</i>
      <i>ورقة إكسل تُفتح من اليمين</i>
      <i>ترويسة، حدود، وأعمدة مضبوطة</i>
    </div>
  </div>
</section>
"""


def _arabic_error(exc: BaseException) -> str:
    text = str(exc).strip()
    lowered = text.lower()
    for needle, message in _ERROR_RULES:
        if needle in lowered:
            return message
    if text:
        return f"تعذّر إكمال التحويل. {text}"
    return "تعذّر إكمال التحويل."


def _safe_stem(filename: str) -> str:
    stem = Path(filename or "document").stem
    cleaned = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', " ", stem)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned[:80] or "document"


def _size_label(num: int) -> str:
    if num < 1024:
        return f"{num} بايت"
    if num < 1024 * 1024:
        return f"{num / 1024:.0f} ك.ب"
    return f"{num / (1024 * 1024):.1f} م.ب"


def _require_pdf(data: bytes) -> None:
    if not data:
        raise ValueError("The PDF is empty.")
    if b"%PDF" not in data[:1024]:
        raise ValueError("That file is not a PDF.")


def _remember_upload(slot: str, uploaded) -> None:
    token = None
    if uploaded is not None:
        identity = getattr(uploaded, "file_id", None) or uploaded.name
        token = f"{identity}:{uploaded.size}"
    token_key = _WORD_TOKEN if slot == "word" else _EXCEL_TOKEN
    result_key = _WORD_RESULT if slot == "word" else _EXCEL_RESULT
    if st.session_state.get(token_key) != token:
        st.session_state[token_key] = token
        st.session_state.pop(result_key, None)


def _convert_word(data: bytes, filename: str, password: str, options: WordOptions):
    _require_pdf(data)
    stem = _safe_stem(filename)
    with conversion_workspace("pdf-word-") as root:
        source = root / f"{stem}.pdf"
        target = root / f"{stem}.docx"
        source.write_bytes(data)
        build = read_word_build(source, password or "", options)
        counts = write_word_build(build, target)
        return read_output(target), counts, build


def _rebuild_word(result: dict) -> None:
    """Write the .docx again from the pages stored for this conversion."""
    build = result.get("build")
    if build is None:
        return
    with conversion_workspace("pdf-word-edit-") as root:
        target = root / "edited.docx"
        result["counts"] = write_word_build(build, target)
        result["bytes"] = read_output(target)
    result["revision"] = int(result.get("revision") or 0) + 1


def _word_options_guide() -> None:
    """Arabic notes for the choices before conversion and the fixes after it."""
    with st.expander("دليل خيارات وورد", expanded=False):
        st.markdown(
            """
**قبل التحويل**

**اللغة والاتجاه.** «تلقائي / مختلط» يقرر اتجاه كل فقرة وجدول من حروفها، فالعربي من اليمين والإنجليزي من اليسار في الملف نفسه. «عربي» يُلزم الكتابة من اليمين في الفقرات والجداول والهوامش. «إنجليزي» يُلزم الكتابة من اليسار.

**الهيدر والفوتر.** يُقرآن من الصفحات المختارة معاً، لا من الصفحة الأولى وحدها. ما يتكرر هو الرأس والذيل. حدّ الجدول وخط أزرق تحت الشعار لا يُحسبان رأساً. الشريط الملوّن يُقصّ صورةً ويُثبَّت في هامش وورد، من غير خط زائد تحته.

**الحفاظ على الصفحة الأولى / الغلاف.** فعّله عندما تكون الصفحة الأولى غلافاً أو لوحة بصرية. تبقى صورة واحدة بلا هيدر ولا فوتر ولا إعادة سطر. أوقفه عندما تكون الصفحة الأولى نصاً عادياً تريد تحريره.

**تثبيت الهيدر.** يثبّت شريط الرأس بعرض الصفحة، بارتفاع محسوب، ومن غير خط زائد أو فراغ كبير تحته. أوقفه إذا كان المستند بلا رأس متكرر.

**تثبيت الفوتر.** يثبّت مسافة الفوتر عن حافة الصفحة وعن المتن، ويبقي رقم الصفحة والنص الجاري على طرفين متقابلين من غير تداخل. أوقفه إذا لم يكن في الملف فوتر.

**نطاق الصفحات.** اتركه فارغاً لتحويل كل الصفحات. `1-5` صفحة متتالية، و`3,4,7` صفحات متفرقة.

**دمج الفقرات.** يضم الأسطر المتتالية في فقرة عربية متصلة. أوقفه إذا كان كل سطر عنواناً أو صفاً يجب أن يبقى وحده.

**الجداول والرسوم.** «جدول قابل للتعديل» يحوّل الشبكة إلى خلايا وورد. «كتلة بصرية» تُبقي الجدول صورة عندما يكون التخطيط متداخلاً ولا تريد أن يتفكك.

**بعد التحويل**

اختر صفحة واحدة. باقي الصفحات تبقى كما قُرئت. إعادة القراءة والتراجع يبقيان لغة التحويل والهيدر والفوتر نفسيهما.

**إصلاح اتجاه النص وتداخل الكلمات.** يعيد ترتيب الكتل من اليمين، ويزيل الكلمة المنسوخة فوق كلمة أخرى، ويعيد الفقرة المتصلة.

**إعادة استخراج الصور والإنفوجرافيك.** يستبدل الصور والرسوم في الصفحة بدقة أعلى، ويترك النص والجداول.

**إعادة بناء الجداول فقط.** يعيد قراءة جداول هذه الصفحة وحدها.

**إعادة فحص شامل للصفحة.** يعيد بناء الصفحة من الملف بالخيارات نفسها.

**الهيدر والفوتر لهذه الصفحة.** يخفي الرأس أو الذيل في صفحة الفهرس أو اللوحة الكاملة. يعمل عندما كان التثبيت مفعّلاً قبل التحويل.

**تراجع عن تعديل الصفحة.** يعيد هذه الصفحة إلى ما كانت عليه قبل آخر إعادة قراءة أو إخفاء.
            """
        )


_WORD_FIXES = (
    ("rtl", "إصلاح اتجاه النص وتداخل الكلمات"),
    ("images", "إعادة استخراج الصور والإنفوجرافيك"),
    ("tables", "إعادة بناء الجداول فقط"),
    ("full", "إعادة فحص شامل للصفحة"),
)


def _convert_excel(data: bytes, filename: str, password: str, progress, strict: bool = True, direction: str = "auto") -> dict:
    """Read every page once and keep each row tagged with that page number."""
    _require_pdf(data)
    opened = time.perf_counter()
    cached = load_cached_workbook(data, password or "", bool(strict), direction or "auto")
    if cached:
        cached["pdf"] = data
        cached["password"] = password or ""
        cached["direction"] = direction or "auto"
        cached["print_date"] = read_print_date(data, password or "")
        cached["revision"] = 0
        cached["built_revision"] = 0
        cached["dirty"] = False
        cached["review_cursor"] = 0
        cached["review_open"] = False
        cached["timings"] = [{"name": "cache", "seconds": round(time.perf_counter() - opened, 3)}]
        if not cached.get("page_kinds"):
            cached["page_kinds"] = describe_pages(data)
        ensure_page_bank(data)
        return cached
    stem = _safe_stem(filename)
    with conversion_workspace("pdf-excel-") as root:
        source = root / f"{stem}.pdf"
        source.write_bytes(data)
        state = read_pdf_workbook(source, password or "", progress=progress, strict=bool(strict))
        target = root / f"{stem}.xlsx"
        write_started = time.perf_counter()
        if state["rows"] or state["figures"]:
            write_edited_excel(
                state["columns"],
                state["rows"],
                state["page_numbers"],
                target,
                state.get("kind") or "",
                state.get("figures"),
                strict=bool(strict),
                validation=state.get("validation"),
                direction=direction or "auto",
                provenance=state.get("provenance"),
                suggestions=state.get("suggestions"),
                raw_readings=state.get("raw_readings"),
            )
            payload = read_output(target)
        else:
            payload = b""
        write_seconds = round(time.perf_counter() - write_started, 3)
    state["bytes"] = payload
    timings = list(state.get("timings") or [])
    timings.append({"name": "build_excel", "seconds": write_seconds})
    state["timings"] = timings
    save_cached_workbook(data, password or "", bool(strict), direction or "auto", state)
    state["pdf"] = data
    state["password"] = password or ""
    state["direction"] = direction or "auto"
    state["print_date"] = read_print_date(data, password or "")
    state["revision"] = 0
    state["built_revision"] = 0
    state["review_cursor"] = 0
    state["review_open"] = False
    state["page_kinds"] = describe_pages(data)
    ensure_page_bank(data)
    return state


def _clip(text: str, limit: int = 320) -> str:
    compact = " ".join(text.split())
    if len(compact) <= limit:
        return compact
    return compact[: limit - 1].rstrip() + "…"


def _word_preview(payload: bytes, limit: int = 5) -> list[tuple[str, bool]]:
    document = Document(BytesIO(payload))
    lines: list[tuple[str, bool]] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        heading = any(run.bold for run in paragraph.runs)
        lines.append((_clip(text), heading))
        if len(lines) >= limit:
            break
    return lines


def _note(title: str, items: list[tuple[str, str]]) -> str:
    rows = []
    for index, (heading, body) in enumerate(items, start=1):
        rows.append(
            "<li>"
            f'<span class="n">{index}</span>'
            f"<div><strong>{html.escape(heading)}</strong><span>{html.escape(body)}</span></div>"
            "</li>"
        )
    return (
        '<div class="note">'
        f"<h3>{html.escape(title)}</h3>"
        f"<ol>{''.join(rows)}</ol>"
        "</div>"
    )


def _fail(status, exc: BaseException, unexpected: bool = False) -> None:
    status.update(label="تعذّر التحويل", state="error", expanded=True)
    if unexpected:
        st.error("حدث خطأ غير متوقع أثناء التحويل.")
        st.exception(exc)
        return
    st.error(_arabic_error(exc))


def _render_word() -> None:
    form, aside = st.columns([1.55, 0.85], gap="large", vertical_alignment="top")
    with form:
        st.markdown(
            """
            <div class="tool-head">
              <p class="eyebrow">الأداة الأولى</p>
              <h2>نص وفقرات إلى وورد</h2>
              <p>يقرأ العناوين والفقرات والجداول من PDF ويكتبها في مستند وورد. قبل التحويل تُحدَّد الصفحات والهوامش، وبعده يمكن إعادة قراءة صفحة واحدة من غير المساس بباقي المستند.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            "ملف PDF",
            type=["pdf"],
            key="word_pdf",
            help="يُقرأ الملف على هذا الجهاز، ويُحذف الملف المؤقت بعد إنشاء مستند وورد.",
        )
        _remember_upload("word", uploaded)
        if uploaded is not None:
            st.markdown(
                f'<p class="file-meta">{html.escape(uploaded.name)} · {_size_label(uploaded.size)}</p>',
                unsafe_allow_html=True,
            )
        password = st.text_input(
            "كلمة المرور، إن كان الملف محمياً",
            type="password",
            key="word_password",
            help="اترك الحقل فارغاً إذا كان الملف يُفتح من غير كلمة مرور.",
        )
        st.markdown('<div class="section-label">قبل التحويل</div>', unsafe_allow_html=True)
        direction = st.radio(
            "اللغة والاتجاه",
            options=("auto", "rtl", "ltr"),
            format_func=lambda item: {
                "auto": "تلقائي / مختلط",
                "rtl": "عربي - من اليمين لليسار",
                "ltr": "إنجليزي - من اليسار لليمين",
            }[item],
            horizontal=True,
            key="word_direction",
            help="تلقائي يميّز العربي من الإنجليزي داخل الملف. عربي يُلزم اليمين. إنجليزي يُلزم اليسار.",
        )
        preserve = st.checkbox(
            "الحفاظ على الصفحة الأولى / الغلاف",
            value=False,
            key="word_preserve",
            help="الصفحة الأولى تبقى صورة كاملة، بلا هيدر ولا فوتر ولا إعادة سطر.",
        )
        lock_header = st.checkbox(
            "تثبيت الهيدر",
            value=True,
            key="word_lock_header",
            help="شريط الرأس بعرض الصفحة، بارتفاع ثابت، ومن غير خط زائد تحته.",
        )
        lock_footer = st.checkbox(
            "تثبيت الفوتر",
            value=True,
            key="word_lock_footer",
            help="مسافة ثابتة عن حافة الصفحة، ورقم الصفحة مقابل النص الجاري.",
        )
        page_spec = st.text_input(
            "نطاق الصفحات",
            value="",
            key="word_pages",
            placeholder="كل الصفحات، أو 1-5، أو 3,4,7",
            help="اتركه فارغاً لتحويل كل الصفحات.",
        )
        merge = st.checkbox(
            "دمج الأسطر في فقرات متصلة",
            value=True,
            key="word_merge",
            help="الأسطر المتتالية تصبح فقرة واحدة. أوقفه إذا كان كل سطر عنصراً مستقلاً.",
        )
        tables_mode = st.radio(
            "الجداول والرسوم",
            options=("editable", "visual"),
            format_func=lambda item: "جدول قابل للتعديل" if item == "editable" else "كتلة بصرية ثابتة",
            horizontal=True,
            key="word_tables",
        )
        if st.button(
            "تحويل إلى وورد",
            key="word_convert",
            type="primary",
            width="stretch",
            disabled=uploaded is None,
        ):
            _run_word(
                uploaded,
                password,
                WordOptions(
                    preserve_first_page=bool(preserve),
                    lock_header=bool(lock_header),
                    lock_footer=bool(lock_footer),
                    page_spec=page_spec or "",
                    merge_paragraphs=bool(merge),
                    tables_as_images=tables_mode == "visual",
                    fit_a4=True,
                    direction=str(direction or "auto"),
                ),
            )
        _show_word_result()
    with aside:
        st.markdown(
            _note(
                "في ملف وورد",
                [
                    ("اتجاه الفقرة", "الكتابة العربية تبدأ من اليمين وتجري باتجاهها الطبيعي."),
                    ("صفحة واحدة", "إعادة القراءة تستبدل الصفحة المختارة وتترك الباقي."),
                    ("الهيدر والفوتر", "يمكن إخفاؤهما في صفحة الفهرس أو اللوحة الكاملة."),
                    ("التنزيل", "الملف يُبنى من آخر حالة محفوظة للصفحات."),
                ],
            ),
            unsafe_allow_html=True,
        )
        _word_options_guide()


def _run_word(uploaded, password: str, options: WordOptions) -> None:
    with st.status("جارٍ استخراج النص والفقرات...", expanded=False) as status:
        try:
            payload, counts, build = _convert_word(uploaded.getvalue(), uploaded.name, password, options)
        except (WordError, ValueError) as exc:
            _fail(status, exc)
            return
        except Exception as exc:
            _fail(status, exc, unexpected=True)
            return
        status.update(label="اكتمل استخراج النص", state="complete")
    st.session_state[_WORD_RESULT] = {
        "bytes": payload,
        "name": f"{_safe_stem(uploaded.name)}.docx",
        "counts": counts,
        "build": build,
        "pdf": uploaded.getvalue(),
        "password": password or "",
        "revision": 0,
    }
    st.toast("تم إنشاء ملف وورد")


def _word_page_tools(result: dict) -> None:
    """Re-read one stored page, or hide its header and footer, then rebuild the file."""
    build = result.get("build")
    pages = list(getattr(build, "pages", []) or [])
    if not pages or not result.get("pdf"):
        return
    st.markdown('<div class="section-label">بعد التحويل — صفحة واحدة</div>', unsafe_allow_html=True)
    st.caption("العمل يجري على الصفحة المختارة. الصفحات الأخرى تبقى كما هي في الملف.")
    labels = {page.number: f"صفحة {page.number}" for page in pages}
    numbers = [page.number for page in pages]
    token = st.session_state.get(_WORD_TOKEN) or "word"
    revision = int(result.get("revision") or 0)
    selected = st.selectbox(
        "الصفحة",
        options=numbers,
        format_func=lambda number: labels[number],
        key=f"word_page_{token}",
    )
    record = next(page for page in pages if page.number == int(selected))
    mode = st.selectbox(
        "نوع المعالجة",
        options=[item[0] for item in _WORD_FIXES],
        format_func=lambda item: dict(_WORD_FIXES)[item],
        key=f"word_fix_{token}",
    )
    lock_header = bool(build.options.lock_header)
    lock_footer = bool(build.options.lock_footer)
    header_col, footer_col = st.columns(2)
    with header_col:
        show_header = st.checkbox(
            "إظهار الهيدر في هذه الصفحة",
            value=bool(record.show_header and lock_header and not record.cover),
            key=f"word_page_header_{token}_{record.number}_{revision}",
            disabled=not lock_header or record.cover,
        )
    with footer_col:
        show_footer = st.checkbox(
            "إظهار الفوتر في هذه الصفحة",
            value=bool(record.show_footer and lock_footer and not record.cover),
            key=f"word_page_footer_{token}_{record.number}_{revision}",
            disabled=not lock_footer or record.cover,
        )
    rescan_col, chrome_col, undo_col = st.columns(3)
    with rescan_col:
        rescan = st.button("إعادة معالجة الصفحة", key="word_rescan", width="stretch")
    with chrome_col:
        apply_chrome = st.button(
            "تطبيق الهيدر والفوتر",
            key="word_chrome",
            width="stretch",
            disabled=record.cover or (not lock_header and not lock_footer),
        )
    with undo_col:
        undo = st.button(
            "تراجع عن تعديل الصفحة",
            key="word_undo",
            width="stretch",
            disabled=not record.undo,
        )
    if not (rescan or apply_chrome or undo):
        return
    try:
        if undo:
            undo_word_page(build, int(selected))
            note = "تم التراجع عن تعديل الصفحة"
        elif apply_chrome:
            before = (record.show_header, record.show_footer)
            set_word_page_chrome(build, int(selected), bool(show_header), bool(show_footer))
            if (record.show_header, record.show_footer) == before:
                st.toast("الهيدر والفوتر على هذه الصفحة كما هما")
                return
            note = "تم تحديث الهيدر والفوتر لهذه الصفحة"
        else:
            with conversion_workspace("pdf-word-page-") as root:
                source = root / "page.pdf"
                source.write_bytes(result["pdf"])
                rescan_word_page(source, build, int(selected), str(mode), result.get("password") or "")
            note = "أُعيدت معالجة الصفحة"
        _rebuild_word(result)
    except (WordError, ValueError) as exc:
        st.error(_arabic_error(exc))
        return
    except Exception as exc:
        st.error("حدث خطأ غير متوقع أثناء تعديل الصفحة.")
        st.exception(exc)
        return
    st.session_state[_WORD_RESULT] = result
    st.toast(note)
    st.rerun()


def _score_color(score: float | None, threshold: float = 97.0) -> str:
    if score is None:
        return "#9aa4b2"
    if score >= threshold:
        return "#1f8a4c"
    if score >= threshold - 10:
        return "#c48a1a"
    return "#c0392b"


def _sim_persist_dir(token: str) -> Path:
    import re
    import tempfile

    # Upload tokens look like "{uuid}:{size}"; ":" is illegal in Windows paths.
    safe = re.sub(r'[<>:"/\\|?*]+', "_", str(token or "word")).strip(" ._") or "word"
    root = Path(tempfile.gettempdir()) / "pdf_tool_sim_ui" / safe
    root.mkdir(parents=True, exist_ok=True)
    return root


def _word_similarity_preview(result: dict) -> None:
    """
    Minimal Round 7 preview: score thumbnails, side-by-side + diff, reconvert, undo.
    Uses core.similarity only; conversion stays faithful by default.
    """
    pdf_bytes = result.get("pdf")
    docx_bytes = result.get("bytes")
    if not pdf_bytes or not docx_bytes:
        return
    st.markdown('<div class="section-label">معاينة التشابه قبل التنزيل</div>', unsafe_allow_html=True)
    st.caption(
        "الصور منخفضة الدقة افتراضياً. الخطوط الناقصة على الجهاز تؤثر على الدرجة. "
        "الوضع الافتراضي يبقى faithful دون ملاءمة A4."
    )
    token = st.session_state.get(_WORD_TOKEN) or "word"
    persist = _sim_persist_dir(token)
    if st.button("حساب درجات الصفحات", key=f"sim_score_{token}"):
        try:
            from core import similarity
            from core.visual_check import RenderError

            pdf_path = persist / "source.pdf"
            docx_path = persist / "result.docx"
            project_path = persist / "job.project.json"
            out = persist / "sim"
            pdf_path.write_bytes(pdf_bytes)
            docx_path.write_bytes(docx_bytes)
            report = similarity.run_report(
                pdf_path,
                docx_path,
                out,
                scale=similarity.PREVIEW_SCALE,
                workers=2,
                project_path=project_path,
            )
            st.session_state[f"sim_report_{token}"] = {
                "pages": [page.as_dict() for page in report.pages],
                "mean": report.mean_score,
                "renderer": report.renderer,
                "warnings": report.warnings,
                "out": str(out),
                "project": str(project_path),
            }
        except RenderError as exc:
            st.warning(str(exc))
            return
        except Exception as exc:
            st.warning(f"تعذّر حساب التشابه: {exc}")
            return

    payload = st.session_state.get(f"sim_report_{token}")
    if not payload:
        return
    st.caption(
        f"العارض: {payload.get('renderer')} · متوسط الدرجة: {float(payload.get('mean') or 0):.1f}"
    )
    for warning in payload.get("warnings") or []:
        st.caption(warning)
    pages = payload.get("pages") or []
    page_numbers = [int(p["number"]) for p in pages]
    if not page_numbers:
        return
    cols = st.columns(min(6, len(page_numbers)))
    selected = st.selectbox(
        "صفحة المعاينة",
        options=page_numbers,
        key=f"sim_page_{token}",
    )
    for index, page in enumerate(pages):
        col = cols[index % len(cols)]
        score = page.get("score")
        color = _score_color(score, float(page.get("threshold") or 97))
        with col:
            st.markdown(
                f"<div style='border:2px solid {color};padding:0.35rem;border-radius:6px;"
                f"text-align:center;font-size:0.85rem'>ص {page['number']}<br>"
                f"<b style='color:{color}'>{float(score or 0):.0f}</b></div>",
                unsafe_allow_html=True,
            )
    page_info = next((p for p in pages if int(p["number"]) == int(selected)), None)
    if not page_info:
        return
    out = Path(payload["out"])
    pdf_img = out / f"page-{int(selected):03d}-pdf.png"
    docx_img = out / f"page-{int(selected):03d}-docx.png"
    diff_img = out / f"page-{int(selected):03d}-diff.png"
    show_diff = st.toggle("إظهار فروقات (diff)", value=False, key=f"sim_diff_{token}_{selected}")
    left, right = st.columns(2)
    with left:
        st.caption("المصدر PDF")
        if pdf_img.is_file():
            st.image(str(pdf_img), use_container_width=True)
    with right:
        st.caption("وورد بعد الطباعة" + (" + diff" if show_diff else ""))
        img = diff_img if show_diff and diff_img.is_file() else docx_img
        if img.is_file():
            st.image(str(img), use_container_width=True)
    place = page_info.get("placement") or {}
    st.caption(
        f"درجة {page_info.get('score')} · "
        f"إزاحة وسيطة {place.get('median_shift_pt')}pt · "
        f"≤1pt {place.get('pct_within_1pt')}% · ≤2pt {place.get('pct_within_2pt')}%"
    )
    mode = st.selectbox(
        "إعادة تحويل الصفحة",
        options=["", "faithful", "hybrid", "editable", "ocr_spans", "alt_tables", "picture"],
        format_func=lambda m: {
            "": "بحث تلقائي عن أفضل وضع",
            "faithful": "faithful",
            "hybrid": "hybrid (faithful + إصلاح RTL)",
            "editable": "editable",
            "ocr_spans": "OCR للمقاطع التالفة",
            "alt_tables": "جداول بديلة (editable)",
            "picture": "الصفحة كصورة (غير قابل للتحرير)",
        }.get(m, m),
        key=f"sim_mode_{token}_{selected}",
    )
    target = st.number_input(
        "الدرجة المستهدفة",
        min_value=50.0,
        max_value=100.0,
        value=float(page_info.get("threshold") or 97.0),
        key=f"sim_target_{token}_{selected}",
    )
    rec_col, undo_col, accept_col = st.columns(3)
    with rec_col:
        do_rec = st.button("إعادة التحويل", key=f"sim_reconvert_{token}_{selected}", width="stretch")
    with undo_col:
        do_undo = st.button("تراجع", key=f"sim_undo_{token}_{selected}", width="stretch")
    with accept_col:
        do_accept = st.button("قبول الصفحة", key=f"sim_accept_{token}_{selected}", width="stretch")
    if not (do_rec or do_undo or do_accept):
        return
    try:
        from core import similarity
        from core.project_state import load_project, save_project

        pdf_path = persist / "source.pdf"
        docx_path = persist / "result.docx"
        project_path = Path(payload["project"])
        pdf_path.write_bytes(result["pdf"])
        docx_path.write_bytes(result["bytes"])
        if do_accept:
            project = load_project(project_path)
            choice = project.page(int(selected))
            choice.accepted = True
            choice.needs_review = False
            save_project(project, project_path)
            st.toast(f"تم قبول الصفحة {selected}")
        elif do_undo:
            ok = similarity.undo_reconvert(pdf_path, docx_path, int(selected), project_path)
            if ok:
                result["bytes"] = docx_path.read_bytes()
                st.toast("تم التراجع عن إعادة التحويل")
            else:
                st.caption("لا يوجد تراجع لهذه الصفحة")
        else:
            before = float(page_info.get("score") or 0)
            rec = similarity.reconvert_page(
                pdf_path,
                docx_path,
                int(selected),
                mode=mode or None,
                target=float(target),
                project_path=project_path,
                out_dir=persist / "trials",
            )
            result["bytes"] = docx_path.read_bytes()
            st.toast(
                f"الصفحة {selected}: {before:.1f} → {rec.score:.1f} ({rec.mode})"
                + (" · تحتاج مراجعة" if rec.needs_review else "")
            )
        st.session_state[_WORD_RESULT] = result
        st.session_state.pop(f"sim_report_{token}", None)
        st.rerun()
    except Exception as exc:
        st.error(f"تعذّر تعديل الصفحة: {exc}")


def _show_word_result() -> None:
    result = st.session_state.get(_WORD_RESULT)
    if not result:
        return
    counts = result["counts"]
    skipped = list(counts.get("skipped") or [])
    st.success("المستند جاهز للتنزيل.", title="تم إنشاء ملف وورد")
    pages, paragraphs, headings, tables = st.columns(4)
    pages.metric("الصفحات", counts.get("pages", 0))
    paragraphs.metric("الفقرات", counts.get("paragraphs", 0))
    headings.metric("العناوين", counts.get("headings", 0))
    tables.metric("الجداول", counts.get("tables", 0))
    if skipped:
        st.warning("بقيت بعض الصفحات خارج المستند:\n\n" + "\n".join(f"• {item}" for item in skipped))
    retries = list(counts.get("retries") or [])
    if retries:
        st.info("أُعيدت قراءة بعض الصفحات قبل الحفظ:\n\n" + "\n".join(f"• {item}" for item in retries))
    _word_page_tools(result)
    _word_similarity_preview(result)
    st.download_button(
        "تنزيل ملف وورد",
        data=result["bytes"],
        file_name=result["name"],
        mime=_WORD_MIME,
        type="primary",
        width="stretch",
        key="word_download",
        icon=":material/download:",
        icon_position="right",
    )
    try:
        lines = _word_preview(result["bytes"])
    except Exception:
        st.caption("الملف جاهز للتنزيل. تعذّر رسم معاينة النص.")
        return
    if not lines:
        st.caption("النص الظاهر قليل. الجداول والصور داخل الملف نفسه.")
        return
    body = []
    for text, heading in lines:
        tag = "h4" if heading else "p"
        body.append(f"<{tag}>{html.escape(text)}</{tag}>")
    st.markdown(
        '<div class="section-label">معاينة النص</div>'
        f'<div class="preview">{"".join(body)}</div>',
        unsafe_allow_html=True,
    )


def _render_excel() -> None:
    form, aside = st.columns([1.55, 0.85], gap="large", vertical_alignment="top")
    with form:
        st.markdown(
            """
            <div class="tool-head">
              <p class="eyebrow">الأداة الثانية</p>
              <h2>جداول وفواتير إلى إكسل</h2>
              <p>بعد التحويل يظهر العدد، ثم خلية واحدة للمراجعة مع صورة الصف، ثم زر تنزيل واحد.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            "ملف PDF",
            type=["pdf"],
            key="excel_pdf",
            help="الجداول الرقمية تُقرأ من النص. النسخة الممسوحة تُقرأ على هذا الجهاز.",
        )
        _remember_upload("excel", uploaded)
        if uploaded is not None:
            st.markdown(
                f'<p class="file-meta">{html.escape(uploaded.name)} · {_size_label(uploaded.size)}</p>',
                unsafe_allow_html=True,
            )
        password = st.text_input(
            "كلمة المرور، إن كان الملف محمياً",
            type="password",
            key="excel_password",
            help="اترك الحقل فارغاً إذا كان الملف يُفتح من غير كلمة مرور.",
        )
        direction = st.radio(
            "اللغة والاتجاه",
            ["auto", "rtl", "ltr"],
            format_func=lambda item: {"auto": "تلقائي / مختلط", "rtl": "عربي - من اليمين لليسار", "ltr": "إنجليزي - من اليسار لليمين"}[item],
            horizontal=True,
            key="excel_direction",
        )
        with st.expander("متقدم", expanded=False):
            st.caption("بدون تخمين: الخانة الفارغة تبقى فارغة ولا تُنسخ من الصف الأعلى.")
            strict = st.checkbox("بدون تخمين", value=True, key="excel_strict")
            st.caption("الوضعان يقرآن الصفحة بالطريقة نفسها. السريع يقلل عدد العمليات والذاكرة.")
            mode = st.radio(
                "وضع القراءة",
                ["accurate", "fast"],
                format_func=lambda item: "دقيق" if item == "accurate" else "سريع",
                horizontal=True,
                key="excel_mode",
            )
        os.environ["PDF_TOOL_MODE"] = mode
        if uploaded is not None:
            try:
                from engine.benchmark import estimate_text

                guess = estimate_text(uploaded.getvalue())
            except Exception:
                guess = ""
            if guess:
                st.caption(guess)
        if st.button(
            "تحويل إلى إكسل",
            key="excel_convert",
            type="primary",
            width="stretch",
            disabled=uploaded is None,
        ):
            _run_excel(uploaded, password, strict, str(direction))
        _show_excel_result()
    with aside:
        st.markdown(
            _note(
                "ثلاث خطوات",
                [
                    ("ارفع الملف", "ثم اضغط تحويل."),
                    ("راجع الخلايا", "صورة الصف، ثم اختر القيمة أو اكتبها."),
                    ("نزّل الملف", "زر واحد، بعد المراجعة أو قبلها."),
                ],
            ),
            unsafe_allow_html=True,
        )


def _run_excel(uploaded, password: str, strict: bool = True, direction: str = "auto") -> None:
    with st.status("جارٍ قراءة الجداول...", expanded=True) as status:
        bar = st.progress(0.0, text="بدء القراءة")

        def report(message: str) -> None:
            match = re.search(r"Reading page (\d+) of (\d+)", message or "")
            eta = re.search(r"ETA (\d+)", message or "")
            if match:
                done = int(match.group(1))
                total = max(int(match.group(2)), 1)
                text = f"قراءة الصفحة {done} من {total}"
                if eta:
                    text += f" — المتبقي حوالي {eta.group(1)} ث"
                bar.progress(min(done / total, 1.0), text=text)
                status.write(text)
            elif message:
                status.write(message)

        started = time.perf_counter()
        try:
            state = _convert_excel(uploaded.getvalue(), uploaded.name, password, report, strict, direction)
            state["elapsed"] = round(time.perf_counter() - started, 1)
        except (ExcelError, ValueError) as exc:
            _fail(status, exc)
            return
        except Exception as exc:
            _fail(status, exc, unexpected=True)
            return
        status.update(label="اكتمل استخراج الجداول", state="complete", expanded=False)
    state["name"] = f"{_safe_stem(uploaded.name)}.xlsx"
    state["summary"] = {
        "rows": len(state.get("rows") or []),
        "figures": len(state.get("figures") or []),
    }
    if _restore_review_session(state):
        st.toast("تم استئناف جلسة المراجعة")
    st.session_state[_EXCEL_RESULT] = state
    st.toast("تم إنشاء ملف إكسل")


_EXCEL_MODES = (
    ("auto", "تلقائي"),
    ("grid", "شبكة الخطوط"),
    ("words", "إحداثيات الكلمات"),
    ("ocr", "مسح ضوئي"),
)


def _extraction_modes_guide() -> None:
    """A short Arabic explanation of the four extraction modes."""
    with st.expander("دليل أوضاع الاستخراج", expanded=False):
        st.markdown(
            """
**اللغة والاتجاه.** «تلقائي / مختلط» يتبع حروف الجدول. «عربي» يفتح الورقة من اليمين ويحاذي النص يميناً. «إنجليزي» يفتحها من اليسار.

**تلقائي.** الخيار الأول. يختار لكل صفحة الطريقة الأنسب حسب النص المضمّن، أو خطوط الجدول، أو كون الصفحة صورة.

**شبكة الخطوط.** عندما تكون حدود الجدول مرسومة بوضوح. الخانة الفارغة تبقى في عمودها، والفراغ في منتصف الصفحة أو صف المجموع أو ترويسة مكررة لا يزحزح الأعمدة المجاورة.

**إحداثيات الكلمات.** عندما تكون الخطوط مفقودة أو باهتة أو يقطعها فراغ أو عدّاد فرعي. تُثبَّت الكلمة داخل حدود عمودها الأفقي، ولا تنتقل إلى العمود المجاور إذا كان فارغاً.

**مسح ضوئي.** للصفحات المصورة، أو عندما تضيع خانة أو رقم أو تاريخ في القراءة العادية.

**بدون تخمين.** فراغ أو خط أو صف مجموع في منتصف الصفحة يقسمها إلى شريط أعلى وشريط أسفل. الترويسة والفاصل يُعزلان قبل تصحيح الأسماء. التاريخ واللون والصنع لا يتكرران من الصف المجاور، والخانة الفارغة تبقى فارغة وتظهر في الملاحظات.

إعادة قراءة نطاق صفوف تعيد استخراج الشريط الرأسي لتلك الصفوف فقط، وتترك باقي الصفحة وباقي الصفحات كما هي. عدد الصفوف المستخرجة وعدد المركبات ذات اللوحة يظهر لكل صفحة.
            """
        )


def _page_column_name(columns: list[str]) -> str | None:
    for name in columns:
        folded = str(name).replace(" ", "")
        if folded in {"رقمالصفحة", "Page"} or folded.lower() in {"page", "pageno"}:
            return str(name)
    return None


def _frame_for_page(result: dict, page) -> "pd.DataFrame":
    columns = list(result.get("columns") or [])
    rows = result.get("rows") or []
    numbers = result.get("page_numbers") or []
    if page == "all":
        visible = rows
    else:
        visible = [row for row, number in zip(rows, numbers) if int(number) == int(page)]
    padded = [list(row) + [""] * max(0, len(columns) - len(row)) for row in visible]
    if not columns:
        return pd.DataFrame()
    frame = pd.DataFrame(padded, columns=columns).fillna("").astype(str)
    if len(frame):
        frame.index = range(1, len(frame) + 1)
    return frame


def _commit_excel_edit(result: dict, edited: "pd.DataFrame", page) -> None:
    """Write the editor back into the page-tagged rows. Other pages stay as they are."""
    columns = list(result.get("columns") or [])
    if not columns or edited is None:
        return
    records = edited.fillna("").astype(str).values.tolist()
    records = [list(row)[: len(columns)] + [""] * max(0, len(columns) - len(row)) for row in records]
    if page == "all":
        page_name = _page_column_name(columns)
        page_at = columns.index(page_name) if page_name in columns else None
        numbers = []
        previous = list(result.get("page_numbers") or [])
        for index, row in enumerate(records):
            fallback = previous[index] if index < len(previous) else (previous[-1] if previous else 1)
            if page_at is None:
                numbers.append(int(fallback))
                continue
            digits = re.sub(r"\D", "", row[page_at])
            numbers.append(int(digits) if digits else int(fallback))
        result["rows"] = records
        result["page_numbers"] = numbers
        return
    page_no = int(page)
    # The selected page is replaced. Rows the user added stay on that page.
    edited_table = ExtractedTable(columns, records, None, [page_no] * len(records))
    merged, rows, numbers = replace_page_rows(
        columns,
        list(result.get("rows") or []),
        list(result.get("page_numbers") or []),
        [edited_table],
        page_no,
    )
    result["columns"] = merged
    result["rows"] = rows
    result["page_numbers"] = numbers


def _rebuild_excel_bytes(result: dict) -> bytes:
    if not result.get("rows") and not result.get("figures"):
        result["bytes"] = b""
        return b""
    with conversion_workspace("pdf-excel-edit-") as root:
        target = root / "edited.xlsx"
        write_edited_excel(
            list(result.get("columns") or []),
            list(result.get("rows") or []),
            list(result.get("page_numbers") or []),
            target,
            result.get("kind") or "",
            result.get("figures"),
            strict=bool(result.get("strict", True)),
            validation=result.get("validation"),
            direction=str(result.get("direction") or "auto"),
            provenance=result.get("provenance"),
            suggestions=result.get("suggestions"),
            raw_readings=result.get("raw_readings"),
            developer=bool(result.get("developer")),
            audit=result.get("audit"),
        )
        payload = read_output(target)
    result["bytes"] = payload
    result["summary"] = {
        "rows": len(result.get("rows") or []),
        "figures": len(result.get("figures") or []),
    }
    return payload


def _rescan_excel_page(result: dict, page: int, mode: str) -> None:
    data = result.get("pdf") or b""
    _require_pdf(data)
    with conversion_workspace("pdf-excel-page-") as root:
        source = root / "page.pdf"
        source.write_bytes(data)
        tables = extract_page_tables(
            source,
            int(page),
            result.get("password") or "",
            mode,
            result.get("kind") or "register",
            strict=bool(result.get("strict", True)),
        )
    columns, rows, numbers = replace_page_rows(
        list(result.get("columns") or []),
        list(result.get("rows") or []),
        list(result.get("page_numbers") or []),
        tables,
        int(page),
    )
    result["columns"] = columns
    result["rows"] = rows
    result["page_numbers"] = numbers
    result["revision"] = int(result.get("revision") or 0) + 1


def _page_rows_of(result: dict, page: int) -> list[list[str]]:
    return [
        list(row)
        for row, number in zip(result.get("rows") or [], result.get("page_numbers") or [])
        if int(number) == int(page)
    ]


def _rescan_excel_rows(result: dict, page: int, mode: str, start: int, end: int) -> None:
    """Re-read one vertical band and replace only that row span."""
    data = result.get("pdf") or b""
    _require_pdf(data)
    page_rows = _page_rows_of(result, page)
    with conversion_workspace("pdf-excel-band-") as root:
        source = root / "page.pdf"
        source.write_bytes(data)
        tables = extract_row_band(
            source,
            int(page),
            page_rows,
            list(result.get("columns") or []),
            int(start),
            int(end),
            result.get("password") or "",
            mode,
            result.get("kind") or "register",
            strict=bool(result.get("strict", True)),
        )
    columns, rows, numbers = replace_row_range(
        list(result.get("columns") or []),
        list(result.get("rows") or []),
        list(result.get("page_numbers") or []),
        int(page),
        int(start),
        int(end),
        tables,
    )
    result["columns"] = columns
    result["rows"] = rows
    result["page_numbers"] = numbers
    result["revision"] = int(result.get("revision") or 0) + 1


def _excel_row_tools(token: str, row_count: int, pdf_ready: bool, strict: bool = False):
    """Row span inputs and the quick actions for the selected page."""
    st.markdown('<div class="section-label">صف أو نطاق داخل الصفحة</div>', unsafe_allow_html=True)
    st.caption("الرقم 1 هو أول صف ظاهر في هذه الصفحة. النطاق يشمل البداية والنهاية.")
    start_box, end_box = st.columns(2)
    start_key = f"excel_row_start_{token}"
    end_key = f"excel_row_end_{token}"
    start_cap = max(row_count, int(st.session_state.get(start_key, 1) or 1))
    end_cap = max(row_count, int(st.session_state.get(end_key, 1) or 1))
    with start_box:
        start = st.number_input("من صف", min_value=1, max_value=start_cap, value=1, step=1, key=start_key)
    with end_box:
        end = st.number_input("إلى صف", min_value=1, max_value=end_cap, value=1, step=1, key=end_key)
    band_col, date_col, fill_col = st.columns(3)
    merge_col, drop_col, _spare = st.columns(3)
    with band_col:
        band = st.button(
            "إعادة قراءة النطاق",
            key="excel_rescan_rows",
            width="stretch",
            disabled=not pdf_ready,
            help="يعيد استخراج الشريط الرأسي لهذه الصفوف بالوضع المختار.",
        )
    with date_col:
        fix_dates = st.button("تصحيح التواريخ", key="excel_fix_dates", width="stretch")
    with fill_col:
        fill_above = st.button(
            "ملء الفراغ من الأعلى",
            key="excel_fill_above",
            width="stretch",
            disabled=strict,
            help="ينسخ اللون أو الصنع أو السنة من الصف الأعلى داخل هذه الصفحة. يتوقف عندما يكون «بدون تخمين» مفعّلاً.",
        )
    with merge_col:
        merge_rows = st.button("دمج في الصف الأعلى", key="excel_merge_rows", width="stretch")
    with drop_col:
        drop_rows = st.button("حذف النطاق", key="excel_drop_rows", width="stretch")
    return start, end, band, fix_dates, fill_above, merge_rows, drop_rows


def _excel_count_frame(result: dict) -> pd.DataFrame:
    """Extracted rows, plated vehicles, and header/gap rows for every page."""
    audits = {
        item["page"]: item
        for item in page_audit(
            list(result.get("columns") or []),
            list(result.get("rows") or []),
            list(result.get("page_numbers") or []),
        )
    }
    page_count = int(result.get("page_count") or 0)
    numbers = range(1, page_count + 1) if page_count else sorted(audits)
    records = []
    for number in numbers:
        item = audits.get(int(number), {})
        records.append(
            {
                "الصفحة": int(number),
                "صفوف مستخرجة": int(item.get("extracted") or 0),
                "مركبات بلوحة": int(item.get("plated") or 0),
                "ترويسة أو فاصل": int(item.get("separators") or 0),
            }
        )
    return pd.DataFrame(records)


def _tiers(result: dict) -> dict[str, list[dict]]:
    """Green, amber, and red. Cached until a cell changes, then the dismissed red cells drop out."""
    stamp = int(result.get("revision") or 0)
    cached = result.get("tier_cache")
    if isinstance(cached, dict) and cached.get("stamp") == stamp:
        tiers = cached["tiers"]
    else:
        if not result.get("print_date") and result.get("pdf"):
            result["print_date"] = read_print_date(result.get("pdf") or b"", result.get("password") or "")
        tiers = classify_cells(
            list(result.get("columns") or []),
            list(result.get("rows") or []),
            list(result.get("page_numbers") or []),
            str(result.get("print_date") or ""),
            result.get("provenance") or {},
            result.get("suggestions") or {},
            result.get("raw_readings") or {},
            result.get("pdf") or b"",
            result.get("baselines") or {},
            result.get("column_spans") or {},
        )
        result["tier_cache"] = {"stamp": stamp, "tiers": tiers}
    red_done = set(result.get("red_done") or [])
    yellow_done = set(result.get("yellow_done") or [])
    if not red_done and not yellow_done:
        return tiers
    return {
        **tiers,
        "red": [
            item
            for item in tiers["red"]
            if f"{item['page']}|{item['plate']}|{item['field']}" not in red_done
        ],
        "yellow": [
            item
            for item in tiers.get("yellow") or []
            if f"{item['page']}|{item['plate']}|{item['field']}" not in yellow_done
        ],
    }


def _issues(result: dict) -> list[dict]:
    """Mandatory review only. Green and amber are accepted."""
    return list(_tiers(result)["red"])


def _sheet_edits(columns: list[str], rows: list[list[str]], before: list[list[str]]) -> list[tuple[int, int, str]]:
    """Cells whose text changed. Excluded rows stay on their own sheet."""
    found: list[tuple[int, int, str]] = []
    for row_index, (old, new) in enumerate(zip(before, rows)):
        width = max(len(old), len(new))
        for column in range(width):
            old_text = blank_text(old[column] if column < len(old) else "")
            new_text = blank_text(new[column] if column < len(new) else "")
            if old_text != new_text:
                found.append((row_index, column, new_text))
    del columns
    return found


def _patch_excel_bytes(payload: bytes, columns: list[str], rows: list[list[str]], edits: list[tuple[int, int, str]]) -> bytes:
    """Change the edited cells in the workbook that is already on screen."""
    from openpyxl import load_workbook

    book = load_workbook(BytesIO(payload))
    sheet = book[book.sheetnames[0]]
    headers = [str(cell.value or "").replace("\u200e", "") for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    excel_row = 1
    excel_row_of: dict[int, int] = {}
    for index, row in enumerate(rows):
        if _nondata_kind(list(row), columns):
            continue
        excel_row += 1
        excel_row_of[index] = excel_row
    for row_index, column, value in edits:
        target = excel_row_of.get(row_index)
        if target is None or column >= len(columns):
            continue
        header = str(columns[column])
        try:
            sheet_column = headers.index(header) + 1
        except ValueError:
            sheet_column = column + 1
        kind = _export_column_kind(header)
        if kind == "date":
            written = _protect_excel_date(value) or blank_text(value)
        elif kind == "number":
            number = _coerce_excel_number(value)
            written = number if number is not None else blank_text(value)
        else:
            written = blank_text(value)
        sheet.cell(target, sheet_column).value = written
    output = BytesIO()
    book.save(output)
    return output.getvalue()


def _mark_action(result: dict, started: float, edits: list[tuple[int, int, str]] | None = None) -> None:
    """An edit stays in the table. The workbook is built when Download is pressed."""
    if edits:
        result["dirty"] = True
    elapsed = int((time.perf_counter() - started) * 1000)
    result["last_ms"] = elapsed
    recent = [int(item) for item in (result.get("action_ms") or [])]
    recent.append(elapsed)
    result["action_ms"] = recent[-12:]
    if elapsed >= 300:
        result["done_note"] = f"تم ✓ ({max(1, round(elapsed / 1000))} ث)"


def _notes_column(columns: list[str]) -> int | None:
    return next((index for index, name in enumerate(columns) if "ملاحظ" in str(name)), None)


def _clear_changed_notes(columns: list[str], rows: list[list[str]], edits: list[tuple[int, int, str]]) -> None:
    fields = {index: name for name, index in _vehicle_field_at(columns).items()}
    seen: set[tuple[int, str]] = set()
    for row_index, column, _value in edits:
        field = fields.get(column)
        if not field or (row_index, field) in seen or row_index < 0 or row_index >= len(rows):
            continue
        seen.add((row_index, field))
        clear_field_note(columns, rows[row_index], field)


def _write_cells(
    result: dict,
    changes: list[tuple[dict, str]],
    *,
    reclassify: bool = True,
    learn: bool = True,
) -> list[dict]:
    """Write the chosen text into the listed cells. One undo step covers the batch."""
    rows = result.get("rows") or []
    columns = list(result.get("columns") or [])
    notes_at = _notes_column(columns)
    batch: list[dict] = []
    original_notes: dict[int, str] = {}
    for item, new_value in changes:
        row_index = int(item["row_index"])
        column = int(item["index"])
        if row_index < 0 or row_index >= len(rows):
            continue
        row = rows[row_index]
        while len(row) <= column:
            row.append("")
        before = blank_text(row[column])
        cleaned = blank_text(new_value)
        if before == cleaned:
            continue
        if notes_at is not None and row_index not in original_notes:
            original_notes[row_index] = blank_text(row[notes_at] if notes_at < len(row) else "")
        row[column] = cleaned
        field = str(item.get("field") or "")
        if field:
            clear_field_note(columns, row, field)
        page_no = int(item["page"])
        plate = str(item.get("plate") or "")
        source_key = f"{page_no}|{plate}|{field}" if field and plate else ""
        provenance = result.setdefault("provenance", {})
        suggestions = result.setdefault("suggestions", {})
        raw_readings = result.setdefault("raw_readings", {})
        previous_source = provenance.get(source_key, "") if source_key else ""
        previous_suggestion = suggestions.get(source_key, "") if source_key else ""
        if source_key:
            if source_key not in raw_readings and before:
                raw_readings[source_key] = before
            provenance[source_key] = "manual"
            suggestions[source_key] = cleaned
        batch.append(
            {
                "row_index": row_index,
                "page": page_no,
                "plate": plate,
                "field": field,
                "column": column,
                "before": before,
                "after": cleaned,
                "notes_at": notes_at,
                "notes_before": original_notes.get(row_index, ""),
                "provenance_before": previous_source,
                "suggestion_before": previous_suggestion,
                "source_key": source_key,
            }
        )
        _remember_audit(result, page_no, plate, field, before, cleaned)
        if learn:
            remember_correction(field, before, cleaned)
    if not batch:
        return []
    result.setdefault("undo", []).append(batch)
    _touch_edit(result, reclassify)
    if reclassify:
        result["pending_undo"] = False
    _save_review_session(result)
    return batch


def _undo_last(result: dict, reclassify: bool = True) -> list[tuple[int, int, str]]:
    """Put back the last saved batch and return the cells that changed."""
    batches = result.get("undo") or []
    if not batches:
        return []
    batch = batches.pop()
    rows = result.get("rows") or []
    edits: list[tuple[int, int, str]] = []
    for change in batch:
        row_index = change.get("row_index")
        if row_index is None or int(row_index) >= len(rows):
            continue
        row = rows[int(row_index)]
        column = int(change["column"])
        while len(row) <= column:
            row.append("")
        restored = change.get("before") or ""
        row[column] = restored
        notes_at = change.get("notes_at")
        if notes_at is not None:
            while len(row) <= int(notes_at):
                row.append("")
            row[int(notes_at)] = change.get("notes_before") or ""
        source_key = str(change.get("source_key") or "")
        if source_key:
            provenance = result.setdefault("provenance", {})
            suggestions = result.setdefault("suggestions", {})
            previous = str(change.get("provenance_before") or "")
            if previous:
                provenance[source_key] = previous
            else:
                provenance.pop(source_key, None)
            previous_idea = str(change.get("suggestion_before") or "")
            if previous_idea:
                suggestions[source_key] = previous_idea
            else:
                suggestions.pop(source_key, None)
        edits.append((int(row_index), column, str(restored)))
        _remember_audit(
            result,
            change.get("page"),
            str(change.get("plate") or ""),
            str(change.get("field") or ""),
            str(change.get("after") or ""),
            str(restored),
        )
    _touch_edit(result, reclassify)
    _save_review_session(result)
    return edits


def _plates_on_page(result: dict, page: int) -> list[str]:
    columns = list(result.get("columns") or [])
    plate_at = next((index for index, name in enumerate(columns) if "لوحة" in str(name)), None)
    pages = list(result.get("page_numbers") or [])
    found: list[str] = []
    for index, row in enumerate(result.get("rows") or []):
        if index >= len(pages) or int(pages[index]) != int(page):
            continue
        if plate_at is None or plate_at >= len(row):
            found.append("")
        else:
            found.append(str(row[plate_at]).strip())
    return found


def _review_image(result: dict, page: int, plate: str, field: str) -> bytes:
    """Full-width row strip. A missing Y is taken from the other rows on the page."""
    del field
    data = result.get("pdf") or b""
    y_value = interpolate_row_y(
        result.get("baselines") or {},
        int(page),
        _plates_on_page(result, int(page)),
        str(plate),
    )
    if not data or y_value <= 0:
        return b""
    digest = hashlib.sha1(data).hexdigest()
    return _cached_row_image(digest, data, int(page), round(y_value, 2))


@st.cache_data(show_spinner=False)
def _cached_row_image(digest: str, data: bytes, page: int, y_value: float) -> bytes:
    del digest
    return crop_context_png(data, page, y_value, contrast=True)


def _prefetch_crop(result: dict, items: list[dict], cursor: int) -> None:
    """The next row's picture is ready before the user presses the button."""
    for item in items[cursor : cursor + 2]:
        _review_image(result, int(item["page"]), str(item["plate"]), str(item["field"]))


_FIELD_LABELS = {
    "plate": "رقم اللوحة",
    "make": "الصنع",
    "year": "السنة",
    "color": "اللون",
    "date": "انتهاء التأمين",
    "purpose": "غرض الترخيص",
    "admin": "رقم إداري",
}


def _clip_batch(result: dict, items: list[dict]) -> None:
    """Accept the safe track only: two readings of the same cell agree."""
    pairs = [(item, value) for item, value in clip_acceptances(items) if item.get("track") in {None, "safe"}]
    pairs = [(item, value) for item, value in pairs if item.get("track") != "exception"]
    if not pairs:
        st.caption("لا توجد قراءة آمنة. الاستثناء يُراجع من صورته، بلا قبول جماعي.")
        return
    st.caption(f"{len(pairs)} قراءة متفق عليها بين قراءتين. الاستثناءات ليست في هذا الزر.")
    if st.button("قبول كل القراءات الآمنة", type="primary", key="accept_all_clips", width="stretch"):
        started = time.perf_counter()
        written = _write_cells(result, [(item, value) for item, value in pairs])
        result["review_cursor"] = 0
        _mark_action(
            result,
            started,
            [(int(change["row_index"]), int(change["column"]), str(change["after"])) for change in written],
        )
        st.session_state[_EXCEL_RESULT] = result
        st.rerun()
    frame = pd.DataFrame(
        [
            {
                "تحديد": False,
                "الصفحة": int(item["page"]),
                "اللوحة": item.get("plate") or "",
                "الحقل": _FIELD_LABELS.get(item.get("field"), item.get("field")),
                "الاقتراح": value,
            }
            for item, value in pairs
        ]
    )
    edited = st.data_editor(
        frame,
        hide_index=True,
        width="stretch",
        height=220,
        num_rows="fixed",
        column_config={
            "تحديد": st.column_config.CheckboxColumn("تحديد"),
            "الاقتراح": st.column_config.TextColumn("الاقتراح", disabled=True),
        },
        key=f"clip_picker_{result.get('revision', 0)}",
    )
    if st.button("قبول القراءات الآمنة المحددة", key="accept_picked_clips"):
        flags = edited["تحديد"].tolist() if edited is not None and "تحديد" in getattr(edited, "columns", []) else []
        picked = [
            (item, value)
            for flag, (item, value) in zip(flags, pairs)
            if flag is True or str(flag).strip().lower() in {"true", "1", "yes"}
        ]
        if not picked:
            st.warning("لم يُحدَّد صف.")
            return
        started = time.perf_counter()
        written = _write_cells(result, picked)
        result["review_cursor"] = 0
        _mark_action(
            result,
            started,
            [(int(change["row_index"]), int(change["column"]), str(change["after"])) for change in written],
        )
        st.session_state[_EXCEL_RESULT] = result
        st.rerun()


def _shown_value(item: dict) -> str:
    ideas = [blank_text(value) for value in item.get("suggestions") or [] if blank_text(value)]
    return ideas[0] if ideas else blank_text(item.get("value"))


def _red_groups(items: list[dict]) -> list[dict]:
    """One crop per row. The first cell keeps the risk order of the red list."""
    groups: list[dict] = []
    seen: dict[tuple[int, str], dict] = {}
    for item in items:
        key = (int(item["page"]), str(item["plate"]))
        group = seen.get(key)
        if group is None:
            group = {"page": key[0], "plate": key[1], "row_index": item["row_index"], "cells": []}
            seen[key] = group
            groups.append(group)
        group["cells"].append(item)
    return groups


def _big(text: str, size: int = 22) -> None:
    st.markdown(f'<p style="font-size:{size}px;margin:0.2rem 0">{text}</p>', unsafe_allow_html=True)


def _boxed_review_image(result: dict, item: dict) -> bytes:
    """The row, plus the lines above and below, with a red box on this cell."""
    page = int(item["page"])
    plate = str(item["plate"])
    field = str(item["field"])
    data = result.get("pdf") or b""
    y_value = interpolate_row_y(
        result.get("baselines") or {},
        page,
        _plates_on_page(result, page),
        plate,
    )
    spans = (result.get("column_spans") or {}).get(str(page)) or {}
    box = spans.get(field)
    if field == "date" and not box and spans.get("color"):
        box = (8.0, max(16.0, float(spans["color"][0]) - 8.0))
    if not data or y_value <= 0:
        return b""
    if box:
        return crop_context_png(data, page, y_value, float(box[0]), float(box[1]), contrast=True)
    return _review_image(result, page, plate, field)


def _quick_picks(result: dict, item: dict) -> list[str]:
    """Other values from the same page, then the suggestions. The current reading is left out."""
    current = blank_text(item.get("value")).replace("\u200e", "")
    found: list[str] = []
    for value in item.get("suggestions") or []:
        text = blank_text(value).replace("\u200e", "")
        if text and text != current and text not in found:
            found.append(text)
    columns = list(result.get("columns") or [])
    fields = _vehicle_field_at(columns)
    index = fields.get(str(item.get("field") or ""))
    counts: dict[str, int] = {}
    for row, page in zip(result.get("rows") or [], result.get("page_numbers") or []):
        if int(page) != int(item["page"]) or index is None or index >= len(row):
            continue
        text = blank_text(row[index]).replace("\u200e", "")
        if text and text != current:
            counts[text] = counts.get(text, 0) + 1
    for text in sorted(counts, key=lambda name: (-counts[name], name)):
        if text not in found:
            found.append(text)
    return found[:8]


def ui_text(value: object) -> str:
    """Dates and digits stay left-to-right inside the Arabic page."""
    text = blank_text(value).replace("\u200e", "").replace("\u2066", "").replace("\u2069", "")
    if any(char.isdigit() for char in text):
        return f"\u2066{text}\u2069"
    return text


def _card_header(item: dict) -> str:
    label = _FIELD_LABELS.get(item.get("field"), item.get("field"))
    return f"لوحة {ui_text(item.get('plate'))} · صفحة {ui_text(item.get('page'))} · {label}"


def _row_y(result: dict, item: dict) -> float:
    """The plate's own baseline. A missing one is not borrowed from another row."""
    stored = float(item.get("y") or 0)
    if stored > 0:
        return stored
    return float((result.get("baselines") or {}).get(f"{int(item['page'])}|{item.get('plate')}") or 0)


def _page_bank(result: dict):
    return ensure_page_bank(result.get("pdf") or b"")


def _card_image(result: dict, item: dict) -> bytes:
    """Slice the cached page. A click does not open the PDF."""
    bank = _page_bank(result)
    if bank is None or not bank.ready:
        return b""
    page = int(item["page"])
    spans = (result.get("column_spans") or {}).get(str(page)) or {}
    return card_png(bank, page, _row_y(result, item), str(item.get("field") or ""), spans, result.get("baselines") or {})


def _prefetch_cards(result: dict, items: list[dict], cursor: int) -> None:
    for item in items[cursor : cursor + 4]:
        _card_image(result, item)


def _images_ready(result: dict) -> bool:
    bank = _page_bank(result)
    if bank is not None and bank.ready:
        return True
    st.info("جاري تجهيز صور الصفحات، يرجى الانتظار")

    @st.fragment(run_every=1)
    def _refresh_images() -> None:
        current = _page_bank(result)
        if current is not None and current.ready:
            st.rerun(scope="app")

    _refresh_images()
    return False


def _remember_item_crop(result: dict, item: dict, value: str) -> None:
    """Keep the crop of a cell the user just confirmed, from the cached page."""
    if not blank_text(value):
        return
    bank = _page_bank(result)
    if bank is None or not bank.ready:
        return
    try:
        page = int(item["page"])
        spans = (result.get("column_spans") or {}).get(str(page)) or {}
        image = cell_ink(bank, page, _row_y(result, item), str(item.get("field") or ""), spans, result.get("baselines") or {})
        remember_crop(str(item.get("field") or ""), value, image)
    except Exception:
        return


def _advance(result: dict, item: dict, new_value: str | None) -> None:
    started = time.perf_counter()
    confirmed = blank_text(item.get("value")) if new_value is None else blank_text(new_value)
    _remember_item_crop(result, item, confirmed)
    identity = f"{item['page']}|{item['plate']}|{item['field']}"
    done = list(result.get("red_done") or [])
    if identity not in done:
        done.append(identity)
    result["red_done"] = done
    result["red_asking"] = ""
    result["red_cursor"] = int(result.get("red_cursor") or 0) + 1
    written = []
    if new_value is not None and blank_text(new_value) != blank_text(item.get("value")):
        written = _write_cells(result, [(item, new_value)])
    _mark_action(
        result,
        started,
        [(int(change["row_index"]), int(change["column"]), str(change["after"])) for change in written],
    )
    st.session_state[_EXCEL_RESULT] = result
    st.rerun()


def _vision_log_path() -> Path:
    override = os.environ.get("PDF_TOOL_VISION_LOG", "").strip()
    if override:
        return Path(override)
    try:
        from engine.paths import using_app_data, vision_dir

        if using_app_data():
            return vision_dir() / "vision_log.jsonl"
    except Exception:
        pass
    return Path(__file__).resolve().parent / "vision_log.jsonl"


def _vision_cache_dir() -> Path:
    try:
        from engine.paths import using_app_data, vision_dir

        if using_app_data():
            return vision_dir()
    except Exception:
        pass
    return Path(__file__).resolve().parent / ".vision_cache"


def _run_vision_batch(result: dict, items: list[dict], api_key: str) -> None:
    """Read red and yellow cell crops only. Results are cached and every call is logged."""
    from core.excel_converter import vision_read_crop

    bank = _page_bank(result)
    if bank is None or not bank.ready:
        result["vision_running"] = False
        return
    folder = _vision_cache_dir()
    folder.mkdir(exist_ok=True)
    readings = dict(result.get("vision") or {})
    log_path = _vision_log_path()
    for item in items:
        identity = _item_identity(item)
        page = int(item["page"])
        spans = (result.get("column_spans") or {}).get(str(page)) or {}
        png = cell_png(bank, page, _row_y(result, item), str(item.get("field") or ""), spans, result.get("baselines") or {})
        if not png:
            continue
        digest = hashlib.sha1(png).hexdigest()
        cached = folder / f"{digest}.txt"
        was_cached = cached.is_file()
        if was_cached:
            text = cached.read_text(encoding="utf-8").strip()
        else:
            text = vision_read_crop(png, str(item.get("field") or ""), api_key).strip()
            cached.write_text(text, encoding="utf-8")
        readings[identity] = text
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(
                json.dumps(
                    {
                        "cell": identity,
                        "field": item.get("field"),
                        "ocr": blank_text(item.get("value")),
                        "vision": text,
                        "cached": was_cached,
                    },
                    ensure_ascii=False,
                )
                + "\n"
            )
    result["vision"] = readings
    result["vision_running"] = False


def _offer_vision(result: dict, items: list[dict]) -> None:
    """Off until the reviewer agrees. The desktop app never sends a crop anywhere."""
    if os.environ.get("PDF_TOOL_OFFLINE") == "1":
        st.caption("التطبيق يعمل دون اتصال. لا تُرسل صور الخلايا إلى أي خدمة.")
        return
    st.caption("نموذج الرؤية اختياري ومطفأ. لا يُرسل شيء قبل الموافقة.")
    opted = st.checkbox("إرسال صور الخلايا الحمراء والصفراء إلى نموذج رؤية", value=False, key="vision_opt_in")
    if not opted:
        result["vision_consent"] = False
        return
    st.warning(
        "بالموافقة تُرسل فقط صورة الخلية الحمراء أو الصفراء، وليس الصفحة. "
        "النتيجة تُحفظ هنا وتُسجَّل في سجل محلي، وتظهر كإشارة: توافق القراءة أو تختلف عنها."
    )
    api_key = st.text_input("مفتاح الواجهة", type="password", key="vision_api_key")
    if st.button("أوافق على إرسال صور هذه الخلايا فقط", key="vision_consent_button"):
        result["vision_consent"] = True
        result["vision_key"] = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
        st.session_state[_EXCEL_RESULT] = result
    if not result.get("vision_consent"):
        return
    key = str(result.get("vision_key") or os.environ.get("ANTHROPIC_API_KEY", "")).strip()
    if not key:
        st.error("لا يوجد مفتاح، ولم يُرسل شيء.")
        return
    if result.get("vision_running"):
        st.info("جاري قراءة الصور.")
        return
    if result.get("vision"):
        st.caption(f"وصلت قراءة الرؤية لـ {len(result.get('vision') or {})} خلية.")
        return
    result["vision_running"] = True

    def work() -> None:
        try:
            _run_vision_batch(result, list(items), key)
        except Exception as exc:
            result["vision_error"] = str(exc)
            result["vision_running"] = False

    threading.Thread(target=work, daemon=True).start()
    st.info("جاري قراءة الصور.")


def _vision_line(result: dict, item: dict) -> None:
    readings = result.get("vision") or {}
    seen = blank_text(readings.get(f"{item['page']}|{item['plate']}|{item['field']}"))
    if not seen:
        return
    if seen == blank_text(item.get("value")):
        _big("الرؤية توافق القراءة", 18)
    else:
        _big(f"الرؤية تختلف عن القراءة: {ui_text(seen)}", 18)


def _yellow_one(result: dict, items: list[dict]) -> None:
    """Unclear or empty text. The choices are on the card. There is no accept-all."""
    cursor = max(0, min(int(result.get("yellow_cursor") or 0), len(items) - 1))
    result["yellow_cursor"] = cursor
    item = items[cursor]
    label = _FIELD_LABELS.get(item["field"], item["field"])
    shown = ui_text(item.get("value")) or "فارغة"
    _big(f"مراجعة سريعة {ui_text(cursor + 1)} من {ui_text(len(items))}", 22)
    _big(_card_header(item), 28)
    if _images_ready(result):
        _prefetch_cards(result, items, cursor)
        image = _card_image(result, item)
        if image:
            st.image(image, width="stretch")
    _big(f"{label} المقروء: {shown}", 26)
    _big(str(item.get("reason") or ""), 18)
    _vision_line(result, item)
    identity = f"{item['page']}|{item['plate']}|{item['field']}"
    for index, value in enumerate(_quick_picks(result, item)):
        if st.button(value, key=f"yellow_{identity}_{index}", width="stretch"):
            _finish_yellow(result, item, value)
    typed = st.text_input("أو اكتب القيمة", key=f"yellow_type_{identity}")
    if st.button("حفظ والانتقال", key="yellow_save", width="stretch"):
        _finish_yellow(result, item, blank_text(typed))


def _finish_yellow(result: dict, item: dict, new_value: str) -> None:
    started = time.perf_counter()
    _remember_item_crop(result, item, blank_text(new_value) or blank_text(item.get("value")))
    identity = f"{item['page']}|{item['plate']}|{item['field']}"
    done = list(result.get("yellow_done") or [])
    if identity not in done:
        done.append(identity)
    result["yellow_done"] = done
    result["yellow_cursor"] = int(result.get("yellow_cursor") or 0) + 1
    written = []
    if blank_text(new_value) != blank_text(item.get("value")):
        written = _write_cells(result, [(item, new_value)])
    _mark_action(
        result,
        started,
        [(int(change["row_index"]), int(change["column"]), str(change["after"])) for change in written],
    )
    st.session_state[_EXCEL_RESULT] = result
    st.rerun()


def _review_one(result: dict, items: list[dict]) -> None:
    """One cell. A wide row, a red box, and two large answers."""
    st.markdown(
        """
        <style>
        div[class*="st-key-red_"] button,
        div[class*="st-key-pick_"] button,
        div[class*="st-key-yellow_"] button {
          font-size: 22px;
          min-height: 3.2rem;
        }
        div[class*="st-key-type_"] input {font-size: 20px;}
        div[class*="st-key-cluster_"] button {
          white-space: nowrap;
          min-height: 2.6rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    cursor = max(0, min(int(result.get("red_cursor") or 0), len(items) - 1))
    result["red_cursor"] = cursor
    item = items[cursor]
    label = _FIELD_LABELS.get(item["field"], item["field"])
    shown = ui_text(item.get("value")) or "فارغة"
    _big(f"{ui_text(cursor + 1)} من {ui_text(len(items))}", 22)
    _big(_card_header(item), 28)
    if _images_ready(result):
        _prefetch_cards(result, items, cursor)
        image = _card_image(result, item)
        if image:
            st.image(image, width="stretch")
    _big(f"{label} المقروء: {shown}. صح؟", 26)
    _big(str(item.get("reason") or ""), 18)
    _vision_line(result, item)
    identity = f"{item['page']}|{item['plate']}|{item['field']}"
    if result.get("red_asking") == identity:
        _big("اختر القيمة الصحيحة", 22)
        picks = _quick_picks(result, item)
        for index, value in enumerate(picks):
            if st.button(value, key=f"pick_{identity}_{index}", width="stretch"):
                _advance(result, item, value)
        typed = st.text_input("اكتب القيمة", key=f"type_{identity}")
        save_col, back_col = st.columns(2)
        with save_col:
            save = st.button("حفظ", type="primary", key="red_type_save", width="stretch")
        with back_col:
            back = st.button("رجوع", key="red_type_back", width="stretch")
        if save:
            _advance(result, item, blank_text(typed))
        if back:
            result["red_asking"] = ""
            st.session_state[_EXCEL_RESULT] = result
            st.rerun()
        return
    yes_col, no_col = st.columns(2)
    with yes_col:
        yes = st.button("✔ صح", key="red_yes", width="stretch")
    with no_col:
        no = st.button("✘ غلط", key="red_no", width="stretch")
    if yes:
        _advance(result, item, None)
    if no:
        result["red_asking"] = identity
        st.session_state[_EXCEL_RESULT] = result
        st.rerun()


def _follow_up(result: dict) -> None:
    pending = list(result.get("red_follow") or [])
    if not pending:
        return
    st.caption("هذه الخلايا علّمتها خاطئة. اكتب القيمة الصحيحة، أو اترك الحقل فارغاً.")
    revision = int(result.get("revision") or 0)
    boxes = []
    for item in pending:
        label = _FIELD_LABELS.get(item["field"], item["field"])
        key = f"red_fix_{item['page']}_{item['plate']}_{item['field']}_{revision}"
        boxes.append((item, key))
        st.text_input(f"لوحة {item['plate']} · {label}", key=key)
    if st.button("حفظ التصحيحات", key="red_save_fixes"):
        started = time.perf_counter()
        changes = [(item, blank_text(st.session_state.get(key))) for item, key in boxes]
        written = _write_cells(result, changes)
        result["red_follow"] = []
        _mark_action(
            result,
            started,
            [(int(change["row_index"]), int(change["column"]), str(change["after"])) for change in written],
        )
        st.session_state[_EXCEL_RESULT] = result
        st.rerun()


def _accepted_tiers(result: dict, tiers: dict[str, list[dict]]) -> None:
    """Accept-all stays on green and amber. Red has no such button."""
    green = list(tiers.get("green") or [])
    amber = list(tiers.get("amber") or [])
    sample = list(tiers.get("amber_sample") or [])
    if green and not result.get("green_done"):
        _big(f"متفق عليه: {len(green)} خلية. بلا مراجعة.", 18)
        if st.button("قبول كل المتفق عليه", key="accept_green"):
            result["green_done"] = True
            st.session_state[_EXCEL_RESULT] = result
            st.rerun()
    if result.get("amber_done"):
        _big(f"مقبول: {len(amber)} خلية.", 18)
        return
    _big(f"مقبول: {len(amber)} خلية. عينة للمراجعة: {len(sample)}.", 18)
    if amber and st.button("قبول كل المقبول", key="accept_amber"):
        result["amber_done"] = True
        st.session_state[_EXCEL_RESULT] = result
        st.rerun()
    if not sample:
        return
    with st.expander("عينة العنبر", expanded=False):
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "الصفحة": int(item["page"]),
                        "اللوحة": item["plate"],
                        "الحقل": _FIELD_LABELS.get(item["field"], item["field"]),
                        "القيمة": blank_text(item.get("value")) or "فارغة",
                    }
                    for item in sample
                ]
            ),
            hide_index=True,
            width="stretch",
        )


def _item_identity(item: dict) -> str:
    return f"{item['page']}|{item['plate']}|{item['field']}"


def _ink_bank(result: dict, items: list[dict]) -> dict:
    """Ink maps for clustering, sliced from the cached pages."""
    bank = _page_bank(result)
    if bank is None or not bank.ready:
        return {}
    found = {}
    for item in items:
        if str(item.get("field") or "") not in {"color", "make", "date"}:
            continue
        page = int(item["page"])
        spans = (result.get("column_spans") or {}).get(str(page)) or {}
        image = cell_ink(bank, page, _row_y(result, item), str(item.get("field") or ""), spans, result.get("baselines") or {})
        if image is not None:
            found[(str(page), str(item.get("plate")), str(item.get("field")))] = image
    return found


_SESSION_DIR = Path(__file__).resolve().parent / ".review_sessions"
_CLUSTER_PAGE = 6


def _reviewer_name() -> str:
    try:
        return getpass.getuser() or "المراجع"
    except Exception:
        return "المراجع"


def _remember_audit(result: dict, page: object, plate: str, field: str, before: str, after: str) -> None:
    label = _FIELD_LABELS.get(field, field)
    result.setdefault("audit", []).append(
        audit_record(_reviewer_name(), page, plate, label or field, before, after)
    )


def _touch_edit(result: dict, reclassify: bool) -> None:
    """A cluster accept must not throw away the tier cache. Download still sees a new serial."""
    result["edit_serial"] = int(result.get("edit_serial") or 0) + 1
    result["dirty"] = True
    if reclassify:
        result["revision"] = int(result.get("revision") or 0) + 1
        result.pop("review_cache", None)


def _session_payload(result: dict) -> dict:
    return {
        "rows": result.get("rows") or [],
        "red_done": list(result.get("red_done") or []),
        "yellow_done": list(result.get("yellow_done") or []),
        "cluster_out": list(result.get("cluster_out") or []),
        "audit": list(result.get("audit") or []),
        "undo": list(result.get("undo") or []),
        "revision": int(result.get("revision") or 0),
        "edit_serial": int(result.get("edit_serial") or 0),
    }


def _save_review_session(result: dict) -> None:
    pdf = result.get("pdf")
    if not isinstance(pdf, (bytes, bytearray)) or not pdf:
        return
    try:
        save_review_session(_SESSION_DIR, bytes(pdf), _session_payload(result))
    except (OSError, TypeError, ValueError):
        return


def _restore_review_session(state: dict) -> bool:
    pdf = state.get("pdf")
    if not isinstance(pdf, (bytes, bytearray)) or not pdf:
        return False
    data = load_review_session(_SESSION_DIR, bytes(pdf))
    saved = data.get("rows") if isinstance(data, dict) else None
    if not isinstance(saved, list) or len(saved) != len(state.get("rows") or []):
        return False
    state["rows"] = saved
    state["red_done"] = list(data.get("red_done") or [])
    state["yellow_done"] = list(data.get("yellow_done") or [])
    state["cluster_out"] = list(data.get("cluster_out") or [])
    state["audit"] = list(data.get("audit") or [])
    state["undo"] = list(data.get("undo") or [])
    state["revision"] = int(data.get("revision") or 0)
    state["edit_serial"] = int(data.get("edit_serial") or 1)
    state["resumed"] = True
    state["dirty"] = True
    return True


def _cluster_stamp(result: dict) -> tuple:
    return (
        int(result.get("revision") or 0),
        tuple(result.get("cluster_out") or []),
        tuple(result.get("red_done") or []),
        tuple(result.get("yellow_done") or []),
    )


def _ensure_review_total(result: dict, tiers: dict) -> int:
    stored = int(result.get("review_total") or 0)
    if stored:
        return stored
    done = set(result.get("red_done") or []) | set(result.get("yellow_done") or [])
    total = len(tiers.get("red") or []) + len(tiers.get("yellow") or []) + len(done)
    result["review_total"] = total
    return total


def _rerun_cluster() -> None:
    st.session_state[_EXCEL_RESULT] = st.session_state.get(_EXCEL_RESULT)
    st.rerun(scope="fragment")


def _review_clusters(result: dict, items: list[dict]) -> list[dict]:
    """Group the queued crops once. Accepting a group updates this list in place."""
    queued = [item for item in items if str(item.get("field") or "") in {"color", "make", "date"}]
    stamp = _cluster_stamp(result)
    if result.get("cluster_stamp") == stamp and isinstance(result.get("cluster_cache"), list):
        return result["cluster_cache"]
    if not queued or not result.get("pdf"):
        result["cluster_cache"] = []
        result["cluster_stamp"] = stamp
        return []
    bank = _page_bank(result)
    if bank is None or not bank.ready:
        return []
    revision = int(result.get("revision") or 0)
    if result.get("crop_bank_revision") != revision or not isinstance(result.get("crop_bank"), dict):
        result["crop_bank"] = _ink_bank(result, queued)
        result["crop_bank_revision"] = revision
    held = set(result.get("cluster_out") or [])
    done = set(result.get("red_done") or []) | set(result.get("yellow_done") or [])
    visible = [
        item
        for item in queued
        if _item_identity(item) not in held and _item_identity(item) not in done
    ]
    clusters = cluster_review_items(visible, result.get("crop_bank") or {})
    result["cluster_cache"] = clusters
    result["cluster_stamp"] = stamp
    return clusters


def _drop_cluster(result: dict, cluster: dict, identities: set[str]) -> None:
    index = int(result.get("cluster_cursor") or 0)
    kept = []
    for item in result.get("cluster_cache") or []:
        if item.get("id") == cluster.get("id"):
            continue
        members = [member for member in item.get("members") or [] if _item_identity(member) not in identities]
        outliers = [member for member in item.get("outliers") or [] if _item_identity(member) not in identities]
        if len(members) < 2:
            continue
        kept.append({**item, "members": members, "outliers": outliers})
    result["cluster_cache"] = kept
    result["cluster_cursor"] = min(index, max(0, len(kept) - 1)) if kept else 0
    result["cluster_grid"] = 0
    result["cluster_zoom"] = ""
    result["cluster_stamp"] = _cluster_stamp(result)


def _perform_accept(result: dict, cluster: dict, label: str) -> None:
    """Write one reading across the group. Outliers stay as they were. No reclassify."""
    started = time.perf_counter()
    text = blank_text(label).replace("\u200e", "")
    chosen = "" if text == "فارغة" else text
    members = [
        member
        for member in cluster.get("members") or []
        if _item_identity(member) not in set(result.get("cluster_out") or [])
    ]
    red_done = list(result.get("red_done") or [])
    yellow_done = list(result.get("yellow_done") or [])
    identities: list[str] = []
    pairs = []
    for member in members:
        identity = _item_identity(member)
        identities.append(identity)
        if member.get("tier") == "yellow":
            if identity not in yellow_done:
                yellow_done.append(identity)
        elif identity not in red_done:
            red_done.append(identity)
        current = blank_text(member.get("value")).replace("\u200e", "")
        if chosen != current:
            pairs.append((member, chosen))
    result["cluster_last"] = {
        "cluster": cluster,
        "index": int(result.get("cluster_cursor") or 0),
        "ids": list(identities),
    }
    result["red_done"] = red_done
    result["yellow_done"] = yellow_done
    written = _write_cells(result, pairs, reclassify=False, learn=False) if pairs else []
    _drop_cluster(result, cluster, set(identities))
    elapsed = int((time.perf_counter() - started) * 1000)
    _note_cluster_time(result, elapsed, len(identities))
    _mark_action(
        result,
        started,
        [(int(change["row_index"]), int(change["column"]), str(change["after"])) for change in written],
    )
    field_label = _FIELD_LABELS.get(str(cluster.get("field") or ""), cluster.get("field"))
    result["cluster_notice"] = f"✔ حُفظ {len(identities)} خلايا: {field_label} = {text or 'فارغة'} (تراجع)"
    result["cluster_notice_shown"] = False
    result["pending_undo"] = True
    result["cluster_scroll"] = True
    _save_review_session(result)
    _ensure_workbook_build(result)


def _undo_cluster(result: dict) -> list[tuple[int, int, str]]:
    last = result.pop("cluster_last", None)
    restored = _undo_last(result, reclassify=False)
    identities = set(last.get("ids") or []) if isinstance(last, dict) else set()
    if identities:
        result["red_done"] = [item for item in result.get("red_done") or [] if item not in identities]
        result["yellow_done"] = [item for item in result.get("yellow_done") or [] if item not in identities]
    if isinstance(last, dict) and last.get("cluster"):
        cache = list(result.get("cluster_cache") or [])
        cache.insert(min(int(last.get("index") or 0), len(cache)), last["cluster"])
        result["cluster_cache"] = cache
        result["cluster_cursor"] = int(last.get("index") or 0)
        result["cluster_stamp"] = _cluster_stamp(result)
    else:
        result.pop("cluster_cache", None)
        result.pop("cluster_stamp", None)
    result["pending_undo"] = False
    result["cluster_notice"] = ""
    result["cluster_scroll"] = True
    _save_review_session(result)
    return restored


def _note_cluster_time(result: dict, elapsed_ms: int, cells: int) -> None:
    bucket = [int(item) for item in (result.get("cluster_action_ms") or [])]
    bucket.append(int(elapsed_ms))
    result["cluster_action_ms"] = bucket[-12:]
    if cells <= 0:
        return
    per = [float(item) for item in (result.get("cell_ms") or [])]
    per.append(elapsed_ms / cells)
    result["cell_ms"] = per[-12:]
    result["ms_per_cell"] = sum(result["cell_ms"]) / len(result["cell_ms"])


def _cluster_is_slow(result: dict) -> bool:
    samples = [int(item) for item in (result.get("cluster_action_ms") or [])]
    return bool(samples) and (sum(samples) / len(samples)) >= 300


def _eta_seconds(result: dict, cells: int) -> int:
    per = float(result.get("ms_per_cell") or 0)
    if per <= 0:
        samples = [int(item) for item in (result.get("cluster_action_ms") or [])]
        per = (sum(samples) / len(samples)) if samples else 400
        return max(1, int(round(per / 1000)))
    return max(1, int(round(per * max(cells, 1) / 1000)))


def _queue_accept(result: dict, cluster: dict, label: str) -> None:
    cells = len(cluster.get("members") or [])
    if _cluster_is_slow(result):
        result["ui_locked"] = True
        result["ui_job"] = {
            "action": "accept",
            "cluster_id": cluster.get("id"),
            "label": label,
            "cells": cells,
            "go": False,
        }
        st.session_state[_EXCEL_RESULT] = result
        _rerun_cluster()
    _perform_accept(result, cluster, label)
    st.session_state[_EXCEL_RESULT] = result
    _rerun_cluster()


def _word_image(result: dict, item: dict) -> bytes:
    """The colour or make word only, at 4×. Repeat views read the bytes already sliced."""
    cache = result.setdefault("word_png", {})
    key = _item_identity(item)
    cached = cache.get(key)
    if isinstance(cached, (bytes, bytearray)) and cached:
        return bytes(cached)
    bank = _page_bank(result)
    if bank is None or not bank.ready:
        return b""
    page = int(item["page"])
    spans = (result.get("column_spans") or {}).get(str(page)) or {}
    png = cell_png(
        bank,
        page,
        _row_y(result, item),
        str(item.get("field") or ""),
        spans,
        result.get("baselines") or {},
        scale=4,
    )
    if png:
        if len(cache) > 80:
            cache.clear()
        cache[key] = png
    return png or b""


def _eject_from_cluster(result: dict, cluster: dict, identity: str) -> None:
    held = list(result.get("cluster_out") or [])
    if identity not in held:
        held.append(identity)
    result["cluster_out"] = held
    kept = []
    for item in result.get("cluster_cache") or []:
        if item.get("id") != cluster.get("id"):
            kept.append(item)
            continue
        members = [member for member in item.get("members") or [] if _item_identity(member) != identity]
        outliers = [member for member in item.get("outliers") or [] if _item_identity(member) != identity]
        if len(members) >= 2:
            kept.append({**item, "members": members, "outliers": outliers})
    result["cluster_cache"] = kept
    if kept:
        result["cluster_cursor"] = min(int(result.get("cluster_cursor") or 0), len(kept) - 1)
    result["cluster_stamp"] = _cluster_stamp(result)
    _save_review_session(result)


def _mark_scanned(result: dict, cluster: dict, display_len: int, offset: int) -> bool:
    seen = list(result.get("cluster_seen") or [])
    if cluster.get("id") in seen:
        return True
    if offset + _CLUSTER_PAGE >= display_len:
        seen.append(cluster.get("id"))
        result["cluster_seen"] = seen
        return True
    return False


def _scroll_to_cluster() -> None:
    st.components.v1.html(
        """
        <script>
        const node = window.parent.document.getElementById("cluster-top");
        if (node) { node.scrollIntoView({block: "start"}); }
        </script>
        """,
        height=0,
    )


def _cluster_across(field: str, count: int) -> int:
    """Dates are a long line, so each one gets the full row. Other cells share a row of three."""
    if count <= 1 or field == "date":
        return 1
    return min(3, count)


def _trim_crop(png: bytes) -> bytes:
    """Drop the blank margin and scale the ink up so a date line stays readable."""
    if not png:
        return b""
    from PIL import Image

    image = Image.open(BytesIO(png)).convert("RGB")
    try:
        mask = image.convert("L").point(lambda value: 255 if value < 215 else 0)
        box = mask.getbbox()
        if not box:
            return png
        pad = 14
        left = max(0, box[0] - pad)
        top = max(0, box[1] - pad)
        right = min(image.width, box[2] + pad)
        bottom = min(image.height, box[3] + pad)
        chip = image.crop((left, top, right, bottom))
        if chip.height < 96:
            factor = 96 / max(chip.height, 1)
            chip = chip.resize((max(1, int(chip.width * factor)), 96), Image.Resampling.LANCZOS)
        output = BytesIO()
        chip.save(output, format="PNG")
        return output.getvalue()
    finally:
        image.close()


def _show_crop(png: bytes, height: int) -> None:
    encoded = base64.b64encode(png).decode("ascii")
    st.markdown(
        (
            '<div style="overflow-x:auto; border:1px solid #e5e7eb; border-radius:12px; '
            'padding:12px 14px; background:#ffffff;">'
            f'<img alt="خلية" src="data:image/png;base64,{encoded}" '
            f'style="height:{int(height)}px; width:auto; max-width:none; display:block;" />'
            "</div>"
        ),
        unsafe_allow_html=True,
    )


def _paint_cluster(result: dict) -> None:
    """The group on screen. Only this fragment runs again after an accept."""
    st.markdown('<div id="cluster-top"></div>', unsafe_allow_html=True)
    if result.pop("cluster_scroll", None):
        _scroll_to_cluster()
    notice = str(result.get("cluster_notice") or "")
    if notice and not result.get("cluster_notice_shown"):
        st.toast(notice)
        result["cluster_notice_shown"] = True
    tiers = _tiers(result)
    done_ids = set(result.get("red_done") or []) | set(result.get("yellow_done") or [])
    total = _ensure_review_total(result, tiers)
    st.markdown(
        """
        <style>
        .review-progress {position: sticky; top: 0; z-index: 100; background: #ffffff; padding: 0.35rem 0;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    header = progress_text(len(done_ids), total, float(result.get("ms_per_cell") or 0))
    st.markdown(f'<div class="review-progress"><strong>{html.escape(header)}</strong></div>', unsafe_allow_html=True)
    if result.get("pending_undo") and st.button("تراجع", key="cluster_undo"):
        _undo_cluster(result)
        st.session_state[_EXCEL_RESULT] = result
        _rerun_cluster()
    clusters = list(result.get("cluster_cache") or [])
    covered = {
        _item_identity(member)
        for cluster in clusters
        for member in cluster.get("members") or []
    }
    open_red = [item for item in tiers.get("red") or [] if _item_identity(item) not in covered]
    red_rows = len({(item["page"], item["plate"]) for item in open_red})
    _big(f"أحمر بعد التجميع: {ui_text(len(open_red))} خلية في {ui_text(red_rows)} صف.", 20)
    if not clusters:
        return
    _big(f"مجموعات متشابهة: {ui_text(len(clusters))}. كل مجموعة سؤال واحد.", 20)
    cursor = max(0, min(int(result.get("cluster_cursor") or 0), len(clusters) - 1))
    result["cluster_cursor"] = cursor
    cluster = clusters[cursor]
    label = str(cluster.get("label") or "فارغة")
    field_label = _FIELD_LABELS.get(cluster["field"], cluster["field"])
    members = list(cluster.get("members") or [])
    outliers = list(cluster.get("outliers") or [])
    display = outliers + members
    offset = max(0, int(result.get("cluster_grid") or 0))
    if display and offset >= len(display):
        offset = 0
        result["cluster_grid"] = 0
    scanned = _mark_scanned(result, cluster, len(display), offset)
    decision = cluster_decision(cluster, set(result.get("cluster_out") or []), scanned)
    _big(f"مجموعة {ui_text(cursor + 1)} من {ui_text(len(clusters))}", 22)
    _big(f"{field_label}: هذه كلها: {ui_text(label)}؟", 28)
    _big(
        f"{ui_text(decision['count'])} خلية بنفس القراءة، صفحة {ui_text(cluster['page'])}.",
        18,
    )
    if outliers:
        _big(f"{ui_text(len(outliers))} خلية بقراءة مختلفة في أول الصف. القبول لا يكتبها.", 18)
    if not _images_ready(result):
        st.info("جاري تجهيز صور الخلايا")
        return
    window = display[offset : offset + _CLUSTER_PAGE]
    enlarged = str(result.get("cluster_zoom") or "")
    across = _cluster_across(str(cluster.get("field") or ""), len(window))
    for start in range(0, len(window), across):
        row = window[start : start + across]
        slots = st.columns(len(row))
        for slot, member in enumerate(row):
            with slots[slot]:
                identity = _item_identity(member)
                reading = blank_text(member.get("value")) or "فارغة"
                odd = any(_item_identity(item) == identity for item in outliers)
                st.caption(f"مختلفة: {reading} · {member.get('plate')}" if odd else str(member.get("plate") or ""))
                image = _trim_crop(_word_image(result, member))
                if image:
                    _show_crop(image, 168 if enlarged == identity else 96)
                eject_col, zoom_col = st.columns(2)
                with eject_col:
                    if st.button("إخراج", key=f"cluster_out_{cluster['id']}_{identity}", width="stretch"):
                        _eject_from_cluster(result, cluster, identity)
                        st.session_state[_EXCEL_RESULT] = result
                        _rerun_cluster()
                with zoom_col:
                    if st.button("تكبير", key=f"cluster_zoom_{cluster['id']}_{identity}", width="stretch"):
                        result["cluster_zoom"] = "" if enlarged == identity else identity
                        st.session_state[_EXCEL_RESULT] = result
                        _rerun_cluster()
    if len(display) > _CLUSTER_PAGE:
        back_col, next_col = st.columns(2)
        with back_col:
            if offset and st.button("السابق", key=f"cluster_prev_{cluster['id']}", width="stretch"):
                result["cluster_grid"] = max(0, offset - _CLUSTER_PAGE)
                st.session_state[_EXCEL_RESULT] = result
                _rerun_cluster()
        with next_col:
            if offset + _CLUSTER_PAGE < len(display) and st.button("التالي", key=f"cluster_next_{cluster['id']}", width="stretch"):
                result["cluster_grid"] = offset + _CLUSTER_PAGE
                st.session_state[_EXCEL_RESULT] = result
                _rerun_cluster()
    if not decision["ready"]:
        st.caption("أخرج الخلايا المختلفة، أو اعرض آخر الصفحة، قبل الحفظ.")
    if st.button(decision["label"], key=f"cluster_yes_{cluster['id']}", width="stretch", disabled=not decision["ready"]):
        _queue_accept(result, cluster, label)
    typed = st.text_input("أو قيمة واحدة للخلايا بنفس القراءة", key=f"cluster_type_{cluster['id']}")
    typed_ready = bool(decision["ready"] and blank_text(typed))
    if st.button(
        "تطبيق القيمة المكتوبة على هذه القراءة",
        key=f"cluster_apply_{cluster['id']}",
        width="stretch",
        disabled=not typed_ready,
    ):
        _queue_accept(result, cluster, blank_text(typed))


def _run_cluster_job(result: dict, job: dict) -> None:
    cluster = next(
        (item for item in result.get("cluster_cache") or [] if item.get("id") == job.get("cluster_id")),
        None,
    )
    if cluster is None:
        return
    _perform_accept(result, cluster, str(job.get("label") or cluster.get("label") or ""))


@st.fragment
def _cluster_board() -> None:
    result = st.session_state.get(_EXCEL_RESULT)
    if not isinstance(result, dict):
        return
    job = result.get("ui_job")
    if result.get("ui_locked") and isinstance(job, dict) and not job.get("go"):
        job["go"] = True
        result["ui_job"] = job
        st.markdown(
            "<style>div[data-testid='stButton'] button {pointer-events:none; opacity:0.45}</style>",
            unsafe_allow_html=True,
        )
        seconds = _eta_seconds(result, int(job.get("cells") or 1))
        st.warning(f"جاري الحفظ… (حوالي {ui_text(seconds)} ثانية)")
        st.session_state[_EXCEL_RESULT] = result
        st.rerun(scope="fragment")
    if isinstance(job, dict) and job.get("go"):
        result.pop("ui_job", None)
        result["ui_locked"] = False
        _run_cluster_job(result, job)
        st.session_state[_EXCEL_RESULT] = result
    _paint_cluster(result)


def _start_report(result: dict, tiers: dict) -> None:
    """Text or scan, difficulty, how many cells to review, and whether the columns were named."""
    kinds = list(result.get("page_kinds") or [])
    if not kinds and result.get("pdf"):
        kinds = describe_pages(result.get("pdf") or b"")
        result["page_kinds"] = kinds
    warning = str(result.get("column_warning") or "")
    missing = ("تعذر" in warning) or ("لا تعني" in warning)
    text_pages = sum(1 for item in kinds if item.get("kind") == "طبقة نص")
    scan_pages = sum(1 for item in kinds if item.get("kind") == "مسح")
    weak_pages = sum(1 for item in kinds if item.get("kind") == "نص غير مقروء")
    total = _ensure_review_total(result, tiers)
    score = difficulty_score(kinds, missing) if kinds else 0
    columns = "الأعمدة غير معروفة" if missing else "الأعمدة معروفة"
    line = (
        f"تقرير البداية: {text_pages} صفحة بطبقة نص، {scan_pages} صفحة مسح، {weak_pages} صفحة نص غير مقروء. "
        f"الصعوبة {score}. المتوقع مراجعته {total} خلية. {columns}."
    )
    if result.get("resumed"):
        line += " تم استئناف جلسة المراجعة."
    if missing:
        st.error(line + " لا تراجع الخلايا قبل تصحيح الأعمدة. قائمة فارغة لا تعني أن الملف سليم. " + warning)
    else:
        st.info(line)
    if kinds:
        st.caption(" · ".join(f"صفحة {item['page']}: {item['kind']}" for item in kinds))


def _excel_workspace(result: dict) -> None:
    """Upload is above this. Red cells are a contact sheet. Green and amber can be accepted together."""
    if result.get("done_note"):
        st.success(str(result.pop("done_note")))
        st.session_state[_EXCEL_RESULT] = result
    if result.get("bulk_wait"):
        recent = [int(item) for item in (result.get("action_ms") or [])]
        average = int(sum(recent) / len(recent)) if recent else 1000
        seconds = max(1, round(average / 1000))
        st.info(f"جاري الحفظ… (حوالي {ui_text(seconds)} ثانية)")
        st.progress(0.35)
        result["bulk_wait"] = False
        st.session_state[_EXCEL_RESULT] = result
        st.rerun()
    tiers = _tiers(result)
    items = list(tiers["red"])
    facts = banner_facts(list(result.get("validation") or []), len(items))
    if facts["count_ok"]:
        st.success(facts["count_line"])
    else:
        st.warning(facts["count_line"])
    status = queue_status(facts["issue_count"], str(result.get("column_warning") or ""))
    if "تعذر" in status or "لا تعني" in status:
        st.error(status)
    elif status.startswith("✅"):
        st.success(status)
    else:
        st.warning(status)
    visual_note = str(tiers.get("visual_note") or "")
    if visual_note:
        st.warning(visual_note + " — تُستخدم مكتبة الأشكال المؤكدة كمرجع.")
    _start_report(result, tiers)
    for page in facts["pages"]:
        st.write(page["line"])
        if st.button(f"راجع صفحة {page['page']}", key=f"goto_page_{page['page']}"):
            result["red_cursor"] = 0
            result["red_asking"] = ""
            for index, item in enumerate(items):
                if int(item["page"]) == int(page["page"]):
                    result["red_cursor"] = index
                    break
            st.session_state[_EXCEL_RESULT] = result
            st.rerun()
    _accepted_tiers(result, tiers)
    _follow_up(result)
    yellow = list(tiers.get("yellow") or [])
    review_items = list(items) + yellow
    if any(str(item.get("field") or "") in {"color", "make", "date"} for item in review_items) and not result.get("crop_bank"):
        with st.spinner("تجميع الخلايا المتشابهة"):
            _review_clusters(result, review_items)
    else:
        _review_clusters(result, review_items)
    _cluster_board()
    covered = {
        _item_identity(member)
        for cluster in (result.get("cluster_cache") or [])
        for member in cluster.get("members") or []
    }
    red_cards = [item for item in items if _item_identity(item) not in covered]
    yellow_cards = [item for item in yellow if _item_identity(item) not in covered]
    if red_cards:
        _review_one(result, red_cards)
    if yellow_cards:
        rows = len({(item["page"], item["plate"]) for item in yellow_cards})
        _big(f"أصفر بعد التجميع: {len(yellow_cards)} خلية في {rows} صف. قراءة غير واضحة أو فارغة.", 20)
        _yellow_one(result, yellow_cards)
    if result.get("undo") and st.button("تراجع عن آخر تعديل", key="excel_undo"):
        started = time.perf_counter()
        restored = _undo_cluster(result) if result.get("pending_undo") else _undo_last(result)
        _mark_action(result, started, restored)
        st.session_state[_EXCEL_RESULT] = result
        st.toast(f"أُلغي تعديل {len(restored)} خلية خلال {result['last_ms']} مللي ثانية")
        st.rerun()
    _download_excel(result)
    if result.get("last_ms") is not None:
        st.caption(f"آخر إجراء: {int(result['last_ms'])} مللي ثانية")
    _bulk_panel(result, items)
    _advanced_excel(result, list(items) + list(tiers.get("yellow") or []))


def _review_card(result: dict, items: list[dict], cursor: int) -> None:
    item = items[cursor]
    label = _FIELD_LABELS.get(item["field"], item["field"])
    st.caption(f"{cursor + 1} من {len(items)}")
    plate = item["plate"] or "بدون لوحة"
    st.markdown(f"**لوحة {plate} — صفحة {item['page']} — {label}**")
    st.caption(item["reason"])
    _prefetch_crop(result, items, cursor)
    image = _review_image(result, int(item["page"]), str(item["plate"]), str(item["field"]))
    if image:
        st.image(image, width="stretch")
    else:
        st.caption("لا تتوفر صورة لهذا الصف. القيمة الحالية ظاهرة أدناه.")
    current = blank_text(item.get("value"))
    st.caption("القيمة الحالية: " + (current if current else "فارغة"))
    field_key = f"review_text_{item['row_index']}_{item['field']}_{result.get('revision', 0)}"
    reasons = list(item.get("suggestion_reasons") or [])
    paired = [
        (blank_text(value), reasons[index] if index < len(reasons) else "")
        for index, value in enumerate(item.get("suggestions") or [])
        if blank_text(value)
    ]
    clips = [(value, reason) for value, reason in paired if suggestion_kind(reason) == "clip"]
    neighbours = [(value, reason) for value, reason in paired if suggestion_kind(reason) == "neighbour"]
    others = [
        (value, reason)
        for value, reason in paired
        if suggestion_kind(reason) not in {"clip", "neighbour"}
    ]
    if item.get("track") == "exception":
        st.caption("استثناء. لا يُقبل مع الدفعة. الصورة هي صف هذه اللوحة.")
    elif clips and item.get("track") == "safe":
        st.caption("قراءتان متفقتان. يمكن قبولها من زر القراءات الآمنة.")
    if clips and item.get("track") != "exception":
        if item.get("track") != "safe":
            st.caption("قراءة واحدة. لا تُقبل إلا من هذا الصف.")
        buttons = st.columns(len(clips))
        for index, (suggestion, reason) in enumerate(clips):
            with buttons[index]:
                if st.button(suggestion, key=f"suggest_{field_key}_clip_{index}", width="stretch"):
                    st.session_state[field_key] = suggestion
                    st.rerun()
                if reason:
                    st.caption(reason)
    if neighbours:
        st.caption("جوار نفس الصفحة. لا يُقبل إلا من هذا الصف.")
        buttons = st.columns(len(neighbours))
        for index, (suggestion, reason) in enumerate(neighbours):
            with buttons[index]:
                if st.button(suggestion, key=f"suggest_{field_key}_near_{index}", width="stretch"):
                    st.session_state[field_key] = suggestion
                    st.rerun()
                if reason:
                    st.caption(reason)
    if others:
        buttons = st.columns(len(others))
        for index, (suggestion, reason) in enumerate(others):
            with buttons[index]:
                if st.button(suggestion, key=f"suggest_{field_key}_other_{index}", width="stretch"):
                    st.session_state[field_key] = suggestion
                    st.rerun()
                if reason:
                    st.caption(reason)
    typed = st.text_input("أو اكتب القيمة", key=field_key)
    save_col, skip_col, close_col = st.columns(3)
    with save_col:
        save = st.button("حفظ والتالي", type="primary", key="excel_save_next", width="stretch")
    with skip_col:
        skip = st.button("تخطي", key="excel_skip", width="stretch")
    with close_col:
        close = st.button("إغلاق المراجعة", key="excel_close_review", width="stretch")
    if save:
        started = time.perf_counter()
        chosen = blank_text(typed)
        if not chosen:
            st.warning("اختر اقتراحاً أو اكتب قيمة. التخطي يترك الخلية كما هي.")
        else:
            current_key = (int(item["row_index"]), str(item["field"]))
            written = _write_cells(result, [(item, chosen)])
            fresh = _issues(result)
            keys = [(int(entry["row_index"]), str(entry["field"])) for entry in fresh]
            if current_key in keys:
                result["review_cursor"] = min(keys.index(current_key) + 1, max(len(fresh) - 1, 0))
            else:
                result["review_cursor"] = min(cursor, max(len(fresh) - 1, 0))
            _mark_action(result, started, [(int(change["row_index"]), int(change["column"]), str(change["after"])) for change in written])
            st.session_state[_EXCEL_RESULT] = result
            st.rerun()
    if skip:
        started = time.perf_counter()
        if cursor >= len(items) - 1:
            result["review_open"] = False
            result["review_cursor"] = 0
        else:
            result["review_cursor"] = cursor + 1
        _mark_action(result, started)
        st.session_state[_EXCEL_RESULT] = result
        st.rerun()
    if close:
        result["review_open"] = False
        st.session_state[_EXCEL_RESULT] = result
        st.rerun()


def _bulk_panel(result: dict, items: list[dict]) -> None:
    """Several cells can take one value, after the affected rows are shown."""
    if not items:
        return
    with st.expander("تعديل عدة صفوف", expanded=False):
        st.caption("حدّد الصفوف، اكتب قيمة واحدة، ثم أكّد بعد ظهور القائمة.")
        records = [
            {
                "الصفحة": int(item["page"]),
                "اللوحة": item["plate"],
                "الحقل": _FIELD_LABELS.get(item["field"], item["field"]),
                "القيمة": blank_text(item.get("value")),
                "السبب": item["reason"],
            }
            for item in items
        ]
        frame = pd.DataFrame(records).fillna("").astype(str)
        frame.insert(0, "تحديد", False)
        edited = st.data_editor(
            frame,
            hide_index=True,
            width="stretch",
            height=280,
            num_rows="fixed",
            column_config={
                "تحديد": st.column_config.CheckboxColumn("تحديد"),
                "القيمة": st.column_config.TextColumn("القيمة", disabled=True),
            },
            key=f"bulk_editor_{result.get('revision', 0)}",
        )
        new_value = st.text_input("القيمة لكل الصفوف المحددة", key=f"bulk_value_{result.get('revision', 0)}")
        if st.button("عرض الصفوف المحددة", key="bulk_preview"):
            chosen = []
            flags = edited["تحديد"].tolist() if edited is not None and "تحديد" in edited.columns else []
            for flag, item in zip(flags, items):
                if flag is True or str(flag).strip().lower() in {"true", "1", "yes"}:
                    chosen.append(item)
            if not chosen or not blank_text(new_value):
                st.warning("حدّد صفاً واحداً على الأقل واكتب القيمة.")
            else:
                result["bulk_pending"] = [
                    {
                        "row_index": item["row_index"],
                        "page": item["page"],
                        "plate": item["plate"],
                        "field": item["field"],
                        "index": item["index"],
                        "y": item.get("y") or 0,
                        "before": blank_text(item.get("value")),
                        "after": blank_text(new_value),
                    }
                    for item in chosen
                ]
                result["bulk_wait"] = True
                st.session_state[_EXCEL_RESULT] = result
                st.rerun()
        pending = list(result.get("bulk_pending") or [])
        if pending:
            preview = pd.DataFrame(
                [
                    {
                        "الصفحة": item["page"],
                        "اللوحة": item["plate"],
                        "الحقل": _FIELD_LABELS.get(item["field"], item["field"]),
                        "قبل": blank_text(item["before"]),
                        "بعد": blank_text(item["after"]),
                    }
                    for item in pending
                ]
            ).fillna("").astype(str)
            st.caption(f"تعديل {ui_text(len(pending))} خلية. كل خلية بصورتها.")
            with st.container(height=560):
                for item in pending:
                    _big(_card_header(item), 20)
                    image = _card_image(result, item)
                    if image:
                        st.image(image, width="stretch")
                    st.caption(f"{ui_text(item.get('before')) or 'فارغة'} ← {ui_text(item.get('after'))}")
            yes_col, no_col, undo_col = st.columns(3)
            with yes_col:
                confirm = st.button("تأكيد التعديل", type="primary", key="bulk_confirm", width="stretch")
            with no_col:
                cancel = st.button("إلغاء", key="bulk_cancel", width="stretch")
            if confirm:
                started = time.perf_counter()
                written = _write_cells(
                    result,
                    [(item, item["after"]) for item in pending],
                )
                result["bulk_pending"] = []
                _mark_action(result, started, [(int(change["row_index"]), int(change["column"]), str(change["after"])) for change in written])
                st.session_state[_EXCEL_RESULT] = result
                st.rerun()
            if cancel:
                result["bulk_pending"] = []
                st.session_state[_EXCEL_RESULT] = result
                st.rerun()
            with undo_col:
                if result.get("undo") and st.button("تراجع", key="bulk_undo", width="stretch"):
                    _undo_last(result)
                    result["bulk_pending"] = []
                    st.session_state[_EXCEL_RESULT] = result
                    st.rerun()


def _workbook_snapshot(result: dict) -> dict:
    """A copy the background build can read while the screen keeps editing."""
    return {
        "columns": list(result.get("columns") or []),
        "rows": [list(row) for row in result.get("rows") or []],
        "page_numbers": list(result.get("page_numbers") or []),
        "kind": result.get("kind") or "",
        "figures": result.get("figures"),
        "strict": bool(result.get("strict", True)),
        "direction": str(result.get("direction") or "auto"),
        "provenance": dict(result.get("provenance") or {}),
        "suggestions": dict(result.get("suggestions") or {}),
        "raw_readings": dict(result.get("raw_readings") or {}),
        "stated_total": result.get("stated_total"),
        "printed_groups": dict(result.get("printed_groups") or {}),
        "detected_pages": dict(result.get("detected_pages") or {}),
        "name": result.get("name") or "tables.xlsx",
        "developer": bool(result.get("developer")),
        "audit": [dict(item) for item in (result.get("audit") or []) if isinstance(item, dict)],
    }


def _ensure_workbook_build(result: dict) -> dict:
    """Start the workbook when review opens. A later edit starts it again."""
    revision = int(result.get("revision") or 0)
    serial = int(result.get("edit_serial") or 0)
    holder = st.session_state.get("_excel_bg")
    if (
        isinstance(holder, dict)
        and holder.get("revision") == revision
        and holder.get("serial") == serial
        and holder.get("developer") == bool(result.get("developer"))
        and (holder.get("running") or holder.get("ready"))
    ):
        return holder
    snapshot = _workbook_snapshot(result)
    holder = {
        "ready": False,
        "running": True,
        "bytes": b"",
        "error": "",
        "revision": revision,
        "serial": serial,
        "developer": bool(result.get("developer")),
        "started": time.perf_counter(),
    }

    def work() -> None:
        started = time.perf_counter()
        try:
            snapshot["validation"] = validate_register(
                snapshot["columns"],
                snapshot["rows"],
                snapshot["page_numbers"],
                snapshot["stated_total"],
                snapshot["printed_groups"],
                snapshot["detected_pages"],
            )
            holder["bytes"] = _rebuild_excel_bytes(snapshot)
            holder["seconds"] = round(time.perf_counter() - started, 2)
        except Exception as exc:
            holder["error"] = str(exc)
        holder["ready"] = True
        holder["running"] = False

    st.session_state["_excel_bg"] = holder
    threading.Thread(target=work, daemon=True).start()
    return holder


def _prepare_download(result: dict) -> None:
    """Build the workbook from the table as it is now, including a fresh validation sheet."""
    started = time.perf_counter()
    result["validation"] = validate_register(
        list(result.get("columns") or []),
        list(result.get("rows") or []),
        list(result.get("page_numbers") or []),
        result.get("stated_total"),
        result.get("printed_groups") or {},
        result.get("detected_pages") or {},
    )
    validated = time.perf_counter()
    result["bytes"] = _rebuild_excel_bytes(result)
    built = time.perf_counter()
    result["dirty"] = False
    result["built_revision"] = int(result.get("revision") or 0)
    result["last_ms"] = int((built - started) * 1000)
    timings = list(result.get("timings") or [])
    timings.append({"name": "validation_after_review", "seconds": round(validated - started, 3)})
    timings.append({"name": "build_excel", "seconds": round(built - validated, 3)})
    result["timings"] = timings


def _download_excel(result: dict) -> None:
    holder = _ensure_workbook_build(result)
    ready = bool(holder.get("ready") and holder.get("bytes") and not holder.get("error"))
    if holder.get("error"):
        st.error(holder["error"])
    if not ready:
        recent = [int(item) for item in (result.get("action_ms") or [])]
        average = int(sum(recent) / len(recent)) if recent else 2500
        seconds = max(1, round(average / 1000))
        st.info(f"جاري الحفظ… (حوالي {ui_text(seconds)} ثانية)")
        started_at = float(holder.get("started") or time.perf_counter())
        holder["started"] = started_at
        st.progress(min(0.95, (time.perf_counter() - started_at) / max(1.0, seconds)))

        @st.fragment(run_every=1)
        def _refresh_download() -> None:
            current = st.session_state.get("_excel_bg") or {}
            if current.get("ready"):
                st.rerun(scope="app")

        _refresh_download()
        return
    result["bytes"] = holder["bytes"]
    result["dirty"] = False
    seconds = holder.get("seconds")
    if seconds is not None:
        _big(f"تجهّز الملف خلال {seconds} ثانية", 18)
    st.download_button(
        "تنزيل ملف إكسل",
        data=holder.get("bytes") or b"",
        file_name=result.get("name") or "tables.xlsx",
        mime=_EXCEL_MIME,
        type="primary",
        width="stretch",
        key="excel_download",
    )


def _timing_rows(events: list[dict]) -> list[tuple[str, float]]:
    """One line per page, and one line for every weak-cell retry."""
    named = {
        "open_pdf": "فتح الملف",
        "cache": "من الذاكرة",
        "pages_elapsed": "قراءة الصفحات معاً",
        "figures": "الصور",
        "validation": "التحقق",
        "validation_after_review": "التحقق بعد المراجعة",
        "build_excel": "بناء ملف إكسل",
    }
    head: list[tuple[str, float]] = []
    tail: list[tuple[str, float]] = []
    pages: dict[str, dict[str, float]] = {}
    retry = second = ladder = 0.0
    retry_n = second_n = ladder_n = 0
    for item in events:
        name = str(item.get("name") or "")
        seconds = round(float(item.get("seconds") or 0), 2)
        if "_retry_" in name:
            retry += seconds
            retry_n += 1
        elif "_second_" in name:
            second += seconds
            second_n += 1
        elif "_ladder_" in name:
            ladder += seconds
            ladder_n += 1
        elif name.startswith("page_") and name.endswith("_ocr"):
            pages.setdefault(name.split("_")[1], {})["ocr"] = seconds
        elif name.endswith("_arabic_columns"):
            pages.setdefault(name.split("_")[1], {})["arabic"] = seconds
        elif name.endswith("_plate_band"):
            pages.setdefault(name.split("_")[1], {})["plate"] = seconds
        elif name.endswith("_color_band"):
            pages.setdefault(name.split("_")[1], {})["color"] = seconds
        elif name in {"validation", "validation_after_review", "build_excel"}:
            tail.append((named[name], seconds))
        elif name in named:
            head.append((named[name], seconds))
    body: list[tuple[str, float]] = []
    for page in sorted(pages, key=lambda item: int(item) if str(item).isdigit() else 0):
        info = pages[page]
        body.append((f"صفحة {page}: قراءة", info.get("ocr", 0)))
        if info.get("plate") or info.get("color"):
            body.append((f"صفحة {page}: عمود اللوحة واللون", round(info.get("plate", 0) + info.get("color", 0), 2)))
        if "arabic" in info:
            body.append((f"صفحة {page}: الصنع واللون", info.get("arabic", 0)))
    if retry_n:
        body.append((f"إعادة خلايا ضعيفة ({retry_n})", round(retry, 2)))
    if second_n:
        body.append((f"المحرك الثاني على خلايا ضعيفة ({second_n})", round(second, 2)))
    if ladder_n:
        body.append((f"تباين وتصحيح الميل ({ladder_n})", round(ladder, 2)))
    return head + body + tail


def _advanced_excel(result: dict, items: list[dict]) -> None:
    """Tools that are not part of the normal review."""
    with st.expander("متقدم", expanded=False):
        developer = st.checkbox(
            "ملف المطوّر: أوراق المصدر والصفوف المستبعدة",
            value=bool(result.get("developer")),
            key="excel_developer",
        )
        if bool(developer) != bool(result.get("developer")):
            result["developer"] = bool(developer)
            st.session_state[_EXCEL_RESULT] = result
            st.rerun()
        timings = list(result.get("timings") or [])
        if timings:
            st.caption("أوقات آخر استخراج بالثواني. قراءة الصفحات معاً هي الانتظار، ومجموع الصفحات قد يكون أكبر لأن العمل يتداخل.")
            st.dataframe(pd.DataFrame(_timing_rows(timings), columns=["المرحلة", "ثوانٍ"]), hide_index=True, width="stretch", height=280)
        columns = list(result.get("columns") or [])
        rows = list(result.get("rows") or [])
        pages = list(result.get("page_numbers") or [])
        st.caption("ملء من الأعلى: ينسخ الصنع أو اللون الفارغ من أقرب صف أعلاه في الصفحة نفسها.")
        if st.button("ملء من الأعلى", key="excel_fill_above_view"):
            started = time.perf_counter()
            before = [list(row) for row in rows]
            updated, _changed = fill_blank_fields_from_above(columns, rows, pages, ["make", "color"], None)
            result["rows"] = updated
            edits = _sheet_edits(columns, updated, before)
            _clear_changed_notes(columns, updated, edits)
            result["revision"] = int(result.get("revision") or 0) + 1
            result.pop("review_cache", None)
            _mark_action(result, started, edits)
            st.session_state[_EXCEL_RESULT] = result
            st.rerun()
        st.caption("تصحيح التواريخ: يوحّد التاريخ المكتمل إلى YYYY-MM-DD ولا يخترع سنة ولا يبدّل تاريخاً خاطئاً.")
        if st.button("تصحيح التواريخ", key="excel_normalize_dates"):
            started = time.perf_counter()
            before = [list(row) for row in rows]
            updated, _changed = normalize_edited_dates(columns, rows, None)
            result["rows"] = updated
            edits = _sheet_edits(columns, updated, before)
            _clear_changed_notes(columns, updated, edits)
            result["revision"] = int(result.get("revision") or 0) + 1
            result.pop("review_cache", None)
            _mark_action(result, started, edits)
            st.session_state[_EXCEL_RESULT] = result
            st.rerun()
        st.caption("إعادة قراءة صفحة: تعيد استخراج صفحة واحدة فقط، وتترك باقي الصفحات.")
        page_count = int(result.get("page_count") or 0) or max(pages or [1])
        mode = st.selectbox(
            "طريقة القراءة",
            [code for code, _label in _EXCEL_MODES],
            format_func=lambda code: dict(_EXCEL_MODES)[code],
            key="excel_mode_advanced",
        )
        st.caption("تلقائي يختار الطريقة. شبكة الخطوط للجداول المرسومة. إحداثيات الكلمات عندما تبهت الخطوط. مسح ضوئي للصفحة المصورة.")
        rescan_page = st.selectbox("الصفحة", list(range(1, int(page_count) + 1)), key="excel_rescan_page")
        if st.button("إعادة قراءة الصفحة", key="excel_rescan", disabled=not result.get("pdf")):
            with st.status(f"إعادة قراءة الصفحة {rescan_page}...", expanded=True) as status_box:
                try:
                    _rescan_excel_page(result, int(rescan_page), str(mode))
                except (ExcelError, ValueError) as exc:
                    _fail(status_box, exc)
                    return
                except Exception as exc:
                    _fail(status_box, exc, unexpected=True)
                    return
                result.pop("review_cache", None)
                result["dirty"] = True
                status_box.update(label=f"اكتملت قراءة الصفحة {rescan_page}. الملف يُبنى عند التنزيل.", state="complete", expanded=False)
            st.session_state[_EXCEL_RESULT] = result
            st.rerun()
        _data, _data_pages, excluded, excluded_pages = partition_rows(columns, rows, pages)
        st.caption("صفوف مستبعدة: ترويسات وفراغات لا تُحسب ضمن المركبات.")
        if excluded:
            excluded_frame = pd.DataFrame(
                [[blank_text(cell) for cell in list(row) + [""] * max(0, len(columns) - len(row))][: len(columns)] for row in excluded],
                columns=columns,
            )
            excluded_frame.insert(0, "الصفحة", excluded_pages)
            st.dataframe(excluded_frame.fillna("").astype(str), width="stretch", hide_index=True, height=220)
        else:
            st.caption("لا توجد صفوف مستبعدة.")
        _offer_vision(result, items)
        st.caption(
            f"مكتبة الأشكال: {crop_library_count()} شكلاً أكده المراجع. تبقى على هذا الجهاز، وتُستخدم عندما يقل تكرار القراءة عن 8."
        )
        st.download_button(
            "تصدير مكتبة الأشكال",
            export_crop_library(),
            file_name="learned_crops.zip",
            mime="application/zip",
            key="excel_crop_library_download",
        )
        if st.button("حذف مكتبة الأشكال", key="excel_crop_library_delete"):
            clear_crop_library()
            st.caption("حُذفت مكتبة الأشكال.")
        st.caption("القاموس: يحفظ تصحيحاً لإعادة استخدامه في ملف لاحق.")
        learned = load_corrections()
        learned_bytes = json.dumps(learned, ensure_ascii=False, indent=2).encode("utf-8")
        st.download_button(
            "تصدير القاموس",
            learned_bytes,
            file_name="learned_corrections.json",
            mime="application/json",
            key="excel_dictionary_download",
        )
        uploaded = st.file_uploader("استيراد القاموس", type=["json"], key="learned_import")
        if uploaded is not None:
            token = uploaded.getvalue()
            if st.session_state.get("learned_token") != token:
                try:
                    incoming = json.loads(token.decode("utf-8"))
                    if isinstance(incoming, dict):
                        merged = load_corrections()
                        for field, pairs in incoming.items():
                            if isinstance(pairs, dict):
                                merged.setdefault(str(field), {}).update({str(key): str(value) for key, value in pairs.items()})
                        save_corrections(merged)
                        st.session_state["learned_token"] = token
                        st.caption("تم دمج القاموس.")
                except (UnicodeDecodeError, json.JSONDecodeError):
                    st.caption("ملف القاموس غير صالح.")
    del items


def _show_excel_result() -> None:
    result = st.session_state.get(_EXCEL_RESULT)
    if not result:
        return
    result["strict"] = bool(st.session_state.get("excel_strict", result.get("strict", True)))
    if "rows" not in result:
        st.success("الورقة تُفتح من اليمين داخل إكسل.", title="تم إنشاء ملف إكسل")
        if result.get("bytes"):
            st.download_button(
                "تنزيل ملف إكسل",
                data=result["bytes"],
                file_name=result["name"],
                mime=_EXCEL_MIME,
                type="primary",
                width="stretch",
                key="excel_download_legacy",
            )
        return
    _excel_workspace(result)



def main() -> None:
    st.set_page_config(
        page_title="محوّل PDF إلى وورد وإكسل",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="collapsed",
        menu_items={
            "Get Help": None,
            "Report a bug": None,
            "About": "تحويل محلي من PDF إلى وورد وإكسل. الملفات تُعالج على هذا الجهاز.",
        },
    )
    st.markdown(_CSS, unsafe_allow_html=True)
    st.markdown(_HERO, unsafe_allow_html=True)
    try:
        from engine.ocr_status import describe
        from engine.version import VERSION

        level, text = describe()
        st.caption(f"الإصدار {VERSION}. التحديث: ثبّت ملف التثبيت الأحدث فوق هذه النسخة. الإعدادات تبقى على هذا الجهاز.")
        if level == "error":
            st.error(text)
        elif level == "warn":
            st.warning(text)
        else:
            st.caption(text)
    except Exception as exc:
        st.caption(f"تعذّر فحص محركات القراءة. {exc}")
    if _IMPORT_ERROR is not None:
        st.error("تعذّر تحميل أدوات التحويل. ثبّت المتطلبات ثم أعد تشغيل التطبيق: pip install -r requirements.txt")
        st.exception(_IMPORT_ERROR)
        return
    word_tab, excel_tab = st.tabs(["PDF إلى وورد", "PDF إلى إكسل"])
    with word_tab:
        _render_word()
    with excel_tab:
        _render_excel()
    st.markdown(
        '<p class="foot">تُعالج الملفات على هذا الجهاز.</p>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
