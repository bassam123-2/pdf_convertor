# -*- coding: utf-8 -*-
"""
General quality checks for a Word file made from a PDF.

The visual similarity score in visual_check rewards a page that *looks* like
the PDF, even when every line is a separate text box that nobody can edit.
These checks look at what a person meets when they open and edit the file:

    editable     share of the words that sit in normal paragraphs or tables
    spacing      runs whose letter spacing was stretched to fake a position
    duplicates   the same text drawn twice at the same spot (fake bold)
    stray        words that land in the margin or off the sheet when printed
    overlap      printed words that sit on top of other words
    text         words of the PDF missing from the Word file, or added to it
    pages        page count of the printed Word file against the PDF

None of the checks knows anything about a particular document. Each finding
names the page and quotes the text so it can be found in the file.

    python -m core.quality_check source.pdf converted.docx
    python -m core.quality_check source.pdf converted.docx --json report.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
import unicodedata
import zipfile
from collections import Counter
from dataclasses import dataclass, field, asdict
from pathlib import Path

import fitz

# Limits. A finding is raised when a measure crosses its limit.
SPACING_LIMIT_TWIPS = 100          # 5 pt of extra room after one letter or space
EDITABLE_MIN = 0.80                # share of words outside floating text boxes
STRAY_MARGIN_PT = 6.0              # how far past the PDF's text area a word may land
OVERLAP_MIN = 0.35                 # share of the smaller word covered by another
TEXT_RECALL_MIN = 0.97             # share of the PDF's words found in Word
TEXT_PRECISION_MIN = 0.97          # share of Word's words found in the PDF
DUPLICATE_TOLERANCE_EMU = 3 * 12700  # 3 pt

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_ARABIC = re.compile(r"[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]")
_DIACRITICS = re.compile(r"[ؐ-ًؚ-ٰٟۖ-ۭـ]")
_PUNCT = re.compile(r"[^\w؀-ۿ]+", re.UNICODE)


@dataclass
class Finding:
    check: str
    severity: str  # "error" or "warning"
    message: str
    page: int | None = None
    sample: str = ""


@dataclass
class QualityReport:
    pdf: str
    docx: str
    renderer: str = ""
    measures: dict = field(default_factory=dict)
    findings: list[Finding] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not any(f.severity == "error" for f in self.findings)

    def counts(self) -> dict:
        result: dict[str, int] = {}
        for finding in self.findings:
            result[finding.check] = result.get(finding.check, 0) + 1
        return result

    def to_dict(self) -> dict:
        data = asdict(self)
        data["ok"] = self.ok
        data["counts"] = self.counts()
        return data


# --------------------------------------------------------------------------
# Words


def normalize_word(word: str) -> str:
    """One spelling per word: no tashkeel, no tatweel, no punctuation, NFKC."""
    word = unicodedata.normalize("NFKC", word)
    word = _DIACRITICS.sub("", word)
    word = _PUNCT.sub("", word)
    word = word.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    return word.lower()


def _words_of(text: str) -> list[str]:
    found = []
    for raw in text.split():
        word = normalize_word(raw)
        if word:
            found.append(word)
    return found


def _bag(words: list[str]) -> Counter:
    """
    Count words so that a word read backwards still matches.

    Some PDFs store Arabic in visual order, so the same word comes out of the
    PDF reversed. Each word is keyed by the smaller of itself and its reverse.
    """
    return Counter(min(word, word[::-1]) if _ARABIC.search(word) else word for word in words)


_READABLE = re.compile(r"^[0-9A-Za-z\u00C0-\u024F\u0600-\u06FF\u0750-\u077F]+$")


def page_words(page: fitz.Page) -> tuple[list[str], int]:
    """
    The words of one page, and how many of them could not be read.

    A word drawn twice at the same spot (a common way to fake bold) counts
    once. A word with letters outside Arabic, Latin and digits came from a font
    with a broken character map; it is left out, because no converter can be
    held to a spelling the PDF itself does not give.
    """
    kept: list[str] = []
    spots: dict[str, list[fitz.Rect]] = {}
    unreadable = 0
    from core.ligature_repair import _FOREIGN, truncated_ligature

    for text, rect in _trace_words(page):
        text = "".join(truncated_ligature(ch) or ch for ch in text)
        if _FOREIGN.search(text) and _ARABIC.search(text):
            # A ligature whose code the PDF lost: the PDF does not say which
            # letters it is, so the word cannot be held against the output.
            unreadable += 1
            continue
        word = normalize_word(text)
        if not word:
            continue
        if any(abs(r.x0 - rect.x0) < 2.5 and abs(r.y0 - rect.y0) < 2.5 for r in spots.get(word, [])):
            continue
        spots.setdefault(word, []).append(rect)
        if not _READABLE.match(word):
            unreadable += 1
            continue
        kept.append(word)
    return kept, unreadable


def _trace_words(page: fitz.Page) -> list[tuple[str, fitz.Rect]]:
    """
    Words with their boxes, read glyph by glyph.

    A lam-alef ligature is one glyph that stands for two letters; the second
    letter comes back with no glyph of its own (glyph id -1). A right-to-left
    run is stored in visual order, so the run is turned around to read it,
    but the ligature's two letters must stay in their own order. Plain text
    extraction turns them around too and reports التعامالت for التعاملات, a
    fault in the reading, not in the file.
    """
    words: list[tuple[str, fitz.Rect]] = []
    try:
        traces = page.get_texttrace() or []
    except Exception:
        traces = []
    for trace in traces:
        if int(trace.get("type") or 0) != 0:
            continue
        units: list[tuple[str, fitz.Rect]] = []
        for raw in trace.get("chars") or []:
            char = chr(raw[0])
            rect = fitz.Rect(raw[3])
            if raw[1] < 0 and units:
                text, box = units[-1]
                units[-1] = (text + char, box | rect)
                continue
            units.append((char, rect))
        if int(trace.get("bidi_lvl") or 0) % 2 == 1 or (
            units and sum(1 for t, _ in units if _ARABIC.search(t)) > len(units) / 2
            and units[0][1].x0 > units[-1][1].x0
        ):
            units.reverse()
        size = float(trace.get("size") or 10)
        current, box = "", None
        previous: fitz.Rect | None = None
        for text, rect in units + [(" ", None)]:
            # A writer may leave the space out and just move the pen: a gap
            # wider than a quarter of the type size ends the word too.
            gap = rect is not None and previous is not None and (
                max(rect.x0, previous.x0) - min(rect.x1, previous.x1) > size * 0.25
            )
            if text.isspace() or gap:
                if current.strip():
                    words.append((current, box))
                current, box = "", None
                if text.isspace():
                    previous = None
                    continue
            current += text
            box = rect if box is None else box | rect
            previous = rect
    if not words:
        words = [(str(w[4]), fitz.Rect(w[:4])) for w in page.get_text("words")]
    return words


def pdf_words(pdf: fitz.Document) -> list[list[str]]:
    return [page_words(page)[0] for page in pdf]


# --------------------------------------------------------------------------
# The Word package


def _xml_parts(docx: Path) -> dict[str, str]:
    with zipfile.ZipFile(docx) as package:
        names = [n for n in package.namelist() if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n)]
        return {name: package.read(name).decode("utf-8") for name in names}


def _text_of(xml: str) -> str:
    """Visible text of an XML fragment, one paragraph per line."""
    paragraphs = re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
    lines = []
    for paragraph in paragraphs:
        # A paragraph inside a text box is also inside its anchor's paragraph;
        # keep only the innermost text by skipping paragraphs that hold boxes.
        if "<w:txbxContent" in paragraph:
            continue
        lines.append("".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", paragraph)))
    return "\n".join(lines)


def _strip_boxes(xml: str) -> str:
    xml = re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", xml, flags=re.S)
    return re.sub(r"<w:txbxContent>.*?</w:txbxContent>", "", xml, flags=re.S)


def _box_texts(xml: str) -> list[str]:
    xml = re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", xml, flags=re.S)
    return [_text_of(box) for box in re.findall(r"<w:txbxContent>(.*?)</w:txbxContent>", xml, re.S)]


def check_structure(docx: Path, report: QualityReport, source_word_count: int, expect_editable: bool = True) -> None:
    parts = _xml_parts(docx)
    body = parts.get("word/document.xml", "")
    box_words = sum(len(_words_of(t)) for xml in parts.values() for t in _box_texts(xml))
    flow_words = sum(len(_words_of(_text_of(_strip_boxes(xml)))) for xml in parts.values())
    total = box_words + flow_words
    editable = flow_words / total if total else 1.0
    boxes = len(re.findall(r"<w:txbxContent>", _strip_boxes_fallback(body)))
    report.measures.update(
        text_boxes=boxes,
        words_in_boxes=box_words,
        words_in_flow=flow_words,
        editable_share=round(editable, 3),
        source_words=source_word_count,
    )
    if total and editable < EDITABLE_MIN:
        report.findings.append(Finding(
            "editable", "error" if expect_editable else "warning",
            f"Only {editable:.0%} of the words are in normal paragraphs; {box_words} words sit in "
            f"{boxes} floating text boxes. Word counts, search, editing and reflow will not work as expected.",
        ))


def _strip_boxes_fallback(xml: str) -> str:
    return re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", xml, flags=re.S)


def check_spacing(docx: Path, report: QualityReport) -> None:
    parts = _xml_parts(docx)
    stretched = []
    for name, xml in parts.items():
        xml = _strip_boxes_fallback(xml)
        for paragraph in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
            if "<w:txbxContent" in paragraph:
                continue
            runs = re.findall(r"<w:r[ >].*?</w:r>", paragraph, re.S)
            line = "".join("".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", r)) for r in runs)
            for run in runs:
                match = re.search(r'<w:spacing w:val="(-?\d+)"', run)
                if match and abs(int(match.group(1))) > SPACING_LIMIT_TWIPS:
                    stretched.append((int(match.group(1)), line.strip()))
    report.measures["stretched_runs"] = len(stretched)
    report.measures["max_spacing_pt"] = round(max((abs(v) for v, _ in stretched), default=0) / 20, 1)
    for value, line in sorted(stretched, key=lambda item: -abs(item[0]))[:15]:
        report.findings.append(Finding(
            "spacing", "error" if abs(value) > SPACING_LIMIT_TWIPS * 4 else "warning",
            f"A gap was stretched by {value / 20:.1f} pt to push the next word into place.",
            sample=line[:90],
        ))
    if len(stretched) > 15:
        report.findings.append(Finding(
            "spacing", "warning", f"{len(stretched) - 15} more stretched gaps not listed.",
        ))


def check_duplicates(docx: Path, report: QualityReport) -> None:
    """Two text boxes with the same text at (nearly) the same spot."""
    xml = _strip_boxes_fallback(_xml_parts(docx).get("word/document.xml", ""))
    # One sheet at a time: the same heading at the same spot on two pages is fine.
    doubles = []
    for sheet in re.split(r'<w:br w:type="page"/>|<w:sectPr\b', xml):
        doubles += _doubles_on_sheet(sheet)
    report.measures["duplicate_boxes"] = len(doubles)
    for text in doubles[:10]:
        report.findings.append(Finding("duplicates", "error", "The same text is drawn twice at one spot.", sample=text[:60]))


def _doubles_on_sheet(xml: str) -> list[str]:
    anchors = re.findall(r"<wp:anchor.*?</wp:anchor>", xml, re.S)
    placed = []
    for anchor in anchors:
        h = re.search(r"<wp:positionH[^>]*>\s*<wp:posOffset>(-?\d+)</wp:posOffset>", anchor)
        v = re.search(r"<wp:positionV[^>]*>\s*<wp:posOffset>(-?\d+)</wp:posOffset>", anchor)
        texts = [t for t in _box_texts(anchor) if t.strip()]
        if h and v and texts:
            placed.append((int(h.group(1)), int(v.group(1)), normalize_word("".join(texts))))
    seen: dict[str, list[tuple[int, int]]] = {}
    doubles = []
    for x, y, text in placed:
        if not text:
            continue
        for ox, oy in seen.get(text, []):
            if abs(ox - x) <= DUPLICATE_TOLERANCE_EMU and abs(oy - y) <= DUPLICATE_TOLERANCE_EMU:
                doubles.append(text)
                break
        seen.setdefault(text, []).append((x, y))
    return doubles


# --------------------------------------------------------------------------
# The printed Word file


def _text_area(page: fitz.Page) -> fitz.Rect | None:
    area = None
    for item in page.get_text("words"):
        rect = fitz.Rect(item[:4])
        area = rect if area is None else area | rect
    for drawing in page.get_drawings():
        rect = drawing.get("rect")
        if rect is not None and rect.width < page.rect.width * 0.98:
            area = rect if area is None else area | rect
    return area


def _overlaps(words: list) -> list[tuple[str, str]]:
    boxes = sorted(((fitz.Rect(w[:4]), str(w[4])) for w in words), key=lambda item: item[0].y0)
    found = []
    for index, (rect, text) in enumerate(boxes):
        if rect.is_empty or rect.width < 1:
            continue
        for other, other_text in boxes[index + 1:]:
            if other.y0 > rect.y1:
                break
            cross = rect & other
            if cross.is_empty:
                continue
            smaller = min(rect.get_area(), other.get_area()) or 1.0
            if cross.get_area() / smaller >= OVERLAP_MIN:
                found.append((text, other_text))
    return found


def check_printed(pdf: fitz.Document, docx: Path, report: QualityReport, renderer: str, flowing: bool = False) -> None:
    from core.visual_check import render_docx

    with tempfile.TemporaryDirectory(prefix="pdf_tool_qc_") as work:
        printed_path = Path(work) / "printed.pdf"
        report.renderer = render_docx(docx, printed_path, renderer=renderer)
        printed = fitz.open(printed_path)
        try:
            _check_printed_pages(pdf, printed, report, flowing)
        finally:
            printed.close()


def _check_printed_pages(pdf: fitz.Document, printed: fitz.Document, report: QualityReport, flowing: bool = False) -> None:
    """
    `flowing`: the file is meant to reflow (editable mode). Text may then move
    from its PDF spot and a page may run a little longer; only words off the
    sheet count as stray, and a small change in page count is a warning.
    """
    report.measures["pdf_pages"] = pdf.page_count
    report.measures["word_pages"] = printed.page_count
    if printed.page_count != pdf.page_count:
        close = abs(printed.page_count - pdf.page_count) <= max(1, round(pdf.page_count * 0.15))
        report.findings.append(Finding(
            "pages", "warning" if flowing and close else "error",
            f"The PDF has {pdf.page_count} pages; the Word file prints on {printed.page_count}.",
        ))

    stray_total = overlap_total = 0
    for index in range(min(pdf.page_count, printed.page_count)):
        source, result = pdf[index], printed[index]
        area = _text_area(source)
        words = [(*tuple(rect), text) for text, rect in _trace_words(result) if text.strip()]
        if flowing:
            sheet = result.rect + (-1, -1, 1, 1)
            stray = [str(w[4]) for w in words if not sheet.contains(fitz.Rect(w[:4]))]
            stray_total += len(stray)
            if stray:
                report.findings.append(Finding(
                    "stray", "error", f"{len(stray)} words print off the sheet.",
                    page=index + 1, sample=" ".join(stray[:8])[:90],
                ))
        elif area is not None:
            # The printed sheet may be a different size; scale the PDF's area onto it.
            sx = result.rect.width / source.rect.width
            sy = result.rect.height / source.rect.height
            allowed = fitz.Rect(area.x0 * sx, area.y0 * sy, area.x1 * sx, area.y1 * sy)
            allowed = allowed + (-STRAY_MARGIN_PT, -STRAY_MARGIN_PT, STRAY_MARGIN_PT, STRAY_MARGIN_PT)
            stray = [str(w[4]) for w in words if not fitz.Rect(w[:4]).intersects(allowed)
                     or fitz.Rect(w[:4]).x1 > result.rect.width or fitz.Rect(w[:4]).x0 < 0]
            stray_total += len(stray)
            if stray:
                report.findings.append(Finding(
                    "stray", "error",
                    f"{len(stray)} words print outside the area the PDF uses (margin or off the sheet).",
                    page=index + 1, sample=" ".join(stray[:8])[:90],
                ))
        overlaps = _overlaps(words)
        overlap_total += len(overlaps)
        if overlaps:
            report.findings.append(Finding(
                "overlap", "error" if len(overlaps) >= 3 else "warning",
                f"{len(overlaps)} printed words sit on top of other words.",
                page=index + 1, sample=" | ".join(f"{a} / {b}" for a, b in overlaps[:4])[:90],
            ))
    report.measures["stray_words"] = stray_total
    report.measures["overlapping_words"] = overlap_total

    source_words, unreadable = [], 0
    for page in pdf:
        words, bad = page_words(page)
        source_words += words
        unreadable += bad
    if len(source_words) < 5 * max(1, pdf.page_count) * 0.2 and not unreadable:
        report.findings.append(Finding(
            "text", "warning",
            "The PDF has no text layer (a scan), so its words cannot be compared with Word.",
        ))
        return
    source_bag = _bag(source_words)
    result_bag = _bag([w for page in printed for w in page_words(page)[0]])
    report.measures["source_unreadable_words"] = unreadable
    common = sum((source_bag & result_bag).values())
    recall = common / (sum(source_bag.values()) or 1)
    # Words the PDF itself garbled cannot be matched either way; leave the
    # same number out of Word's side.
    precision = common / (max(1, sum(result_bag.values()) - unreadable))
    report.measures["text_recall"] = round(recall, 3)
    report.measures["text_precision"] = round(precision, 3)
    if recall < TEXT_RECALL_MIN:
        missing = list((source_bag - result_bag).elements())[:12]
        report.findings.append(Finding(
            "text", "error", f"{1 - recall:.1%} of the PDF's words are missing or misspelt in Word.",
            sample=" ".join(missing)[:90],
        ))
    if precision < TEXT_PRECISION_MIN:
        extra = list((result_bag - source_bag).elements())[:12]
        report.findings.append(Finding(
            "text", "error", f"{1 - precision:.1%} of Word's words are not in the PDF (doubled, split or garbled).",
            sample=" ".join(extra)[:90],
        ))


# --------------------------------------------------------------------------


def check_word(
    pdf_path: str | Path,
    docx_path: str | Path,
    renderer: str = "auto",
    printed: bool = True,
    expect_editable: bool = True,
) -> QualityReport:
    """
    Run every check. `expect_editable` is False for the faithful mode, where
    text boxes are the design; there a low editable share is only a warning.
    """
    pdf_path, docx_path = Path(pdf_path), Path(docx_path)
    report = QualityReport(str(pdf_path), str(docx_path))
    pdf = fitz.open(pdf_path)
    try:
        source_words = sum(len(words) for words in pdf_words(pdf))
        check_structure(docx_path, report, source_words, expect_editable)
        check_spacing(docx_path, report)
        check_duplicates(docx_path, report)
        if printed:
            try:
                check_printed(pdf, docx_path, report, renderer, flowing=expect_editable)
            except Exception as exc:  # no renderer, or it failed
                report.findings.append(Finding("render", "warning", f"Printed checks skipped: {exc}"))
    finally:
        pdf.close()
    return report


def format_report(report: QualityReport) -> str:
    lines = [f"{'PASS' if report.ok else 'FAIL'}  {Path(report.docx).name}  (printed with {report.renderer or '-'})"]
    lines.append("  " + ", ".join(f"{k}={v}" for k, v in report.measures.items()))
    for finding in report.findings:
        where = f"p{finding.page} " if finding.page else ""
        sample = f"  «{finding.sample}»" if finding.sample else ""
        lines.append(f"  [{finding.severity}] {finding.check}: {where}{finding.message}{sample}")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Check a Word file made from a PDF.")
    parser.add_argument("pdf")
    parser.add_argument("docx")
    parser.add_argument("--renderer", default="auto", choices=("auto", "word", "libreoffice"))
    parser.add_argument("--no-print", action="store_true", help="Skip the checks that print the Word file")
    parser.add_argument("--faithful", action="store_true", help="The file was made in faithful mode (text boxes allowed)")
    parser.add_argument("--json", help="Write the report as JSON to this path")
    args = parser.parse_args(argv)
    report = check_word(args.pdf, args.docx, renderer=args.renderer, printed=not args.no_print,
                        expect_editable=not args.faithful)
    print(format_report(report))
    if args.json:
        Path(args.json).write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if report.ok else 1


if __name__ == "__main__":
    sys.exit(main())
