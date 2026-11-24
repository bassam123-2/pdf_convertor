# -*- coding: utf-8 -*-
"""
Word-placement debug: compare PDF words with the Word page after conversion.

    python -m core.placement_debug source.pdf result.docx --page 6 --out report_folder

For each PDF word, finds the nearest same-text word on the rendered Word page
(by OCR of the docx render, falling back to the engine's written line boxes).
Prints dx/dy/width deltas and flags words that left their line, cell, or page.
Draws an overlay of PDF boxes on the docx render with displaced/duplicated marks.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path

import fitz

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.visual_check import RenderError, render_docx

SCALE = 2.0


@dataclass
class WordBox:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    source: str = ""

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2.0

    @property
    def width(self) -> float:
        return max(0.0, self.x1 - self.x0)

    @property
    def height(self) -> float:
        return max(0.0, self.y1 - self.y0)


@dataclass
class WordDelta:
    text: str
    pdf: WordBox
    docx: WordBox | None
    dx: float
    dy: float
    dwidth: float
    flags: list[str] = field(default_factory=list)
    score: float = 0.0


@dataclass
class PagePlacement:
    number: int
    words: int
    matched: int
    mean_shift: float
    median_shift: float
    pct_within_1pt: float
    pct_within_2pt: float
    displaced: int
    duplicated_pdf: int
    worst: list[WordDelta] = field(default_factory=list)
    overlay: str = ""
    cell_count: int = 0


def _norm(text: str) -> str:
    text = re.sub(r"\s+", "", (text or "").strip())
    return text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))


def pdf_words(page: fitz.Page) -> list[WordBox]:
    words: list[WordBox] = []
    try:
        for item in page.get_text("words") or []:
            x0, y0, x1, y1, text = item[0], item[1], item[2], item[3], item[4]
            text = (text or "").strip()
            if not text:
                continue
            words.append(WordBox(text, float(x0), float(y0), float(x1), float(y1), "pdf"))
    except Exception:
        pass
    return words


def pdf_duplicate_pairs(words: list[WordBox], max_shift: float = 2.5) -> list[tuple[WordBox, WordBox]]:
    """Identical text with nearly the same center — classic fake-bold overprint."""
    pairs = []
    for i, a in enumerate(words):
        for b in words[i + 1 : i + 40]:
            if _norm(a.text) != _norm(b.text) or not _norm(a.text):
                continue
            if abs(a.cx - b.cx) <= max_shift and abs(a.cy - b.cy) <= max_shift:
                if abs(a.cx - b.cx) > 0.05 or abs(a.cy - b.cy) > 0.05 or abs(a.width - b.width) > 0.5:
                    pairs.append((a, b))
    return pairs


def cell_rects(page: fitz.Page) -> list[tuple[float, float, float, float]]:
    """Large stroked/filled rectangles that look like table cells or column frames."""
    cells = []
    page_area = max(1.0, float(page.rect.width) * float(page.rect.height))
    try:
        drawings = page.get_drawings() or []
    except Exception:
        drawings = []
    for drawing in drawings:
        rect = drawing.get("rect")
        if rect is None:
            continue
        x0, y0, x1, y1 = float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1)
        w, h = x1 - x0, y1 - y0
        if w < 40 or h < 40:
            continue
        if w * h < page_area * 0.04:
            continue
        if w * h > page_area * 0.95:
            continue
        cells.append((x0, y0, x1, y1))
    # Also rebuild cells from upright rules + outer frame (bilingual column tables).
    xs: list[float] = []
    ys: list[float] = []
    for drawing in drawings:
        for item in drawing.get("items") or []:
            kind = item[0]
            if kind == "l":
                first, second = item[1], item[2]
                if abs(float(first.x) - float(second.x)) <= 0.75 and abs(float(first.y) - float(second.y)) >= 40:
                    xs.append((float(first.x) + float(second.x)) / 2.0)
                    ys.extend([float(first.y), float(second.y)])
                elif abs(float(first.y) - float(second.y)) <= 0.75 and abs(float(first.x) - float(second.x)) >= 40:
                    ys.append((float(first.y) + float(second.y)) / 2.0)
                    xs.extend([float(first.x), float(second.x)])
            elif kind in {"re", "qu"}:
                box = item[1] if kind == "re" else item[1].rect
                if float(box.x1) - float(box.x0) >= 40 and float(box.y1) - float(box.y0) >= 40:
                    xs.extend([float(box.x0), float(box.x1)])
                    ys.extend([float(box.y0), float(box.y1)])
    xs = sorted({round(x, 1) for x in xs})
    ys = sorted({round(y, 1) for y in ys})
    # Collapse near-duplicates.
    def _unique(vals: list[float], tol: float = 3.0) -> list[float]:
        out: list[float] = []
        for value in vals:
            if not out or abs(value - out[-1]) > tol:
                out.append(value)
        return out

    xs = _unique(xs)
    ys = _unique(ys)
    if len(xs) >= 2 and len(ys) >= 2:
        for i in range(len(xs) - 1):
            for j in range(len(ys) - 1):
                x0, x1 = xs[i], xs[i + 1]
                y0, y1 = ys[j], ys[j + 1]
                if x1 - x0 < 40 or y1 - y0 < 40:
                    continue
                if (x1 - x0) * (y1 - y0) < page_area * 0.03:
                    continue
                cells.append((x0, y0, x1, y1))
    # Fall back: upright rules alone define column strips (bilingual frames).
    if not cells and len(xs) >= 2:
        top = min(ys) if ys else 0.0
        bottom = max(ys) if ys else float(page.rect.height)
        if bottom - top < 40:
            top, bottom = 0.0, float(page.rect.height)
        for i in range(len(xs) - 1):
            x0, x1 = xs[i], xs[i + 1]
            if x1 - x0 < 40:
                continue
            cells.append((x0, top, x1, bottom))
    # Deduplicate nearly identical rects.
    kept: list[tuple[float, float, float, float]] = []
    for cell in sorted(cells, key=lambda c: (c[1], c[0])):
        if any(abs(cell[0] - k[0]) < 2 and abs(cell[1] - k[1]) < 2 and abs(cell[2] - k[2]) < 2 and abs(cell[3] - k[3]) < 2 for k in kept):
            continue
        kept.append(cell)
    return kept


def containing_cell(box: WordBox, cells: list[tuple[float, float, float, float]]):
    for cell in cells:
        if cell[0] - 1 <= box.cx <= cell[2] + 1 and cell[1] - 1 <= box.cy <= cell[3] + 1:
            return cell
    return None


def engine_words_from_faithful(page: fitz.Page) -> list[WordBox]:
    """Where the converter intends to place words (faithful line/chunk boxes)."""
    from core import faithful

    payload = faithful.read_page(page)
    words: list[WordBox] = []
    for line in payload.lines:
        if not line.chunks:
            text = (line.text or "").strip()
            if text:
                words.append(WordBox(text, line.x0, line.y0, line.x1, line.y1, "engine"))
            continue
        # Reconstruct chunk positions along the line from recorded widths.
        if line.rtl:
            cursor = line.x1
            for chunk in line.chunks:
                width = chunk.width if chunk.width is not None else max(1.0, len(chunk.text) * line.size * 0.5)
                x1 = cursor
                x0 = cursor - width
                cursor = x0
                if chunk.space or not chunk.text.strip():
                    continue
                words.append(WordBox(chunk.text.strip(), x0, line.y0, x1, line.y1, "engine"))
        else:
            cursor = line.x0
            for chunk in line.chunks:
                width = chunk.width if chunk.width is not None else max(1.0, len(chunk.text) * line.size * 0.5)
                x0 = cursor
                x1 = cursor + width
                cursor = x1
                if chunk.space or not chunk.text.strip():
                    continue
                words.append(WordBox(chunk.text.strip(), x0, line.y0, x1, line.y1, "engine"))
    return words


def ocr_words_from_image(png_path: Path, page_width: float, page_height: float) -> list[WordBox]:
    """OCR the rendered Word page; map pixel boxes back to PDF points."""
    try:
        from PIL import Image
    except Exception:
        return []
    image = Image.open(png_path)
    pw, ph = image.size
    scale_x = page_width / max(pw, 1)
    scale_y = page_height / max(ph, 1)
    data = png_path.read_bytes()
    # Prefer Windows OCR eng+ar; fall back to Tesseract.
    words = _windows_ocr(data, scale_x, scale_y)
    if not words:
        words = _tesseract_ocr(data, scale_x, scale_y)
    return words


def _windows_ocr(png: bytes, scale_x: float, scale_y: float) -> list[WordBox]:
    try:
        from winrt.runtime import ApartmentType, init_apartment
        from winrt.windows.globalization import Language
        from winrt.windows.graphics.imaging import BitmapAlphaMode, BitmapDecoder, BitmapPixelFormat
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    except Exception:
        return []
    try:
        init_apartment(ApartmentType.MTA)
    except Exception:
        pass
    engine = OcrEngine.try_create_from_language(Language("ar-SA")) or OcrEngine.try_create_from_language(
        Language("en-US")
    )
    if engine is None:
        return []
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(png)
    writer.store_async().get()
    writer.detach_stream()
    stream.seek(0)
    decoder = BitmapDecoder.create_async(stream).get()
    bitmap = decoder.get_software_bitmap_async(BitmapPixelFormat.BGRA8, BitmapAlphaMode.PREMULTIPLIED).get()
    result = engine.recognize_async(bitmap).get()
    words: list[WordBox] = []
    for line in result.lines:
        for item in line.words:
            text = (getattr(item, "text", "") or "").strip()
            if not text:
                continue
            box = item.bounding_rect
            words.append(
                WordBox(
                    text,
                    float(box.x) * scale_x,
                    float(box.y) * scale_y,
                    (float(box.x) + float(box.width)) * scale_x,
                    (float(box.y) + float(box.height)) * scale_y,
                    "docx-ocr",
                )
            )
    return words


def _tesseract_ocr(png: bytes, scale_x: float, scale_y: float) -> list[WordBox]:
    import subprocess
    import tempfile

    try:
        from engine.tesseract_bin import find_exe, ara_tessdata_dir
    except Exception:
        return []
    exe = find_exe()
    if not exe:
        return []
    folder = None
    try:
        folder = ara_tessdata_dir()
    except Exception:
        folder = None
    picture = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    try:
        picture.write(png)
        picture.close()
        cmd = [exe, picture.name, "stdout", "-l", "ara+eng", "--psm", "6", "tsv"]
        if folder is not None:
            cmd[4:4] = ["--tessdata-dir", str(folder)]
        result = subprocess.run(cmd, capture_output=True, timeout=120, check=False)
    except Exception:
        return []
    finally:
        Path(picture.name).unlink(missing_ok=True)
    words: list[WordBox] = []
    for line in result.stdout.decode("utf-8", errors="replace").splitlines()[1:]:
        parts = line.split("\t")
        if len(parts) < 12 or parts[0] != "5":
            continue
        raw = parts[11].strip()
        if not raw:
            continue
        try:
            left, top, width, height = map(float, parts[6:10])
        except ValueError:
            continue
        words.append(
            WordBox(
                raw,
                left * scale_x,
                top * scale_y,
                (left + width) * scale_x,
                (top + height) * scale_y,
                "docx-ocr",
            )
        )
    return words


def _match(pdf_word: WordBox, candidates: list[WordBox], used: set[int]) -> WordBox | None:
    key = _norm(pdf_word.text)
    if not key:
        return None
    best = None
    best_dist = 1e9
    best_i = -1
    for i, cand in enumerate(candidates):
        if i in used:
            continue
        if _norm(cand.text) != key and _norm(cand.text) not in key and key not in _norm(cand.text):
            # Allow short suffix/prefix matches for OCR noise.
            if len(key) >= 3 and (_norm(cand.text).endswith(key) or key.endswith(_norm(cand.text))):
                pass
            else:
                continue
        dist = abs(cand.cx - pdf_word.cx) + abs(cand.cy - pdf_word.cy)
        if dist < best_dist:
            best_dist = dist
            best = cand
            best_i = i
    if best is None or best_dist > 120:
        return None
    used.add(best_i)
    return best


def analyze_page(
    pdf_page: fitz.Page,
    docx_words: list[WordBox],
    overlay_bg: Path | None = None,
    out_overlay: Path | None = None,
) -> PagePlacement:
    source = pdf_words(pdf_page)
    cells = cell_rects(pdf_page)
    dups = pdf_duplicate_pairs(source)
    used: set[int] = set()
    deltas: list[WordDelta] = []
    shifts: list[float] = []
    for word in source:
        matched = _match(word, docx_words, used)
        if matched is None:
            delta = WordDelta(word.text, word, None, 999.0, 999.0, 999.0, ["unmatched"], 999.0)
        else:
            dx = matched.cx - word.cx
            dy = matched.cy - word.cy
            dw = matched.width - word.width
            flags = []
            cell = containing_cell(word, cells)
            if cell is not None:
                if matched.cx < cell[0] - 1 or matched.cx > cell[2] + 1 or matched.cy < cell[1] - 1 or matched.cy > cell[3] + 1:
                    flags.append("outside_cell")
                if matched.x0 < cell[0] - 2 or matched.x1 > cell[2] + 2:
                    flags.append("extends_past_cell")
            page = pdf_page.rect
            if matched.x0 < -1 or matched.y0 < -1 or matched.x1 > page.width + 1 or matched.y1 > page.height + 1:
                flags.append("outside_page")
            if abs(dy) > max(word.height, 4.0) * 0.85:
                flags.append("outside_line")
            if abs(dx) > 2 or abs(dy) > 2:
                flags.append("displaced")
            shift = (dx * dx + dy * dy) ** 0.5
            shifts.append(shift)
            delta = WordDelta(word.text, word, matched, dx, dy, dw, flags, shift)
        deltas.append(delta)

    matched_n = sum(1 for d in deltas if d.docx is not None)
    shifts_sorted = sorted(shifts)
    mean_shift = sum(shifts) / len(shifts) if shifts else 0.0
    median_shift = shifts_sorted[len(shifts_sorted) // 2] if shifts_sorted else 0.0
    within1 = 100.0 * sum(1 for s in shifts if s <= 1.0) / max(len(shifts), 1)
    within2 = 100.0 * sum(1 for s in shifts if s <= 2.0) / max(len(shifts), 1)
    displaced = sum(1 for d in deltas if "displaced" in d.flags or "outside_cell" in d.flags or "outside_page" in d.flags)
    worst = sorted(
        [d for d in deltas if d.docx is not None],
        key=lambda d: (-(1 if d.flags else 0), -d.score),
    )[:20]

    overlay = ""
    if out_overlay is not None:
        overlay = str(_draw_overlay(pdf_page, source, deltas, cells, dups, overlay_bg, out_overlay))

    return PagePlacement(
        number=int(pdf_page.number) + 1,
        words=len(source),
        matched=matched_n,
        mean_shift=mean_shift,
        median_shift=median_shift,
        pct_within_1pt=within1,
        pct_within_2pt=within2,
        displaced=displaced,
        duplicated_pdf=len(dups),
        worst=worst,
        overlay=overlay,
        cell_count=len(cells),
    )


def _draw_overlay(
    page: fitz.Page,
    source: list[WordBox],
    deltas: list[WordDelta],
    cells: list[tuple],
    dups: list[tuple[WordBox, WordBox]],
    bg: Path | None,
    out: Path,
) -> Path:
    from PIL import Image, ImageDraw

    zoom = SCALE
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    base = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if bg is not None and bg.is_file():
        try:
            rendered = Image.open(bg).convert("RGB").resize(base.size)
            base = Image.blend(base, rendered, 0.45)
        except Exception:
            pass
    draw = ImageDraw.Draw(base)
    for cell in cells:
        draw.rectangle(
            [cell[0] * zoom, cell[1] * zoom, cell[2] * zoom, cell[3] * zoom],
            outline=(0, 160, 0),
            width=2,
        )
    flagged = {id(d.pdf) for d in deltas if d.flags}
    for word in source:
        color = (220, 40, 40) if id(word) in flagged else (40, 90, 220)
        draw.rectangle(
            [word.x0 * zoom, word.y0 * zoom, word.x1 * zoom, word.y1 * zoom],
            outline=color,
            width=2 if id(word) in flagged else 1,
        )
    for a, b in dups:
        draw.line(
            [a.cx * zoom, a.cy * zoom, b.cx * zoom, b.cy * zoom],
            fill=(255, 140, 0),
            width=2,
        )
    out.parent.mkdir(parents=True, exist_ok=True)
    base.save(out)
    return out


def run(
    pdf_path: Path,
    docx_path: Path,
    out_dir: Path,
    pages: list[int] | None = None,
    use_engine_fallback: bool = True,
) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    document = fitz.open(pdf_path)
    selected = pages or list(range(1, document.page_count + 1))
    # Render docx → pdf → page images when possible.
    rendered_pdf = out_dir / "docx_print.pdf"
    renderer = "engine-fallback"
    docx_page_words: dict[int, list[WordBox]] = {}
    try:
        render_docx(docx_path, rendered_pdf)
        printed = fitz.open(rendered_pdf)
        renderer = "docx-render"
        for number in selected:
            if number < 1 or number > printed.page_count or number > document.page_count:
                continue
            page = printed[number - 1]
            png = out_dir / f"page-{number:03d}-docx.png"
            pix = page.get_pixmap(matrix=fitz.Matrix(SCALE, SCALE), alpha=False)
            pix.save(str(png))
            src = document[number - 1]
            docx_page_words[number] = ocr_words_from_image(png, float(src.rect.width), float(src.rect.height))
        printed.close()
    except Exception as exc:
        (out_dir / "render_error.txt").write_text(str(exc), encoding="utf-8")
        if not use_engine_fallback:
            document.close()
            raise
        for number in selected:
            if number < 1 or number > document.page_count:
                continue
            docx_page_words[number] = engine_words_from_faithful(document[number - 1])

    reports = []
    for number in selected:
        if number < 1 or number > document.page_count:
            continue
        page = document[number - 1]
        bg = out_dir / f"page-{number:03d}-docx.png"
        overlay = out_dir / f"page-{number:03d}-overlay.png"
        report = analyze_page(page, docx_page_words.get(number, []), bg if bg.is_file() else None, overlay)
        reports.append(report)
        # Human-readable worst-20 dump
        lines = [
            f"page {report.number} words={report.words} matched={report.matched} "
            f"mean_shift={report.mean_shift:.2f}pt median={report.median_shift:.2f}pt "
            f"within_1pt={report.pct_within_1pt:.1f}% within_2pt={report.pct_within_2pt:.1f}% "
            f"displaced={report.displaced} pdf_dup_pairs={report.duplicated_pdf} cells={report.cell_count}",
            "worst 20:",
        ]
        for delta in report.worst:
            lines.append(
                f"  {delta.text!r} dx={delta.dx:.2f} dy={delta.dy:.2f} dw={delta.dwidth:.2f} "
                f"shift={delta.score:.2f} flags={delta.flags} "
                f"pdf=({delta.pdf.x0:.1f},{delta.pdf.y0:.1f})-({delta.pdf.x1:.1f},{delta.pdf.y1:.1f}) "
                f"docx="
                + (
                    f"({delta.docx.x0:.1f},{delta.docx.y0:.1f})-({delta.docx.x1:.1f},{delta.docx.y1:.1f})"
                    if delta.docx
                    else "None"
                )
            )
        (out_dir / f"page-{number:03d}-report.txt").write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        print("overlay", report.overlay)

    document.close()
    summary = {
        "renderer": renderer,
        "pages": [
            {
                **{k: v for k, v in asdict(page).items() if k != "worst"},
                "worst": [
                    {
                        "text": w.text,
                        "dx": w.dx,
                        "dy": w.dy,
                        "dwidth": w.dwidth,
                        "shift": w.score,
                        "flags": w.flags,
                    }
                    for w in page.worst
                ],
            }
            for page in reports
        ],
    }
    (out_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Debug PDF→Word word placement")
    parser.add_argument("pdf")
    parser.add_argument("docx")
    parser.add_argument("--page", type=int, nargs="*", default=None)
    parser.add_argument("--out", type=Path, default=Path("placement_report"))
    args = parser.parse_args(argv)
    run(Path(args.pdf), Path(args.docx), args.out, args.page)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
