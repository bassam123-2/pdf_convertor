# -*- coding: utf-8 -*-
"""
Similarity engine and preview support for PDF → Word.

Public API (UI calls these later; this round is engine-only):

    render_pdf_page(pdf, page, scale=...) -> RenderedPage
    render_docx_page(docx, page, scale=...) -> RenderedPage
    compare_page(pdf_page_image, docx_page_image, ..., pdf_page, docx_print_page) -> PageCompare
    reconvert_page(pdf, docx, page, mode=None, target=97, project=...) -> ReconvertResult
    run_report(pdf, docx, out_dir, ...) -> Report  (CLI exit code via main)

Faithful mode stays the default. fit_a4 stays off unless asked.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from pathlib import Path

import fitz
import numpy

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.project_state import (
    ProjectState,
    load_project,
    remember_before_reconvert,
    save_project,
    undo_page_choice,
)
from core.visual_check import RenderError, available_renderers, render_docx

_THRESHOLDS_PATH = Path(__file__).with_name("similarity_thresholds.json")
_CACHE: dict[str, Path] = {}

# Preview is low-res by default; detail scale is for explicit requests.
PREVIEW_SCALE = 1.25
DETAIL_SCALE = 2.0


@dataclass
class RenderedPage:
    image_path: str
    width_px: int
    height_px: int
    width_pt: float
    height_pt: float
    scale: float
    renderer: str
    warnings: list[str] = field(default_factory=list)


@dataclass
class PlacementMetrics:
    words: int = 0
    matched: int = 0
    mean_shift_pt: float = 0.0
    median_shift_pt: float = 0.0
    pct_within_1pt: float = 0.0
    pct_within_2pt: float = 0.0
    score: float = 0.0  # 0–100


@dataclass
class ContentMetrics:
    missing: int = 0
    extra: int = 0
    merged_or_split: int = 0
    garbled_arabic: int = 0
    score: float = 0.0  # 0–100
    samples: list[str] = field(default_factory=list)


@dataclass
class VisualMetrics:
    ssim: float = 0.0  # 0–1
    ink_overlap: float = 0.0
    score: float = 0.0  # 0–100
    diff_image: str = ""


@dataclass
class PageCompare:
    number: int
    score: float  # 0–100 composite
    visual: VisualMetrics
    placement: PlacementMetrics
    content: ContentMetrics
    size_matches: bool = True
    document_type: str = "text_document"
    threshold: float = 97.0
    pass_: bool = False
    renderer: str = ""
    warnings: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        data = asdict(self)
        data["pass"] = data.pop("pass_")
        return data


@dataclass
class ReconvertResult:
    page: int
    mode: str
    score: float
    target: float
    reached: bool
    needs_review: bool
    tried: list[dict] = field(default_factory=list)
    compare: PageCompare | None = None
    docx_path: str = ""
    notes: list[str] = field(default_factory=list)


@dataclass
class SimilarityReport:
    pdf: str
    docx: str
    renderer: str
    document_type: str
    pages: list[PageCompare] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def mean_score(self) -> float:
        if not self.pages:
            return 0.0
        return sum(page.score for page in self.pages) / len(self.pages)

    def failed_pages(self, project: ProjectState | None = None) -> list[PageCompare]:
        failed = []
        for page in self.pages:
            if page.pass_:
                continue
            if project is not None and project.page(page.number).accepted:
                continue
            failed.append(page)
        return failed


def load_thresholds() -> dict:
    if _THRESHOLDS_PATH.is_file():
        return json.loads(_THRESHOLDS_PATH.read_text(encoding="utf-8"))
    return {"defaults": {"target_score": 97.0, "fail_below": 85.0, "weights": {"visual": 0.45, "placement": 0.35, "content": 0.20}}, "document_types": {}}


def thresholds_for(document_type: str) -> dict:
    cfg = load_thresholds()
    base = dict(cfg.get("defaults") or {})
    weights = dict(base.get("weights") or {})
    typed = (cfg.get("document_types") or {}).get(document_type) or {}
    merged = {**base, **typed}
    if "weights" in typed:
        weights = {**weights, **typed["weights"]}
    merged["weights"] = weights
    return merged


def infer_document_type(pdf_path: str | Path) -> str:
    """Cheap type guess from the first pages (general, not file-name based)."""
    from core.page_strategy import classify_document

    document = fitz.open(str(pdf_path))
    try:
        plans = classify_document(document)
    finally:
        document.close()
    if not plans:
        return "text_document"
    kinds = [plan.strategy for plan in plans]
    if sum(1 for k in kinds if k == "scanned") >= max(1, len(kinds) // 2):
        return "scanned"
    if sum(1 for k in kinds if k == "table") >= max(1, len(kinds) // 3):
        return "table_heavy"
    landscape = sum(1 for plan in plans if plan.width > plan.height * 1.15)
    if landscape >= max(1, len(plans) // 2) and sum(plan.text_chars for plan in plans) / max(len(plans), 1) < 800:
        return "slides"
    return "text_document"


def _cache_key(*parts: object) -> str:
    blob = "|".join(str(part) for part in parts)
    return hashlib.sha1(blob.encode("utf-8", errors="replace")).hexdigest()[:24]


def _cache_dir() -> Path:
    root = Path(tempfile.gettempdir()) / "pdf_tool_similarity_cache"
    root.mkdir(parents=True, exist_ok=True)
    return root


def render_pdf_page(
    pdf_path: str | Path,
    page: int,
    scale: float = PREVIEW_SCALE,
    out_path: str | Path | None = None,
) -> RenderedPage:
    """Rasterize one PDF page. `page` is 1-based."""
    from PIL import Image

    pdf = Path(pdf_path)
    document = fitz.open(str(pdf))
    try:
        if page < 1 or page > document.page_count:
            raise ValueError(f"Page {page} out of range (1..{document.page_count})")
        leaf = document.load_page(page - 1)
        key = _cache_key("pdf", pdf.resolve(), pdf.stat().st_mtime_ns, page, round(scale, 4))
        target = Path(out_path) if out_path else _cache_dir() / f"{key}.png"
        if not target.is_file():
            pix = leaf.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            target.parent.mkdir(parents=True, exist_ok=True)
            pix.save(str(target))
        image = Image.open(target)
        return RenderedPage(
            image_path=str(target),
            width_px=image.width,
            height_px=image.height,
            width_pt=float(leaf.rect.width),
            height_pt=float(leaf.rect.height),
            scale=scale,
            renderer="pymupdf",
            warnings=[],
        )
    finally:
        document.close()


def _docx_print_pdf(docx_path: Path, renderer: str = "auto") -> tuple[Path, str, list[str]]:
    """Print the whole docx once; cache by mtime. Returns (pdf, renderer, warnings)."""
    docx = Path(docx_path).resolve()
    warnings: list[str] = []
    key = _cache_key("docx_print", docx, docx.stat().st_mtime_ns, renderer)
    target = _cache_dir() / f"{key}.pdf"
    used = renderer
    if not target.is_file():
        used = render_docx(docx, target, renderer=renderer)
    else:
        # Infer which renderer would be preferred without re-printing.
        choices = available_renderers()
        used = choices[0] if renderer == "auto" and choices else (renderer if renderer != "auto" else "cached")
    if used == "libreoffice":
        warnings.append(
            "Rendered with LibreOffice. Layout may differ from Microsoft Word; "
            "fonts missing on this machine also affect the result."
        )
    else:
        warnings.append(
            "Fonts missing on this machine are substituted and can lower the similarity score."
        )
    return target, used, warnings


def render_docx_page(
    docx_path: str | Path,
    page: int,
    scale: float = PREVIEW_SCALE,
    renderer: str = "auto",
    out_path: str | Path | None = None,
    match_size: tuple[float, float] | None = None,
) -> RenderedPage:
    """
    Render one Word page to an image via Word COM when available, else LibreOffice.

    Reports which renderer was used. Missing fonts on the machine affect results.
    """
    from PIL import Image

    printed, used, warnings = _docx_print_pdf(Path(docx_path), renderer=renderer)
    document = fitz.open(str(printed))
    try:
        if page < 1 or page > document.page_count:
            raise ValueError(f"Word page {page} out of range (1..{document.page_count})")
        leaf = document.load_page(page - 1)
        if match_size is not None:
            zoom_x = (match_size[0] * scale) / max(leaf.rect.width, 1.0)
            zoom_y = (match_size[1] * scale) / max(leaf.rect.height, 1.0)
            matrix = fitz.Matrix(zoom_x, zoom_y)
        else:
            matrix = fitz.Matrix(scale, scale)
        key = _cache_key("docx_page", printed, printed.stat().st_mtime_ns, page, round(scale, 4), match_size)
        target = Path(out_path) if out_path else _cache_dir() / f"{key}.png"
        if not target.is_file():
            pix = leaf.get_pixmap(matrix=matrix, alpha=False)
            target.parent.mkdir(parents=True, exist_ok=True)
            pix.save(str(target))
        image = Image.open(target)
        return RenderedPage(
            image_path=str(target),
            width_px=image.width,
            height_px=image.height,
            width_pt=float(leaf.rect.width),
            height_pt=float(leaf.rect.height),
            scale=scale,
            renderer=used,
            warnings=warnings,
        )
    finally:
        document.close()


def _ssim(left: "numpy.ndarray", right: "numpy.ndarray") -> float:
    """Structural similarity on grayscale float arrays in [0, 1]."""
    if left.shape != right.shape:
        raise ValueError("SSIM input shapes differ")
    c1 = (0.01) ** 2
    c2 = (0.03) ** 2
    mu_x = left.mean()
    mu_y = right.mean()
    sigma_x = left.var()
    sigma_y = right.var()
    sigma_xy = ((left - mu_x) * (right - mu_y)).mean()
    num = (2 * mu_x * mu_y + c1) * (2 * sigma_xy + c2)
    den = (mu_x**2 + mu_y**2 + c1) * (sigma_x + sigma_y + c2)
    if den <= 0:
        return 1.0
    return float(max(0.0, min(1.0, num / den)))


def _visual_compare(source_path: Path, result_path: Path, diff_path: Path | None) -> VisualMetrics:
    from PIL import Image, ImageFilter

    source = Image.open(source_path).convert("RGB")
    result = Image.open(result_path).convert("RGB")
    if result.size != source.size:
        result = result.resize(source.size, Image.LANCZOS)
    left = numpy.asarray(source.filter(ImageFilter.GaussianBlur(0.8)), dtype=numpy.float32) / 255.0
    right = numpy.asarray(result.filter(ImageFilter.GaussianBlur(0.8)), dtype=numpy.float32) / 255.0
    gray_l = left.mean(axis=2)
    gray_r = right.mean(axis=2)
    ssim = _ssim(gray_l, gray_r)
    delta = numpy.abs(left - right).max(axis=2)
    # Ink overlap vs paper.
    def ink(rgb: numpy.ndarray) -> numpy.ndarray:
        flat = (rgb.reshape(-1, 3)[:: max(1, rgb.shape[0] * rgb.shape[1] // 20000)] * 255).astype(numpy.int16)
        values, counts = numpy.unique(flat, axis=0, return_counts=True)
        paper = values[counts.argmax()] / 255.0
        return numpy.abs(rgb - paper).max(axis=2) > (40 / 255.0)

    ink_l = ink(left)
    ink_r = ink(right)
    union = int((ink_l | ink_r).sum())
    overlap = float((ink_l & ink_r).sum()) / union if union else 1.0

    diff_file = ""
    if diff_path is not None:
        faded = (numpy.asarray(source.convert("L"), dtype=numpy.float32) * 0.35 + 255 * 0.65).astype(numpy.uint8)
        picture = numpy.stack([faded, faded, faded], axis=2)
        # Heat: yellow → red by delta magnitude.
        hot = delta > (32 / 255.0)
        heat = numpy.clip(delta * 255.0, 0, 255).astype(numpy.uint8)
        picture[hot, 0] = 255
        picture[hot, 1] = numpy.clip(255 - heat[hot], 0, 180)
        picture[hot, 2] = 30
        diff_path.parent.mkdir(parents=True, exist_ok=True)
        Image.fromarray(picture, "RGB").save(diff_path)
        diff_file = str(diff_path)

    score = 100.0 * (0.75 * ssim + 0.25 * overlap)
    return VisualMetrics(ssim=round(ssim, 4), ink_overlap=round(overlap, 4), score=round(score, 2), diff_image=diff_file)


def _norm_word(text: str) -> str:
    text = re.sub(r"\s+", "", (text or "").strip())
    return text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))


def _pdf_words(page: fitz.Page) -> list[tuple[str, float, float, float, float]]:
    words = []
    for item in page.get_text("words") or []:
        text = str(item[4]).strip()
        if text:
            words.append((text, float(item[0]), float(item[1]), float(item[2]), float(item[3])))
    return words


def _placement_metrics(source: fitz.Page, printed: fitz.Page) -> PlacementMetrics:
    scale_x = source.rect.width / max(printed.rect.width, 1.0)
    scale_y = source.rect.height / max(printed.rect.height, 1.0)
    src = _pdf_words(source)
    dst = _pdf_words(printed)
    buckets: dict[str, list[tuple[float, float]]] = {}
    for text, x0, y0, x1, y1 in dst:
        key = _norm_word(text)
        if not key:
            continue
        buckets.setdefault(key, []).append(((x0 + x1) / 2.0 * scale_x, (y0 + y1) / 2.0 * scale_y))
    shifts: list[float] = []
    used: dict[str, set[int]] = {}
    for text, x0, y0, x1, y1 in src:
        key = _norm_word(text)
        places = buckets.get(key) or []
        sx, sy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        best = None
        best_i = -1
        best_d = 1e9
        seen = used.setdefault(key, set())
        for i, (dx, dy) in enumerate(places):
            if i in seen:
                continue
            dist = ((sx - dx) ** 2 + (sy - dy) ** 2) ** 0.5
            if dist < best_d:
                best_d = dist
                best = dist
                best_i = i
        if best is None:
            continue
        seen.add(best_i)
        shifts.append(best)
    if not shifts:
        return PlacementMetrics(words=len(src), matched=0, score=0.0)
    ordered = sorted(shifts)
    mean = sum(shifts) / len(shifts)
    median = ordered[len(ordered) // 2]
    within1 = 100.0 * sum(1 for s in shifts if s <= 1.0) / len(shifts)
    within2 = 100.0 * sum(1 for s in shifts if s <= 2.0) / len(shifts)
    # Score: full marks when median ≤ 1pt and most words within 2pt.
    score = max(0.0, min(100.0, 0.5 * within2 + 0.3 * within1 + 0.2 * max(0.0, 100.0 - median * 15.0)))
    return PlacementMetrics(
        words=len(src),
        matched=len(shifts),
        mean_shift_pt=round(mean, 3),
        median_shift_pt=round(median, 3),
        pct_within_1pt=round(within1, 2),
        pct_within_2pt=round(within2, 2),
        score=round(score, 2),
    )


def _is_arabic(char: str) -> bool:
    return "\u0600" <= char <= "\u06FF" or "\u0750" <= char <= "\u077F"


def _garbled_arabic(text: str) -> bool:
    """Latin / private-use / replacement letters inside an otherwise Arabic token."""
    arabic = sum(1 for char in text if _is_arabic(char))
    if arabic < 2:
        return False
    bad = sum(
        1
        for char in text
        if (char.isascii() and char.isalpha())
        or ord(char) == 0xFFFD
        or 0xE000 <= ord(char) <= 0xF8FF
    )
    return bad > 0


def _content_metrics(source: fitz.Page, printed: fitz.Page) -> ContentMetrics:
    src = [_norm_word(t) for t, *_ in _pdf_words(source) if _norm_word(t)]
    dst = [_norm_word(t) for t, *_ in _pdf_words(printed) if _norm_word(t)]
    from collections import Counter

    left = Counter(src)
    right = Counter(dst)
    missing = sum((left - right).values())
    extra = sum((right - left).values())
    # Merged/split heuristic: tokens that appear only as substrings of the other side.
    src_set = set(src)
    dst_set = set(dst)
    merged = 0
    samples: list[str] = []
    for token in list(src_set - dst_set)[:200]:
        if len(token) < 3:
            continue
        if any(token != other and token in other for other in dst_set):
            merged += 1
            if len(samples) < 8:
                samples.append(f"merged/split:{token}")
    garbled = 0
    for text, *_ in _pdf_words(printed):
        if _garbled_arabic(text):
            garbled += 1
            if len(samples) < 12:
                samples.append(f"garbled:{text[:40]}")
    total = max(len(src), 1)
    penalty = (missing + extra) * 100.0 / total + merged * 2.0 + garbled * 3.0
    score = max(0.0, min(100.0, 100.0 - penalty))
    return ContentMetrics(
        missing=int(missing),
        extra=int(extra),
        merged_or_split=int(merged),
        garbled_arabic=int(garbled),
        score=round(score, 2),
        samples=samples,
    )


def compare_page(
    pdf_path: str | Path,
    docx_path: str | Path,
    page: int,
    *,
    scale: float = PREVIEW_SCALE,
    document_type: str | None = None,
    out_dir: str | Path | None = None,
    renderer: str = "auto",
    printed_pdf: str | Path | None = None,
) -> PageCompare:
    """
    Compare one PDF page with the same page printed from the Word file.

    Returns a 0–100 score plus visual / placement / content breakdown and a
    diff heatmap path when `out_dir` is set.
    """
    pdf = Path(pdf_path)
    doc_type = document_type or infer_document_type(pdf)
    thr = thresholds_for(doc_type)
    weights = thr.get("weights") or {"visual": 0.45, "placement": 0.35, "content": 0.20}
    target = float(thr.get("target_score") or 97.0)

    folder = Path(out_dir) if out_dir else None
    if folder is not None:
        folder.mkdir(parents=True, exist_ok=True)

    pdf_render = render_pdf_page(
        pdf,
        page,
        scale=scale,
        out_path=(folder / f"page-{page:03d}-pdf.png") if folder else None,
    )
    if printed_pdf is None:
        docx_render = render_docx_page(
            docx_path,
            page,
            scale=scale,
            renderer=renderer,
            out_path=(folder / f"page-{page:03d}-docx.png") if folder else None,
            match_size=(pdf_render.width_pt, pdf_render.height_pt),
        )
        printed_path, used, warn = _docx_print_pdf(Path(docx_path), renderer=renderer)
        warnings = list(docx_render.warnings) + warn
        renderer_name = docx_render.renderer or used
    else:
        printed_path = Path(printed_pdf)
        used = renderer if renderer != "auto" else "provided"
        # Render the given print PDF page into the same pixel grid.
        from PIL import Image

        src_doc = fitz.open(str(printed_path))
        try:
            leaf = src_doc.load_page(page - 1)
            zoom_x = (pdf_render.width_pt * scale) / max(leaf.rect.width, 1.0)
            zoom_y = (pdf_render.height_pt * scale) / max(leaf.rect.height, 1.0)
            pix = leaf.get_pixmap(matrix=fitz.Matrix(zoom_x, zoom_y), alpha=False)
            out = folder / f"page-{page:03d}-docx.png" if folder else _cache_dir() / f"cmp-{_cache_key(printed_path, page, scale)}.png"
            out.parent.mkdir(parents=True, exist_ok=True)
            pix.save(str(out))
            docx_render = RenderedPage(
                str(out), pix.width, pix.height, float(leaf.rect.width), float(leaf.rect.height), scale, used, []
            )
        finally:
            src_doc.close()
        warnings = [
            "Fonts missing on this machine are substituted and can lower the similarity score."
        ]
        renderer_name = used

    diff_path = (folder / f"page-{page:03d}-diff.png") if folder else None
    visual = _visual_compare(Path(pdf_render.image_path), Path(docx_render.image_path), diff_path)

    source = fitz.open(str(pdf))
    printed = fitz.open(str(printed_path))
    try:
        left = source.load_page(page - 1)
        right = printed.load_page(min(page, printed.page_count) - 1)
        placement = _placement_metrics(left, right)
        content = _content_metrics(left, right)
        size_matches = abs(left.rect.width - right.rect.width) <= 1.0 and abs(left.rect.height - right.rect.height) <= 1.0
        if not size_matches:
            warnings.append(
                f"Page size differs: PDF {left.rect.width:.1f}×{left.rect.height:.1f} pt vs "
                f"Word print {right.rect.width:.1f}×{right.rect.height:.1f} pt (fit_a4 should be off)."
            )
    finally:
        source.close()
        printed.close()

    score = (
        float(weights.get("visual", 0.45)) * visual.score
        + float(weights.get("placement", 0.35)) * placement.score
        + float(weights.get("content", 0.20)) * content.score
    )
    # Size mismatch is a hard penalty — scaling violates the no-scale rule.
    if not size_matches:
        score *= 0.85
        visual.score = round(visual.score * 0.9, 2)

    return PageCompare(
        number=page,
        score=round(score, 2),
        visual=visual,
        placement=placement,
        content=content,
        size_matches=size_matches,
        document_type=doc_type,
        threshold=target,
        pass_=score >= target,
        renderer=renderer_name,
        warnings=warnings,
    )


def _convert_page_candidate(
    pdf: Path,
    page: int,
    mode: str,
    out_docx: Path,
) -> list[str]:
    """Convert a single page with a candidate strategy. Returns notes."""
    from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx
    from core import faithful

    notes: list[str] = []
    if mode == "picture":
        # Last resort: whole page as one picture (visual match, not editable).
        options = WordOptions(mode="faithful", fit_a4=False, page_spec=str(page))
        # Capture as picture by writing through a tiny helper path.
        document = fitz.open(str(pdf))
        try:
            payload = faithful.picture_page(
                document[page - 1],
                note=f"Page {page}: kept as a picture (last-resort visual match; not editable).",
            )
        finally:
            document.close()
        # Use normal convert then overwrite that page's faithful payload via build.
        from core.docx_converter import read_word_build, write_word_build

        build = read_word_build(pdf, options=options)
        for record in build.pages:
            if record.number == page:
                record.faithful = payload
                record.blocks = []
                record.kind = "picture"
                notes.append("page-as-picture")
        write_word_build(build, out_docx)
        return notes

    mode_map = {
        "faithful": dict(mode="faithful", fit_a4=False),
        "hybrid": dict(mode="faithful", fit_a4=False),  # faithful + rtl repair via rescan
        "editable": dict(mode="editable", fit_a4=False),
        "ocr_spans": dict(mode="faithful", fit_a4=False),
        "alt_tables": dict(mode="editable", fit_a4=False, tables_as_images=False),
        "adjusted_boxes": dict(mode="faithful", fit_a4=False),
    }
    kwargs = mode_map.get(mode) or dict(mode="faithful", fit_a4=False)
    options = WordOptions(page_spec=str(page), **kwargs)
    convert_arabic_pdf_to_docx(pdf, out_docx, options=options)
    if mode in {"hybrid", "ocr_spans"}:
        from core.docx_converter import read_word_build, write_word_build, rescan_word_page

        build = read_word_build(pdf, options=WordOptions(mode="faithful", fit_a4=False, page_spec=str(page)))
        try:
            rescan_word_page(build, page, mode="rtl")
            notes.append("rtl-repair rescan")
        except Exception as exc:
            notes.append(f"rtl-repair skipped: {exc}")
        write_word_build(build, out_docx)
    if mode == "adjusted_boxes":
        notes.append("faithful with column clamp / spacing fit (current engine)")
    if mode == "alt_tables":
        notes.append("editable flowing tables")
    return notes


def reconvert_page(
    pdf_path: str | Path,
    docx_path: str | Path,
    page: int,
    mode: str | None = None,
    target: float | None = None,
    project_path: str | Path | None = None,
    document_type: str | None = None,
    scale: float = PREVIEW_SCALE,
    out_dir: str | Path | None = None,
) -> ReconvertResult:
    """
    Re-convert one page, optionally searching strategies until `target` is met.

    When no strategy reaches the target, marks needs_review and can keep the
    last-resort page-as-picture result (100% visual intent, not editable).
    """
    pdf = Path(pdf_path)
    docx = Path(docx_path)
    doc_type = document_type or infer_document_type(pdf)
    thr = thresholds_for(doc_type)
    goal = float(target if target is not None else thr.get("target_score") or 97.0)

    project: ProjectState | None = None
    if project_path:
        project = load_project(project_path)
        project.pdf = str(pdf)
        project.docx = str(docx)
        project.document_type = doc_type
        choice = project.page(page)
        remember_before_reconvert(choice)

    candidates = [mode] if mode else [
        "faithful",
        "hybrid",
        "ocr_spans",
        "editable",
        "alt_tables",
        "adjusted_boxes",
        "picture",
    ]

    folder = Path(out_dir) if out_dir else Path(tempfile.mkdtemp(prefix="pdf_tool_reconvert_"))
    folder.mkdir(parents=True, exist_ok=True)
    tried: list[dict] = []
    best: ReconvertResult | None = None

    for candidate in candidates:
        trial = folder / f"page-{page:03d}-{candidate}.docx"
        try:
            notes = _convert_page_candidate(pdf, page, candidate, trial)
            compare = compare_page(
                pdf,
                trial,
                page,
                scale=scale,
                document_type=doc_type,
                out_dir=folder / candidate,
            )
            # Picture mode: visual score dominates the reported score.
            score = compare.score
            if candidate == "picture":
                score = max(score, compare.visual.score)
                notes.append("page as picture — visual match, not editable")
            entry = {"mode": candidate, "score": round(score, 2), "notes": notes, "pass": score >= goal}
            tried.append(entry)
            result = ReconvertResult(
                page=page,
                mode=candidate,
                score=round(score, 2),
                target=goal,
                reached=score >= goal,
                needs_review=False,
                tried=list(tried),
                compare=compare,
                docx_path=str(trial),
                notes=notes,
            )
            if best is None or result.score > best.score:
                best = result
            if result.reached:
                break
        except Exception as exc:
            tried.append({"mode": candidate, "score": 0.0, "notes": [str(exc)], "pass": False})

    if best is None:
        best = ReconvertResult(
            page=page,
            mode=mode or "faithful",
            score=0.0,
            target=goal,
            reached=False,
            needs_review=True,
            tried=tried,
            notes=["no candidate produced a document"],
        )

    if not best.reached:
        best.needs_review = True
        if best.mode != "picture":
            best.notes.append(
                f"No strategy reached target {goal:.1f}. Best was {best.mode} at {best.score:.1f}. "
                "Page marked needs review; page-as-picture remains available."
            )

    # Apply winner into the real docx using page_strategies + full rewrite of that page's mode.
    if best.docx_path and Path(best.docx_path).is_file():
        from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx, read_word_build, write_word_build

        strategies = {}
        if project is not None:
            strategies = project.strategies()
        strategies[page] = "faithful" if best.mode in {"hybrid", "ocr_spans", "adjusted_boxes", "picture"} else best.mode
        if best.mode == "editable" or best.mode == "alt_tables":
            # Full-document editable for one page only is awkward; rewrite just this page file into place
            # by merging builds: read full faithful, replace page from trial build.
            full = read_word_build(pdf, options=WordOptions(mode="faithful", fit_a4=False, page_strategies=strategies))
            trial_build = read_word_build(
                pdf,
                options=WordOptions(mode="editable", fit_a4=False, page_spec=str(page)),
            )
            trial_page = next((p for p in trial_build.pages if p.number == page), None)
            for record in full.pages:
                if record.number == page and trial_page is not None:
                    record.blocks = trial_page.blocks
                    record.faithful = None
                    record.kind = trial_page.kind
            write_word_build(full, docx)
        elif best.mode == "picture":
            full = read_word_build(pdf, options=WordOptions(mode="faithful", fit_a4=False))
            from core import faithful as faithful_mod

            document = fitz.open(str(pdf))
            try:
                payload = faithful_mod.picture_page(
                    document[page - 1],
                    note=f"Page {page}: page as picture (not editable).",
                )
            finally:
                document.close()
            for record in full.pages:
                if record.number == page:
                    record.faithful = payload
                    record.blocks = []
                    record.kind = "picture"
            write_word_build(full, docx)
        else:
            convert_arabic_pdf_to_docx(
                pdf,
                docx,
                options=WordOptions(mode="faithful", fit_a4=False, page_strategies=strategies),
            )
            if best.mode in {"hybrid", "ocr_spans"}:
                from core.docx_converter import rescan_word_page

                build = read_word_build(pdf, options=WordOptions(mode="faithful", fit_a4=False))
                try:
                    rescan_word_page(build, page, mode="rtl")
                    write_word_build(build, docx)
                except Exception:
                    pass

    if project is not None and project_path:
        choice = project.page(page)
        choice.mode = best.mode
        choice.score = best.score
        choice.target = goal
        choice.needs_review = best.needs_review
        choice.tried = tried
        choice.notes = list(best.notes)
        if best.reached:
            choice.accepted = False
        save_project(project, project_path)

    best.tried = tried
    return best


def undo_reconvert(
    pdf_path: str | Path,
    docx_path: str | Path,
    page: int,
    project_path: str | Path,
) -> bool:
    """Undo the last reconvert for a page using the project file, then rewrite."""
    from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx

    project = load_project(project_path)
    choice = project.page(page)
    if not undo_page_choice(choice):
        return False
    save_project(project, project_path)
    strategies = project.strategies()
    mode = choice.mode if choice.mode in {"faithful", "editable"} else "faithful"
    convert_arabic_pdf_to_docx(
        pdf_path,
        docx_path,
        options=WordOptions(mode=mode if len(strategies) <= 1 else "faithful", fit_a4=False, page_strategies=strategies),
    )
    return True


def run_report(
    pdf_path: str | Path,
    docx_path: str | Path,
    out_dir: str | Path,
    *,
    pages: list[int] | None = None,
    document_type: str | None = None,
    scale: float = PREVIEW_SCALE,
    renderer: str = "auto",
    project_path: str | Path | None = None,
    workers: int = 2,
) -> SimilarityReport:
    """Score every (or selected) page; write JSON + diff images."""
    pdf = Path(pdf_path)
    docx = Path(docx_path)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    doc_type = document_type or infer_document_type(pdf)
    project = load_project(project_path) if project_path else None

    source = fitz.open(str(pdf))
    try:
        total = source.page_count
    finally:
        source.close()
    selected = pages or list(range(1, total + 1))

    # Print docx once for all pages.
    printed, used, warn = _docx_print_pdf(docx, renderer=renderer)
    report = SimilarityReport(str(pdf), str(docx), used, doc_type, warnings=list(warn))

    def _one(number: int) -> PageCompare:
        return compare_page(
            pdf,
            docx,
            number,
            scale=scale,
            document_type=doc_type,
            out_dir=out,
            renderer=renderer,
            printed_pdf=printed,
        )

    results: dict[int, PageCompare] = {}
    workers = max(1, int(workers))
    if workers == 1 or len(selected) == 1:
        for number in selected:
            results[number] = _one(number)
    else:
        with ThreadPoolExecutor(max_workers=min(workers, len(selected))) as pool:
            futures = {pool.submit(_one, number): number for number in selected}
            for future in as_completed(futures):
                number = futures[future]
                results[number] = future.result()

    for number in selected:
        page = results[number]
        if project is not None:
            choice = project.page(number)
            if choice.score is None:
                choice.score = page.score
                choice.mode = choice.mode or "faithful"
            # Accepted pages always pass the gate.
            if choice.accepted:
                page.pass_ = True
            page.warnings = list(page.warnings) + list(choice.notes)
        report.pages.append(page)

    if project is not None and project_path:
        project.pdf = str(pdf)
        project.docx = str(docx)
        project.document_type = doc_type
        save_project(project, project_path)

    (out / "report.json").write_text(
        json.dumps(
            {
                "pdf": report.pdf,
                "docx": report.docx,
                "renderer": report.renderer,
                "document_type": report.document_type,
                "mean_score": report.mean_score,
                "warnings": report.warnings,
                "pages": [page.as_dict() for page in report.pages],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    lines = [
        f"renderer={report.renderer} type={report.document_type} mean={report.mean_score:.2f}",
        *report.warnings,
    ]
    for page in report.pages:
        lines.append(
            f"page {page.number}: score={page.score:.2f} visual={page.visual.score:.1f} "
            f"place={page.placement.score:.1f} (med {page.placement.median_shift_pt:.2f}pt, "
            f"≤1={page.placement.pct_within_1pt:.0f}% ≤2={page.placement.pct_within_2pt:.0f}%) "
            f"content={page.content.score:.1f} size_ok={page.size_matches} "
            f"{'PASS' if page.pass_ else 'FAIL'} thr={page.threshold}"
        )
        if page.visual.diff_image:
            lines.append(f"  diff {page.visual.diff_image}")
    (out / "report.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PDF↔Word similarity engine")
    sub = parser.add_subparsers(dest="cmd", required=True)

    report_p = sub.add_parser("report", help="Score pages and write diff images")
    report_p.add_argument("pdf")
    report_p.add_argument("docx")
    report_p.add_argument("--out", required=True)
    report_p.add_argument("--page", type=int, nargs="*", default=None)
    report_p.add_argument("--scale", type=float, default=PREVIEW_SCALE)
    report_p.add_argument("--detail", action="store_true", help="Use high-resolution scale")
    report_p.add_argument("--renderer", default="auto", choices=("auto", "word", "libreoffice"))
    report_p.add_argument("--type", dest="document_type", default="")
    report_p.add_argument("--project", default="")
    report_p.add_argument("--workers", type=int, default=2)

    rec_p = sub.add_parser("reconvert", help="Re-convert one page toward a target score")
    rec_p.add_argument("pdf")
    rec_p.add_argument("docx")
    rec_p.add_argument("--page", type=int, required=True)
    rec_p.add_argument("--mode", default="")
    rec_p.add_argument("--target", type=float, default=None)
    rec_p.add_argument("--project", required=True)
    rec_p.add_argument("--out", default="")

    undo_p = sub.add_parser("undo", help="Undo last page reconvert from project state")
    undo_p.add_argument("pdf")
    undo_p.add_argument("docx")
    undo_p.add_argument("--page", type=int, required=True)
    undo_p.add_argument("--project", required=True)

    accept_p = sub.add_parser("accept", help="Mark a page accepted (exempt from fail exit)")
    accept_p.add_argument("--project", required=True)
    accept_p.add_argument("--page", type=int, required=True)

    args = parser.parse_args(argv)

    if args.cmd == "report":
        scale = DETAIL_SCALE if args.detail else args.scale
        report = run_report(
            args.pdf,
            args.docx,
            args.out,
            pages=args.page,
            document_type=args.document_type or None,
            scale=scale,
            renderer=args.renderer,
            project_path=args.project or None,
            workers=args.workers,
        )
        project = load_project(args.project) if args.project else None
        failed = report.failed_pages(project)
        return 1 if failed else 0

    if args.cmd == "reconvert":
        result = reconvert_page(
            args.pdf,
            args.docx,
            args.page,
            mode=args.mode or None,
            target=args.target,
            project_path=args.project,
            out_dir=args.out or None,
        )
        print(
            f"page {result.page} mode={result.mode} score={result.score:.2f} "
            f"target={result.target:.1f} reached={result.reached} review={result.needs_review}"
        )
        for item in result.tried:
            print(f"  tried {item}")
        return 0 if result.reached or result.mode == "picture" else 1

    if args.cmd == "undo":
        ok = undo_reconvert(args.pdf, args.docx, args.page, args.project)
        print("undone" if ok else "nothing to undo")
        return 0 if ok else 1

    if args.cmd == "accept":
        project = load_project(args.project)
        choice = project.page(args.page)
        choice.accepted = True
        choice.needs_review = False
        save_project(project, args.project)
        print(f"page {args.page} accepted")
        return 0

    return 2


if __name__ == "__main__":
    raise SystemExit(main())
