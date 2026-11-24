# -*- coding: utf-8 -*-
"""
Convert every corpus sample and score the Word pages against the PDF.

    python -m tests.run_corpus
    python -m tests.run_corpus --only hamza_resume_ppt

Exit code 1 when any page falls below the similarity threshold in the manifest.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx, read_word_build
from core.docx_validate import validate_docx
from core.page_strategy import classify_document
from core import table_rebuild
from core.visual_check import RenderError, compare_pdf_with_docx

CORPUS = Path(__file__).resolve().parent / "corpus"
MANIFEST = CORPUS / "manifest.json"


def _load_manifest() -> dict:
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def _table_check(pdf: Path, document, mode: str) -> str:
    """Compare rebuilt table shape/text with PDF regions; return a short note."""
    parts = []
    for index in range(document.page_count):
        page = document.load_page(index)
        regions = table_rebuild.classify_regions(page)
        rebuilt = table_rebuild.rebuild_tables(page, regions)
        table_regions = [r for r in regions if r.kind in {"vector_table", "raster_table"}]
        for region, table in zip(table_regions, rebuilt):
            rows = table.rows or []
            cols = max((len(row) for row in rows), default=0)
            header = ""
            if rows:
                header = " | ".join((rows[0][i] if i < len(rows[0]) else "") for i in range(min(cols, 3)))
            parts.append(
                f"p{index + 1}:{region.kind} {len(rows)}x{cols} header={header[:40]!r} "
                f"low_conf={len(table.low_confidence_cells)}"
            )
        if table_regions and not rebuilt:
            parts.append(f"p{index + 1}: classified table but rebuild empty")
    try:
        build = read_word_build(pdf, options=WordOptions(mode=mode, fit_a4=False))
        carried = sum(len(getattr(p.faithful, "tables", None) or []) for p in build.pages if p.faithful)
        parts.append(f"word_tables={carried}")
    except Exception as exc:
        parts.append(f"word_tables_err={exc}")
    return "; ".join(parts) if parts else "no tables detected"


def _html_report(rows: list[dict], out: Path) -> None:
    lines = [
        "<!doctype html><html><head><meta charset='utf-8'><title>Corpus report</title>",
        "<style>body{font-family:Segoe UI,Arial,sans-serif;margin:1.5rem}table{border-collapse:collapse;width:100%}",
        "th,td{border:1px solid #ccc;padding:0.4rem 0.6rem;text-align:start;vertical-align:top}",
        "img{max-width:220px;border:1px solid #ddd}.fail{background:#ffe8e8}.pass{background:#e8ffe8}</style></head><body>",
        "<h1>PDF → Word corpus</h1><table><tr><th>Sample</th><th>Page</th><th>Score</th><th>Size</th><th>Strategy</th><th>Diff</th><th>Notes</th></tr>",
    ]
    for row in rows:
        klass = "pass" if row.get("pass") else "fail"
        diff = row.get("diff_image") or ""
        img = f"<img src='{diff}' alt='diff'>" if diff else ""
        lines.append(
            f"<tr class='{klass}'><td>{row['id']}</td><td>{row['page']}</td>"
            f"<td>{row['similarity']:.3f}</td><td>{row['size_match']}</td>"
            f"<td>{row.get('strategy','')}</td><td>{img}</td><td>{row.get('notes','')}</td></tr>"
        )
    lines.append("</table></body></html>")
    out.write_text("\n".join(lines), encoding="utf-8")


def run(only: set[str] | None = None, out_dir: Path | None = None) -> int:
    manifest = _load_manifest()
    defaults = manifest.get("defaults") or {}
    out_dir = out_dir or (CORPUS / "report")
    out_dir.mkdir(parents=True, exist_ok=True)
    rows: list[dict] = []
    failed = 0
    for sample in manifest.get("samples") or []:
        sample_id = str(sample["id"])
        if only and sample_id not in only:
            continue
        pdf = CORPUS / sample["file"]
        if not pdf.is_file():
            rows.append(
                {
                    "id": sample_id,
                    "page": 0,
                    "similarity": 0.0,
                    "size_match": False,
                    "pass": False,
                    "notes": f"missing file {pdf}",
                    "strategy": "",
                    "diff_image": "",
                    "seconds": 0.0,
                }
            )
            failed += 1
            continue
        mode = str(sample.get("mode") or defaults.get("mode") or "faithful")
        threshold = float(sample.get("min_similarity", defaults.get("min_similarity", 0.55)))
        docx_path = out_dir / f"{sample_id}.docx"
        started = time.perf_counter()
        try:
            counts = convert_arabic_pdf_to_docx(pdf, docx_path, options=WordOptions(mode=mode, fit_a4=False))
            import fitz

            document = fitz.open(pdf)
            plans = {plan.number: plan for plan in classify_document(document, default_mode=mode)}
            table_notes = ""
            if sample.get("check_tables"):
                table_notes = _table_check(pdf, document, mode)
            document.close()
            validation = counts.get("docx_validation") or validate_docx(docx_path)
            if not validation.get("ok"):
                failed += 1
                rows.append(
                    {
                        "id": sample_id,
                        "page": 0,
                        "similarity": 0.0,
                        "size_match": False,
                        "pass": False,
                        "notes": f"docx invalid: {validation.get('errors')}",
                        "strategy": "",
                        "diff_image": "",
                        "seconds": 0.0,
                    }
                )
            report = compare_pdf_with_docx(pdf, docx_path, out_dir / sample_id)
            elapsed = time.perf_counter() - started
            for page in report.pages:
                ok = page.similarity >= threshold and (page.size_matches or not defaults.get("min_size_match", True))
                if not ok:
                    failed += 1
                plan = plans.get(page.number)
                note = (
                    f"ink={page.ink_overlap:.3f}; words_in_place={page.words_in_place:.3f}; "
                    f"renderer={report.renderer}"
                )
                if table_notes and page.number == 1:
                    note = f"{note}; {table_notes}"
                rows.append(
                    {
                        "id": sample_id,
                        "page": page.number,
                        "similarity": page.similarity,
                        "size_match": page.size_matches,
                        "pass": ok,
                        "notes": note,
                        "strategy": plan.strategy if plan else "",
                        "diff_image": str(Path(page.diff_image).relative_to(out_dir)).replace("\\", "/")
                        if page.diff_image
                        else "",
                        "seconds": round(elapsed, 2),
                    }
                )
            if report.source_pages != report.result_pages:
                failed += 1
                rows.append(
                    {
                        "id": sample_id,
                        "page": 0,
                        "similarity": 0.0,
                        "size_match": False,
                        "pass": False,
                        "notes": f"page count PDF={report.source_pages} DOCX={report.result_pages}",
                        "strategy": "",
                        "diff_image": "",
                        "seconds": round(elapsed, 2),
                    }
                )
        except RenderError as exc:
            failed += 1
            rows.append(
                {
                    "id": sample_id,
                    "page": 0,
                    "similarity": 0.0,
                    "size_match": False,
                    "pass": False,
                    "notes": f"render failed: {exc}",
                    "strategy": "",
                    "diff_image": "",
                    "seconds": round(time.perf_counter() - started, 2),
                }
            )
        except Exception as exc:
            failed += 1
            rows.append(
                {
                    "id": sample_id,
                    "page": 0,
                    "similarity": 0.0,
                    "size_match": False,
                    "pass": False,
                    "notes": f"{type(exc).__name__}: {exc}",
                    "strategy": "",
                    "diff_image": "",
                    "seconds": round(time.perf_counter() - started, 2),
                }
            )
            (out_dir / f"{sample_id}.error.txt").write_text(traceback.format_exc(), encoding="utf-8")

    csv_path = out_dir / "report.csv"
    with csv_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["id", "page", "similarity", "size_match", "pass", "strategy", "seconds", "notes", "diff_image"],
        )
        writer.writeheader()
        writer.writerows(rows)
    _html_report(rows, out_dir / "report.html")
    summary = {
        "failed": failed,
        "rows": len(rows),
        "missing_categories": manifest.get("missing_categories") or [],
        "csv": str(csv_path),
        "html": str(out_dir / "report.html"),
    }
    (out_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 1 if failed else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the PDF→Word corpus and fail on regressions.")
    parser.add_argument("--only", nargs="*", default=None, help="Sample ids to run")
    parser.add_argument("--out", type=Path, default=None, help="Report folder")
    args = parser.parse_args(argv)
    only = set(args.only) if args.only else None
    return run(only=only, out_dir=args.out)


if __name__ == "__main__":
    raise SystemExit(main())
