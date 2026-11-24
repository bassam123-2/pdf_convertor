# -*- coding: utf-8 -*-
"""
Convert every corpus sample in each Word mode and run the quality checks.

    python -m tests.run_quality
    python -m tests.run_quality --only related_party_bilingual --modes faithful

Writes tests/corpus/_quality/report.md and report.json. Exit code 1 when any
conversion has an error-level finding, so the run can gate a change: a fix is
only kept when it lowers the findings across the whole corpus.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.docx_converter import WordOptions, convert_arabic_pdf_to_docx
from core.quality_check import check_word

CORPUS = ROOT / "tests" / "corpus"
OUT = CORPUS / "_quality"
COLUMNS = ("editable_share", "stretched_runs", "max_spacing_pt", "duplicate_boxes",
           "stray_words", "overlapping_words", "text_recall", "text_precision", "pdf_pages", "word_pages")


def run(only: list[str], modes: list[str], renderer: str) -> list[dict]:
    manifest = json.loads((CORPUS / "manifest.json").read_text(encoding="utf-8"))
    OUT.mkdir(parents=True, exist_ok=True)
    rows = []
    for sample in manifest["samples"]:
        if only and sample["id"] not in only:
            continue
        pdf = CORPUS / sample["file"]
        for mode in modes:
            docx = OUT / f"{sample['id']}.{mode}.docx"
            row = {"id": sample["id"], "mode": mode, "language": sample.get("language", "")}
            started = time.time()
            try:
                convert_arabic_pdf_to_docx(pdf, docx, options=WordOptions(mode=mode))
                report = check_word(pdf, docx, renderer=renderer, expect_editable=(mode == "editable")).to_dict()
                row.update(ok=report["ok"], measures=report["measures"], counts=report["counts"],
                           findings=report["findings"])
            except Exception as exc:
                row.update(ok=False, crash=f"{type(exc).__name__}: {exc}", trace=traceback.format_exc()[-1500:])
            row["seconds"] = round(time.time() - started, 1)
            rows.append(row)
            print(f"{'PASS' if row['ok'] else 'FAIL'} {sample['id']} [{mode}] {row.get('counts') or row.get('crash')}", flush=True)
    return rows


def write_reports(rows: list[dict]) -> None:
    (OUT / "report.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = ["# Word quality report", "",
             "| sample | mode | ok | " + " | ".join(COLUMNS) + " |",
             "|---|---|---|" + "---|" * len(COLUMNS)]
    for row in rows:
        measures = row.get("measures", {})
        cells = [str(measures.get(c, "")) for c in COLUMNS]
        status = "✅" if row["ok"] else ("💥" if row.get("crash") else "❌")
        lines.append(f"| {row['id']} | {row['mode']} | {status} | " + " | ".join(cells) + " |")
    totals: dict[str, int] = {}
    for row in rows:
        for check, count in (row.get("counts") or {}).items():
            totals[check] = totals.get(check, 0) + count
    lines += ["", "## Findings by check", ""] + [f"- **{k}**: {v}" for k, v in sorted(totals.items(), key=lambda i: -i[1])]
    lines += ["", "## Details", ""]
    for row in rows:
        lines.append(f"### {row['id']} [{row['mode']}]")
        if row.get("crash"):
            lines.append(f"- crash: `{row['crash']}`")
        for finding in row.get("findings", []):
            where = f"p{finding['page']} " if finding.get("page") else ""
            sample = f" — «{finding['sample']}»" if finding.get("sample") else ""
            lines.append(f"- [{finding['severity']}] **{finding['check']}** {where}{finding['message']}{sample}")
        lines.append("")
    (OUT / "report.md").write_text("\n".join(lines), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument("--only", nargs="*", default=[])
    parser.add_argument("--modes", nargs="*", default=["faithful", "editable"])
    parser.add_argument("--renderer", default="auto", choices=("auto", "word", "libreoffice"))
    args = parser.parse_args(argv)
    rows = run(args.only, args.modes, args.renderer)
    write_reports(rows)
    print(f"\nReport: {OUT / 'report.md'}")
    return 0 if all(row["ok"] for row in rows) else 1


if __name__ == "__main__":
    sys.exit(main())
