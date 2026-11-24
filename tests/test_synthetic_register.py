"""Golden comparison for the generated Arabic and English registers."""

from __future__ import annotations

import csv
import os
from pathlib import Path

import pytest

from core.excel_converter import read_pdf_workbook, _vehicle_field_at
from excel_review import classify_cells, cluster_plan, cluster_review_items, read_print_date, review_crops
from tests.synthetic_register import FIELDS, write_synthetic


def _text(value) -> str:
    return str(value or "").replace("\u200e", "").replace("\u200f", "").strip()


def score_result(result: dict, truth_path: Path) -> tuple[dict, list[tuple]]:
    columns = list(result.get("columns") or [])
    fields = _vehicle_field_at(columns)
    actual: dict[str, dict[str, str]] = {}
    plate_at = fields.get("plate")
    for row, page in zip(result.get("rows") or [], result.get("page_numbers") or []):
        plate = _text(row[plate_at]) if plate_at is not None and plate_at < len(row) else ""
        if not plate:
            continue
        actual[plate] = {
            name: _text(row[index]) if (index := fields.get(name if name != "insurance_expiry" else "date")) is not None and index < len(row) else ""
            for name in FIELDS
        }
    with truth_path.open(encoding="utf-8", newline="") as handle:
        expected_rows = list(csv.DictReader(handle))
    counts = {key: {"correct": 0, "empty": 0, "WRONG": 0} for key in FIELDS}
    wrong = []
    for record in expected_rows:
        plate = _text(record.get("plate"))
        got = actual.get(plate)
        for key in FIELDS:
            wanted = _text(record.get(key))
            bucket = counts[key]
            found = "" if got is None else got.get(key, "")
            if got is None:
                bucket["WRONG"] += 1
                wrong.append(("1", plate, key, "", wanted))
            elif found == wanted:
                bucket["correct"] += 1
            elif not found:
                bucket["empty"] += 1
            else:
                bucket["WRONG"] += 1
                wrong.append(("1", plate, key, found, wanted))
    return counts, wrong


def _report(language: str, pdf_path: Path, truth_path: Path) -> str:
    os.environ["PDF_TOOL_LEARNED"] = str(pdf_path.with_suffix(".learned.json"))
    Path(os.environ["PDF_TOOL_LEARNED"]).write_text("{}", encoding="utf-8")
    os.environ["PDF_TOOL_CROPS"] = str(pdf_path.with_suffix(".crops"))
    result = read_pdf_workbook(pdf_path, strict=True)
    pdf = pdf_path.read_bytes()
    tiers = classify_cells(
        list(result.get("columns") or []),
        list(result.get("rows") or []),
        list(result.get("page_numbers") or []),
        read_print_date(pdf),
        result.get("provenance") or {},
        result.get("suggestions") or {},
        result.get("raw_readings") or {},
        pdf,
        result.get("baselines") or {},
        result.get("column_spans") or {},
    )
    counts, wrong = score_result(result, truth_path)
    red = {(str(item["page"]), str(item["plate"]), "date" if item["field"] == "date" else item["field"]) for item in tiers["red"]}
    yellow = {(str(item["page"]), str(item["plate"]), "date" if item["field"] == "date" else item["field"]) for item in tiers.get("yellow") or []}

    def placed(page, plate, field):
        key_field = "date" if field == "insurance_expiry" else field
        key = (page, plate, key_field)
        if key in red:
            return "red"
        if key in yellow:
            return "yellow"
        return "outside"

    in_red = sum(1 for page, plate, field, _found, _wanted in wrong if placed(page, plate, field) == "red")
    in_yellow = sum(1 for page, plate, field, _found, _wanted in wrong if placed(page, plate, field) == "yellow")
    items = list(tiers["red"]) + list(tiers.get("yellow") or [])
    crops = review_crops(
        items,
        pdf,
        list(result.get("columns") or []),
        list(result.get("rows") or []),
        list(result.get("page_numbers") or []),
        result.get("baselines") or {},
        result.get("column_spans") or {},
    )
    clusters = cluster_review_items(items, crops)
    plan = cluster_plan(clusters, items)
    lines = [
        f"== {language}",
        f"columns={result.get('columns')}",
        f"warning={result.get('column_warning') or ''}",
        f"visual_note={tiers.get('visual_note') or ''}",
        f"rows={len(result.get('rows') or [])}",
        f"red_cells={len(tiers['red'])} red_rows={len({(item['page'], item['plate']) for item in tiers['red']})}",
        f"yellow_cells={len(tiers.get('yellow') or [])} yellow_rows={len({(item['page'], item['plate']) for item in tiers.get('yellow') or []})}",
        f"known_wrong_in_red={in_red}/{len(wrong)} known_wrong_in_yellow={in_yellow}/{len(wrong)}",
        f"clusters={plan['clusters']} expected_clicks={plan['expected_clicks']} red_cells_after={plan['red_cells_after']} red_rows_after={plan['red_rows_after']}",
    ]
    lines.extend(f"{key}: correct={item['correct']} empty={item['empty']} WRONG={item['WRONG']}" for key, item in counts.items())
    return "\n".join(lines)


@pytest.mark.parametrize("language", ["arabic", "english"])
def test_synthetic_register_reports_coverage(language: str, tmp_path: Path, capsys):
    pdf_path = tmp_path / f"{language}.pdf"
    truth_path = tmp_path / f"{language}.csv"
    write_synthetic(language, pdf_path, truth_path)
    report = _report(language, pdf_path, truth_path)
    print(report)
    assert "columns=" in report
