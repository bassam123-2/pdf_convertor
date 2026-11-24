"""Every vehicle in the sample register, compared with the supplied expected file."""

from __future__ import annotations

import csv
import os
from pathlib import Path

import pytest

from core.excel_converter import convert_pdf_to_excel, read_pdf_workbook
from excel_review import classify_cells, read_print_date

SAMPLE = Path(r"C:\Users\jouan\Downloads\كشف سيارات الافضل 2026.pdf")
GOLDEN = Path(__file__).parent / "golden" / "sample_expected.csv"
_LRM = "\u200e"

# English names in the expected file, Arabic names in the workbook.
_FIELDS = {
    "purpose": "غرض الترخيص",
    "admin": "رقم إداري",
    "make": "الصنع",
    "year": "السنة",
    "color": "اللون",
    "insurance_expiry": "انتهاء التأمين",
}


def _text(value) -> str:
    return str(value or "").replace(_LRM, "").replace("\u200f", "").strip()


def _fold(value: str) -> str:
    text = _text(value)
    return (
        text.replace("أ", "ا")
        .replace("إ", "ا")
        .replace("آ", "ا")
        .replace("ة", "ه")
        .replace("ى", "ي")
    )


def _header_index(header: list[str], *needles: str) -> int | None:
    folded_needles = [_fold(needle) for needle in needles]
    for index, name in enumerate(header):
        folded = _fold(name)
        if any(needle in folded for needle in folded_needles):
            return index
    return None


def score_grid(header: list[str], body: list, golden_path: Path) -> tuple[dict, list[tuple]]:
    """Compare one grid with the expected file. A non-empty mismatch is WRONG."""
    plate_at = _header_index(header, "لوح")
    page_at = _header_index(header, "صفح")
    columns = {key: _header_index(header, label) for key, label in _FIELDS.items()}
    if plate_at is None:
        raise AssertionError(f"plate column missing from {header}")
    actual: dict[str, dict[str, str]] = {}
    for row in body:
        plate = _text(row[plate_at])
        if not plate:
            continue
        page = _text(row[page_at]) if page_at is not None else ""
        page = "".join(char for char in page if char.isdigit())
        actual[f"{page}|{plate}"] = {
            key: _text(row[index]) if index is not None and index < len(row) else ""
            for key, index in columns.items()
        }
        actual[f"{page}|{plate}"]["page"] = page
    with golden_path.open(encoding="utf-8-sig", newline="") as handle:
        expected_rows = list(csv.DictReader(handle))
    counts = {key: {"correct": 0, "empty": 0, "WRONG": 0} for key in _FIELDS}
    wrong: list[tuple[str, str, str, str]] = []
    for record in expected_rows:
        page = "".join(char for char in _text(record.get("page")) if char.isdigit())
        plate = _text(record.get("plate"))
        got = actual.get(f"{page}|{plate}")
        for key in _FIELDS:
            wanted = _text(record.get(key))
            bucket = counts[key]
            if got is None:
                bucket["WRONG"] += 1
                wrong.append((page, plate, key, "", wanted))
                continue
            found = got.get(key, "")
            if key == "make" and wanted in {"جي ام سي", "جي إم سي"} and found in {"جي ام سي", "جي إم سي"}:
                bucket["correct"] += 1
            elif found == wanted:
                bucket["correct"] += 1
            elif not found:
                bucket["empty"] += 1
            else:
                bucket["WRONG"] += 1
                wrong.append((page, plate, key, found, wanted))
    return counts, wrong


def score_workbook(excel_path: Path, golden_path: Path) -> tuple[dict, list[tuple]]:
    from openpyxl import load_workbook

    book = load_workbook(excel_path, data_only=True)
    sheet = book[book.sheetnames[0]]
    grid = list(sheet.iter_rows(values_only=True))
    book.close()
    header = [_text(cell) for cell in grid[0]]
    return score_grid(header, grid[1:], golden_path)


def test_golden_register(tmp_path: Path, capsys):
    if not SAMPLE.is_file() or not GOLDEN.is_file():
        pytest.skip("sample PDF or expected file is not on this machine")
    os.environ["PDF_TOOL_LEARNED"] = str(tmp_path / "learned.json")
    excel_path = tmp_path / "golden.xlsx"
    convert_pdf_to_excel(SAMPLE, excel_path, strict=True)
    counts, wrong = score_workbook(excel_path, GOLDEN)
    from openpyxl import load_workbook

    book = load_workbook(excel_path, read_only=True)
    review_name = "مراجعة" if "مراجعة" in book.sheetnames else "exceptions"
    exception_rows = book[review_name].max_row - 1 if review_name in book.sheetnames else 0
    book.close()
    result = read_pdf_workbook(SAMPLE, strict=True)
    print_date = read_print_date(SAMPLE.read_bytes())
    tiers = classify_cells(
        list(result.get("columns") or []),
        list(result.get("rows") or []),
        list(result.get("page_numbers") or []),
        print_date,
        result.get("provenance") or {},
        result.get("suggestions") or {},
        result.get("raw_readings") or {},
        SAMPLE.read_bytes(),
        result.get("baselines") or {},
        result.get("column_spans") or {},
    )
    def keys(name: str) -> set[tuple[str, str, str]]:
        return {
            (str(item["page"]), str(item["plate"]), item["field"])
            for item in tiers.get(name) or []
        }

    red = keys("red")
    yellow = keys("yellow")
    tier_field = {"insurance_expiry": "date"}

    def placed(page: str, plate: str, field: str) -> str:
        key = (page, plate, tier_field.get(field, field))
        if key in red:
            return "red"
        if key in yellow:
            return "yellow"
        return ""

    outside = [
        (page, plate, field)
        for page, plate, field, _found, _wanted in wrong
        if not placed(page, plate, field)
    ]
    in_red = sum(1 for page, plate, field, _found, _wanted in wrong if placed(page, plate, field) == "red")
    in_yellow = sum(1 for page, plate, field, _found, _wanted in wrong if placed(page, plate, field) == "yellow")
    red_rows = len({(item["page"], item["plate"]) for item in tiers["red"]})
    yellow_rows = len({(item["page"], item["plate"]) for item in tiers.get("yellow") or []})
    lines = [
        f"red_cells={len(tiers['red'])} red_rows={red_rows}",
        f"yellow_cells={len(tiers.get('yellow') or [])} yellow_rows={yellow_rows}",
        f"known_wrong_in_red={in_red}/{len(wrong)}",
        f"known_wrong_in_yellow={in_yellow}/{len(wrong)}",
        f"exceptions_sheet={exception_rows}",
    ]
    lines.extend(
        f"{key}: correct={item['correct']} empty={item['empty']} WRONG={item['WRONG']}"
        for key, item in counts.items()
    )
    for page, plate, field, found, wanted in wrong:
        lines.append(
            f"WRONG page={page} plate={plate} field={field} got={found} expected={wanted} tier={placed(page, plate, field) or 'outside'}"
        )
    report = "\n".join(lines)
    print(report)
    assert not outside, report
