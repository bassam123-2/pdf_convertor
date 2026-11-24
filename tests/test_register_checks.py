"""Validation and OCR repairs. The sample scan is checked when the file is present."""
import csv
from pathlib import Path

import fitz

from core.excel_converter import (
    ExtractedTable,
    _VEHICLE_COLUMNS,
    _confidence_band,
    _fill_purpose_down,
    _isolate_nondata_rows,
    _match_manufacturer,
    _match_official_color,
    _repair_group_dates,
    _dates_in_text,
    _footer_rule_y,
    _repair_plates,
    _repair_text_columns,
    _row_band_limits,
    convert_pdf_to_excel,
    forget_correction,
    load_corrections,
    neighbourhood_make,
    remember_correction,
    score_cell,
    validate_register,
    vote_readings,
    write_edited_excel,
)

HEADERS = [name for _key, name in _VEHICLE_COLUMNS]


def _row(date, color, make, year, plate, purpose):
    return [date, color, "", make, year, plate, purpose]


def test_validation_flags_a_short_group_and_a_short_page():
    columns = HEADERS
    rows = [
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "77309", "خصوصي"),
        _row("2026-04-10", "أسود", "نيسان", "2024", "18018", "خصوصي"),
    ]
    checks = validate_register(
        columns,
        rows,
        [1, 1],
        printed_total=3,
        printed_groups={"خصوصي": 2},
        detected_pages={1: 3},
    )
    by_name = {item["name"]: item["status"] for item in checks}
    assert by_name["المجموع"] == "fail"
    assert by_name["خصوصي"] == "ok"
    assert by_name["صفحة 1"] == "fail"
    weak = validate_register(
        columns,
        rows,
        [1, 1],
        detected_pages={1: 1},
    )
    page = next(item for item in weak if item["name"] == "صفحة 1")
    assert page["status"] == "weak"
    assert page["detail"] == "المرور الأول وجد 1، والباقي استُعيد بإعادة القراءة"


def test_cell_confidence_bands():
    assert score_cell("plate", "77309") >= 0.8
    assert _confidence_band(score_cell("plate", "77309")) == "green"
    assert score_cell("make", "") == 0.0
    assert _confidence_band(score_cell("make", "")) == "red"
    assert _confidence_band(score_cell("color", "ايبدي")) == "amber" or score_cell("color", "ايبدي") < 0.8
    assert score_cell("make", "تويوتا") >= 0.8
    assert score_cell("date", "04-09") < 0.8


def test_purpose_make_color_and_date_repairs_are_logged():
    table = ExtractedTable(
        HEADERS + ["ملاحظات"],
        [
            _row("2026-04-09", "أبيض", "تويوتا", "2024", "77309", "خصوصي") + [""],
            _row("2026-04-11", "أسود", "نيسان", "2024", "18011", "خصوصي") + [""],
            _row("2026-04-12", "أبيض", "تويوتا", "2024", "18012", "خصوصي") + [""],
            _row("04-09", "رمادي غ", "توو", "2024", "895C5", "") + [""],
            _row("2026-09-04", "ايبدي", "جي ام سي", "2024", "18018", "خصوصي") + [""],
            _row("2026-04-08", "أبيض", "تويوتا", "2024", "18019", "نفل خاص") + [""],
        ],
    )
    _fill_purpose_down(table)
    _repair_plates(table)
    assert table.rows[3][6] == "خصوصي"
    assert table.rows[5][6] == "نقل خاص"
    assert "نفل خاص -> نقل خاص" in table.rows[5][-1]
    assert table.rows[1][3] == "تويوتا" or _match_manufacturer("توو") == "تويوتا"
    assert _match_manufacturer("توو") == "تويوتا"
    assert _match_manufacturer("جي ام سي") == "جي إم سي"
    assert _match_official_color("رمادي غ") == "رمادي غامق"
    _repair_group_dates(table)
    assert table.rows[3][0] == "04-09"
    assert "2026-04-09" not in table.rows[3][-1]
    assert table.rows[4][0] == "2026-09-04"


def test_leading_banner_is_kept_and_not_counted():
    table = ExtractedTable(
        HEADERS + ["ملاحظات"],
        [
            _row("2028", "صدفي", "", "2025", "", "") + [""],
            _row("2026-04-09", "أبيض", "تويوتا", "2024", "77309", "خصوصي") + [""],
        ],
        None,
        [2, 2],
    )
    _isolate_nondata_rows(table)
    assert table.rows[0][-1] == "ترويسة"
    assert table.rows[1][5] == "77309"
    checks = validate_register(HEADERS + ["ملاحظات"], table.rows, [2, 2], detected_pages={2: 1})
    page = next(item for item in checks if item["name"] == "صفحة 2")
    assert page["status"] == "ok"
    assert page["actual"] == "1"


def test_english_table_still_converts(tmp_path: Path):
    pdf_path = tmp_path / "fleet.pdf"
    excel_path = tmp_path / "fleet.xlsx"
    document = fitz.open()
    page = document.new_page()
    columns = ["Plate", "Color", "Make", "Year"]
    xs = [72, 200, 320, 460]
    for x, word in zip(xs, columns):
        page.insert_text((x, 80), word)
    for y, values in ((110, ["18018", "White", "Toyota", "2024"]), (140, ["77309", "Black", "Nissan", "2023"])):
        for x, word in zip(xs, values):
            page.insert_text((x, y), word)
    document.save(pdf_path)
    document.close()
    count = convert_pdf_to_excel(pdf_path, excel_path)
    assert int(count) >= 1
    assert excel_path.is_file()


def test_vote_accepts_agreement_and_rejects_a_split(monkeypatch, tmp_path: Path):
    monkeypatch.setenv("PDF_TOOL_LEARNED", str(tmp_path / "learned.json"))
    chosen, reason = vote_readings("make", "جي ام سي", "جي إم سي")
    assert chosen == "جي إم سي"
    assert reason == "اتفاق المحركين"
    empty, why = vote_readings("make", "تويوتا", "نيسان")
    assert empty == ""
    assert why.startswith("لم يتفق المحركان")
    single, single_why = vote_readings("make", "جي ام سي", "", 0, 0)
    assert single == ""
    assert "جي إم سي" in single_why
    assert single_why.startswith("محرك واحد")
    remember_correction("color", "ابسنيحن", "أبيض")
    entry = load_corrections()["color"]["ابسنيحن"]
    assert entry["value"] == "أبيض"
    assert entry["source"] == "manual"
    assert entry["date"]
    forget_correction("color", "ابسنيحن")
    assert "ابسنيحن" not in load_corrections().get("color", {})


def test_spaced_arabic_is_collapsed_and_kept():
    from core.excel_converter import _collapse_spaced_arabic, _keep_arabic_value

    assert _collapse_spaced_arabic("ت و ي و ت ا") == "تويوتا"
    assert _collapse_spaced_arabic("ص د ف ى") == "صدفي"
    assert _collapse_spaced_arabic("ا ب ي ض") == "ابيض"
    assert _collapse_spaced_arabic("باص خاص") == "باص خاص"
    assert _keep_arabic_value("make", "ت و ي و ت ا") == "تويوتا"
    assert _keep_arabic_value("color", "ا ب ي ض") == "أبيض"
    assert _keep_arabic_value("color", "ص د ف ى") == "صدفي"
    assert _keep_arabic_value("color", "وردي") == "وردي"
    assert _keep_arabic_value("make", "ورديات") == "ورديات"


def test_unmatched_arabic_is_not_wiped():
    from core.excel_converter import fill_blank_fields_from_above, normalize_edited_dates

    table = ExtractedTable(
        HEADERS + ["ملاحظات"],
        [
            _row("2026 - 04 - 09", "ا ب ي ض", "ت و ي و ت ا", "2024", "18011", "خصوصي") + [""],
            _row("04-09", "وردي", "توو", "2024", "18012", "خصوصي") + [""],
            _row("2026-04-09", "", "", "2024", "18013", "خصوصي") + [""],
        ],
        None,
        [1, 1, 1],
    )
    _repair_text_columns(table)
    assert table.rows[0][1] == "أبيض"
    assert table.rows[0][3] == "تويوتا"
    assert table.rows[1][1] == "وردي"
    assert table.rows[1][3] == ""
    updated, changed = fill_blank_fields_from_above(HEADERS + ["ملاحظات"], table.rows, [1, 1, 1], ["make", "color"], [2])
    assert changed == 2
    assert updated[2][1] == "وردي"
    assert updated[2][3] == "تويوتا"
    dated, date_changes = normalize_edited_dates(HEADERS + ["ملاحظات"], table.rows, [0, 1])
    assert dated[0][0] == "2026-04-09"
    assert dated[1][0] == "04-09"
    assert date_changes >= 1


def test_gmc_spaced_forms_are_the_same_make():
    from core.excel_converter import _accept_make, _exact_manufacturer, _keep_arabic_value

    for spelling in ("جي إم سي", "جي ام سي", "جي  ام  سي", "جي أم سي"):
        assert _exact_manufacturer(spelling) == "جي إم سي"
        assert _accept_make(spelling) == "جي إم سي"
        assert _keep_arabic_value("make", spelling) == "جي إم سي"
    for spelling in ("ج ي ام سي", "جيامسي", "جيمسي"):
        assert _exact_manufacturer(spelling) == ""
        assert _accept_make(spelling) != "جي إم سي"
    for spelling in ("ام جى", "ام جي", "يو لو سى ام جى", "يولوسي امجي", "gmc", "GMC"):
        assert _accept_make(spelling) == "جي إم سي", spelling
        assert _keep_arabic_value("make", spelling) == "جي إم سي", spelling
    assert _accept_make("توو") == ""
    assert _match_manufacturer("جي ام سي") == "جي إم سي"


def test_margin_writes_official_names_and_keeps_the_raw_reading():
    from core.excel_converter import _accept_color, _accept_make, _provenance_sheet_rows

    for slip in ("توبو", "تويو", "توسو", "توبوسا"):
        assert _accept_make(slip) == "تويوتا", slip
    assert _accept_make("توو") == ""
    assert _accept_make("نتوسو") == ""
    assert _accept_make("جيمسي") == ""
    for slip, official in (("سدفسح", "صدفي"), ("صدفسح", "صدفي"), ("ابيض", "أبيض"), ("اسود", "أسود")):
        assert _accept_color(slip) == official
    assert _accept_color("سدفسح") == "صدفي"
    assert _accept_color("صدفسح") == "صدفي"
    table = ExtractedTable(
        HEADERS + ["ملاحظات"],
        [
            _row("2026-04-09", "سدفسح", "توبو", "2024", "18011", "خصوصي") + [""],
            _row("2026-04-09", "لون غريب", "نتوسو", "2024", "18012", "خصوصي") + [""],
        ],
        None,
        [1, 1],
    )
    _repair_text_columns(table)
    assert table.rows[0][3] == "تويوتا"
    assert table.rows[0][1] == "صدفي"
    assert table.rows[1][3] == "نتوسو"
    assert table.rows[1][1] == "لون غريب"
    assert table.raw_readings["1|18011|make"] == "توبو"
    assert table.provenance["1|18011|make"] == "normalized"
    assert table.provenance["1|18012|make"] == "one-engine"
    assert "تويوتا" in table.suggestions["1|18011|make"]
    assert "تطبيع" in table.suggestions["1|18011|make"]
    assert "تويوتا" in table.suggestions["1|18012|make"]
    sheet = _provenance_sheet_rows(table)
    make_row = next(row for row in sheet if row[1] == "18011" and row[2] == "الصنع")
    assert make_row[3] == "توبو"
    assert make_row[4] == "normalized"
    assert make_row[5]
    assert all(row[4] in {"agreed", "one-engine", "normalized", "suggested", "manual"} for row in sheet)


def test_a_year_on_the_same_band_completes_the_date():
    from core.excel_converter import _complete_row_date, _repair_group_dates

    assert _complete_row_date("2026 - 04 - 09") == "2026-04-09"
    assert _complete_row_date("\u202b04-09\u202c", "انتهاء 2026 04 09") == "2026-04-09"
    assert _complete_row_date("04-09", "") == "04-09"
    table = ExtractedTable(
        HEADERS + ["ملاحظات"],
        [
            _row("04-09", "أبيض", "تويوتا", "2024", "18011", "خصوصي") + [""],
            _row("04-09", "أبيض", "تويوتا", "2024", "18012", "خصوصي") + [""],
        ],
        None,
        [1, 1],
    )
    table.lines = ["2026-04-09 تويوتا", "لا سنة هنا"]
    _repair_group_dates(table)
    assert table.rows[0][0] == "2026-04-09"
    assert table.rows[1][0] == "04-09"


def test_a_missing_date_year_is_suggested_and_not_filled():
    from core.excel_converter import _infer_group_years

    columns = HEADERS + ["ملاحظات"]
    table = ExtractedTable(
        columns,
        [
            _row("2026-04-09", "أبيض", "تويوتا", "2024", "18011", "خصوصي") + [""],
            _row("2026-04-09", "أبيض", "تويوتا", "2024", "18012", "خصوصي") + [""],
            _row("2026-04-10", "أبيض", "تويوتا", "2024", "18013", "خصوصي") + [""],
            _row("04-09", "أبيض", "تويوتا", "", "18014", "خصوصي") + [""],
        ],
        None,
        [1, 1, 1, 1],
    )
    _infer_group_years(table)
    assert table.rows[3][0] == "04-09"
    assert table.rows[3][4] == ""
    assert "2026-04-09" in table.suggestions["1|18014|date"]
    assert "جوار" in table.suggestions["1|18014|date"]
    assert table.provenance["1|18014|date"] == "suggested"
    assert table.rows[3][0] == "04-09"


def test_column_comparison_counts_correct_empty_and_wrong():
    expected = {"لوحة": "18011", "الصنع": "تويوتا", "اللون": "أبيض"}
    actual = {"لوحة": "18011", "الصنع": "", "اللون": "أسود"}
    counts = {name: {"correct": 0, "empty": 0, "WRONG": 0} for name in expected}
    for name, wanted in expected.items():
        got = actual[name]
        if got == wanted:
            counts[name]["correct"] += 1
        elif not got:
            counts[name]["empty"] += 1
        else:
            counts[name]["WRONG"] += 1
    assert counts["لوحة"]["correct"] == 1
    assert counts["الصنع"]["empty"] == 1
    assert counts["اللون"]["WRONG"] == 1


def test_a_missing_year_is_not_copied_from_the_group():
    from core.excel_converter import _infer_group_years

    columns = HEADERS + ["ملاحظات"]
    table = ExtractedTable(
        columns,
        [
            _row("2026-04-09", "أبيض", "تويوتا", "2024", "18011", "خصوصي") + [""],
            _row("2026-04-09", "أبيض", "تويوتا", "2024", "18012", "خصوصي") + [""],
            _row("2026-04-09", "أبيض", "تويوتا", "2024", "18013", "خصوصي") + [""],
            _row("04-09", "أبيض", "تويوتا", "", "18014", "خصوصي") + [""],
        ],
        None,
        [1, 1, 1, 1],
    )
    _infer_group_years(table)
    _repair_group_dates(table)
    assert table.rows[3][4] == ""
    assert table.rows[3][0] == "04-09"


def test_last_row_band_stops_on_the_footer_rule():
    top, bottom = _row_band_limits(669.0, [640.0, 657.0, 669.0], 701.0, 842.0)
    assert abs(top - 663.0) < 0.01
    assert abs(bottom - 700.0) < 0.01
    middle_top, middle_bottom = _row_band_limits(657.0, [640.0, 657.0, 669.0], 701.0, 842.0)
    assert abs(middle_top - 648.5) < 0.01
    assert abs(middle_bottom - (657.0 + 669.0) / 2.0) < 0.01
    assert _dates_in_text("2026 - 04 - 09 2026 - 04 - 09") == "2026-04-09"
    assert _dates_in_text("2028 - 02 - 22 2025") == "2028-02-22"
    assert _dates_in_text("2026 - 04 - 09 2028 - 02 - 22") == ""
    document = fitz.open()
    page = document.new_page()
    page.draw_line((40, 700), (520, 700), width=1.4, color=(0, 0, 0))
    found = _footer_rule_y(page)
    document.close()
    assert found is not None
    assert abs(found - 700.0) < 3.0


def test_neighbour_make_is_suggested_and_not_written():
    columns = HEADERS + ["ملاحظات"]
    table = ExtractedTable(
        columns,
        [
            _row("2026-04-09", "أبيض", "نيسان", "2024", "18011", "خصوصي") + [""],
            _row("2026-04-09", "أبيض", "توو", "2024", "18012", "خصوصي") + [""],
            _row("2026-04-09", "أبيض", "نيسان", "2024", "18013", "خصوصي") + [""],
        ],
        None,
        [1, 1, 1],
    )
    assert neighbourhood_make(columns, table.rows, [1, 1, 1], 1) == "نيسان"
    _repair_text_columns(table)
    assert table.rows[1][3] == ""
    assert "نيسان" in table.suggestions["1|18012|make"]
    assert "جوار" in table.suggestions["1|18012|make"]
    assert "اقتراح" not in table.rows[1][-1]
    assert table.rows[1][3] != "نيسان"


def test_default_workbook_keeps_summary_and_numbers(tmp_path: Path):
    columns = ["رقم الصفحة", "غرض الترخيص", "رقم اللوحة", "رقم إداري", "الصنع", "السنة", "اللون", "انتهاء التأمين"]
    rows = [
        ["1", "خصوصي", "18011", "22", "تويوتا", "2024", "أبيض", "2026-04-09"],
        ["1", "خصوصي", "18012", "23", "تويوتا", "2024", "رمادي", "2026-04-09"],
        ["1", "خصوصي", "18013", "24", "شيفروليه", "2024", "أبيض", "2026-04-09"],
        ["1", "باص خاص", "18014", "25", "تاتويو", "2024", "أبيض", "2026-04-09"],
    ]
    target = tmp_path / "plain.xlsx"
    write_edited_excel(columns, rows, [1, 1, 1, 1], target, kind="register", provenance={"1|18011|color": "one-engine"})
    from openpyxl import load_workbook

    book = load_workbook(target)
    assert "ملخص" in book.sheetnames
    assert "provenance" not in book.sheetnames
    assert "excluded" not in book.sheetnames
    sheet = book[book.sheetnames[0]]
    header = [str(cell.value or "") for cell in sheet[1]]
    plate_at = header.index("رقم اللوحة") + 1
    page_at = header.index("رقم الصفحة") + 1
    assert sheet.cell(2, plate_at).value == 18011
    assert sheet.cell(2, header.index("السنة") + 1).value == 2024
    assert sheet.cell(2, header.index("رقم إداري") + 1).value == 22
    assert sheet.cell(2, page_at).value == 1
    assert sheet.cell(2, page_at).number_format == '"صفحة "0'
    assert "تاريخ للفرز" in header
    hidden = sheet.column_dimensions[sheet.cell(1, header.index("تاريخ للفرز") + 1).column_letter].hidden
    assert hidden is True
    summary = list(book["ملخص"].iter_rows(values_only=True))
    legend = [str(value or "") for row in summary for value in row]
    assert "قراءة واحدة" in legend
    assert "تصحيح إملائي" in legend
    assert "اقتراح" in legend
    assert any(str(row[0] or "") == "غير معروف: 1" for row in summary)
    assert any(str(row[0] or "") == "خصوصي" and str(row[2] or "") == "لم يُقرأ" for row in summary)
    book.close()
    checked = tmp_path / "checked.xlsx"
    write_edited_excel(
        columns,
        rows,
        [1, 1, 1, 1],
        checked,
        kind="register",
        validation=[{"name": "خصوصي", "expected": "485", "status": "ok", "actual": "3", "detail": ""}],
    )
    book = load_workbook(checked)
    summary = list(book["ملخص"].iter_rows(values_only=True))
    assert any(str(row[0] or "") == "خصوصي" and str(row[2] or "") == "485" for row in summary)
    book.close()


def test_header_rows_go_to_the_excluded_sheet(tmp_path: Path):
    columns = HEADERS + ["ملاحظات"]
    rows = [
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "77309", "خصوصي") + ["ترويسة"],
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "18018", "خصوصي") + [""],
    ]
    target = tmp_path / "split.xlsx"
    write_edited_excel(columns, rows, [1, 1], target, developer=True)
    from openpyxl import load_workbook

    book = load_workbook(target, data_only=True)
    assert "excluded" in book.sheetnames
    data = list(book[book.sheetnames[0]].iter_rows(values_only=True))
    assert all(str(row[-1] or "") not in {"ترويسة", "فاصل"} for row in data)
    excluded = list(book["excluded"].iter_rows(values_only=True))
    assert any("ترويسة" in [str(cell or "") for cell in row] for row in excluded)


ARABIC_SCAN = Path(r"C:\Users\jouan\Downloads\كشف تجريبي لترخيص وسيارات الأفضل 2026.pdf")
ENGLISH_DOC = Path(r"C:\Users\jouan\Downloads\english_learning_plan.pdf")


def test_other_arabic_document_converts(tmp_path: Path):
    if not ARABIC_SCAN.is_file():
        return
    excel_path = tmp_path / "arabic.xlsx"
    count = convert_pdf_to_excel(ARABIC_SCAN, excel_path, strict=True)
    assert int(count) >= 1
    assert excel_path.is_file()


def test_real_english_document_converts(tmp_path: Path):
    if not ENGLISH_DOC.is_file():
        return
    excel_path = tmp_path / "english.xlsx"
    count = convert_pdf_to_excel(ENGLISH_DOC, excel_path, strict=True)
    assert int(count) >= 1
    assert excel_path.is_file()


SAMPLE = Path(r"C:\Users\jouan\Downloads\كشف سيارات الافضل 2026.pdf")


def test_sample_register_matches_printed_totals(tmp_path: Path):
    if not SAMPLE.is_file():
        return
    excel_path = tmp_path / "sample.xlsx"
    convert_pdf_to_excel(SAMPLE, excel_path, strict=True, developer=True)
    from openpyxl import load_workbook

    book = load_workbook(excel_path, data_only=True)
    assert "validation" in book.sheetnames
    sheet = book[book.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))
    header = [str(cell or "") for cell in rows[0]]
    plate_at = next(index for index, name in enumerate(header) if "لوحة" in name)
    purpose_at = next(index for index, name in enumerate(header) if "غرض" in name)
    notes_at = next(index for index, name in enumerate(header) if "ملاحظات" in name)
    plated = []
    for row in rows[1:]:
        notes = str(row[notes_at] or "")
        if notes in {"ترويسة", "فاصل"}:
            continue
        plate = str(row[plate_at] or "")
        if plate.isdigit() and len(plate) >= 4:
            plated.append((plate, str(row[purpose_at] or "")))
            assert not any(char.isalpha() for char in plate)
    purposes = {}
    for _plate, purpose in plated:
        purposes[purpose] = purposes.get(purpose, 0) + 1
    assert len(plated) == 564
    assert purposes.get("باص خاص") == 24
    assert purposes.get("خصوصي") == 485
    assert purposes.get("نقل خاص") == 55
    page_at = next(index for index, name in enumerate(header) if "صفحة" in name)
    per_page = {}
    for row in rows[1:]:
        notes = str(row[notes_at] or "")
        if notes in {"ترويسة", "فاصل"}:
            continue
        plate = str(row[plate_at] or "")
        if not (plate.isdigit() and len(plate) >= 4):
            continue
        digits = "".join(char for char in str(row[page_at] or "") if char.isdigit())
        per_page[int(digits or 0)] = per_page.get(int(digits or 0), 0) + 1
    assert per_page[1] == 38
    for page in range(2, 13):
        assert per_page[page] == 41
    assert per_page[13] == 38
    assert per_page[14] == 37
    assert 15 not in per_page
    assert "excluded" in book.sheetnames
    assert all(str(row[notes_at] or "") not in {"ترويسة", "فاصل"} for row in rows[1:])
    assert all("محاولة" not in str(row[notes_at] or "") for row in rows[1:])
    validation = list(book["validation"].iter_rows(values_only=True))
    for item in validation[1:]:
        name, status, expected, actual = [str(cell or "") for cell in item[:4]]
        if not name.startswith("صفحة"):
            continue
        page = int(name.split()[-1])
        assert int(actual) == per_page.get(page, 0)
        if expected.isdigit() and int(actual) > int(expected):
            assert status == "weak"
        if expected.isdigit() and int(actual) < int(expected):
            assert status == "fail"
    make_at = next(index for index, name in enumerate(header) if "صنع" in name)
    year_at = next(index for index, name in enumerate(header) if "سنة" in name or name == "السنة")
    color_at = next(index for index, name in enumerate(header) if "لون" in name)
    date_at = next(index for index, name in enumerate(header) if "تأمين" in name or "تامين" in name)
    missing = 0
    plate_only = 0
    for row in rows[1:]:
        plate = str(row[plate_at] or "")
        if not (plate.isdigit() and len(plate) >= 4):
            continue
        blanks = sum(not str(row[index] or "").strip() for index in (make_at, year_at, color_at, date_at))
        if blanks:
            missing += 1
        if blanks == 4:
            plate_only += 1
    by_plate = {}
    for row in rows[1:]:
        plate = str(row[plate_at] or "")
        if plate.isdigit() and len(plate) >= 4:
            by_plate[plate] = row
    gmc = "جي إم سي"
    for plate in ("68994", "81365", "81544", "81649"):
        assert str(by_plate[plate][make_at] or "") in {"", gmc}, plate
    for plate in ("98447", "98452", "98689"):
        assert str(by_plate[plate][make_at] or "") in {"", "تويوتا"}, plate
    assert str(by_plate["98569"][make_at] or "") in {"", "تويوتا"}
    assert str(by_plate["16300"][make_at] or "") in {"", "شيفروليه"}
    assert str(by_plate["16300"][make_at] or "") != "تويوتا"
    for plate in ("98447", "98452", "98569", "98689", "16300"):
        assert str(by_plate[plate][make_at] or "") != gmc, plate
    for plate in ("89505", "89819", "89849"):
        make = str(by_plate[plate][make_at] or "")
        assert make not in {gmc, "شيفروليه", "رينو", "إيسوزو", "فورد", "نيسان", "هوندا", "هيونداي", "لكزس"}, plate
    assert str(by_plate["96995"][make_at] or "") not in {gmc, "تويوتا", "شيفروليه", "فورد"}
    assert "provenance" in book.sheetnames
    provenance = list(book["provenance"].iter_rows(values_only=True))
    sources = {str(row[4] or "") for row in provenance[1:]}
    assert sources <= {"agreed", "one-engine", "normalized", "suggested", "manual", ""}
    assert "inferred" not in sources
    assert "one-engine" not in sources or sources != {"one-engine"}
    assert any(str(row[5] or "").strip() for row in provenance[1:])
    assert str(by_plate["77309"][date_at] or "") != "2027-04-09"
    assert all(str(row[make_at] or "") != "فورد" for row in rows[1:])
    assert all(str(row[make_at] or "") not in {"توبو", "تويو", "توسو", "توبوسا", "توو"} for row in rows[1:])
    assert all(str(row[color_at] or "") not in {"سدفسح", "صدفسح", "صدفسج"} for row in rows[1:])


def _cell_text(value: object) -> str:
    marks = "\u200e\u200f\u202a\u202b\u202c\u2066\u2067\u2068\u2069"
    return "".join(char for char in str(value or "") if char not in marks).strip()


def test_catalog_prior_is_not_written_for_the_regressed_plates():
    """Every plate from the regression. A prior stays out of the cell; empty is allowed."""
    from core.excel_converter import _repair_text_columns

    make_truth = {
        "68994": "جي إم سي",
        "81365": "جي إم سي",
        "81544": "جي إم سي",
        "81649": "جي إم سي",
        "97307": "شيفروليه",
        "97500": "شيفروليه",
        "16300": "شيفروليه",
    }
    color_truth = {
        "15415": "صدفي",
        "89505": "صدفي",
        "89702": "صدفي",
        "89705": "صدفي",
        "89805": "صدفي",
        "89819": "صدفي",
        "89844": "صدفي",
        "89929": "صدفي",
        "98221": "صدفي",
        "98447": "صدفي",
        "98452": "صدفي",
        "98706": "صدفي",
        "98803": "صدفي",
        "98809": "صدفي",
        "98854": "صدفي",
        "98861": "صدفي",
        "98392": "صدفي",
        "98689": "أسود",
        "16300": "أحمر",
        "98420": "رمادي",
        "97991": "أبيض",
    }
    date_truth = {
        "15366": "2028-02-22",
        "15415": "2028-02-22",
        "89705": "2028-02-22",
        "89819": "2028-02-22",
        "89843": "2028-02-22",
        "89844": "2028-02-22",
        "89925": "2028-02-22",
        "89929": "2028-02-22",
        "98214": "2028-02-22",
        "97799": "2028-04-01",
        "97821": "2028-04-01",
        "98221": "2028-02-25",
        "98706": "2028-02-25",
        "98452": "2028-02-24",
        "98689": "2028-03-18",
    }
    plates = list(dict.fromkeys([*make_truth, *color_truth, *date_truth]))
    columns = HEADERS + ["ملاحظات"]
    golden = Path(__file__).parent / "golden" / "regression_plates.csv"
    listed = {row["رقم اللوحة"] for row in csv.DictReader(golden.open(encoding="utf-8-sig"))}
    assert listed == set(plates)
    rows = []
    for plate in plates:
        rows.append(_row("", "أبيض", "", "2025", plate, "خصوصي") + [""])
    table = ExtractedTable(columns, rows, None, [1] * len(rows))
    for plate in plates:
        table.raw_readings[f"1|{plate}|color"] = "ماد" if plate in color_truth else ""
        table.rows[plates.index(plate)][1] = "أبيض"
    _repair_text_columns(table)
    for index, plate in enumerate(plates):
        assert table.rows[index][3] == ""
        assert table.provenance.get(f"1|{plate}|make") != "inferred"
        assert "تويوتا 100" not in table.suggestions.get(f"1|{plate}|make", "")
        if plate in color_truth:
            assert table.rows[index][1] == "ماد"
            assert table.rows[index][1] != "أبيض"
    for token in ("ماد", "امنير", "ابنيضر", "اير", "ابسنيحن", "انير", "مأد", "صده"):
        held = ExtractedTable(
            columns,
            [_row("2026-04-09", token, "نتوسو", "2024", "18011", "خصوصي") + [""]],
            None,
            [1],
        )
        held.raw_readings["1|18011|color"] = token
        held.raw_readings["1|18011|make"] = "نتوسو"
        _repair_text_columns(held)
        assert held.rows[0][1] == token
        assert held.rows[0][1] != "أبيض"
        assert held.rows[0][3] == "نتوسو"
    dated = ExtractedTable(
        columns,
        [
            _row("2028-02-23", "أبيض", "تويوتا", "2025", "15366", "خصوصي") + [""],
            _row("", "أبيض", "تويوتا", "2025", "15415", "خصوصي") + [""],
        ],
        None,
        [1, 1],
    )
    dated.lines = ["2028-02-23", "2028-02-23"]
    from core.excel_converter import _fill_blanks_from_line, _infer_group_years

    _fill_blanks_from_line(dated)
    _infer_group_years(dated)
    assert dated.rows[1][0] == ""
    assert dated.provenance.get("1|15415|date") != "inferred"


def test_golden_sample_matches_expected_cells(tmp_path: Path):
    """Cell-by-cell check. The expected file is supplied separately."""
    golden = Path(__file__).parent / "golden" / "regression_plates.csv"
    if not golden.is_file() or not SAMPLE.is_file():
        return
    excel_path = tmp_path / "golden.xlsx"
    convert_pdf_to_excel(SAMPLE, excel_path, strict=True)
    from openpyxl import load_workbook

    book = load_workbook(excel_path, data_only=True)
    sheet = book[book.sheetnames[0]]
    grid = list(sheet.iter_rows(values_only=True))
    header = [_cell_text(cell) for cell in grid[0]]
    with golden.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        expected_rows = list(reader)
        expected_header = [_cell_text(name) for name in (reader.fieldnames or [])]
    plate_name = next(name for name in expected_header if "لوحة" in name)
    page_name = next((name for name in expected_header if "صفحة" in name), "")

    def identity(record: dict[str, str]) -> tuple[str, str]:
        page = _cell_text(record.get(page_name, "")) if page_name else ""
        page_digits = "".join(char for char in page if char.isdigit())
        return page_digits, _cell_text(record.get(plate_name, ""))

    actual: dict[tuple[str, str], dict[str, str]] = {}
    by_plate: dict[str, dict[str, str]] = {}
    for row in grid[1:]:
        record = {name: _cell_text(cell) for name, cell in zip(header, row)}
        key = identity(record)
        actual[key] = record
        by_plate.setdefault(key[1], record)
    expected = {identity({name: _cell_text(row.get(raw, "")) for name, raw in zip(expected_header, reader.fieldnames or [])}): {
        name: _cell_text(row.get(raw, "")) for name, raw in zip(expected_header, reader.fieldnames or [])
    } for row in expected_rows}
    counts: dict[str, dict[str, int]] = {}
    wrong: dict[str, list[str]] = {}
    for key, wanted in expected.items():
        got = actual.get(key) if key[0] else by_plate.get(key[1])
        if got is None:
            wrong.setdefault(plate_name, []).append(f"صفحة {key[0]} لوحة {key[1]} مفقودة من الناتج")
            continue
        for name in expected_header:
            if name in {plate_name, page_name} or not wanted.get(name):
                continue
            bucket = counts.setdefault(name, {"correct": 0, "empty": 0, "WRONG": 0})
            if name not in got:
                bucket["WRONG"] += 1
                wrong.setdefault(name, []).append(f"صفحة {key[0]} لوحة {key[1]}: العمود غير موجود")
                continue
            if got[name] == wanted[name] or (name == "الصنع" and wanted[name] in {"جي ام سي", "جي إم سي"} and got[name] in {"جي ام سي", "جي إم سي"}):
                bucket["correct"] += 1
            elif not got[name]:
                bucket["empty"] += 1
            else:
                bucket["WRONG"] += 1
                wrong.setdefault(name, []).append(
                    f"لوحة {key[1]}: المتوقع {wanted[name]} الناتج {got[name]}"
                )
    lines = [f"{name}: correct={item['correct']} empty={item['empty']} WRONG={item['WRONG']}" for name, item in counts.items()]
    total_wrong = sum(item["WRONG"] for item in counts.values()) + len(wrong.get(plate_name, []))
    if total_wrong:
        detail = []
        for name, items in wrong.items():
            detail.append(f"{name}: {len(items)}")
            detail.extend(f"  {item}" for item in items[:8])
        raise AssertionError("\n".join(lines + detail))
