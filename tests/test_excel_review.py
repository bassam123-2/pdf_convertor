"""The review list is the only count. It does not rewrite a cell."""

from excel_review import banner_facts, blank_text, clear_field_note, print_date_from_text, review_cells

COLUMNS = [
    "انتهاء التأمين",
    "اللون السائد",
    "رقم اداري",
    "الصنع",
    "السنة",
    "رقم اللوحة",
    "غرض الترخيص",
    "ملاحظات",
]


def _row(date, color, make, year, plate, purpose="خصوصي", admin="101"):
    return [date, color, admin, make, year, plate, purpose, ""]


def test_blank_text_never_says_none():
    assert blank_text(None) == ""
    assert blank_text("None") == ""
    assert blank_text(" nan ") == ""
    assert blank_text("أبيض") == "أبيض"


def test_footer_clock_is_the_print_date_not_a_vehicle_date():
    text = "2028-03-23 1 08.47.42 2026-04-06"
    assert print_date_from_text(text) == "2026-04-06"


def test_plate_77309_is_flagged_and_the_cell_stays():
    rows = [
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "18011"),
        _row("2023-09-04", "أبيض", "تويوتا", "2024", "77309"),
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "18013"),
    ]
    issues = review_cells(COLUMNS, rows, [1, 1, 1], "2026-04-06")
    assert [(item["plate"], item["field"]) for item in issues] == [("77309", "date")]
    assert issues[0]["suggestions"] == ["2026-04-09"]
    assert "2026-04-06" in issues[0]["reason"]
    assert rows[1][0] == "2023-09-04"


def test_empty_and_unknown_share_one_list():
    rows = [
        _row("2026-04-09", None, "نتوسو", "2024", "18011"),
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "18012"),
    ]
    issues = review_cells(COLUMNS, rows, [1, 1], "2026-04-06")
    assert [(item["field"], item["reason"]) for item in issues] == [
        ("make", "قيمة غير معروفة"),
        ("color", "فارغ"),
    ]
    assert all(item["value"] != "None" for item in issues)
    assert issues[1]["value"] == ""


def test_year_after_the_print_year_is_flagged_and_not_written():
    rows = [
        _row("2026-04-09", "أبيض", "تويوتا", "2025", "89928"),
        _row("2028-03-23", "أبيض", "تويوتا", "2028", "89929"),
        _row("2026-04-09", "أبيض", "تويوتا", "2025", "89930"),
    ]
    issues = review_cells(COLUMNS, rows, [3, 3, 3], "2026-04-06")
    year_issues = [item for item in issues if item["plate"] == "89929" and item["field"] == "year"]
    assert len(year_issues) == 1
    assert "2026" in year_issues[0]["reason"]
    assert year_issues[0]["suggestions"] == ["2025"]
    assert "2028" not in year_issues[0]["suggestions"]
    assert rows[1][4] == "2028"


def test_empty_admin_is_only_a_suggestion():
    rows = [
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "18011", admin="100"),
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "18012", admin=""),
        _row("2026-04-09", "أبيض", "تويوتا", "2024", "18013", admin="100"),
    ]
    issues = review_cells(COLUMNS, rows, [1, 1, 1], "2026-04-06")
    admin = [item for item in issues if item["field"] == "admin"]
    assert [(item["plate"], item["suggestions"]) for item in admin] == [("18012", ["100"])]
    assert rows[1][2] == ""
    filled = _row("2026-04-09", "أبيض", "تويوتا", "2024", "18014", admin="100")
    assert [item["field"] for item in review_cells(COLUMNS, [filled], [1], "2026-04-06")] == []


def test_clip_suggestions_are_separate_from_neighbours():
    from excel_review import clip_acceptances, filter_review_items

    items = [
        {
            "plate": "1",
            "suggestions": ["تويوتا", "نيسان"],
            "suggestion_reasons": ["قراءة المقطع", "جوار نفس الصفحة والمجموعة"],
        },
        {
            "plate": "2",
            "suggestions": ["شيفروليه"],
            "suggestion_reasons": ["جوار نفس الصفحة والمجموعة"],
        },
        {
            "plate": "3",
            "track": "safe",
            "suggestions": ["نيسان"],
            "suggestion_reasons": ["اتفاق المحركين"],
        },
        {
            "plate": "4",
            "track": "exception",
            "suggestions": ["رمادي"],
            "suggestion_reasons": ["اتفاق المحركين"],
        },
    ]
    accepted = clip_acceptances(items)
    assert [item["plate"] for item, _value in accepted] == ["3"]
    assert accepted[0][1] == "نيسان"
    assert [item["plate"] for item in filter_review_items(items, "neighbour")] == ["2"]
    assert [item["plate"] for item in filter_review_items(items, "clip")] == ["1"]


def test_a_missing_row_y_comes_from_the_neighbours_on_the_page():
    from core.excel_converter import interpolate_row_y

    baselines = {"11|100": 100.0, "11|300": 140.0}
    y_value = interpolate_row_y(baselines, 11, ["100", "98689", "300"], "98689")
    assert abs(y_value - 120.0) < 0.01
    assert interpolate_row_y(baselines, 11, ["100", "98689", "300"], "100") == 100.0


def test_unknown_colors_are_flagged_like_unknown_makes():
    for token in ("ماد", "امنير", "ابنيضر", "اير", "ابسنيحن", "انير", "مأد", "صده"):
        rows = [_row("2026-04-09", token, "تويوتا", "2024", "18011")]
        issues = review_cells(COLUMNS, rows, [1], "2026-04-06")
        color = [item for item in issues if item["field"] == "color"]
        assert color and color[0]["reason"] == "قيمة غير معروفة", token
        assert rows[0][1] == token


def test_editing_a_flagged_cell_clears_that_note():
    row = _row("2026-04-09", "أبيض", "تويوتا", "2024", "18011")
    row[-1] = "سنة مفقودة؛ تاريخ مفقود؛ صنع مفقود"
    assert clear_field_note(COLUMNS, row, "make") == "سنة مفقودة؛ تاريخ مفقود"
    assert "صنع مفقود" not in row[-1]
    clear_field_note(COLUMNS, row, "year")
    clear_field_note(COLUMNS, row, "date")
    assert row[-1] == "تمت المراجعة"
    header = list(row)
    header[-1] = "ترويسة"
    assert clear_field_note(COLUMNS, header, "make") == "ترويسة"


def test_validation_follows_the_edited_rows():
    from core.excel_converter import validate_register

    columns = ["غرض الترخيص", "رقم اللوحة", "ملاحظات"]
    rows = [["خصوصي", "12345", ""], ["خصوصي", "12346", ""]]
    pages = [1, 1]
    groups = {"خصوصي": 2, "نقل خاص": 0}
    before = {item["name"]: item["actual"] for item in validate_register(columns, rows, pages, 2, groups, {1: 2})}
    rows[0][0] = "نقل خاص"
    after = {item["name"]: item["actual"] for item in validate_register(columns, rows, pages, 2, groups, {1: 2})}
    assert before["خصوصي"] == "2"
    assert after["خصوصي"] == "1"
    assert after["نقل خاص"] == "1"


def test_banner_uses_the_review_list_once():
    validation = [
        {"name": "المجموع", "status": "ok", "expected": "564", "actual": "564"},
        {"name": "باص خاص", "status": "ok", "expected": "24", "actual": "24"},
        {"name": "خصوصي", "status": "ok", "expected": "485", "actual": "485"},
        {"name": "نقل خاص", "status": "ok", "expected": "55", "actual": "55"},
        {"name": "صفحة 3", "status": "fail", "expected": "41", "actual": "40"},
        {"name": "صفحة 4", "status": "ok", "expected": "41", "actual": "41"},
    ]
    facts = banner_facts(validation, 34)
    assert facts["count_line"] == "✅ العدد مطابق: 564 من 564 (24/485/55)"
    assert facts["issue_line"] == "⚠ 34 خلية تحتاج مراجعة"
    assert facts["issue_count"] == 34
    assert [page["page"] for page in facts["pages"]] == [3]
    assert "المرور الأول وجد 41" in facts["pages"][0]["line"]


def test_a_glyph_on_the_cut_does_not_belong_to_either_row():
    from core.excel_converter import _glyph_owns_row

    baselines = [100.0, 114.0]
    assert _glyph_owns_row(110.0, 114.0, baselines)
    assert not _glyph_owns_row(107.0, 114.0, baselines)
    assert not _glyph_owns_row(107.0, 100.0, baselines)
    assert _glyph_owns_row(102.0, 100.0, baselines)


def test_a_rare_value_on_two_adjacent_rows_is_not_written():
    from core.excel_converter import ExtractedTable, _repair_text_columns

    columns = ["اللون السائد", "رقم اللوحة", "غرض الترخيص", "ملاحظات"]
    rows = [
        ["صدفي", "18011", "خصوصي", ""],
        ["رمادي", "18012", "خصوصي", ""],
        ["رمادي", "18013", "خصوصي", ""],
        ["صدفي", "18014", "خصوصي", ""],
        ["صدفي", "18015", "خصوصي", ""],
    ]
    table = ExtractedTable(columns, rows, None, [1, 1, 1, 1, 1], kind="register")
    _repair_text_columns(table)
    assert [row[0] for row in table.rows] == ["صدفي", "", "", "صدفي", "صدفي"]
    assert "قيمة نادرة" in table.suggestions.get("1|18012|color", "")


def test_an_outlier_is_not_a_safe_batch():
    from excel_review import assign_review_tracks

    columns = ["اللون السائد", "رقم اللوحة"]
    rows = [["أبيض", "1"], ["أبيض", "2"], ["رمادي", "3"]]
    items = [
        {"page": 1, "field": "color", "value": "رمادي", "suggestions": ["رمادي"], "suggestion_reasons": ["اتفاق المحركين"]},
        {"page": 1, "field": "color", "value": "", "suggestions": ["أبيض"], "suggestion_reasons": ["اتفاق المحركين"]},
    ]
    grouped = assign_review_tracks(items, rows, [1, 1, 1], columns)
    assert [item["value"] for item in grouped["exception"]] == ["رمادي"]
    assert grouped["safe"][0]["suggestions"] == ["أبيض"]


def test_a_rare_color_is_mandatory_and_agreement_does_not_hide_it():
    from excel_review import classify_cells

    columns = ["اللون", "رقم اللوحة"]
    rows = [["أبيض", "18011"], ["أبيض", "18012"], ["أبيض", "18013"], ["رمادي", "18014"]]
    provenance = {"1|18011|color": "agreed", "1|18013|color": "agreed"}
    tiers = classify_cells(columns, rows, [1, 1, 1, 1], "", provenance, {}, {})
    red = {(item["plate"], item["field"]) for item in tiers["red"]}
    green = {(item["plate"], item["field"]) for item in tiers["green"]}
    assert ("18014", "color") in red
    assert ("18013", "color") in red
    assert ("18013", "color") not in green
    assert ("18011", "color") in green
    assert ("18011", "color") not in red


def test_unclear_text_is_yellow_not_red():
    from excel_review import classify_cells

    columns = ["اللون", "رقم اللوحة"]
    rows = [["فيصد", "18011"], ["أبيض", "18012"], ["أبيض", "18013"], ["أبيض", "18014"]]
    tiers = classify_cells(columns, rows, [1, 1, 1, 1], "", {}, {}, {})
    yellow = {(item["plate"], item["field"]) for item in tiers["yellow"]}
    red = {(item["plate"], item["field"]) for item in tiers["red"]}
    assert ("18011", "color") in yellow
    assert ("18011", "color") not in red


def test_a_crop_unlike_its_group_is_flagged():
    import numpy as np

    from excel_review import dissimilar_keys

    same = np.zeros((16, 32), dtype="float32")
    same[:, 8:24] = 1
    odd = np.ones((16, 32), dtype="float32") - same
    records = [(1, str(18011 + index), "color", "أبيض", same.copy()) for index in range(10)]
    records.append((1, "18099", "color", "أبيض", odd))
    flagged = {(plate, field) for _page, plate, field in dissimilar_keys(records)}
    assert ("18099", "color") in flagged
    assert ("18011", "color") not in flagged


def test_matching_crops_share_one_question():
    import numpy as np

    from excel_review import cluster_plan, cluster_review_items

    image = np.zeros((16, 32), dtype="float32")
    image[:, 8:24] = 1
    items = []
    crops = {}
    for index in range(6):
        item = {
            "page": 1,
            "plate": str(18011 + index),
            "field": "color",
            "value": "" if index == 0 else "صدفي",
            "tier": "yellow" if index == 0 else "red",
            "row_index": index,
        }
        items.append(item)
        crops[("1", item["plate"], "color")] = image.copy()
    odd = np.ones((16, 32), dtype="float32")
    items.append(
        {"page": 1, "plate": "18099", "field": "color", "value": "أسود", "tier": "red", "row_index": 9}
    )
    crops[("1", "18099", "color")] = odd
    clusters = cluster_review_items(items, crops)
    assert len(clusters) == 1
    assert clusters[0]["label"] == "صدفي"
    plates = {member["plate"] for member in clusters[0]["members"]}
    assert plates == {str(18011 + index) for index in range(1, 6)}
    assert {item["plate"] for item in clusters[0]["outliers"]} == {"18011"}
    plan = cluster_plan(clusters, items)
    assert plan["clusters"] == 1
    assert plan["expected_clicks"] == 3
    assert plan["red_cells_after"] == 1
    assert plan["red_rows_after"] == 1


def test_a_different_reading_is_not_accepted_with_the_group():
    import numpy as np

    from excel_review import cluster_decision, cluster_review_items, remaining_cells

    image = np.zeros((16, 32), dtype="float32")
    image[:, 8:24] = 1
    specs = [("18001", "صدفي"), ("18002", "صدفي"), ("18003", "صدفي"), ("27191", "فضي"), ("21804", "رمادي")]
    items = []
    crops = {}
    for plate, color in specs:
        item = {"page": 3, "plate": plate, "field": "color", "value": color, "tier": "red", "row_index": 0}
        items.append(item)
        crops[("3", plate, "color")] = image.copy()
    clusters = cluster_review_items(items, crops)
    assert len(clusters) == 1
    assert {member["plate"] for member in clusters[0]["members"]} == {"18001", "18002", "18003"}
    assert {item["plate"] for item in clusters[0]["outliers"]} == {"27191", "21804"}
    blocked = cluster_decision(clusters[0], set(), scanned=False)
    assert blocked["ready"] is False
    assert blocked["write"] == []
    assert blocked["label"] == "نعم، الـ 3 المتبقية كلها صدفي"
    opened = cluster_decision(clusters[0], set(), scanned=True)
    assert [item["plate"] for item in opened["write"]] == ["18001", "18002", "18003"]
    ejected = cluster_decision(clusters[0], {"3|27191|color", "3|21804|color"}, scanned=False)
    assert ejected["ready"] is True
    assert {item["plate"] for item in ejected["write"]} == {"18001", "18002", "18003"}
    assert remaining_cells(5, [], [f"3|{plate}|color" for plate in ("18001", "18002", "18003")]) == 2


def test_accepting_n_cells_drops_the_remaining_counter_by_exactly_n():
    from excel_review import progress_text, remaining_cells

    total = 19
    done: list[str] = []
    for count in (2, 6, 19):
        accepted = [f"1|{index}|color" for index in range(count)]
        before = remaining_cells(total, done, [])
        after = remaining_cells(total, done, accepted)
        assert before - after == count
    done = [f"1|{index}|color" for index in range(6)]
    assert remaining_cells(total, done, []) == 13
    assert remaining_cells(total, done, [f"1|{index}|color" for index in range(6, 8)]) == 11
    assert "تم 6 من 19" in progress_text(6, 19, 40)


def test_a_short_reading_says_the_document_is_too_small():
    from excel_review import visual_note_for

    assert visual_note_for({(1, "color", "أبيض"): 3}) == "document too small for visual comparison"
    assert visual_note_for({(1, "color", "أبيض"): 8}) == ""


def test_a_card_box_stays_on_one_row():
    import numpy as np
    from PIL import Image
    import io

    from page_images import PageBank, card_png

    bank = PageBank(b"")
    bank.pages[1] = np.full((400, 600), 240, dtype=np.uint8)
    bank.pages[1][180:220, 40:120] = 20
    bank.size[1] = (600 / (200 / 72), 400 / (200 / 72))
    bank.ready = True
    png = card_png(bank, 1, 72.0, "color", {"color": (20.0, 70.0), "make": (200.0, 260.0)}, {"1|1": 72.0, "1|2": 85.0})
    image = Image.open(io.BytesIO(png)).convert("RGB")
    pixels = np.asarray(image)
    red = (pixels[:, :, 0] > 150) & (pixels[:, :, 1] < 80)
    rows = np.where(red.any(1))[0]
    assert len(rows)
    assert rows[-1] - rows[0] < image.height * 0.55


def test_vision_disagreement_moves_a_yellow_cell_to_red():
    from excel_review import vision_queue

    yellow = [{"page": 1, "plate": "11134", "field": "make", "value": "توسوسا", "tier": "yellow"}]
    red = [{"page": 1, "plate": "77309", "field": "date", "value": "2023-09-04", "tier": "red"}]
    moved = vision_queue(red, yellow, {"1|11134|make": "تويوتا", "1|77309|date": "2023-09-04"})
    assert len(moved["red"]) == 2
    assert moved["yellow"] == []
    assert any(item["plate"] == "11134" and item["reason"] == "الرؤية تختلف عن القراءة" for item in moved["red"])


def test_an_empty_queue_is_not_clean_when_columns_are_missing():
    from excel_review import queue_status

    warning = "تعذر التعرف على أعمدة المركبات. قائمة فارغة لا تعني أن الملف سليم."
    assert "لا تعني" in queue_status(0, warning)
    assert queue_status(0, "").startswith("✅")


def test_symbol_font_letters_are_garbage(tmp_path):
    import fitz

    from core.excel_converter import build_column_warning, text_layer_is_garbage

    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 80), "n" * 40, fontname="zadb", fontsize=16)
    page.insert_text((72, 120), "18018 2025", fontname="helv", fontsize=12)
    assert text_layer_is_garbage(page)
    warning = build_column_warning(["IIIIII", "IIII", "Page"], garbage=True, kind="register")
    assert "لا تعني" in warning
    assert build_column_warning(["رقم اللوحة", "الصنع", "اللون"], garbage=False, kind="register") == ""
    document.close()
    del tmp_path


def test_a_confirmed_crop_is_stored_and_can_be_deleted(tmp_path, monkeypatch):
    import numpy as np

    from excel_review import clear_crop_library, crop_library_count, load_library_images, remember_crop

    monkeypatch.setenv("PDF_TOOL_CROPS", str(tmp_path / "crops"))
    image = np.zeros((16, 32), dtype="float32")
    image[:, 4:20] = 1
    remember_crop("color", "صدفي", image)
    remember_crop("color", "", image)
    assert crop_library_count() == 1
    loaded = load_library_images("color")
    assert loaded and loaded[0][1] == "صدفي"
    clear_crop_library()
    assert crop_library_count() == 0
