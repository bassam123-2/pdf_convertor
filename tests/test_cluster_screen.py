# -*- coding: utf-8 -*-
"""Cluster accept: the counter, the workbook, and the golden comparison."""

from pathlib import Path

from openpyxl import load_workbook

from core.excel_converter import _VEHICLE_COLUMNS, write_edited_excel
from excel_review import audit_record, load_review_session, save_review_session
from test_golden import score_workbook

HEADERS = [name for _key, name in _VEHICLE_COLUMNS] + ["رقم الصفحة"]


def _row(color: str, plate: str) -> list[str]:
    return ["2026-04-09", color, "101", "تويوتا", "2024", plate, "خصوصي", "3"]


def test_accepting_a_group_updates_the_workbook_and_the_golden_comparison(tmp_path: Path):
    rows = [
        _row("صدفي", "18001"),
        _row("صدفي", "18002"),
        _row("صدفي", "18003"),
        _row("فضي", "27191"),
        _row("رمادي", "21804"),
    ]
    # The three صدفي cells are the only ones the button is allowed to write.
    rows[0][1] = "صدفي"
    rows[1][1] = "صدفي"
    rows[2][1] = "صدفي"
    audit = [
        audit_record("مراجع", 3, plate, "اللون", "صدفي", "صدفي")
        for plate in ("18001", "18002", "18003")
    ]
    # A real edit is one cell whose reading changes, plus two that already match.
    rows[0][1] = "صدفى"
    audit[0] = audit_record("مراجع", 3, "18001", "اللون", "صدفى", "صدفي")
    rows[0][1] = "صدفي"
    target = tmp_path / "reviewed.xlsx"
    write_edited_excel(HEADERS, rows, [3, 3, 3, 3, 3], target, kind="register", audit=audit)
    book = load_workbook(target, data_only=True)
    assert "سجل" in book.sheetnames
    sheet = book["سجل"]
    saved = list(sheet.iter_rows(values_only=True))
    book.close()
    assert saved[0] == ("الوقت", "المراجع", "الصفحة", "اللوحة", "الحقل", "قبل", "بعد")
    assert len(saved) == 4
    assert saved[1][1] == "مراجع"
    assert saved[1][5] == "صدفى"
    assert saved[1][6] == "صدفي"
    assert {row[3] for row in saved[1:]} == {"18001", "18002", "18003"}
    golden = tmp_path / "expected.csv"
    golden.write_text(
        "page,plate,purpose,admin,make,year,color,insurance_expiry\n"
        "3,18001,خصوصي,101,تويوتا,2024,صدفي,2026-04-09\n"
        "3,18002,خصوصي,101,تويوتا,2024,صدفي,2026-04-09\n"
        "3,18003,خصوصي,101,تويوتا,2024,صدفي,2026-04-09\n"
        "3,27191,خصوصي,101,تويوتا,2024,فضي,2026-04-09\n"
        "3,21804,خصوصي,101,تويوتا,2024,رمادي,2026-04-09\n",
        encoding="utf-8",
    )
    counts, wrong = score_workbook(target, golden)
    assert counts["color"]["WRONG"] == 0
    assert counts["color"]["correct"] == 5
    assert wrong == []


def test_a_review_session_resumes_for_the_same_pdf(tmp_path: Path):
    pdf = b"%PDF-session"
    payload = {
        "rows": [["أ"], ["ب"]],
        "red_done": ["1|18001|color"],
        "yellow_done": [],
        "cluster_out": ["3|27191|color"],
        "audit": [audit_record("مراجع", 1, "18001", "اللون", "فضي", "صدفي")],
        "undo": [],
        "revision": 0,
        "edit_serial": 2,
    }
    save_review_session(tmp_path, pdf, payload)
    loaded = load_review_session(tmp_path, pdf)
    assert loaded is not None
    assert loaded["red_done"] == ["1|18001|color"]
    assert loaded["audit"][0]["before"] == "فضي"
    assert loaded["audit"][0]["after"] == "صدفي"
    assert load_review_session(tmp_path, b"%PDF-other") is None


def test_the_cluster_screen_blocks_accept_until_the_group_is_seen():
    from streamlit.testing.v1 import AppTest

    app = AppTest.from_file(str(Path(__file__).with_name("cluster_screen_app.py")))
    app.run(timeout=30)
    accept = next(button for button in app.button if button.key == "accept")
    assert accept.disabled
    assert "نعم، الـ 17 المتبقية كلها صدفي" in accept.label
    assert any("remaining=19" in line.value for line in app.markdown)
    app.button(key="scan").click()
    app.run(timeout=30)
    accept = next(button for button in app.button if button.key == "accept")
    assert not accept.disabled
    accept.click()
    app.run(timeout=30)
    assert any("remaining=2" in line.value for line in app.markdown)
    assert any("تم 17 من 19" in line.value for line in app.markdown)
