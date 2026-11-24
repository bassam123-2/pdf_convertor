"""Cell review updates one row in memory. The file is not read again."""

from pathlib import Path

from openpyxl import load_workbook

from desktop_review import BUSY_EDIT, ReviewSession

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


def _session() -> ReviewSession:
    session = ReviewSession()
    session.prepare(
        {
            "columns": COLUMNS,
            "rows": [
                ["2026-04-09", "", "22", "تويوتا", "2024", "18011", "خصوصي", "لون مفقود"],
                ["2026-04-09", "أبيض", "23", "تويوتا", "2024", "18012", "خصوصي", ""],
            ],
            "page_numbers": [1, 1],
            "kind": "register",
        }
    )
    return session


def test_suggestions_are_ready_before_the_click():
    session = _session()
    item = session.current()
    assert item is not None
    assert item["plate"] == "18011"
    # Built during prepare. Reading them does not scan the table again.
    assert "suggestions" in item
    assert session.suggestions_for(item) == list(item["suggestions"])


def test_apply_changes_only_the_current_cell(tmp_path: Path):
    session = _session()
    item = session.current()
    rows = session.rows
    assert session.apply_item(item, "أبيض") is True
    assert rows[0][1] == "أبيض"
    assert rows[1][1] == "أبيض"
    assert rows[1][5] == "18012"
    assert all(queued["plate"] != "18011" or queued["field"] != "color" for queued in session.queue)
    target = tmp_path / "review.xlsx"
    session.save(target)
    book = load_workbook(target)
    sheet = book[book.sheetnames[0]]
    header = [str(cell.value or "") for cell in sheet[1]]
    color_at = header.index("اللون السائد") + 1
    assert sheet.cell(2, color_at).value == "أبيض"
    assert sheet.cell(3, color_at).value == "أبيض"
    book.close()
    assert session.undo_last() is True
    assert rows[0][1] == ""


def test_the_wait_text_is_the_one_shown_while_buttons_are_locked():
    assert "لا تكبس" in BUSY_EDIT
