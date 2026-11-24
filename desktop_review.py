# -*- coding: utf-8 -*-
"""Desktop review window. This file does not use Streamlit.

Optimizations
-------------
1. The PDF is read once. Later clicks change one cell in the rows already
   held in memory. They do not open the PDF or rebuild the workbook.
2. Every review cell carries its suggestions from that single pass. A click
   only displays the list that is already on the cell.
3. Conversion, the one-time queue build, and saving run on a worker thread.
   The window only paints status. Action buttons are disabled before the
   thread starts, so a second click cannot write a different cell.
"""

from __future__ import annotations

import threading
from pathlib import Path

from excel_review import blank_text, classify_cells, clear_field_note, read_print_date

# Shown the moment a click is accepted, before the worker runs.
BUSY_EDIT = "استنى شوي، جاري التعديل... لا تكبس على أي شيء"
BUSY_CONVERT = "استنى شوي، جاري التحويل... لا تكبس على أي شيء"
BUSY_SAVE = "استنى شوي، جاري الحفظ... لا تكبس على أي شيء"
DONE_EDIT = "تم التعديل"
DONE_CONVERT = "تم التحويل"
DONE_SAVE = "تم الحفظ"

_FIELD_LABELS = {
    "date": "انتهاء التأمين",
    "make": "الصنع",
    "color": "اللون",
    "year": "السنة",
    "admin": "رقم إداري",
}


class ReviewSession:
    """One converted table, kept in memory for the whole review."""

    def __init__(self) -> None:
        self.columns: list[str] = []
        self.rows: list[list[str]] = []
        self.pages: list[int] = []
        self.queue: list[dict] = []
        self.cursor = 0
        self.kind = ""
        self.strict = True
        self.validation: list[dict] = []
        self.provenance: dict[str, str] = {}
        self.suggestions: dict[str, str] = {}
        self.raw_readings: dict[str, str] = {}
        self.baselines: dict = {}
        self.column_spans: dict = {}
        self.figures = None
        self.print_date = ""
        self.undo: list[dict] = []
        # Serializes cell writes. The window's busy flag stops the clicks.
        self._lock = threading.Lock()

    def prepare(self, result: dict, pdf_bytes: bytes | None = None) -> int:
        """Copy the extract and build the review list once.

        Suggestions are stored on each queued cell here. Nothing in this
        method reads a cell's crop again.
        """
        with self._lock:
            self.columns = [str(name) for name in result.get("columns") or []]
            self.rows = [list(row) for row in result.get("rows") or []]
            self.pages = [int(page) for page in result.get("page_numbers") or []]
            self.kind = str(result.get("kind") or "")
            self.strict = bool(result.get("strict", True))
            self.validation = list(result.get("validation") or [])
            self.provenance = dict(result.get("provenance") or {})
            self.suggestions = dict(result.get("suggestions") or {})
            self.raw_readings = dict(result.get("raw_readings") or {})
            self.baselines = dict(result.get("baselines") or {})
            self.column_spans = dict(result.get("column_spans") or {})
            self.figures = result.get("figures")
            self.print_date = read_print_date(pdf_bytes) if pdf_bytes else ""
            self.undo = []
            self.cursor = 0
            tiers = classify_cells(
                self.columns,
                self.rows,
                self.pages,
                self.print_date,
                self.provenance,
                self.suggestions,
                self.raw_readings,
                pdf_bytes,
                self.baselines,
                self.column_spans,
            )
            # Red first, then the quick yellow cells. Each item already has
            # its suggestion list from classify_cells.
            self.queue = list(tiers.get("red") or []) + list(tiers.get("yellow") or [])
            return len(self.queue)

    def current(self) -> dict | None:
        with self._lock:
            if self.cursor < 0 or self.cursor >= len(self.queue):
                return None
            return self.queue[self.cursor]

    def suggestions_for(self, item: dict | None) -> list[str]:
        """The list built during prepare. This does not recompute anything."""
        if not item:
            return []
        return [blank_text(value) for value in item.get("suggestions") or [] if blank_text(value)]

    def apply_item(self, item: dict, value: str) -> bool:
        """Write one cell. The PDF is not opened and the queue is not rebuilt."""
        cleaned = blank_text(value)
        if not item or not cleaned:
            return False
        with self._lock:
            row_index = int(item["row_index"])
            column = int(item["index"])
            if row_index < 0 or row_index >= len(self.rows) or column < 0:
                return False
            row = self.rows[row_index]
            while len(row) <= column:
                row.append("")
            before = blank_text(row[column])
            if before == cleaned:
                return False
            notes_at = next((index for index, name in enumerate(self.columns) if "ملاحظ" in str(name)), None)
            notes_before = blank_text(row[notes_at]) if notes_at is not None and notes_at < len(row) else ""
            row[column] = cleaned
            field = str(item.get("field") or "")
            if field:
                clear_field_note(self.columns, row, field)
            source_key = f"{int(item.get('page') or 0)}|{item.get('plate') or ''}|{field}"
            previous_source = self.provenance.get(source_key, "")
            if source_key.strip("|"):
                if source_key not in self.raw_readings and before:
                    self.raw_readings[source_key] = before
                self.provenance[source_key] = "manual"
            item["value"] = cleaned
            # Drop this cell from the queue. The following cell keeps its place.
            try:
                position = self.queue.index(item)
            except ValueError:
                position = self.cursor
            else:
                self.queue.pop(position)
                if self.cursor > position:
                    self.cursor -= 1
                self.cursor = min(self.cursor, max(0, len(self.queue) - 1)) if self.queue else 0
            self.undo.append(
                {
                    "item": item,
                    "position": position,
                    "before": before,
                    "notes_at": notes_at,
                    "notes_before": notes_before,
                    "provenance_before": previous_source,
                    "source_key": source_key,
                }
            )
            return True

    def skip(self) -> None:
        """Move to the next cell. The table is not read again."""
        with self._lock:
            if self.queue:
                self.cursor = (self.cursor + 1) % len(self.queue)

    def undo_last(self) -> bool:
        """Put back the last cell only."""
        with self._lock:
            if not self.undo:
                return False
            change = self.undo.pop()
            item = change["item"]
            row = self.rows[int(item["row_index"])]
            column = int(item["index"])
            while len(row) <= column:
                row.append("")
            row[column] = change["before"]
            item["value"] = change["before"]
            notes_at = change.get("notes_at")
            if notes_at is not None:
                while len(row) <= int(notes_at):
                    row.append("")
                row[int(notes_at)] = change.get("notes_before") or ""
            source_key = str(change.get("source_key") or "")
            previous = str(change.get("provenance_before") or "")
            if source_key:
                if previous:
                    self.provenance[source_key] = previous
                else:
                    self.provenance.pop(source_key, None)
            position = max(0, min(int(change["position"]), len(self.queue)))
            self.queue.insert(position, item)
            self.cursor = position
            return True

    def save(self, excel_path: str | Path) -> None:
        """Write the in-memory rows. This does not convert the PDF again."""
        from core.excel_converter import write_edited_excel

        with self._lock:
            columns = list(self.columns)
            rows = [list(row) for row in self.rows]
            pages = list(self.pages)
            kind = self.kind
            strict = self.strict
            validation = list(self.validation)
            provenance = dict(self.provenance)
            suggestions = dict(self.suggestions)
            raw_readings = dict(self.raw_readings)
            figures = self.figures
        write_edited_excel(
            columns,
            rows,
            pages,
            excel_path,
            kind=kind,
            figures=figures,
            strict=strict,
            validation=validation,
            provenance=provenance,
            suggestions=suggestions,
            raw_readings=raw_readings,
        )


def run_desktop() -> None:
    """Open the review window. Safe to call when no command-line path was given."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    from core.excel_converter import ConversionError, read_pdf_workbook

    root = tk.Tk()
    root.title("مراجعة الجدول")
    root.geometry("760x560")
    root.minsize(640, 480)

    session = ReviewSession()
    pdf_var = tk.StringVar()
    excel_var = tk.StringVar()
    status_var = tk.StringVar(value="اختر ملف PDF. المراجعة تشتغل على النسخة المحفوظة في الذاكرة.")
    detail_var = tk.StringVar(value="")
    cell_var = tk.StringVar(value="لا توجد خلية للمراجعة بعد.")
    # Set before the worker starts. Every button checks it and returns.
    busy = {"on": False}
    buttons: list[ttk.Button] = []

    frame = ttk.Frame(root, padding=16)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)

    def paint(message: str, detail: str = "") -> None:
        status_var.set(message)
        if detail:
            detail_var.set(detail)

    def lock(message: str) -> None:
        """Grey every action immediately, then let the caller start the worker."""
        busy["on"] = True
        for button in buttons:
            button.state(["disabled"])
        suggestion_list.configure(state="disabled")
        value_entry.state(["disabled"])
        progress.configure(mode="indeterminate")
        progress.start(12)
        paint(message)

    def unlock(message: str, detail: str = "") -> None:
        busy["on"] = False
        progress.stop()
        progress.configure(mode="determinate", value=0)
        for button in buttons:
            button.state(["!disabled"])
        suggestion_list.configure(state="normal")
        value_entry.state(["!disabled"])
        paint(message, detail)
        show_current()

    def show_current() -> None:
        item = session.current()
        suggestion_list.delete(0, "end")
        if item is None:
            cell_var.set("لا توجد خلية معلّقة." if session.rows else "حوّل الملف أولاً.")
            return
        label = _FIELD_LABELS.get(str(item.get("field") or ""), item.get("field"))
        cell_var.set(
            f"خلية {session.cursor + 1} من {len(session.queue)}  ·  "
            f"صفحة {item.get('page')}  ·  لوحة {item.get('plate')}  ·  {label}\n"
            f"القيمة: {blank_text(item.get('value')) or 'فارغة'}  ·  {item.get('reason')}"
        )
        for idea in session.suggestions_for(item):
            suggestion_list.insert("end", idea)

    def chosen_value() -> str:
        picked = suggestion_list.curselection()
        if picked:
            return str(suggestion_list.get(picked[0]))
        return value_entry.get()

    def choose_pdf() -> None:
        if busy["on"]:
            return
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf"), ("All files", "*.*")])
        if not path:
            return
        pdf_var.set(path)
        if not excel_var.get():
            excel_var.set(str(Path(path).with_suffix(".xlsx")))

    def choose_excel() -> None:
        if busy["on"]:
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            excel_var.set(path)

    def start_convert() -> None:
        if busy["on"]:
            return
        if not pdf_var.get() or not excel_var.get():
            messagebox.showwarning("مراجعة الجدول", "اختر ملف PDF وملف الإكسل.")
            return
        pdf_path = pdf_var.get()
        excel_path = excel_var.get()
        lock(BUSY_CONVERT)

        def report(message: str) -> None:
            # The wait line stays. The stage name is the second line.
            root.after(0, detail_var.set, message)

        def work() -> None:
            try:
                pdf_bytes = Path(pdf_path).read_bytes()
                result = read_pdf_workbook(pdf_path, progress=report, strict=True)
                count = session.prepare(result, pdf_bytes)
                session.save(excel_path)
            except ConversionError as exc:
                root.after(0, lambda: unlock(str(exc)))
                return
            except Exception as exc:
                root.after(0, lambda message=str(exc): unlock(f"تعذر التحويل. {message}"))
                return
            root.after(0, lambda: unlock(DONE_CONVERT, f"{count} خلية للمراجعة. الملف محفوظ، والتعديلات التالية في الذاكرة."))

        threading.Thread(target=work, daemon=True).start()

    def start_apply() -> None:
        if busy["on"]:
            return
        item = session.current()
        value = chosen_value()
        if item is None or not blank_text(value):
            paint("اختر قيمة من القائمة أو اكتبها.")
            return
        # Freeze the cell now. The worker must not follow a cursor that moved.
        target = item
        lock(BUSY_EDIT)

        def work() -> None:
            try:
                changed = session.apply_item(target, value)
            except Exception as exc:
                root.after(0, lambda message=str(exc): unlock(f"تعذر التعديل. {message}"))
                return
            left = len(session.queue)
            text = DONE_EDIT if changed else "القيمة نفسها، لم يتغير شيء"
            root.after(0, lambda: unlock(text, f"باقي {left} خلية. الملف يُحفظ عند الضغط على حفظ."))

        threading.Thread(target=work, daemon=True).start()

    def start_skip() -> None:
        if busy["on"]:
            return
        session.skip()
        show_current()

    def start_undo() -> None:
        if busy["on"]:
            return
        lock(BUSY_EDIT)

        def work() -> None:
            try:
                restored = session.undo_last()
            except Exception as exc:
                root.after(0, lambda message=str(exc): unlock(f"تعذر التراجع. {message}"))
                return
            text = "تم التراجع" if restored else "لا يوجد تعديل للتراجع"
            root.after(0, lambda: unlock(text))

        threading.Thread(target=work, daemon=True).start()

    def start_save() -> None:
        if busy["on"]:
            return
        if not session.rows or not excel_var.get():
            paint("حوّل الملف أولاً.")
            return
        excel_path = excel_var.get()
        lock(BUSY_SAVE)

        def work() -> None:
            try:
                session.save(excel_path)
            except Exception as exc:
                root.after(0, lambda message=str(exc): unlock(f"تعذر الحفظ. {message}"))
                return
            root.after(0, lambda: unlock(DONE_SAVE, excel_path))

        threading.Thread(target=work, daemon=True).start()

    ttk.Label(frame, text="PDF").grid(row=0, column=0, sticky="w", pady=4)
    ttk.Entry(frame, textvariable=pdf_var).grid(row=0, column=1, sticky="ew", padx=8)
    browse_pdf = ttk.Button(frame, text="اختيار", command=choose_pdf)
    browse_pdf.grid(row=0, column=2)
    ttk.Label(frame, text="Excel").grid(row=1, column=0, sticky="w", pady=4)
    ttk.Entry(frame, textvariable=excel_var).grid(row=1, column=1, sticky="ew", padx=8)
    browse_excel = ttk.Button(frame, text="حفظ باسم", command=choose_excel)
    browse_excel.grid(row=1, column=2)

    convert_button = ttk.Button(frame, text="تحويل", command=start_convert)
    convert_button.grid(row=2, column=1, sticky="w", pady=8)
    progress = ttk.Progressbar(frame, mode="indeterminate")
    progress.grid(row=2, column=1, sticky="e", pady=8)
    ttk.Label(frame, textvariable=status_var, wraplength=680).grid(row=3, column=0, columnspan=3, sticky="w")
    ttk.Label(frame, textvariable=detail_var, wraplength=680).grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 8))

    ttk.Label(frame, textvariable=cell_var, wraplength=680, justify="right").grid(
        row=5, column=0, columnspan=3, sticky="ew", pady=6
    )
    suggestion_list = tk.Listbox(frame, height=6, exportselection=False)
    suggestion_list.grid(row=6, column=0, columnspan=3, sticky="ew")
    value_entry = ttk.Entry(frame)
    value_entry.grid(row=7, column=0, columnspan=3, sticky="ew", pady=8)

    actions = ttk.Frame(frame)
    actions.grid(row=8, column=0, columnspan=3, sticky="w")
    apply_button = ttk.Button(actions, text="تطبيق", command=start_apply)
    skip_button = ttk.Button(actions, text="تجاوز", command=start_skip)
    undo_button = ttk.Button(actions, text="تراجع", command=start_undo)
    save_button = ttk.Button(actions, text="حفظ الملف", command=start_save)
    for button in (apply_button, skip_button, undo_button, save_button):
        button.pack(side="left", padx=(0, 8))
    buttons.extend([browse_pdf, browse_excel, convert_button, apply_button, skip_button, undo_button, save_button])

    root.mainloop()
