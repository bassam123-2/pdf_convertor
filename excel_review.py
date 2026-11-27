# -*- coding: utf-8 -*-
"""Review rules for the Excel screen. Nothing here writes the workbook."""

from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime
from pathlib import Path

import fitz

from page_images import row_band
from core.excel_converter import (
    _DATE_VALUE,
    _MANUFACTURERS,
    _OFFICIAL_COLORS,
    _YEAR_VALUE,
    _footer_rule_y,
    _normalize_date_cell,
    _ocr_clip_words,
    _vehicle_field_at,
)

REVIEW_FIELDS = ("date", "make", "color", "year", "admin")
_REVIEWED = "تمت المراجعة"
_FLAG_FOR_FIELD = {
    "year": ("سنة مفقودة", "سنة غير صالحة", "تعارض بين اللوحة والسنة"),
    "date": ("تاريخ مفقود", "تاريخ ناقص", "تاريخ غير مقروء", "تعارض في التاريخ"),
    "make": ("صنع مفقود", "صنع غير معروف"),
    "color": ("لون مفقود", "لون غير معروف"),
    "admin": ("رقم إداري مفقود",),
    "plate": ("لوحة مفقودة أو غير صالحة",),
}
_EMPTY_WORDS = {"none", "nan", "nat", "<na>", "null"}
_PRINT_DATES: dict[str, str] = {}


def blank_text(value: object) -> str:
    """Empty cells stay blank. The word None is not a value."""
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in _EMPTY_WORDS:
        return ""
    return text


def print_date_from_text(text: str) -> str:
    """
    The report date printed beside the footer clock.

    A vehicle date higher on the page is ignored when the clock is present.
    """
    compact = re.sub(r"\s+", "", text or "")
    compact = compact.replace("–", "-").replace("—", "-").replace("/", "-")
    beside_clock = re.search(
        r"\d{2}[.:]\d{2}(?:[.:]\d{2})?((?:19|20)\d{2}-\d{2}-\d{2})",
        compact,
    )
    if beside_clock and _DATE_VALUE.fullmatch(beside_clock.group(1)):
        return beside_clock.group(1)
    found = [
        item
        for item in re.findall(r"(?:19|20)\d{2}-\d{2}-\d{2}", compact)
        if _DATE_VALUE.fullmatch(item)
    ]
    if len(found) == 1:
        return found[0]
    return ""


def read_print_date(pdf_bytes: bytes, password: str = "") -> str:
    """Read the footer date once per file. Later calls reuse that result."""
    if not pdf_bytes:
        return ""
    digest = hashlib.sha1(pdf_bytes).hexdigest()
    if digest in _PRINT_DATES:
        return _PRINT_DATES[digest]
    found = _ocr_footer_date(pdf_bytes, password)
    _PRINT_DATES[digest] = found
    return found


def _ocr_footer_date(pdf_bytes: bytes, password: str) -> str:
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        if document.needs_pass and not document.authenticate(password or ""):
            return ""
        if document.page_count < 1:
            return ""
        page = document.load_page(0)
        rule = _footer_rule_y(page)
        top = float(rule) if rule else max(0.0, float(page.rect.height) - 140.0)
        clip = fitz.Rect(page.rect.x0, top, page.rect.x1, page.rect.y1)
        words = _ocr_clip_words(page, clip, "en-US")
        return print_date_from_text(" ".join(word.text for word in words))
    except Exception:
        return ""
    finally:
        document.close()


def _excluded_row(columns: list[str], row: list[str]) -> bool:
    for index, name in enumerate(columns):
        if "ملاحظ" in str(name) and index < len(row) and blank_text(row[index]) in {"ترويسة", "فاصل"}:
            return True
    return False


def _cell(row: list[str], index: int | None) -> str:
    if index is None or index >= len(row):
        return ""
    return blank_text(row[index])


def _date_reason(iso: str, raw: str, print_date: str, vehicle_year: str) -> str:
    if not raw:
        return "فارغ"
    if not _DATE_VALUE.fullmatch(iso or ""):
        return "تاريخ ناقص أو غير مقروء"
    reasons: list[str] = []
    if print_date and iso < print_date:
        reasons.append(f"أقدم من تاريخ طباعة الكشف {print_date}")
    if vehicle_year and iso < f"{vehicle_year}-01-01":
        reasons.append(f"أقدم من سنة المركبة {vehicle_year}")
    return " و ".join(reasons)


def _print_year(print_date: str) -> int:
    text = (print_date or "")[:4]
    return int(text) if text.isdigit() else 0


def _field_reason(field: str, value: str, print_date: str, vehicle_year: str) -> str:
    if field == "date":
        iso = _normalize_date_cell(value) if value else ""
        return _date_reason(iso, value, print_date, vehicle_year)
    if field == "admin":
        return "" if value else "فارغ"
    if not value:
        return "فارغ"
    if field == "make" and value not in _MANUFACTURERS:
        return "قيمة غير معروفة"
    if field == "color" and value not in _OFFICIAL_COLORS:
        return "قيمة غير معروفة"
    if field == "year":
        if not _YEAR_VALUE.fullmatch(value):
            return "قيمة غير معروفة"
        limit = _print_year(print_date)
        if limit and int(value) > limit:
            return f"أكبر من سنة الكشف {limit}"
    return ""


def clear_field_note(columns: list[str], row: list[str], field: str) -> str:
    """Drop the flags for one edited field. A row with none left is marked reviewed."""
    notes_at = next((index for index, name in enumerate(columns) if "ملاحظ" in str(name)), None)
    if notes_at is None:
        return ""
    while len(row) <= notes_at:
        row.append("")
    note = blank_text(row[notes_at])
    if note in {"ترويسة", "فاصل"}:
        return note
    drop = set(_FLAG_FOR_FIELD.get(field, ()))
    kept: list[str] = []
    for part in re.split(r"[؛;]", note):
        text = part.strip()
        if not text or text in drop:
            continue
        if field == "year" and text.startswith("سنة أكبر"):
            continue
        if text.startswith("محاولة") or text.startswith("اقتراح") or "->" in text:
            continue
        kept.append(text)
    known = {flag for flags in _FLAG_FOR_FIELD.values() for flag in flags}
    if any(part in known or part.startswith("سنة أكبر") for part in kept):
        row[notes_at] = "؛ ".join(kept)
    else:
        row[notes_at] = _REVIEWED
    return str(row[notes_at])


def _plausible_date(iso: str, print_date: str, vehicle_year: str) -> bool:
    if not _DATE_VALUE.fullmatch(iso or ""):
        return False
    if print_date and iso < print_date:
        return False
    if vehicle_year and iso < f"{vehicle_year}-01-01":
        return False
    return True


def _neighbour_values(
    rows: list[list[str]],
    pages: list[int],
    row_index: int,
    column: int,
    purpose_at: int | None = None,
    limit: int = 8,
) -> list[str]:
    """Values from the same page and the same licence group. Other pages are not used."""
    page = int(pages[row_index]) if row_index < len(pages) else -1
    purpose = _cell(rows[row_index], purpose_at) if purpose_at is not None else ""
    found: list[str] = []
    for step in range(1, limit + 1):
        for cursor in (row_index - step, row_index + step):
            if cursor < 0 or cursor >= len(rows):
                continue
            if cursor < len(pages) and int(pages[cursor]) != page:
                continue
            if purpose_at is not None and purpose:
                other = _cell(rows[cursor], purpose_at)
                if other and other != purpose:
                    continue
            text = _cell(rows[cursor], column)
            if text:
                found.append(text)
    return found


def _suggestions(
    field: str,
    current: str,
    neighbours: list[str],
    print_date: str,
    vehicle_year: str,
    evidence: list[tuple[str, str]] | None = None,
) -> tuple[list[str], list[str]]:
    """Ranked values for this row only, each with the reason it was offered."""
    ideas: list[str] = []
    reasons: list[str] = []

    def add(value: str, reason: str) -> None:
        text = blank_text(value)
        if not text or text == current or text in ideas:
            return
        if field == "date":
            iso = _normalize_date_cell(text)
            if not _plausible_date(iso, print_date, vehicle_year):
                return
            text = iso
            if text == current or text in ideas:
                return
        elif field == "make" and text not in _MANUFACTURERS:
            return
        elif field == "color" and text not in _OFFICIAL_COLORS:
            return
        elif field == "year":
            if not _YEAR_VALUE.fullmatch(text):
                return
            limit = _print_year(print_date)
            if limit and int(text) > limit:
                return
        elif field == "admin" and not re.fullmatch(r"\d{1,6}", text):
            return
        ideas.append(text)
        reasons.append(reason)

    for value, reason in evidence or []:
        add(value, reason or "قراءة المقطع")
        if len(ideas) >= 4:
            return ideas, reasons
    if field == "date":
        counts: dict[str, int] = {}
        for value in neighbours:
            iso = _normalize_date_cell(value)
            if _plausible_date(iso, print_date, vehicle_year):
                counts[iso] = counts.get(iso, 0) + 1
        ranked = sorted(counts, key=lambda item: (-counts[item], item))
        for value in ranked[:1]:
            add(value, "جوار نفس الصفحة والمجموعة")
        return ideas[:4], reasons[:4]
    if field == "admin":
        counts = {}
        for value in neighbours:
            if re.fullmatch(r"\d{1,6}", value):
                counts[value] = counts.get(value, 0) + 1
        if counts:
            add(max(counts, key=lambda item: (counts[item], item)), "جوار نفس الصفحة والمجموعة")
        return ideas[:1], reasons[:1]
    counts = {}
    for value in neighbours:
        counts[value] = counts.get(value, 0) + 1
    ranked = sorted(counts, key=lambda item: (-counts[item], item))
    for value in ranked:
        add(value, "جوار نفس الصفحة والمجموعة")
        if len(ideas) >= 4:
            break
    return ideas[:4], reasons[:4]


def suggestion_kind(reason: str) -> str:
    """agreed, clip, neighbour, or other. Only two agreeing readings are a batch."""
    text = str(reason or "").strip()
    if "اتفاق المحركين" in text:
        return "agreed"
    if text == "قراءة المقطع" or text.startswith("قراءة المقطع"):
        return "clip"
    if "جوار" in text:
        return "neighbour"
    return "other"


def _value_counts(rows: list[list[str]], pages: list[int], columns: list[str]) -> dict[tuple[int, str], dict[str, int]]:
    """How often each colour, make, and date appears on one page."""
    fields = _vehicle_field_at(columns)
    counts: dict[tuple[int, str], dict[str, int]] = {}
    for index, row in enumerate(rows):
        page = int(pages[index]) if index < len(pages) else 0
        for field in ("color", "make", "date"):
            column = fields.get(field)
            if column is None or column >= len(row):
                continue
            value = blank_text(row[column]).replace("\u200e", "")
            if not value:
                continue
            bucket = counts.setdefault((page, field), {})
            bucket[value] = bucket.get(value, 0) + 1
    return counts


def outlier_reason(field: str, value: str, page: int, counts: dict[tuple[int, str], dict[str, int]]) -> str:
    """A value that is not the page's common one, or a date that barely appears."""
    text = blank_text(value).replace("\u200e", "")
    bucket = counts.get((int(page), field)) or {}
    if not text or not bucket or field not in {"color", "make", "date"}:
        return ""
    dominant = max(bucket, key=lambda key: (bucket[key], key))
    if bucket[dominant] <= bucket.get(text, 0):
        return ""
    if field in {"color", "make"} and text != dominant:
        return f"يختلف عن السائد في الصفحة ({dominant})"
    if field == "date" and text != dominant and bucket.get(text, 0) <= 2:
        return f"تاريخ نادر {text}"
    return ""


# A value seen this rarely on its page, and not the page's common value, is mandatory.
# Colours are allowed two copies: a single odd colour and its pair still get a look.
# Makes and dates are mandatory at one copy. These numbers are not tied to one file.
_COLOR_RARE = 2
_MAKE_RARE = 1
_DATE_RARE = 1
# An empty colour whose date is this uncommon, and rarer than both neighbours, is mandatory.
_EMPTY_COLOR_DATE = 5
_AMBER_EVERY = 20


def _page_modes(rows, pages, columns):
    """Dominant colour, make, and date inside one licence group, across pages."""
    fields = _vehicle_field_at(columns)
    counts: dict[tuple[str, str], dict[str, int]] = {}
    plated: list[int] = []
    plate_at = fields.get("plate")
    purpose_at = fields.get("purpose")
    for index, row in enumerate(rows):
        if _excluded_row(columns, row):
            continue
        plate = _cell(row, plate_at).replace("\u200e", "")
        if not plate:
            continue
        plated.append(index)
        purpose = _cell(row, purpose_at) if purpose_at is not None else ""
        for field in ("color", "make", "date"):
            column = fields.get(field)
            value = _cell(row, column).replace("\u200e", "")
            if not value:
                continue
            bucket = counts.setdefault((purpose, field), {})
            bucket[value] = bucket.get(value, 0) + 1
    modes = {}
    for key, bucket in counts.items():
        modes[key] = max(bucket, key=lambda name: (bucket[name], name))
    return fields, counts, modes, plated


def _reading_names(suggestions: dict, raw_readings: dict, page, plate, field) -> list[str]:
    key = f"{int(page)}|{plate}|{field}"
    names = []
    for part in re.split(r"[؛;]", str((suggestions or {}).get(key) or "")):
        piece = part.strip()
        if not piece:
            continue
        names.append(piece.split(" — ", 1)[0].strip())
    raw = blank_text((raw_readings or {}).get(key))
    if raw:
        names.append(raw)
    return [name for name in names if name]


# Only a reading that already looks consistent can vote a crop out.
# A weak group (the crops do not agree with each other) is left to the text rules.
_VISUAL_MEDIAN = 0.45
_VISUAL_FLOOR = 0.12
_VISUAL_GAP = 0.30
_VISUAL_GROUP = 8
# Unclear or empty text is a quick pick. It is not the mandatory card.
_YELLOW_REASONS = {"لون غير واضح", "صنع غير واضح", "تاريخ ناقص أو غير مقروء", "فارغة"}


def dissimilar_keys(
    records: list[tuple],
    floor: float = _VISUAL_FLOOR,
    gap: float = _VISUAL_GAP,
    minimum: int = _VISUAL_GROUP,
    median_floor: float = _VISUAL_MEDIAN,
) -> list[tuple]:
    """
    records are (page, plate, field, text, image).

    Images that share one reading are compared with normalised cross-correlation.
    A crop that does not match the rest of its group is returned. Nothing is sent off the machine.
    """
    import numpy as np

    groups: dict[tuple, list] = {}
    for record in records:
        # Same page and same reading. A word that only disagrees with other pages is a different scan, not a bad cell.
        groups.setdefault((record[0], record[2], record[3]), []).append(record)

    def score(image, template):
        left = image - image.mean()
        right = template - template.mean()
        denom = float(np.sqrt((left * left).sum() * (right * right).sum())) + 1e-9
        return float((left * right).sum() / denom)

    found = []
    for items in groups.values():
        if len(items) < minimum:
            continue
        images = np.stack([item[4] for item in items])
        first = np.array([score(image, images.mean(0)) for image in images])
        template = images[first >= np.median(first)].mean(0)
        scores = np.array([score(image, template) for image in images])
        typical = float(np.median(scores))
        if typical < median_floor:
            continue
        for item, value in zip(items, scores):
            if float(value) < floor and float(value) <= typical - gap:
                found.append((item[0], item[1], item[2]))
    return found


def _binarised_crop(image, x0: float, x1: float, y0: float, y1: float, scale: float, keep_blank: bool = False):
    """One cell, trimmed to its ink and resized. The year printed beside a colour is cut off."""
    import numpy as np

    height, width = image.shape
    left = max(0, int(x0 * scale))
    right = min(width, int(x1 * scale))
    top = max(0, int(y0 * scale))
    bottom = min(height, int(y1 * scale))
    if right - left < 8 or bottom - top < 6:
        return None
    chip = image[top:bottom, left:right]
    ink = chip < 150
    if float(ink.mean()) < 0.03:
        if keep_blank:
            return np.zeros((16, 32), dtype="float32")
        return None
    columns = ink.sum(0)
    gap = columns <= 1
    start = None
    for index, flag in enumerate(list(gap) + [False]):
        if flag and start is None:
            start = index
        elif not flag and start is not None:
            if index - start >= 6 and start > 14 and len(columns) - index > 8:
                chip = chip[:, :start]
                ink = chip < 150
                break
            start = None
    ys = np.where(ink.any(1))[0]
    xs = np.where(ink.any(0))[0]
    if len(xs) < 4 or len(ys) < 2:
        return None
    chip = chip[ys[0] : ys[-1] + 1, xs[0] : xs[-1] + 1]
    yy = np.linspace(0, chip.shape[0] - 1, 16).astype(int)
    xx = np.linspace(0, chip.shape[1] - 1, 32).astype(int)
    return (chip[yy][:, xx] < 150).astype("float32")


def _crop_ncc(image, template) -> float:
    import numpy as np

    left = np.asarray(image, dtype="float32")
    right = np.asarray(template, dtype="float32")
    left = left - left.mean()
    right = right - right.mean()
    denom = float(np.sqrt((left * left).sum() * (right * right).sum())) + 1e-9
    return float((left * right).sum() / denom)


def _crop_window(field: str, focus: float, spans: dict) -> tuple[float, float, float, float] | None:
    """The same windows the visual check uses. Make uses its own column."""
    color_box = spans.get("color")
    make_box = spans.get("make")
    if field == "color" and color_box:
        return (float(color_box[0]), min(float(color_box[0]) + 56.0, float(color_box[1])), focus - 8.0, focus - 1.0)
    if field == "make" and make_box:
        return (float(make_box[0]), min(float(make_box[0]) + 72.0, float(make_box[1])), focus - 8.0, focus - 1.0)
    if field == "date" and color_box:
        return (8.0, max(16.0, float(color_box[0]) - 8.0), focus - 3.6, focus + 2.8)
    return None


def _visual_records(
    pdf_bytes: bytes,
    columns: list[str],
    rows: list[list[str]],
    pages: list[int],
    baselines: dict,
    column_spans: dict,
    wanted: set[tuple[str, str, str]] | None = None,
    keep_blank: bool = False,
) -> list[tuple]:
    """(page, plate, field, text, image) for colour, make, and date crops."""
    if not pdf_bytes or not baselines or not column_spans:
        return []
    import numpy as np

    fields = _vehicle_field_at(columns)
    scale = 3.0
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    rendered: dict[int, np.ndarray] = {}
    records = []
    try:
        def page_image(number: int):
            if number not in rendered:
                if number < 1 or number > document.page_count:
                    return None
                page = document.load_page(number - 1)
                pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                rendered[number] = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width, 3).mean(axis=2)
            return rendered[number]

        for index, row in enumerate(rows):
            if _excluded_row(columns, row):
                continue
            plate = _cell(row, fields.get("plate")).replace("\u200e", "")
            if not plate:
                continue
            page_no = int(pages[index]) if index < len(pages) else 0
            focus = float((baselines or {}).get(f"{page_no}|{plate}") or 0)
            spans = (column_spans or {}).get(str(page_no)) or (column_spans or {}).get(page_no) or {}
            if focus <= 0 or not spans:
                continue
            image = page_image(page_no)
            if image is None:
                continue
            for field in ("color", "make", "date"):
                if wanted is not None and (str(page_no), plate, field) not in wanted:
                    continue
                window = _crop_window(field, focus, spans)
                if window is None:
                    continue
                text = _cell(row, fields.get(field)).replace("\u200e", "")
                if wanted is None and field == "make":
                    continue
                if wanted is None and not text:
                    continue
                crop = _binarised_crop(image, window[0], window[1], window[2], window[3], scale, keep_blank=keep_blank)
                if crop is not None:
                    records.append((page_no, plate, field, text, crop))
        return records
    except Exception:
        return records
    finally:
        document.close()


def visual_outliers(
    pdf_bytes: bytes,
    columns: list[str],
    rows: list[list[str]],
    pages: list[int],
    baselines: dict,
    column_spans: dict,
) -> list[tuple]:
    """Crops that do not look like the other cells read as the same text."""
    found, _note = visual_scan(pdf_bytes, columns, rows, pages, baselines, column_spans)
    return found


def visual_scan(
    pdf_bytes: bytes,
    columns: list[str],
    rows: list[list[str]],
    pages: list[int],
    baselines: dict,
    column_spans: dict,
) -> tuple[list[tuple], str]:
    """
    Outliers, and a note when no reading is repeated often enough to compare.

    A reading with fewer than 8 crops is compared with the learned crop library
    when that library has a confirmed example. The within-page thresholds stay put.
    """
    records = _visual_records(pdf_bytes, columns, rows, pages, baselines, column_spans)
    if not records:
        return [], ""
    sizes: dict[tuple, int] = {}
    for record in records:
        key = (record[0], record[2], record[3])
        sizes[key] = sizes.get(key, 0) + 1
    found = list(dissimilar_keys(records))
    found.extend(_library_flags(records, sizes))
    return found, visual_note_for(sizes)


def visual_note_for(sizes: dict) -> str:
    """Said when no reading has enough copies for a within-document comparison."""
    if sizes and max(sizes.values()) < _VISUAL_GROUP:
        return "document too small for visual comparison"
    return ""


def crop_library_dir() -> Path:
    """Confirmed crops live on this machine. They are not the correction dictionary."""
    override = os.environ.get("PDF_TOOL_CROPS", "").strip()
    if override:
        return Path(override)
    try:
        from engine.paths import crops_dir, using_app_data

        if using_app_data():
            return crops_dir()
    except Exception:
        pass
    return Path(__file__).resolve().parent / "learned_crops"


def _library_index_path() -> Path:
    return crop_library_dir() / "index.json"


def _read_library_index() -> list[dict]:
    path = _library_index_path()
    if not path.is_file():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return [item for item in payload if isinstance(item, dict)] if isinstance(payload, list) else []


def _write_library_index(items: list[dict]) -> None:
    root = crop_library_dir()
    root.mkdir(parents=True, exist_ok=True)
    _library_index_path().write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")


def remember_crop(field: str, value: str, image) -> None:
    """Store one confirmed crop with the value the user accepted."""
    text = blank_text(value).replace("\u200e", "")
    if field not in {"color", "make", "date"} or not text or image is None:
        return
    import numpy as np
    from PIL import Image

    array = (np.asarray(image, dtype="float32") > 0.5).astype("uint8")
    if array.shape != (16, 32) or int(array.max()) == 0:
        return
    digest = hashlib.sha1(array.tobytes()).hexdigest()[:16]
    root = crop_library_dir()
    root.mkdir(parents=True, exist_ok=True)
    name = f"{digest}.png"
    Image.fromarray(array * 255, "L").save(root / name)
    items = _read_library_index()
    if any(item.get("file") == name and item.get("field") == field and item.get("value") == text for item in items):
        return
    items.append({"field": field, "value": text, "file": name})
    _write_library_index(items)


def crop_library_count() -> int:
    return len(_read_library_index())


def load_library_images(field: str | None = None) -> list[tuple[str, str, object]]:
    """(field, value, 16×32 ink array) for confirmed crops."""
    import numpy as np
    from PIL import Image

    found = []
    root = crop_library_dir()
    for item in _read_library_index():
        if field and item.get("field") != field:
            continue
        path = root / str(item.get("file") or "")
        if not path.is_file():
            continue
        try:
            array = np.asarray(Image.open(path).convert("L"), dtype="float32") / 255.0
        except OSError:
            continue
        if array.shape != (16, 32):
            continue
        found.append((str(item.get("field") or ""), str(item.get("value") or ""), array))
    return found


def clear_crop_library() -> None:
    root = crop_library_dir()
    if not root.is_dir():
        return
    for path in root.glob("*"):
        if path.is_file():
            path.unlink()


def export_crop_library() -> bytes:
    """A zip of the index and the crops. Nothing is uploaded."""
    import io
    import zipfile

    buffer = io.BytesIO()
    root = crop_library_dir()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("index.json", json.dumps(_read_library_index(), ensure_ascii=False, indent=2))
        written: set[str] = set()
        for item in _read_library_index():
            path = root / str(item.get("file") or "")
            # One crop can back several index entries; it goes in once.
            if path.is_file() and path.name not in written:
                written.add(path.name)
                archive.write(path, arcname=path.name)
    return buffer.getvalue()


def _library_flags(records: list[tuple], sizes: dict[tuple, int]) -> list[tuple]:
    """
    A crop whose reading was seen fewer than 8 times, matched to a confirmed crop.

    The match has to clear the same median used inside a page. An empty library adds nothing.
    """
    try:
        library = load_library_images()
    except Exception:
        return []
    if not library:
        return []
    import numpy as np

    grouped: dict[tuple[str, str], list] = {}
    for field, value, image in library:
        grouped.setdefault((field, value), []).append(image)
    means = {key: np.mean(np.stack(images), axis=0) for key, images in grouped.items() if images}
    found = []
    seen: set[tuple] = set()
    for record in records:
        key = (record[0], record[2], record[3])
        if sizes.get(key, 0) >= _VISUAL_GROUP:
            continue
        best_value = ""
        best_score = -1.0
        for (lib_field, lib_value), template in means.items():
            if lib_field != record[2]:
                continue
            score = _crop_ncc(record[4], template)
            if score > best_score:
                best_score = score
                best_value = lib_value
        if best_value and best_value != record[3] and best_score >= _VISUAL_MEDIAN:
            mark = (record[0], record[1], record[2])
            if mark not in seen:
                seen.add(mark)
                found.append(mark)
    return found


def _cluster_label(texts: list[str]) -> str:
    counts: dict[str, int] = {}
    for text in texts:
        if text:
            counts[text] = counts.get(text, 0) + 1
    if not counts:
        return "فارغة"
    return max(counts, key=lambda name: (counts[name], name))


def _grow_clusters(items: list[dict], join: float) -> list[list[dict]]:
    """Greedy groups. A crop joins the group only when it matches that group's shape."""
    pending = list(items)
    groups: list[list[dict]] = []
    while pending:
        if len(pending) == 1:
            groups.append(pending)
            break
        best_at = 0
        best_near = -1
        for index, item in enumerate(pending):
            near = sum(
                1
                for other in pending
                if other is not item and _crop_ncc(item["image"], other["image"]) >= join
            )
            if near > best_near:
                best_near = near
                best_at = index
        if best_near < 1:
            groups.extend([[item] for item in pending])
            break
        seed = pending.pop(best_at)
        members = [seed]
        import numpy as np

        template = np.asarray(seed["image"], dtype="float32")
        grew = True
        while grew:
            grew = False
            stay = []
            for item in pending:
                if _crop_ncc(item["image"], template) >= join:
                    members.append(item)
                    grew = True
                else:
                    stay.append(item)
            pending = stay
            if grew:
                template = np.mean(np.stack([member["image"] for member in members]), axis=0)
        groups.append(members)
    return groups


def item_identity(item: dict) -> str:
    return f"{item.get('page')}|{item.get('plate')}|{item.get('field')}"


def _same_reading_groups(members: list[dict]) -> list[list[dict]]:
    """Visual groups never mix two OCR readings. A single crop stays out."""
    by_text: dict[str, list[dict]] = {}
    for member in members:
        by_text.setdefault(member["text"], []).append(member)
    groups: list[list[dict]] = []
    for bucket in by_text.values():
        blanks = [member for member in bucket if member["blank"]]
        inked = [member for member in bucket if not member["blank"]]
        if len(blanks) >= 2:
            groups.append(blanks)
        groups.extend(group for group in _grow_clusters(inked, _VISUAL_MEDIAN) if len(group) >= 2)
    return groups


def _attach_outliers(group: list[dict], column_members: list[dict], claimed: set[int]) -> list[dict]:
    """A crop that looks like the group but was read as another word. It is not accepted with the group."""
    import numpy as np

    template = np.mean(np.stack([member["image"] for member in group]), axis=0)
    label = group[0]["text"]
    outliers = []
    for member in column_members:
        if id(member) in claimed or member["text"] == label:
            continue
        if _crop_ncc(member["image"], template) < _VISUAL_MEDIAN:
            continue
        outliers.append(member["item"])
        claimed.add(id(member))
    return outliers


def cluster_review_items(items: list[dict], crops: dict[tuple[str, str, str], object]) -> list[dict]:
    """
    One group per page, column, and OCR reading when the crops look alike.

    A different reading is listed on outliers and is not part of the accept set.
    A lone crop stays out of the groups.
    """
    by_column: dict[tuple[str, str], list[dict]] = {}
    for item in items:
        field = str(item.get("field") or "")
        if field not in {"color", "make", "date"}:
            continue
        key = (str(item.get("page")), str(item.get("plate")), field)
        image = crops.get(key)
        if image is None:
            continue
        import numpy as np

        array = np.asarray(image, dtype="float32")
        by_column.setdefault((str(item.get("page")), field), []).append(
            {
                "item": item,
                "image": array,
                "text": blank_text(item.get("value")).replace("\u200e", ""),
                "blank": float(array.max()) < 0.01,
            }
        )
    clusters = []
    for (page, field), members in by_column.items():
        groups = _same_reading_groups(members)
        claimed = {id(member) for group in groups for member in group}
        for number, group in enumerate(groups):
            chosen = [member["item"] for member in group]
            label = _cluster_label([member["text"] for member in group])
            clusters.append(
                {
                    "id": f"{page}|{field}|{label}|{number}|{chosen[0].get('plate')}",
                    "page": int(page) if str(page).isdigit() else page,
                    "field": field,
                    "label": label,
                    "members": chosen,
                    "outliers": _attach_outliers(group, members, claimed),
                    "images": [member["image"] for member in group],
                }
            )
    clusters.sort(key=lambda cluster: (int(cluster["page"]), cluster["field"], -len(cluster["members"])))
    return clusters


def cluster_plan(clusters: list[dict], items: list[dict]) -> dict:
    """One click per cluster, plus one card for every cell that did not join one."""
    covered = {
        f"{member['page']}|{member['plate']}|{member['field']}"
        for cluster in clusters
        for member in cluster["members"]
    }
    leftovers = [
        item for item in items if f"{item['page']}|{item['plate']}|{item['field']}" not in covered
    ]
    red_left = [item for item in leftovers if item.get("tier") == "red"]
    return {
        "clusters": len(clusters),
        "expected_clicks": len(clusters) + len(leftovers),
        "leftover_cells": len(leftovers),
        "red_cells_after": len(red_left),
        "red_rows_after": len({(item["page"], item["plate"]) for item in red_left}),
    }


def cluster_decision(cluster: dict, ejected: set[str], scanned: bool) -> dict:
    """What the accept button may write. A different OCR reading is never in that list."""
    members = [item for item in cluster.get("members") or [] if item_identity(item) not in ejected]
    outliers = [item for item in cluster.get("outliers") or [] if item_identity(item) not in ejected]
    ready = bool(members) and ((not outliers) or bool(scanned))
    label = str(cluster.get("label") or "فارغة")
    return {
        "ready": ready,
        "label": f"نعم، الـ {len(members)} المتبقية كلها {label}",
        "count": len(members),
        "write": list(members) if ready else [],
        "outliers": outliers,
    }


def remaining_cells(total: int, done: list[str] | set[str], accepted: list[str] | set[str]) -> int:
    """Open cells after this accept. The drop equals the number of new identities."""
    return max(0, int(total) - len(set(done) | set(accepted)))


def progress_text(done: int, total: int, ms_per_cell: float) -> str:
    """Sticky header. Minutes come from the average time of one saved cell."""
    remaining = max(0, int(total) - int(done))
    if remaining <= 0:
        eta = "0"
    elif ms_per_cell <= 0:
        eta = "—"
    else:
        minutes = remaining * float(ms_per_cell) / 60000.0
        if minutes < 0.1:
            eta = "0.1"
        elif minutes < 10:
            eta = f"{minutes:.1f}"
        else:
            eta = str(int(round(minutes)))
    return f"تم {int(done)} من {int(total)} خلية · باقي ≈ {eta} دقيقة"


def audit_record(who: str, page: object, plate: str, field: str, before: str, after: str, when: str = "") -> dict:
    """One workbook audit row: who changed the cell, and the text before and after."""
    return {
        "when": when or datetime.now().isoformat(timespec="seconds"),
        "who": who or "المراجع",
        "page": page,
        "plate": str(plate or ""),
        "field": str(field or ""),
        "before": before,
        "after": after,
    }


def review_session_path(folder: Path, pdf_bytes: bytes) -> Path:
    return Path(folder) / f"{hashlib.sha256(pdf_bytes).hexdigest()}.json"


def save_review_session(folder: Path, pdf_bytes: bytes, payload: dict) -> Path:
    """Resume file for this PDF. The PDF itself is not stored."""
    path = review_session_path(folder, pdf_bytes)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return path


def load_review_session(folder: Path, pdf_bytes: bytes) -> dict | None:
    path = review_session_path(folder, pdf_bytes)
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return data if isinstance(data, dict) else None


def review_crops(
    items: list[dict],
    pdf_bytes: bytes,
    columns: list[str],
    rows: list[list[str]],
    pages: list[int],
    baselines: dict,
    column_spans: dict,
) -> dict[tuple[str, str, str], object]:
    """Crops for the queued cells, including an empty cell that still has a window."""
    wanted = {
        (str(item.get("page")), str(item.get("plate")), str(item.get("field")))
        for item in items
        if str(item.get("field") or "") in {"color", "make", "date"}
    }
    records = _visual_records(
        pdf_bytes, columns, rows, pages, baselines, column_spans, wanted=wanted, keep_blank=True
    )
    return {(str(record[0]), str(record[1]), str(record[2])): record[4] for record in records}


def thumb_png(image) -> bytes:
    """A small black-on-white crop for the contact sheet."""
    import io

    import numpy as np
    from PIL import Image

    array = np.asarray(image, dtype="float32")
    picture = Image.fromarray(((1.0 - array) * 255).astype("uint8"), "L").resize((128, 64))
    buffer = io.BytesIO()
    picture.save(buffer, format="PNG")
    return buffer.getvalue()


def classify_cells(
    columns: list[str],
    rows: list[list[str]],
    pages: list[int],
    print_date: str = "",
    provenance: dict | None = None,
    suggestions: dict | None = None,
    raw_readings: dict | None = None,
    pdf_bytes: bytes | None = None,
    baselines: dict | None = None,
    column_spans: dict | None = None,
) -> dict[str, list[dict]]:
    """
    Green: two readings agree, and the cell is not rare. No review.
    Amber: a filled value that is not rare. A stable 5 percent sample is shown.
    Yellow: unclear or empty text. Quick pick, not the big card.
    Red: the crop does not match its reading, the value is rare, or the date is implausible.
    """
    fields, counts, modes, plated = _page_modes(rows, pages, columns)
    plate_at = fields.get("plate")
    position = {row_index: place for place, row_index in enumerate(plated)}

    def value_at(row_index, field):
        return _cell(rows[row_index], fields.get(field)).replace("\u200e", "")

    def purpose_of(row_index):
        return value_at(row_index, "purpose")

    def rare(purpose, field, value, limit):
        if not value:
            return False
        bucket = counts.get((purpose, field)) or {}
        mode = modes.get((purpose, field), "")
        return value != mode and bucket.get(value, 0) <= limit

    def unknown(field, value):
        if not value:
            return False
        if field == "make":
            return value not in _MANUFACTURERS
        if field == "color":
            return value not in _OFFICIAL_COLORS
        return False

    red_keys: dict[tuple[str, str, str], tuple[int, str]] = {}

    def mark(row_index, field, risk, reason):
        plate = value_at(row_index, "plate")
        page = int(pages[row_index]) if row_index < len(pages) else 0
        key = (str(page), plate, field)
        current = red_keys.get(key)
        if current is None or risk < current[0]:
            red_keys[key] = (risk, reason)

    color_anchors: list[int] = []
    for row_index in plated:
        page = int(pages[row_index]) if row_index < len(pages) else 0
        color = value_at(row_index, "color")
        make = value_at(row_index, "make")
        date = value_at(row_index, "date")
        year = value_at(row_index, "year")
        if unknown("color", color):
            mark(row_index, "color", 0, "لون غير واضح")
            color_anchors.append(row_index)
        elif rare(purpose_of(row_index), "color", color, _COLOR_RARE):
            mark(row_index, "color", 1, "لون قليل في المجموعة")
            color_anchors.append(row_index)
        if unknown("make", make):
            mark(row_index, "make", 0, "صنع غير واضح")
        elif rare(purpose_of(row_index), "make", make, _MAKE_RARE):
            mark(row_index, "make", 2, "صنع قليل في المجموعة")
        if date and not _DATE_VALUE.fullmatch(date):
            mark(row_index, "date", 2, "تاريخ ناقص أو غير مقروء")
        elif print_date and date and date < print_date:
            mark(row_index, "date", 2, "تاريخ أقدم من تاريخ الطباعة")
        elif rare(purpose_of(row_index), "date", date, _DATE_RARE):
            mark(row_index, "date", 2, "تاريخ قليل في المجموعة")
        limit = _print_year(print_date)
        if year and ((not _YEAR_VALUE.fullmatch(year)) or (limit and int(year) > limit)):
            mark(row_index, "year", 2, "سنة غير صالحة")
        if not color:
            mode = modes.get((purpose_of(row_index), "color"), "")
            names = _reading_names(suggestions, raw_readings, page, value_at(row_index, "plate"), "color")
            if any(name != mode for name in names):
                # The empty cell is not a wrong value. It only points at the neighbouring colour and this row's date.
                color_anchors.append(row_index)
                if date:
                    mark(row_index, "date", 5, "تاريخ بجانب لون مختلف")
        if unknown("make", make) and date:
            mark(row_index, "date", 5, "تاريخ بجانب صنع غير واضح")

    for row_index in color_anchors:
        page = int(pages[row_index]) if row_index < len(pages) else 0
        place = position[row_index]
        for delta in (-1, 1):
            other_place = place + delta
            if not (0 <= other_place < len(plated)):
                continue
            other = plated[other_place]
            if (int(pages[other]) if other < len(pages) else 0) != page:
                continue
            if value_at(other, "color"):
                mark(other, "color", 4, "لونه مختلف عن الصفوف حوله")

    for row_index in plated:
        page = int(pages[row_index]) if row_index < len(pages) else 0
        plate = value_at(row_index, "plate")
        if (str(page), plate, "color") in red_keys and value_at(row_index, "date"):
            mark(row_index, "date", 5, "تاريخ بجانب لون مختلف")
        date = value_at(row_index, "date")
        purpose = purpose_of(row_index)
        mode = modes.get((purpose, "date"), "")
        if value_at(row_index, "color") or not date or date == mode:
            continue
        bucket = counts.get((purpose, "date")) or {}
        if bucket.get(date, 0) > _EMPTY_COLOR_DATE:
            continue
        neighbour_counts = []
        place = position[row_index]
        for delta in (-1, 1):
            other_place = place + delta
            if not (0 <= other_place < len(plated)):
                continue
            other = plated[other_place]
            if (int(pages[other]) if other < len(pages) else 0) != page:
                continue
            other_date = value_at(other, "date")
            if other_date:
                neighbour_counts.append(bucket.get(other_date, 0))
        if len(neighbour_counts) == 2 and bucket.get(date, 0) < min(neighbour_counts):
            mark(row_index, "date", 2, "تاريخه مختلف عن الصفين حوله")
    visual_note = ""
    outliers: list[tuple] = []
    if pdf_bytes and baselines and column_spans:
        outliers, visual_note = visual_scan(
            pdf_bytes, columns, rows, pages, baselines, column_spans
        )
    if not visual_note and plated:
        largest_reading = 0
        for bucket in counts.values():
            if bucket:
                largest_reading = max(largest_reading, max(bucket.values()))
        if largest_reading < _VISUAL_GROUP:
            visual_note = "document too small for visual comparison"
    for page_no, plate, field in outliers:
        for row_index in plated:
            if value_at(row_index, "plate") == plate and int(pages[row_index] if row_index < len(pages) else 0) == int(page_no):
                reason = {
                    "color": "شكل اللون لا يشبه بقية الخلايا",
                    "date": "شكل التاريخ لا يشبه بقية الخلايا",
                    "make": "شكل الصنع لا يشبه بقية الخلايا",
                }.get(field, "شكله لا يشبه بقية الخلايا")
                mark(row_index, field, 1, reason)
                break

    def item_for(row_index, field, reason, risk, tier):
        page = int(pages[row_index]) if row_index < len(pages) else 0
        plate = value_at(row_index, "plate")
        column = fields.get(field)
        stored = (suggestions or {}).get(f"{page}|{plate}|{field}", "")
        ideas, why = _suggestions(
            field,
            value_at(row_index, field),
            [],
            print_date,
            "",
            _evidence_pairs(stored),
        )
        y_value = float((baselines or {}).get(f"{page}|{plate}") or 0)
        band = row_band(baselines, page, y_value) if y_value else None
        return {
            "row_index": row_index,
            "page": page,
            "plate": plate,
            "field": field,
            "index": column if column is not None else -1,
            "value": value_at(row_index, field),
            "reason": reason,
            "risk": risk,
            "tier": tier,
            "y": y_value,
            "row_top": band[0] if band else 0,
            "row_bottom": band[1] if band else 0,
            "suggestions": ideas,
            "suggestion_reasons": why,
        }

    red = []
    yellow = []
    green = []
    amber = []
    for row_index in plated:
        page = int(pages[row_index]) if row_index < len(pages) else 0
        plate = value_at(row_index, "plate")
        for field in ("date", "make", "color", "year", "admin"):
            key = (str(page), plate, field)
            source = str((provenance or {}).get(f"{page}|{plate}|{field}") or "")
            value = value_at(row_index, field)
            mode = modes.get((purpose_of(row_index), field), "")
            if key in red_keys:
                risk, reason = red_keys[key]
                tier = "yellow" if reason in _YELLOW_REASONS else "red"
                bucket = yellow if tier == "yellow" else red
                bucket.append(item_for(row_index, field, reason, risk, tier))
            elif source == "agreed":
                green.append(item_for(row_index, field, "قراءتان متفقتان", 9, "green"))
            elif not value:
                yellow.append(item_for(row_index, field, "فارغة", 8, "yellow"))
            elif field in {"color", "make", "date"} and mode and value != mode:
                amber.append(item_for(row_index, field, "مقبول، ضمن العينة", 8, "amber"))
    red.sort(key=lambda item: (item["risk"], item["page"], str(item["plate"]), item["field"]))
    yellow.sort(key=lambda item: (item["risk"], item["page"], str(item["plate"]), item["field"]))
    amber_sample = [
        item for item in amber
        if int(hashlib.sha1(f"{item['page']}|{item['plate']}|{item['field']}".encode("utf-8")).hexdigest(), 16) % _AMBER_EVERY == 0
    ]
    return {
        "red": red,
        "yellow": yellow,
        "green": green,
        "amber": amber,
        "amber_sample": amber_sample,
        "visual_note": visual_note,
    }


def assign_review_tracks(
    items: list[dict],
    rows: list[list[str]],
    pages: list[int],
    columns: list[str],
) -> dict[str, list[dict]]:
    """
    Safe: two readings of this cell agree, and the value is not an outlier.
    Exceptions: the value differs from the page's common value or is rare.
    """
    counts = _value_counts(rows, pages, columns)
    grouped = {"safe": [], "exception": [], "other": []}
    for item in items:
        ideas = [blank_text(value) for value in item.get("suggestions") or [] if blank_text(value)]
        reasons = list(item.get("suggestion_reasons") or [])
        probe = blank_text(item.get("value")) or (ideas[0] if ideas else "")
        reason = outlier_reason(str(item.get("field") or ""), probe, int(item.get("page") or 0), counts)
        agreed = any(suggestion_kind(text) == "agreed" for text in reasons)
        if reason:
            item["track"] = "exception"
            item["reason"] = reason
        elif agreed:
            item["track"] = "safe"
        else:
            item["track"] = "other"
        grouped[item["track"]].append(item)
    return grouped


def clip_acceptances(items: list[dict]) -> list[tuple[dict, str]]:
    """Two readings that agree, and only on the safe track. Exceptions stay out."""
    chosen: list[tuple[dict, str]] = []
    for item in items:
        if item.get("track") not in {None, "safe"}:
            continue
        values = list(item.get("suggestions") or [])
        reasons = list(item.get("suggestion_reasons") or [])
        for value, reason in zip(values, reasons):
            if suggestion_kind(reason) == "agreed" and str(value).strip():
                chosen.append((item, str(value).strip()))
                break
    return chosen


def filter_review_items(items: list[dict], mode: str) -> list[dict]:
    """Group the queue. Neighbour-only cells stay out of the clip batch."""
    if mode == "clip":
        return [
            item
            for item in items
            if any(suggestion_kind(reason) == "clip" for reason in item.get("suggestion_reasons") or [])
        ]
    if mode == "neighbour":
        return [
            item
            for item in items
            if any(suggestion_kind(reason) == "neighbour" for reason in item.get("suggestion_reasons") or [])
            and not any(suggestion_kind(reason) == "clip" for reason in item.get("suggestion_reasons") or [])
        ]
    return list(items)


def _evidence_pairs(text: str) -> list[tuple[str, str]]:
    """Split 'تويوتا — قراءة المقطع؛ شيفروليه — جوار ...' into pairs."""
    pairs: list[tuple[str, str]] = []
    for part in re.split(r"[؛;]", str(text or "")):
        piece = part.strip()
        if not piece:
            continue
        if " — " in piece:
            value, reason = piece.split(" — ", 1)
        else:
            value, reason = piece, ""
        pairs.append((value.strip(), reason.strip()))
    return pairs


def review_cells(
    columns: list[str],
    rows: list[list[str]],
    pages: list[int],
    print_date: str = "",
    evidence: dict[str, str] | None = None,
) -> list[dict]:
    """
    One list of cells that need a person.

    A cell is on the list when it is empty, not a known value, or a date
    that falls before the report was printed or before the vehicle year.
    Suggestions come from this row: its clip, its engines, and neighbours
    on the same page and in the same group.
    """
    fields = _vehicle_field_at(columns)
    plate_at = fields.get("plate")
    year_at = fields.get("year")
    purpose_at = fields.get("purpose")
    stored = evidence or {}
    counts = _value_counts(rows, pages, columns)
    issues: list[dict] = []
    for row_index, row in enumerate(rows):
        if _excluded_row(columns, row):
            continue
        page = int(pages[row_index]) if row_index < len(pages) else 0
        vehicle_year = _cell(row, year_at)
        limit = _print_year(print_date)
        if not _YEAR_VALUE.fullmatch(vehicle_year) or (limit and int(vehicle_year) > limit):
            vehicle_year = ""
        plate = _cell(row, plate_at)
        for field in REVIEW_FIELDS:
            column = fields.get(field)
            if column is None:
                continue
            value = _cell(row, column)
            reason = _field_reason(field, value, print_date, vehicle_year)
            if not reason:
                reason = outlier_reason(field, value, page, counts)
            if not reason and field in {"color", "make", "date"} and value:
                adjacent = []
                for cursor in (row_index - 1, row_index + 1):
                    if cursor < 0 or cursor >= len(rows):
                        continue
                    if cursor < len(pages) and int(pages[cursor]) != page:
                        continue
                    text = _cell(rows[cursor], column).replace("\u200e", "")
                    if text:
                        adjacent.append(text)
                if len(adjacent) == 2 and adjacent[0] != value and adjacent[1] != value:
                    reason = "بين قيمتين مختلفتين عن الجوار"
            if not reason:
                continue
            neighbours = _neighbour_values(rows, pages, row_index, column, purpose_at)
            ideas, why = _suggestions(
                field,
                value,
                neighbours,
                print_date,
                vehicle_year,
                _evidence_pairs(stored.get(f"{page}|{plate}|{field}", "")),
            )
            issues.append(
                {
                    "row_index": row_index,
                    "page": page,
                    "plate": plate,
                    "field": field,
                    "index": column,
                    "value": value,
                    "reason": reason,
                    "suggestions": ideas,
                    "suggestion_reasons": why,
                }
            )
    return issues


def vision_queue(red: list[dict], yellow: list[dict], readings: dict[str, str]) -> dict[str, list[dict]]:
    """
    A vision reading that matches the OCR text leaves the tier as it is.

    A different reading moves that cell into Red. Cells are not dropped.
    """
    promoted = []
    still_yellow = []
    for item in yellow:
        key = f"{item['page']}|{item['plate']}|{item['field']}"
        seen = blank_text((readings or {}).get(key))
        if seen and seen != blank_text(item.get("value")):
            moved = dict(item)
            moved["tier"] = "red"
            moved["reason"] = "الرؤية تختلف عن القراءة"
            promoted.append(moved)
        else:
            still_yellow.append(item)
    noted = []
    for item in red:
        key = f"{item['page']}|{item['plate']}|{item['field']}"
        seen = blank_text((readings or {}).get(key))
        if seen and seen != blank_text(item.get("value")):
            moved = dict(item)
            moved["reason"] = "الرؤية تختلف عن القراءة"
            noted.append(moved)
        else:
            noted.append(item)
    return {"red": noted + promoted, "yellow": still_yellow}


def queue_status(issue_count: int, column_warning: str = "") -> str:
    """An unrecognised column is never reported as a clean, empty queue."""
    warning = str(column_warning or "").strip()
    if warning and ("تعذر" in warning or "لا تعني" in warning):
        return warning
    if issue_count:
        return f"⚠ {issue_count} خلية تحتاج مراجعة"
    if warning:
        return warning
    return "✅ لا توجد خلايا تحتاج مراجعة"


def banner_facts(validation: list[dict], issue_count: int) -> dict:
    """Plain lines for the top of the screen. The issue count is the review list."""
    checks = list(validation or [])
    by_name = {str(item.get("name") or ""): item for item in checks}
    total = by_name.get("المجموع") or {}
    expected = str(total.get("expected") or "")
    actual = str(total.get("actual") or "")
    order = ["باص خاص", "خصوصي", "نقل خاص"]
    groups = []
    for name in order:
        item = by_name.get(name)
        if item is not None:
            groups.append(str(item.get("actual") or "0"))
    for name, item in by_name.items():
        if name in order or name == "المجموع" or name.startswith("صفحة"):
            continue
        groups.append(str(item.get("actual") or "0"))
    group_text = "/".join(groups)
    count_ok = bool(total) and total.get("status") == "ok" and all(
        by_name[name].get("status") == "ok" for name in order if name in by_name
    )
    if expected and actual:
        if count_ok:
            count_line = f"✅ العدد مطابق: {actual} من {expected}"
        else:
            count_line = f"⚠ العدد: {actual} من {expected}"
    elif actual:
        count_line = f"✅ عدد المركبات: {actual}"
        count_ok = True
    else:
        count_line = "✅ تم استخراج الجدول"
        count_ok = True
    if group_text:
        count_line = f"{count_line} ({group_text})"
    if issue_count:
        issue_line = f"⚠ {issue_count} خلية تحتاج مراجعة"
    else:
        issue_line = "✅ لا توجد خلايا تحتاج مراجعة"
    pages = []
    for item in checks:
        name = str(item.get("name") or "")
        if not name.startswith("صفحة") or item.get("status") not in {"fail", "weak"}:
            continue
        number = name.split()[-1]
        pages.append(
            {
                "page": int(number) if number.isdigit() else 0,
                "line": (
                    f"صفحة {number}: المرور الأول وجد {item.get('expected') or item.get('actual') or '0'}، والباقي استُعيد بإعادة القراءة"
                    if item.get("status") == "weak"
                    else f"صفحة {number}: المرور الأول وجد {item.get('expected') or '0'}، والناتج {item.get('actual') or '0'}"
                ),
            }
        )
    return {
        "count_line": count_line,
        "count_ok": count_ok,
        "issue_line": issue_line,
        "issue_count": int(issue_count),
        "pages": pages,
    }
