# -*- coding: utf-8 -*-
"""
Compare a PDF with the Word file made from it, page by page.

The Word file is printed to PDF on this computer, by Microsoft Word when it
is installed and by LibreOffice otherwise. Nothing is uploaded. Each PDF page
and its Word page are drawn at the same scale. The report gives a similarity
score per page, the share of words that sit where they sat in the PDF, and a
picture of the differences.

    python -m core.visual_check source.pdf result.docx --out report_folder
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path

import fitz

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# Points on the page are drawn at this many pixels per point.
SCALE = 1.5
# A word counts as in place when it moved less than this many points.
WORD_TOLERANCE = 2.0


class RenderError(Exception):
    """No program on this computer could print the Word file."""


@dataclass
class PageScore:
    """How one Word page compares with its PDF page."""

    number: int
    similarity: float
    ink_overlap: float
    words: int
    words_in_place: float
    median_shift: float
    worst_shift: float
    size_matches: bool
    diff_image: str = ""


@dataclass
class Report:
    renderer: str
    source_pages: int
    result_pages: int
    pages: list[PageScore] = field(default_factory=list)

    @property
    def similarity(self) -> float:
        if not self.pages:
            return 0.0
        return sum(page.similarity for page in self.pages) / len(self.pages)

    def to_json(self) -> str:
        data = asdict(self)
        data["similarity"] = self.similarity
        return json.dumps(data, ensure_ascii=False, indent=2)


def _word_exe() -> str | None:
    """WINWORD.EXE from the registry or the usual folders."""
    try:
        import winreg

        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, key_path) as key:
                    value = winreg.QueryValue(key, None)
                    if value and Path(value).is_file():
                        return value
            except OSError:
                continue
    except ImportError:
        return None
    return None


def _soffice_exe() -> str | None:
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return found
    for base in (os.environ.get("PROGRAMFILES"), os.environ.get("PROGRAMFILES(X86)")):
        if not base:
            continue
        candidate = Path(base) / "LibreOffice" / "program" / "soffice.exe"
        if candidate.is_file():
            return str(candidate)
    return None


def available_renderers() -> list[str]:
    """The programs that can print a .docx here, best first."""
    names = []
    if _word_exe():
        names.append("word")
    if _soffice_exe():
        names.append("libreoffice")
    return names


_WORD_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$mine = ($word.Documents.Count -eq 0)
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($env:PDF_TOOL_DOCX, $false, $true, $false)
    try {
        $document.ExportAsFixedFormat($env:PDF_TOOL_PDF, 17)
    } finally {
        $document.Close(0)
    }
} finally {
    if ($mine) { $word.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
"""


def _render_with_word(docx: Path, target: Path) -> None:
    env = dict(os.environ)
    env["PDF_TOOL_DOCX"] = str(docx)
    env["PDF_TOOL_PDF"] = str(target)
    done = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", _WORD_SCRIPT],
        env=env,
        capture_output=True,
        text=True,
        timeout=300,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if done.returncode != 0 or not target.is_file():
        raise RenderError(f"Word could not print the file. {done.stderr.strip()[:400]}")


def _render_with_libreoffice(docx: Path, target: Path) -> None:
    exe = _soffice_exe()
    if not exe:
        raise RenderError("LibreOffice is not installed.")
    with tempfile.TemporaryDirectory(prefix="pdf_tool_lo_") as profile:
        out_dir = Path(profile) / "out"
        out_dir.mkdir()
        done = subprocess.run(
            [
                exe,
                f"-env:UserInstallation={Path(profile).as_uri()}/profile",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(out_dir),
                str(docx),
            ],
            capture_output=True,
            text=True,
            timeout=300,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        made = out_dir / (docx.stem + ".pdf")
        if done.returncode != 0 or not made.is_file():
            raise RenderError(f"LibreOffice could not print the file. {done.stderr.strip()[:400]}")
        shutil.copyfile(made, target)


def render_docx(docx_path: str | Path, pdf_path: str | Path, renderer: str = "auto") -> str:
    """
    Print the Word file to PDF and return the name of the program used.

    `renderer` is word, libreoffice, or auto. Auto prefers Word, because Word
    is where the file will be opened.
    """
    docx = Path(docx_path).resolve()
    target = Path(pdf_path).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    choices = available_renderers() if renderer == "auto" else [renderer]
    if not choices:
        raise RenderError("Neither Microsoft Word nor LibreOffice is installed.")
    problems = []
    for name in choices:
        try:
            if name == "word":
                _render_with_word(docx, target)
            elif name == "libreoffice":
                _render_with_libreoffice(docx, target)
            else:
                raise RenderError(f"Unknown renderer: {name}")
            return name
        except (RenderError, subprocess.TimeoutExpired, OSError) as exc:
            problems.append(str(exc))
    raise RenderError(" ".join(problems))


def _page_image(page: fitz.Page, scale: float):
    from PIL import Image

    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def _compare_images(source, result):
    """Similarity, ink overlap, and a picture of where the two pages differ."""
    import numpy
    from PIL import Image, ImageFilter

    if result.size != source.size:
        result = result.resize(source.size, Image.LANCZOS)
    # A one-pixel blur keeps anti-aliasing from counting as a difference.
    left = numpy.asarray(source.filter(ImageFilter.GaussianBlur(1.0)), dtype=numpy.int16)
    right = numpy.asarray(result.filter(ImageFilter.GaussianBlur(1.0)), dtype=numpy.int16)
    delta = numpy.abs(left - right).max(axis=2)
    similarity = 1.0 - float(delta.mean()) / 255.0

    def ink(image) -> "numpy.ndarray":
        """Pixels that differ from the page's most common color."""
        small = numpy.asarray(image.filter(ImageFilter.GaussianBlur(1.0)), dtype=numpy.int16)
        flat = small.reshape(-1, 3)
        values, counts = numpy.unique(flat[:: max(1, len(flat) // 20000)], axis=0, return_counts=True)
        paper = values[counts.argmax()]
        return numpy.abs(small - paper).max(axis=2) > 40

    ink_left = ink(source)
    ink_right = ink(result)
    union = int((ink_left | ink_right).sum())
    overlap = float((ink_left & ink_right).sum()) / union if union else 1.0

    faded = (numpy.asarray(source.convert("L"), dtype=numpy.float32) * 0.35 + 255 * 0.65).astype(numpy.uint8)
    picture = numpy.stack([faded, faded, faded], axis=2)
    changed = delta > 48
    picture[changed] = (220, 30, 30)
    return similarity, overlap, Image.fromarray(picture, "RGB")


def _words(page: fitz.Page) -> list[tuple[str, float, float]]:
    """Each word with the left end of its baseline box and its vertical middle."""
    found = []
    for item in page.get_text("words"):
        text = str(item[4]).strip()
        if text:
            found.append((text, float(item[0]), (float(item[1]) + float(item[3])) / 2.0))
    return found


def _word_shifts(source: fitz.Page, result: fitz.Page) -> list[float]:
    """How far each PDF word moved in the Word page, in points of the PDF page."""
    scale_x = source.rect.width / max(result.rect.width, 1.0)
    scale_y = source.rect.height / max(result.rect.height, 1.0)
    by_text: dict[str, list[tuple[float, float]]] = {}
    for text, x, y in _words(result):
        by_text.setdefault(text, []).append((x * scale_x, y * scale_y))
    shifts = []
    for text, x, y in _words(source):
        places = by_text.get(text)
        if not places:
            shifts.append(float("inf"))
            continue
        shifts.append(min(((x - other_x) ** 2 + (y - other_y) ** 2) ** 0.5 for other_x, other_y in places))
    return shifts


def compare_pdfs(
    source_pdf: str | Path,
    result_pdf: str | Path,
    out_dir: str | Path | None = None,
    renderer: str = "",
    scale: float = SCALE,
) -> Report:
    """Compare the source PDF with the PDF printed from the Word file."""
    source = fitz.open(str(source_pdf))
    result = fitz.open(str(result_pdf))
    folder = Path(out_dir) if out_dir else None
    if folder is not None:
        folder.mkdir(parents=True, exist_ok=True)
    report = Report(renderer, source.page_count, result.page_count)
    try:
        for index in range(min(source.page_count, result.page_count)):
            left = source.load_page(index)
            right = result.load_page(index)
            source_image = _page_image(left, scale)
            # The Word page is drawn into the same pixel grid as the PDF page.
            zoom_x = source_image.width / max(right.rect.width, 1.0)
            zoom_y = source_image.height / max(right.rect.height, 1.0)
            pixmap = right.get_pixmap(matrix=fitz.Matrix(zoom_x, zoom_y), alpha=False)
            from PIL import Image

            result_image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
            similarity, overlap, picture = _compare_images(source_image, result_image)
            shifts = _word_shifts(left, right)
            finite = sorted(value for value in shifts if value != float("inf"))
            in_place = sum(1 for value in shifts if value <= WORD_TOLERANCE)
            diff_path = ""
            if folder is not None:
                diff_file = folder / f"page-{index + 1:03d}-diff.png"
                picture.save(diff_file)
                source_image.save(folder / f"page-{index + 1:03d}-pdf.png")
                result_image.save(folder / f"page-{index + 1:03d}-docx.png")
                diff_path = str(diff_file)
            report.pages.append(
                PageScore(
                    number=index + 1,
                    similarity=round(similarity, 4),
                    ink_overlap=round(overlap, 4),
                    words=len(shifts),
                    words_in_place=round(in_place / len(shifts), 4) if shifts else 1.0,
                    median_shift=round(finite[len(finite) // 2], 2) if finite else 0.0,
                    worst_shift=round(finite[-1], 2) if finite else 0.0,
                    size_matches=abs(left.rect.width - right.rect.width) <= 1.0
                    and abs(left.rect.height - right.rect.height) <= 1.0,
                    diff_image=diff_path,
                )
            )
    finally:
        source.close()
        result.close()
    return report


def compare_pdf_with_docx(
    source_pdf: str | Path,
    docx_path: str | Path,
    out_dir: str | Path | None = None,
    renderer: str = "auto",
    scale: float = SCALE,
) -> Report:
    """
    Print the Word file and compare it with the PDF it came from.

    With `out_dir`, each page leaves three pictures there: the PDF page, the
    Word page, and the differences in red on a faded copy of the PDF page.
    """
    folder = Path(out_dir) if out_dir else Path(tempfile.mkdtemp(prefix="pdf_tool_visual_"))
    folder.mkdir(parents=True, exist_ok=True)
    printed = folder / "docx-printed.pdf"
    used = render_docx(docx_path, printed, renderer)
    report = compare_pdfs(source_pdf, printed, out_dir, used, scale)
    if out_dir:
        (Path(out_dir) / "report.json").write_text(report.to_json(), encoding="utf-8")
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Compare a PDF with the Word file made from it.")
    parser.add_argument("pdf")
    parser.add_argument("docx")
    parser.add_argument("--out", default="", help="Folder for the page pictures and report.json")
    parser.add_argument("--renderer", default="auto", choices=("auto", "word", "libreoffice"))
    args = parser.parse_args(argv)
    try:
        report = compare_pdf_with_docx(args.pdf, args.docx, args.out or None, args.renderer)
    except RenderError as exc:
        print(exc, file=sys.stderr)
        return 1
    print(f"renderer={report.renderer}  pages pdf={report.source_pages} docx={report.result_pages}")
    for page in report.pages:
        print(
            f"page {page.number}: similarity={page.similarity:.4f}  ink overlap={page.ink_overlap:.4f}  "
            f"words in place={page.words_in_place:.1%} of {page.words}  "
            f"median shift={page.median_shift}pt  worst={page.worst_shift}pt  "
            f"same size={'yes' if page.size_matches else 'NO'}"
        )
    print(f"mean similarity={report.similarity:.4f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
