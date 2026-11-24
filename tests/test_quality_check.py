# -*- coding: utf-8 -*-
"""The quality checks find the faults they name, and pass a clean file."""

from __future__ import annotations

import unittest
from pathlib import Path

import fitz
from docx import Document

from core import quality_check as qc

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "tests" / "corpus" / "files" / "related_party_bilingual.pdf"
OLD_WORD = ROOT / "Related_Party_editable.docx"


class WordsTest(unittest.TestCase):
    def test_tashkeel_tatweel_and_hamza_do_not_change_a_word(self):
        self.assertEqual(qc.normalize_word("الإدارةُ"), qc.normalize_word("الادارة"))
        self.assertEqual(qc.normalize_word("مـبـاشـرة"), qc.normalize_word("مباشرة"))

    def test_a_reversed_arabic_word_still_matches(self):
        self.assertEqual(qc._bag(["سياسة"]), qc._bag(["ةسايس"]))

    def test_overprinted_word_counts_once(self):
        page = fitz.open().new_page()
        page.insert_text((72, 72), "Policy", fontsize=12)
        page.insert_text((72.4, 72), "Policy", fontsize=12)
        words, unreadable = qc.page_words(page)
        self.assertEqual(words, ["policy"])
        self.assertEqual(unreadable, 0)


class StructureTest(unittest.TestCase):
    def test_plain_paragraphs_pass(self):
        import tempfile

        with tempfile.TemporaryDirectory() as work:
            path = Path(work) / "plain.docx"
            document = Document()
            document.add_paragraph("سياسة التعاملات مع الأطراف ذوي العلاقة")
            document.add_paragraph("Related Parties Transactions Policy")
            document.save(path)
            report = qc.QualityReport("x.pdf", str(path))
            qc.check_structure(path, report, source_word_count=11)
            qc.check_spacing(path, report)
            qc.check_duplicates(path, report)
            self.assertEqual(report.findings, [])
            self.assertEqual(report.measures["editable_share"], 1.0)

    @unittest.skipUnless(OLD_WORD.is_file() and SAMPLE.is_file(), "sample files not present")
    def test_text_box_file_is_flagged(self):
        report = qc.check_word(SAMPLE, OLD_WORD, printed=False)
        checks = report.counts()
        self.assertIn("editable", checks)
        self.assertIn("spacing", checks)
        self.assertGreater(report.measures["max_spacing_pt"], 100)
        self.assertFalse(report.ok)


if __name__ == "__main__":
    unittest.main()
