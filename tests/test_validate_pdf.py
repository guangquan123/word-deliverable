# -*- coding: utf-8 -*-
import importlib.util
import tempfile
import unittest
from pathlib import Path

import fitz


SKILL_DIR = Path(__file__).resolve().parents[1]
VALIDATOR = SKILL_DIR / "scripts" / "validate_pdf.py"


def load_validator():
    spec = importlib.util.spec_from_file_location("word_deliverable_validate_pdf", VALIDATOR)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ValidatePdfTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.validator = load_validator()

    def make_pdf(self, page_builders):
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        path = Path(temp_dir.name) / "sample.pdf"
        doc = fitz.open()
        for builder in page_builders:
            page = doc.new_page(width=595, height=842)
            builder(page)
        doc.save(path)
        doc.close()
        return path

    @staticmethod
    def normal_page(page):
        for index in range(18):
            page.insert_text((72, 90 + index * 30), f"Line {index + 1}: normal body content for validation.", fontsize=11)

    def test_normal_pdf_passes(self):
        path = self.make_pdf([self.normal_page, self.normal_page])
        report, problems, warnings = self.validator.validate_pdf(path)
        self.assertEqual(problems, [], "\n".join(report + problems + warnings))
        self.assertTrue(any("2 页" in line for line in report))

    def test_blank_middle_page_fails(self):
        path = self.make_pdf([self.normal_page, lambda page: None, self.normal_page])
        _, problems, _ = self.validator.validate_pdf(path)
        self.assertTrue(any("空白页" in item for item in problems))

    def test_text_outside_safe_page_boundary_fails(self):
        def overflowing(page):
            page.insert_text((1, 100), "outside left safe boundary", fontsize=11)
        path = self.make_pdf([overflowing])
        _, problems, _ = self.validator.validate_pdf(path)
        self.assertTrue(any("页面边界" in item for item in problems))

    def test_sparse_last_page_is_reported(self):
        def sparse(page):
            page.insert_text((72, 90), "orphan final line", fontsize=11)
        path = self.make_pdf([self.normal_page, sparse])
        _, problems, _ = self.validator.validate_pdf(path)
        self.assertTrue(any("末页残余" in item for item in problems))

    def test_heading_near_page_bottom_is_reported(self):
        def orphan_heading(page):
            page.insert_text((72, 790), "1. Orphan Heading", fontsize=18)
        path = self.make_pdf([self.normal_page, orphan_heading, self.normal_page])
        _, problems, _ = self.validator.validate_pdf(path)
        self.assertTrue(any("页尾孤立标题" in item for item in problems))


if __name__ == "__main__":
    unittest.main()
