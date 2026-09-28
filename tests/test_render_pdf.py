# -*- coding: utf-8 -*-
import importlib.util
import os
import tempfile
import time
import unittest
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
RENDERER = SKILL_DIR / "scripts" / "render_pdf.py"


def load_renderer():
    spec = importlib.util.spec_from_file_location("word_deliverable_render_pdf", RENDERER)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class RenderPdfTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.renderer = load_renderer()

    def test_prepare_output_removes_stale_pdf(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "old.pdf"
            path.write_bytes(b"old")
            started_at = self.renderer._prepare_output(path)
            self.assertFalse(path.exists())
            self.assertLessEqual(started_at, time.time())

    def test_fresh_output_rejects_missing_empty_and_stale_pdf(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "result.pdf"
            started_at = time.time()
            self.assertFalse(self.renderer._is_fresh_output(path, started_at))
            path.write_bytes(b"")
            self.assertFalse(self.renderer._is_fresh_output(path, started_at))
            path.write_bytes(b"pdf")
            os.utime(path, (started_at - 10, started_at - 10))
            self.assertFalse(self.renderer._is_fresh_output(path, started_at))
            os.utime(path, None)
            self.assertTrue(self.renderer._is_fresh_output(path, started_at))

    def test_visual_review_can_be_read_from_spec(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            spec_path = Path(temp_dir) / "spec.json"
            spec_path.write_text('{"visual_review": "none"}', encoding="utf-8")
            self.assertEqual(self.renderer._resolve_visual_review(spec_path), "none")
            self.assertEqual(self.renderer._resolve_visual_review(None), "pdf-text")


if __name__ == "__main__":
    unittest.main()
