# -*- coding: utf-8 -*-
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
GENERATOR = SKILL_DIR / "scripts" / "generate_docx.py"
VALIDATOR = SKILL_DIR / "scripts" / "validate_docx.py"


def load_validator():
    spec = importlib.util.spec_from_file_location("word_deliverable_validate_docx", VALIDATOR)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ValidateDocxTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.validator = load_validator()

    def generate(self, **overrides):
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        root = Path(temp_dir.name)
        spec = {
            "cover": False,
            "toc": False,
            "doc_control": False,
            "header_footer": False,
            "body": [{"type": "h1", "text": "测试"}, {"type": "p", "text": "正文"}],
        }
        spec.update(overrides)
        spec_path = root / "spec.json"
        output_path = root / "result.docx"
        spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        result = subprocess.run(
            [sys.executable, str(GENERATOR), str(spec_path), str(output_path)],
            capture_output=True,
            text=True,
        )
        self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
        return spec, output_path

    def test_disabled_toc_and_header_footer_are_not_required(self):
        spec, path = self.generate()
        report, problems = self.validator.validate_docx(path, spec)
        self.assertEqual(problems, [], "\n".join(report + problems))
        self.assertTrue(any("按 spec 跳过" in line for line in report))

    def test_enabled_fields_are_found_as_complete_complex_fields(self):
        spec, path = self.generate(
            cover=True,
            toc=True,
            header_footer=True,
            project_name="示例项目",
            doc_name="示例文档",
        )
        report, problems = self.validator.validate_docx(path, spec)
        self.assertFalse(any("缺少 PAGE 字段" in item for item in problems))
        self.assertFalse(any("缺少 NUMPAGES 字段" in item for item in problems))
        self.assertFalse(any("域结构不完整" in item for item in problems))
        self.assertFalse(any("表格#0" in item for item in problems), "\n".join(problems))
        self.assertTrue(any("TOC 仍含占位文本" in item for item in problems))

    def test_validator_without_spec_treats_first_table_as_default_cover_table(self):
        _, path = self.generate(
            cover=True,
            toc=True,
            header_footer=True,
            project_name="示例项目",
            doc_name="示例文档",
        )
        _, problems = self.validator.validate_docx(path)
        self.assertFalse(
            any("表格#0 无重复标题行" in item for item in problems),
            "\n".join(problems),
        )

    def test_generated_table_has_repeat_header_cant_split_and_fixed_layout(self):
        spec, path = self.generate(body=[
            {"type": "h1", "text": "测试"},
            {"type": "table", "headers": ["A", "B"], "rows": [["1", "2"], ["3", "4"]]},
        ])
        report, problems = self.validator.validate_docx(path, spec)
        self.assertFalse(any("表格#0" in item for item in problems), "\n".join(problems))
        self.assertTrue(any("固定布局" in line for line in report))

    def test_all_six_heading_styles_are_bound_to_numbering(self):
        body = [{"type": f"h{i}", "text": f"标题{i}"} for i in range(1, 7)]
        spec, path = self.generate(body=body)
        _, problems = self.validator.validate_docx(path, spec)
        self.assertFalse(any("Heading" in item and "编号" in item for item in problems), "\n".join(problems))


if __name__ == "__main__":
    unittest.main()
