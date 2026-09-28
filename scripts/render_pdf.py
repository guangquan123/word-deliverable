# -*- coding: utf-8 -*-
"""Export a DOCX to a freshly generated PDF and run image-free validation.

Usage:
    python render_pdf.py <input.docx> <output.pdf> [spec.json]

Export priority:
  1. Microsoft Word COM (skip with WD_NO_WORD=1)
  2. LibreOffice UNO (updates TOC and fields)
  3. LibreOffice headless convert-to (basic fallback only)
"""
from __future__ import annotations

import importlib.util
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

SOFFICE = os.environ.get("LO_PATH", r"C:\Program Files\LibreOffice\program\soffice.exe")
LO_PYTHON = os.environ.get("LO_PYTHON", r"C:\Program Files\LibreOffice\program\python.exe")
SCRIPT_DIR = Path(__file__).resolve().parent
UNO_SCRIPT = SCRIPT_DIR / "render_pdf_uno.py"
PDF_VALIDATOR = SCRIPT_DIR / "validate_pdf.py"
VALID_VISUAL_REVIEWS = {"none", "pdf-text", "raster"}


def _com_retry(fn, retries=5, delay=1.5):
    last = None
    for index in range(retries):
        try:
            return fn()
        except Exception as exc:
            last = exc
            if index < retries - 1:
                time.sleep(delay * (index + 1))
    raise last if last else RuntimeError("COM retry exhausted")


def _prepare_output(path):
    """Remove stale output and return the timestamp for freshness checks."""
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    return time.time()


def _is_fresh_output(path, started_at):
    output = Path(path)
    if not output.is_file() or output.stat().st_size <= 0:
        return False
    # Permit a one-second filesystem timestamp granularity difference.
    return output.stat().st_mtime >= started_at - 1.0


def _resolve_visual_review(spec_path=None):
    if spec_path is None:
        return "pdf-text"
    with open(spec_path, "r", encoding="utf-8-sig") as handle:
        value = json.load(handle).get("visual_review", "pdf-text")
    if value not in VALID_VISUAL_REVIEWS:
        raise ValueError(
            f"visual_review 必须为 none、pdf-text 或 raster，当前为: {value!r}"
        )
    return value


def _update_word_fields(doc):
    """Update main-story, header/footer and TOC fields before PDF export."""
    try:
        _com_retry(lambda: doc.Repaginate(), retries=3, delay=0.5)
    except Exception:
        pass

    try:
        _com_retry(lambda: doc.Fields.Update(), retries=3, delay=0.5)
    except Exception:
        pass

    try:
        for index in range(1, doc.TablesOfContents.Count + 1):
            toc = doc.TablesOfContents.Item(index)
            _com_retry(lambda item=toc: item.Update(), retries=3, delay=0.5)
    except Exception:
        pass

    # StoryRanges includes headers, footers, footnotes and text frames.
    try:
        for story_type in range(1, 18):
            try:
                story = doc.StoryRanges.Item(story_type)
            except Exception:
                continue
            while story is not None:
                try:
                    story.Fields.Update()
                except Exception:
                    pass
                try:
                    story = story.NextStoryRange
                except Exception:
                    story = None
    except Exception:
        pass

    try:
        _com_retry(lambda: doc.Repaginate(), retries=3, delay=0.5)
    except Exception:
        pass


def with_word(docx_path, pdf_path):
    import pythoncom
    import win32com.client

    started_at = _prepare_output(pdf_path)
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        try:
            word = win32com.client.gencache.EnsureDispatch("Word.Application")
        except Exception:
            word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = _com_retry(lambda: word.Documents.Open(docx_path, ReadOnly=False))
        _update_word_fields(doc)

        def _export():
            try:
                doc.ExportAsFixedFormat(pdf_path, 17)
            except Exception:
                doc.SaveAs(pdf_path, FileFormat=17)

        _com_retry(_export)
        return _is_fresh_output(pdf_path, started_at)
    finally:
        if doc is not None:
            try:
                _com_retry(lambda: doc.Close(SaveChanges=False), retries=3, delay=1.0)
            except Exception:
                pass
        if word is not None:
            try:
                _com_retry(lambda: word.Quit(), retries=3, delay=1.0)
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def with_uno(docx_path, pdf_path):
    """Use LibreOffice UNO to update TOC/fields and export PDF."""
    if not Path(LO_PYTHON).exists() or not UNO_SCRIPT.exists():
        return False
    started_at = _prepare_output(pdf_path)
    try:
        result = subprocess.run(
            [LO_PYTHON, str(UNO_SCRIPT), docx_path, pdf_path],
            check=True,
            timeout=120,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if _is_fresh_output(pdf_path, started_at):
            return True
        print(result.stdout)
        print(result.stderr)
    except subprocess.CalledProcessError as exc:
        print(f"  UNO 失败: {exc}")
        if exc.stdout:
            print(f"  stdout: {exc.stdout[:500]}")
        if exc.stderr:
            print(f"  stderr: {exc.stderr[:500]}")
    except Exception as exc:
        print(f"  UNO 失败: {exc}")
    return False


def with_libreoffice(docx_path, pdf_path):
    """Use basic LibreOffice convert-to; TOC may remain stale."""
    if not Path(SOFFICE).exists():
        return False
    output = Path(pdf_path).resolve()
    generated = output.parent / (Path(docx_path).stem + ".pdf")
    started_at = _prepare_output(output)
    if generated != output and generated.exists():
        generated.unlink()
    try:
        subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", str(output.parent), docx_path],
            check=True,
            timeout=180,
            capture_output=True,
        )
        if generated != output and _is_fresh_output(generated, started_at):
            shutil.move(str(generated), str(output))
        return _is_fresh_output(output, started_at)
    except Exception as exc:
        print(f"  LibreOffice 失败: {exc}")
        return False


def _load_pdf_validator():
    spec = importlib.util.spec_from_file_location("word_deliverable_validate_pdf", PDF_VALIDATOR)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _validate_export(pdf_path, visual_review):
    if visual_review == "none":
        print("跳过 PDF 文本坐标验收（visual_review=none）")
        return True
    validator = _load_pdf_validator()
    report, problems, warnings = validator.validate_pdf(pdf_path)
    for line in report:
        print(line)
    for item in warnings:
        print(f"提醒: {item}")
    for item in problems:
        print(f"问题: {item}")
    if visual_review == "raster":
        print("提示: raster 为显式人工复检模式；本脚本不自动栅格化页面。")
    return not problems


def main():
    if len(sys.argv) not in (3, 4):
        print("用法: python render_pdf.py <input.docx> <output.pdf> [spec.json]")
        return 1

    docx_path = os.path.abspath(sys.argv[1])
    pdf_path = os.path.abspath(sys.argv[2])
    spec_path = os.path.abspath(sys.argv[3]) if len(sys.argv) == 4 else None
    if not os.path.exists(docx_path):
        print(f"文件不存在: {docx_path}")
        return 1
    if spec_path and not os.path.exists(spec_path):
        print(f"规格文件不存在: {spec_path}")
        return 1

    try:
        visual_review = _resolve_visual_review(spec_path)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"规格文件无效: {exc}")
        return 1

    force_lo = os.environ.get("WD_NO_WORD", "") == "1"
    exporter = None

    if not force_lo:
        try:
            if with_word(docx_path, pdf_path):
                exporter = "Word"
        except Exception as exc:
            print(f"Word COM 失败: {exc}")

    if exporter is None and with_uno(docx_path, pdf_path):
        exporter = "UNO"
    if exporter is None and with_libreoffice(docx_path, pdf_path):
        exporter = "LibreOffice"

    if exporter is None:
        print("失败: 所有导出方式均不可用，或未生成新的非空 PDF")
        return 2

    print(f"OK PDF ({exporter}): {pdf_path}")
    if not _validate_export(pdf_path, visual_review):
        print("失败: PDF 文本坐标验收未通过")
        return 3
    print("OK: PDF 验收通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
