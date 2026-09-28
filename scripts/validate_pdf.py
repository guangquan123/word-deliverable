# -*- coding: utf-8 -*-
"""PDF text-coordinate validation for Word deliverables.

This validator intentionally stays image-free. It inspects the PDF text layer and
page geometry to catch high-confidence layout failures before delivery.
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:  # pragma: no cover - compatibility with older PyMuPDF
    import fitz as pymupdf


SAFE_EDGE_PT = 8.0
SPARSE_TEXT_CHARS = 80
SPARSE_BLOCKS = 2
HEADING_BOTTOM_RATIO = 0.90
HEADING_MIN_SIZE = 15.0


def _text_blocks(page):
    data = page.get_text("dict", sort=True)
    blocks = []
    spans = []
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        block_text = []
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text:
                    continue
                record = {
                    "text": text,
                    "bbox": tuple(span.get("bbox", (0, 0, 0, 0))),
                    "size": float(span.get("size", 0)),
                }
                spans.append(record)
                block_text.append(text)
        if block_text:
            blocks.append({
                "text": " ".join(block_text),
                "bbox": tuple(block.get("bbox", (0, 0, 0, 0))),
            })
    return blocks, spans


def _looks_like_heading(text, size):
    compact = re.sub(r"\s+", "", text)
    if not compact or len(compact) > 80 or size < HEADING_MIN_SIZE:
        return False
    return bool(
        re.match(r"^(第[一二三四五六七八九十百0-9]+章|[0-9]+(?:\.[0-9]+){0,5}[.、]?|[一二三四五六七八九十]+、)", compact)
        or len(compact) <= 30
    )


def validate_pdf(path):
    """Return ``(report, problems, warnings)`` for a PDF path."""
    pdf_path = Path(path)
    report = []
    problems = []
    warnings = []

    if not pdf_path.exists():
        return report, [f"PDF 文件不存在: {pdf_path}"], warnings
    if pdf_path.stat().st_size <= 0:
        return report, [f"PDF 文件为空: {pdf_path}"], warnings

    try:
        doc = pymupdf.open(str(pdf_path))
    except Exception as exc:
        return report, [f"PDF 无法打开或已损坏: {exc}"], warnings

    try:
        page_count = doc.page_count
        report.append(f"PDF 共 {page_count} 页")
        if page_count == 0:
            problems.append("PDF 没有任何页面")
            return report, problems, warnings

        page_stats = []
        for index, page in enumerate(doc):
            page_no = index + 1
            width = float(page.rect.width)
            height = float(page.rect.height)
            blocks, spans = _text_blocks(page)
            text = "".join(span["text"] for span in spans).strip()
            char_count = len(re.sub(r"\s+", "", text))
            page_stats.append((char_count, len(blocks)))

            if not text:
                # A cover may be drawing-only, but internal blank pages are always suspicious.
                if 0 < index < page_count - 1:
                    problems.append(f"第 {page_no} 页为空白页")
                elif page_count == 1:
                    problems.append("PDF 唯一页面没有可检索文本层")
                else:
                    warnings.append(f"第 {page_no} 页没有可检索文本层，请确认是否为设计性页面")
                continue

            overflow = []
            for span in spans:
                x0, y0, x1, y1 = span["bbox"]
                if x0 < SAFE_EDGE_PT or y0 < SAFE_EDGE_PT or x1 > width - SAFE_EDGE_PT or y1 > height - SAFE_EDGE_PT:
                    overflow.append(span["text"][:30])
            if overflow:
                sample = "；".join(overflow[:3])
                problems.append(f"第 {page_no} 页文本触及或越过页面边界: {sample}")

            for span in spans:
                _, y0, _, y1 = span["bbox"]
                if y1 >= height * HEADING_BOTTOM_RATIO and _looks_like_heading(span["text"], span["size"]):
                    problems.append(f"第 {page_no} 页存在页尾孤立标题: {span['text'][:40]}")
                    break

        if page_count > 1:
            last_chars, last_blocks = page_stats[-1]
            previous_chars = max((chars for chars, _ in page_stats[:-1]), default=0)
            if (
                0 < last_chars < SPARSE_TEXT_CHARS
                and last_blocks <= SPARSE_BLOCKS
                and previous_chars >= SPARSE_TEXT_CHARS * 2
            ):
                problems.append(
                    f"末页残余内容过少（{last_chars} 字、{last_blocks} 个文本块），疑似孤立尾行或异常分页"
                )

        # Internal pages with almost no text usually indicate an accidental page or table spill.
        for index, (char_count, block_count) in enumerate(page_stats[1:-1], start=2):
            if 0 < char_count < 30 and block_count <= 1:
                warnings.append(f"第 {index} 页内容异常稀疏，建议检查标题或表格分页")

        report.append(f"文本坐标检查完成：{len(problems)} 个问题，{len(warnings)} 个提醒")
        return report, problems, warnings
    finally:
        doc.close()


def main():
    if len(sys.argv) != 2:
        print("用法: python validate_pdf.py <output.pdf>")
        return 1
    report, problems, warnings = validate_pdf(os.path.abspath(sys.argv[1]))
    print("\n".join(report))
    if warnings:
        print("\n提醒:")
        for item in warnings:
            print(f"- {item}")
    if problems:
        print("\n问题清单:")
        for item in problems:
            print(f"- {item}")
        return 2
    print("\nOK: PDF 文本坐标验收通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
