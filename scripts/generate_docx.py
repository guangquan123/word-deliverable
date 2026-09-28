# -*- coding: utf-8 -*-
"""
word-deliverable / generate_docx.py
按固定规范生成正式交付 Word 文档（.docx）。

用法:
    python generate_docx.py <spec.json> <output.docx>

spec.json 结构见 SKILL.md「输入 JSON Schema」。
生成后需再运行 update_fields.py 更新 TOC 与页码字段。
"""
import json
import os
import sys
import copy

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_LINE_SPACING,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------- 常量 ----------
FONT_HEI = "黑体"      # 黑体
FONT_FANG = "仿宋"     # 仿宋
FONT_SONG = "宋体"     # 宋体（备用）
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_GRAY = "D9D9D9"      # 浅灰底（表头）
COLOR_GRAY_LIGHT = "F2F2F2"  # 极浅灰底（封面信息表左列，复刻样板）
COLOR_HEADER_LINE = "000000"

# 排版 Profile：只调整可维护的版式参数，不改变文档语义与黑白灰基调。
LAYOUT_PROFILES = {
    "government": {
        "label": "政企正式", "body_font": FONT_FANG, "body_size": 14,
        "line_spacing": 28, "first_indent": 28,
        "h_sizes": (22, 16, 14, 14, 14, 14),
        "margins": (3.7, 3.5, 2.8, 2.6), "table_size": 10.5,
    },
    "consulting": {
        "label": "咨询报告", "body_font": FONT_SONG, "body_size": 11,
        "line_spacing": 20, "first_indent": 22,
        "h_sizes": (20, 15, 12.5, 11.5, 11, 10.5),
        "margins": (2.8, 2.8, 2.6, 2.4), "table_size": 9.5,
    },
    "technical": {
        "label": "技术方案", "body_font": FONT_SONG, "body_size": 11,
        "line_spacing": 20, "first_indent": 22,
        "h_sizes": (20, 15, 12.5, 11.5, 11, 10.5),
        "margins": (2.8, 2.8, 2.5, 2.3), "table_size": 9,
    },
    "reader": {
        "label": "简洁阅读", "body_font": FONT_SONG, "body_size": 11.5,
        "line_spacing": 21, "first_indent": 23,
        "h_sizes": (19, 14.5, 12.5, 11.5, 11, 10.5),
        "margins": (2.6, 2.6, 2.6, 2.6), "table_size": 10,
    },
}
LAYOUT_PROFILE_ALIASES = {
    "政企正式": "government", "consulting-report": "consulting", "咨询报告": "consulting",
    "technical-solution": "technical", "技术方案": "technical",
    "simple-reading": "reader", "简洁阅读": "reader", "formal": "government",
}
ACTIVE_PROFILE = LAYOUT_PROFILES["government"]


def _resolve_layout_profile(spec):
    raw = str(spec.get("layout_profile", "government")).strip()
    key = LAYOUT_PROFILE_ALIASES.get(raw, raw)
    if key not in LAYOUT_PROFILES:
        allowed = ", ".join(LAYOUT_PROFILES)
        raise ValueError(f"layout_profile={raw!r} 不合法，可选: {allowed}")
    return key, LAYOUT_PROFILES[key]


# ---------- XML 辅助 ----------
def _set_run_font(run, cn_font, size_pt, bold=False, color=None):
    """设置 run 中英文字体、字号、加粗、颜色（中文需设 eastAsia）。"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color if color is not None else COLOR_BLACK
    run.font.name = cn_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), cn_font)
    rFonts.set(qn("w:hAnsi"), cn_font)
    rFonts.set(qn("w:eastAsia"), cn_font)
    rFonts.set(qn("w:cs"), cn_font)


def _set_cell_bg(cell, hex_color):
    """单元格底色。"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)


def _set_table_borders(table, color="000000", sz="4"):
    """表格全边框（细黑线）。sz 单位 1/8 pt，4 = 0.5pt。"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def _set_table_fixed_layout(table):
    """显式写入固定布局，避免仅设置 autofit=False 时不同渲染器行为不一致。"""
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")


def _set_table_grid_widths(table, widths):
    """固定表格底层网格列宽，确保 Word 与 LibreOffice 均遵守列宽比例。"""
    tbl_grid = table._tbl.tblGrid
    grid_cols = tbl_grid.findall(qn("w:gridCol"))
    for grid_col, width in zip(grid_cols, widths):
        grid_col.set(qn("w:w"), str(width.twips))


def _set_repeat_header(row):
    """设置表格行重复标题。"""
    trPr = row._tr.get_or_add_trPr()
    th = trPr.find(qn("w:tblHeader"))
    if th is None:
        th = OxmlElement("w:tblHeader")
        trPr.append(th)
    th.set(qn("w:val"), "true")


def _set_row_cant_split(row):
    """行不跨页拆分。"""
    trPr = row._tr.get_or_add_trPr()
    cs = trPr.find(qn("w:cantSplit"))
    if cs is None:
        cs = OxmlElement("w:cantSplit")
        trPr.append(cs)


def _add_field(paragraph, instr, placeholder=""):
    """在段落中插入 Word 字段（PAGE/NUMPAGES/TOC）。"""
    run = paragraph.add_run()
    fldBegin = OxmlElement("w:fldChar")
    fldBegin.set(qn("w:fldCharType"), "begin")
    run._r.append(fldBegin)

    run2 = paragraph.add_run()
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = instr
    run2._r.append(instrText)

    run3 = paragraph.add_run()
    fldSep = OxmlElement("w:fldChar")
    fldSep.set(qn("w:fldCharType"), "separate")
    run3._r.append(fldSep)

    run4 = paragraph.add_run(placeholder)

    run5 = paragraph.add_run()
    fldEnd = OxmlElement("w:fldChar")
    fldEnd.set(qn("w:fldCharType"), "end")
    run5._r.append(fldEnd)


def _set_para_border_bottom(paragraph, color="000000", sz="6"):
    """段落底部横线（页眉用）。"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = pBdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        pBdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)


def _set_para_keep_with_next(paragraph):
    paragraph.paragraph_format.keep_with_next = True


def _set_para_keep_lines(paragraph):
    paragraph.paragraph_format.keep_together = True


# ---------- 多级编号 ----------
# 编号 Profile 定义：每种 profile 给出 6 级的 (numFmt, lvlText, isLgl, suff, lvlRestart)
# lvlRestart: OOXML 使用 1 起算的层级值；1 表示上一级（ilvl=0）变化时重置，0 表示永不重置
# formal: 第一章/1.1/1.1.1/1.1.1.1/（1）/1)  —— 长篇正式文档默认
# lightweight: 一、/（一）/1./（1）/1)/① —— 会议纪要/工作记录等轻量文档
# 一级 numFmt 受 chapter_num_format 控制（chinese=chineseCounting / arabic=decimal），仅 formal profile 生效
def _build_levels_def(profile, chapter_num_format, heading_align_left=False):
    """返回 6 级 levels_def 列表。

    heading_align_left=True 时把所有级别 ind_val 归零（编号与正文都从左边距开始），
    并对 formal profile 设 suff=space + lvlText 去尾空格，使编号后统一有一个空格
    （避免 LibreOffice 渲染 CJK 编号时丢失 lvlText 尾空格导致 H1 无间距）。
    """
    if profile == "lightweight":
        # 轻量 Profile：全部用中文/数字编号，不用 isLgl
        # 一级"一、"用 chineseCounting；二级"（一）"用 chineseCounting；
        # 三级"1."用 decimal；四级"（1）"用 decimal；五级"1)"用 decimal；六级"①"用 decimal（圈码靠字体/符号，这里用 decimal 兜底，Word 可后续手动改 chuo）
        ind = 0 if heading_align_left else None
        return [
            # ilvl, start, numFmt, lvlText, ind_val, isLgl, suff, lvlRestart
            (0, 1, "chineseCounting", "一、",   ind or 432, False, None, None),
            (1, 1, "chineseCounting", "（%2）", ind or 575, False, None, 1),
            (2, 1, "decimal",         "%3.",   ind or 720, False, None, 2),
            (3, 1, "decimal",         "（%4）", ind or 864, False, None, 3),
            (4, 1, "decimal",         "%5）",  ind or 1008, False, None, 4),
            (5, 1, "decimal",         "%6.",   ind or 1152, False, None, 5),
        ]
    # formal Profile（默认）
    level0_fmt = "decimal" if chapter_num_format == "arabic" else "chineseCounting"
    if heading_align_left:
        # 归零缩进 + suff=space + lvlText 去尾空格，编号后统一靠空格分隔
        return [
            # ilvl, start, numFmt, lvlText, ind_val, isLgl, suff, lvlRestart
            (0, 1, level0_fmt,      "第%1章",        0, False, "space", None),
            (1, 1, "decimal",       "%1.%2.",        0, True,  "space", 1),
            (2, 1, "decimal",       "%1.%2.%3.",     0, True,  "space", 2),
            (3, 1, "decimal",       "%1.%2.%3.%4.",  0, True,  "space", 3),
            (4, 1, "decimal",       "（%5）",         0, False, "space", 4),
            (5, 1, "decimal",       "%6）",           0, False, "space", 5),
        ]
    return [
        # ilvl, start, numFmt, lvlText, ind_val, isLgl, suff, lvlRestart
        # 一级"第%1章"，suff=nothing（lvlText 已含尾空格）
        (0, 1, level0_fmt,      "第%1章 ",       432, False, "nothing", None),
        # 二三四级"1.1 / 1.1.1 / 1.1.1.1"，decimal + isLgl 强制引用用阿拉伯
        (1, 1, "decimal",       "%1.%2.",        575, True,  None,      1),
        (2, 1, "decimal",       "%1.%2.%3.",     720, True,  None,      2),
        (3, 1, "decimal",       "%1.%2.%3.%4.",  864, True,  None,      3),
        # 五级“（1）”局部编号，每个 H4 下重新编号；OOXML lvlRestart=4 对应 ilvl=3
        (4, 1, "decimal",       "（%5）",       1008, False, None,      4),
        # 六级“1)”局部编号，每个 H5 下重新编号；OOXML lvlRestart=5 对应 ilvl=4
        (5, 1, "decimal",       "%6）",         1152, False, None,      5),
    ]


def _create_multilevel_numbering(document, chapter_num_format="chinese", numbering_profile="formal",
                                  abstract_num_id=100, num_id=100, heading_align_left=False):
    """创建多级编号，支持 H1-H6 全 6 级，两种 Numbering Profile。

    chapter_num_format: "chinese"(第一章) / "arabic"(第1章)，仅 formal profile 一级生效
    numbering_profile:  "formal"(第一章/1.1/1.1.1/1.1.1.1/（1）/1)) / "lightweight"(一、/（一）/1./（1）/1)/①)
    heading_align_left: True 时所有级别缩进归零（编号与正文都从左边距开始）

    关键：formal profile 二三级用 decimal + isLgl=True，isLgl(legal format) 强制 %N 引用上级编号时
    用阿拉伯数字渲染，因此「第一章(中文) + 1.1(阿拉伯)」可在同一编号体系下共存。
    H5/H6 用局部编号（（1）/1）），通过 lvlRestart 在上级变化时重置。
    """
    numbering_part = document.part.numbering_part
    numbering = numbering_part.element

    abstractNum = OxmlElement("w:abstractNum")
    abstractNum.set(qn("w:abstractNumId"), str(abstract_num_id))

    mlp = OxmlElement("w:multiLevelType")
    mlp.set(qn("w:val"), "multilevel")
    abstractNum.append(mlp)

    levels_def = _build_levels_def(numbering_profile, chapter_num_format, heading_align_left)

    # fmt / lvlText / ind(left=hanging) / isLgl / suff / lvlRestart / rFonts
    # isLgl：legal format，强制 %N 引用用阿拉伯数字（解决 %1 引用 chineseCounting 得"一"的问题）
    # suff=nothing：编号与文本间无分隔符（lvlText 已含尾空格）
    # lvlRestart：OOXML 采用 1 起算；值 N 表示第 N 级（对应 ilvl=N-1）变化时本级重置，0 表示永不重置
    for ilvl, start, fmt, text, ind_val, is_lgl, suff, lvl_restart in levels_def:
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))
        startEl = OxmlElement("w:start")
        startEl.set(qn("w:val"), str(start))
        lvl.append(startEl)
        numFmt = OxmlElement("w:numFmt")
        numFmt.set(qn("w:val"), fmt)
        lvl.append(numFmt)
        if suff is not None:
            suffEl = OxmlElement("w:suff")
            suffEl.set(qn("w:val"), suff)
            lvl.append(suffEl)
        if is_lgl:
            isLglEl = OxmlElement("w:isLgl")
            lvl.append(isLglEl)
        if lvl_restart is not None:
            lvlRestartEl = OxmlElement("w:lvlRestart")
            lvlRestartEl.set(qn("w:val"), str(lvl_restart))
            lvl.append(lvlRestartEl)
        lvlText = OxmlElement("w:lvlText")
        lvlText.set(qn("w:val"), text)
        lvl.append(lvlText)
        lvlJc = OxmlElement("w:lvlJc")
        lvlJc.set(qn("w:val"), "left")
        lvl.append(lvlJc)
        pPr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(ind_val))
        ind.set(qn("w:hanging"), str(ind_val))
        pPr.append(ind)
        lvl.append(pPr)
        # rFonts hint=eastAsia（复刻样板）
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:hint"), "eastAsia")
        rPr.append(rFonts)
        lvl.append(rPr)
        abstractNum.append(lvl)

    numbering.insert(0, abstractNum)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    absId = OxmlElement("w:abstractNumId")
    absId.set(qn("w:val"), str(abstract_num_id))
    num.append(absId)
    numbering.append(num)

    return num_id


def _bind_style_numbering(style, num_id, ilvl):
    """把段落样式绑定到多级编号某级别。"""
    pPr = style.element.get_or_add_pPr()
    # 移除已有 numPr
    old = pPr.find(qn("w:numPr"))
    if old is not None:
        pPr.remove(old)
    numPr = OxmlElement("w:numPr")
    ilvlEl = OxmlElement("w:ilvl")
    ilvlEl.set(qn("w:val"), str(ilvl))
    numPr.append(ilvlEl)
    numIdEl = OxmlElement("w:numId")
    numIdEl.set(qn("w:val"), str(num_id))
    numPr.append(numIdEl)
    pPr.append(numPr)


# ---------- 样式创建 ----------
def _ensure_style(styles, name, style_type=WD_STYLE_TYPE.PARAGRAPH):
    """获取或新建样式。"""
    try:
        return styles[name]
    except KeyError:
        return styles.add_style(name, style_type)


def setup_styles(document, chapter_num_format="chinese", numbering_profile="formal", profile=None,
                  heading_align_left=False):
    """创建/配置全部样式。版式参数由 layout_profile 集中控制。"""
    profile = profile or ACTIVE_PROFILE
    body_font = profile["body_font"]
    body_size = profile["body_size"]
    line_spacing = profile["line_spacing"]
    first_indent = profile["first_indent"]
    h_sizes = profile["h_sizes"]
    table_size = profile["table_size"]
    styles = document.styles
    num_id = _create_multilevel_numbering(document, chapter_num_format=chapter_num_format,
                                          numbering_profile=numbering_profile,
                                          heading_align_left=heading_align_left)

    # 正文基础样式，同时提供语义明确的“正文”Style。
    s = styles["Normal"]
    s.font.name = body_font
    s.font.size = Pt(body_size)
    s.font.color.rgb = COLOR_BLACK
    rPr = s.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), body_font)
    pf = s.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Pt(first_indent)
    pf.line_spacing = Pt(line_spacing)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    body_style = _ensure_style(styles, "正文")
    body_style.base_style = s
    body_style.font.name = body_font
    body_style.font.size = Pt(body_size)
    body_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body_style.paragraph_format.first_line_indent = Pt(first_indent)
    body_style.paragraph_format.line_spacing = Pt(line_spacing)
    body_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY

    # 目录条目：覆盖 Normal 的两端对齐与首行缩进，并将页码固定在正文区右侧。
    # Word 更新 TOC 时会应用 TOC 1–TOC 6 样式；若不显式设置，中文目录可能
    # 被分散对齐，页码制表符也可能落到下一行。
    page_top, page_bottom, page_left, page_right = profile["margins"]
    toc_right = Cm(21.0 - page_left - page_right)
    toc_indents = (0.0, 0.6, 1.2, 1.8, 2.4, 3.0)
    for level in range(1, 7):
        toc_style = _ensure_style(styles, f"TOC {level}")
        toc_style.font.name = body_font
        toc_style.font.size = Pt(min(12, body_size))
        toc_style.font.bold = level == 1
        toc_style.font.color.rgb = COLOR_BLACK
        rPr = toc_style.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rFonts.set(qn(attr), body_font)
        pf = toc_style.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Pt(0)
        pf.left_indent = Cm(toc_indents[level - 1])
        pf.right_indent = Pt(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(3 if level == 1 else 0)
        pf.line_spacing = 1.25
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.tab_stops.clear_all()
        pf.tab_stops.add_tab_stop(
            toc_right,
            WD_TAB_ALIGNMENT.RIGHT,
            WD_TAB_LEADER.DOTS,
        )

    # Heading 1（第一章 XXX）—— 复刻样板：sz=44(22pt), line=576(auto=2.4倍),
    # before=340 after=330, ind left=432 hanging=432, keepNext, keepLines
    h1 = styles["Heading 1"]
    h1.font.name = FONT_HEI
    h1.font.size = Pt(h_sizes[0])
    h1.font.bold = True
    h1.font.color.rgb = COLOR_BLACK
    rPr = h1.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_HEI)
    pf = h1.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Pt(0)
    pf.line_spacing = 2.4
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_before = Pt(17)   # 340 twips ≈ 17pt
    pf.space_after = Pt(16.5)  # 330 twips ≈ 16.5pt
    pf.keep_with_next = True
    pf.keep_together = True
    _bind_style_numbering(h1, num_id, 0)

    # Heading 2（1.1. XXX）—— 复刻：黑体 16pt, line=413(auto≈1.72), before=260 after=260,
    # ind left=575 hanging=575
    h2 = styles["Heading 2"]
    h2.font.name = FONT_HEI
    h2.font.size = Pt(h_sizes[1])
    h2.font.bold = True
    h2.font.color.rgb = COLOR_BLACK
    rPr = h2.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_HEI)
    pf = h2.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Pt(0)
    pf.line_spacing = 1.72
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_before = Pt(13)   # 260 twips ≈ 13pt
    pf.space_after = Pt(13)
    pf.keep_with_next = True
    pf.keep_together = True
    _bind_style_numbering(h2, num_id, 1)

    # Heading 3（1.1.1. XXX）—— 黑体 14pt（规范：H3=14pt）, ind left=720 hanging=720
    h3 = styles["Heading 3"]
    h3.font.name = FONT_HEI
    h3.font.size = Pt(h_sizes[2])
    h3.font.bold = True
    h3.font.color.rgb = COLOR_BLACK
    rPr = h3.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_HEI)
    pf = h3.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Pt(0)
    pf.line_spacing = 1.72
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_before = Pt(13)
    pf.space_after = Pt(13)
    pf.keep_with_next = True
    pf.keep_together = True
    _bind_style_numbering(h3, num_id, 2)

    # Heading 4（1.1.1.1. XXX）—— 黑体 14pt 加粗
    try:
        h4 = styles["Heading 4"]
        h4.font.name = FONT_HEI
        h4.font.size = Pt(h_sizes[3])
        h4.font.bold = True
        h4.font.color.rgb = COLOR_BLACK
        rPr = h4.element.get_or_add_rPr()
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rFonts.set(qn(attr), FONT_HEI)
        pf = h4.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(6)
        pf.space_after = Pt(6)
        pf.keep_with_next = True
        _bind_style_numbering(h4, num_id, 3)
    except KeyError:
        pass

    # Heading 5（（1） XXX）—— 正文字号(14pt) 加粗，局部编号，不降级为 Normal
    # 规范：H5 显示"（1）"在每个 H4 下重新编号，但 Word 内部仍是 Heading 5 Style
    try:
        h5 = styles["Heading 5"]
        h5.font.name = FONT_HEI
        h5.font.size = Pt(h_sizes[4])
        h5.font.bold = True
        h5.font.color.rgb = COLOR_BLACK
        rPr = h5.element.get_or_add_rPr()
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rFonts.set(qn(attr), FONT_HEI)
        pf = h5.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(4)
        pf.space_after = Pt(4)
        pf.keep_with_next = True
        _bind_style_numbering(h5, num_id, 4)
    except KeyError:
        pass

    # Heading 6（1） XXX）—— 正文字号(14pt) 可加粗，局部编号，视觉层级低于 H5
    # 规范：H6 显示"1)"在每个 H5 下重新编号，Word 内部仍是 Heading 6 Style
    try:
        h6 = styles["Heading 6"]
        h6.font.name = FONT_HEI
        h6.font.size = Pt(h_sizes[5])
        h6.font.bold = True
        h6.font.color.rgb = COLOR_BLACK
        rPr = h6.element.get_or_add_rPr()
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rFonts.set(qn(attr), FONT_HEI)
        pf = h6.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(2)
        pf.space_after = Pt(2)
        pf.keep_with_next = True
        _bind_style_numbering(h6, num_id, 5)
    except KeyError:
        pass

    # 小节标题（不进目录）
    sn = _ensure_style(styles, "小节标题")
    sn.font.name = FONT_HEI
    sn.font.size = Pt(h_sizes[3])
    sn.font.bold = True
    sn.font.color.rgb = COLOR_BLACK
    rPr = sn.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_HEI)
    pf = sn.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Pt(0)
    pf.line_spacing = Pt(line_spacing)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = Pt(6)
    pf.space_after = Pt(4)
    pf.keep_with_next = True
    # 不绑定编号、不进目录（不基于 Heading）

    # 列表正文
    sl = _ensure_style(styles, "列表正文")
    sl.font.name = body_font
    sl.font.size = Pt(body_size)
    sl.font.color.rgb = COLOR_BLACK
    rPr = sl.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), body_font)
    pf = sl.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = Pt(line_spacing)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    # 表题 / 图题
    caption_size = min(12, max(10, body_size))
    st = _ensure_style(styles, "表题")
    st.font.name = FONT_HEI
    st.font.size = Pt(caption_size)
    st.font.bold = True
    st.font.color.rgb = COLOR_BLACK
    rPr = st.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_HEI)
    pf = st.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(6)
    pf.space_after = Pt(3)
    pf.keep_with_next = True

    sf = _ensure_style(styles, "图题")
    sf.font.name = FONT_HEI
    sf.font.size = Pt(caption_size)
    sf.font.bold = True
    sf.font.color.rgb = COLOR_BLACK
    rPr = sf.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_HEI)
    pf = sf.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(3)
    pf.space_after = Pt(6)
    pf.keep_with_next = True

    # 表头 / 表格正文 / 说明注释
    sh = _ensure_style(styles, "表头")
    sh.font.name = FONT_HEI
    sh.font.size = Pt(table_size)
    sh.font.bold = True
    sh.font.color.rgb = COLOR_BLACK
    rPr = sh.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_HEI)
    pf = sh.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    stb = _ensure_style(styles, "表格正文")
    stb.font.name = body_font
    stb.font.size = Pt(table_size)
    stb.font.color.rgb = COLOR_BLACK
    rPr = stb.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), body_font)
    pf = stb.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    snote = _ensure_style(styles, "说明注释")
    snote.font.name = body_font
    snote.font.size = Pt(table_size)
    snote.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    rPr = snote.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), body_font)
    pf = snote.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    return num_id


# ---------- 页面设置 ----------
def setup_page(section, profile=None, landscape=False):
    profile = profile or ACTIVE_PROFILE
    top, bottom, left, right = profile["margins"]
    section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    section.page_width = Cm(29.7 if landscape else 21.0)
    section.page_height = Cm(21.0 if landscape else 29.7)
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.75)


def _usable_width_cm(section):
    # python-docx 的长度相减结果是 EMU 整数；1 cm = 360000 EMU。
    return (section.page_width - section.left_margin - section.right_margin) / 360000.0


def _copy_header_footer_links(source, target):
    """新 Section 继承前一节的页眉页脚，避免横向页丢失字段。"""
    target.header.is_linked_to_previous = True
    target.footer.is_linked_to_previous = True
    target.first_page_header.is_linked_to_previous = True
    target.first_page_footer.is_linked_to_previous = True
    target.different_first_page_header_footer = False


# ---------- 封面 ----------
def build_cover(document, spec):
    """封面内容（单节架构：封面是首页，由 different_first_page 控制无页眉页脚）。"""
    # 顶部留白（小字号省空间）
    for _ in range(2):
        p = document.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    # 项目名称
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(24)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p.add_run(spec.get("project_name", ""))
    _set_run_font(r, FONT_HEI, 36, bold=True)

    # 文档名称（字间加空格美化，复刻样板"需 求 调 研 大 纲"）
    # doc_name 优先；未设则回退到 document_title（从 Markdown 首 # 识别的文档标题）
    doc_name = spec.get("doc_name") or spec.get("document_title") or "未命名文档"
    spaced_name = " ".join(doc_name) if len(doc_name) <= 12 else doc_name
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(18)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    r = p.add_run(spaced_name)
    _set_run_font(r, FONT_HEI, 36, bold=True)

    # 副标题（可选）
    if spec.get("subtitle"):
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        r = p.add_run(spec["subtitle"])
        _set_run_font(r, FONT_HEI, 18, bold=False)

    # 中部留白（有副标题时少留，无副标题时多留）
    mid_blanks = 3 if spec.get("subtitle") else 5
    for _ in range(mid_blanks):
        p = document.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    # 封面信息表（复刻样板：左列 F2F2F2 底色，不设重复标题行）
    rows = [
        ("版    本", spec.get("version", "V1.0")),
        ("编制单位", spec.get("org", "")),
        ("编制日期", spec.get("date", "")),
    ]
    table = document.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_fixed_layout(table)
    _set_table_borders(table)
    cover_widths = [Cm(5.2), Cm(10.4)]
    _set_table_grid_widths(table, cover_widths)
    # 列宽：左 1/3，右 2/3（正文区宽约 15.6cm）
    for r_idx, (k, v) in enumerate(rows):
        _set_row_cant_split(table.rows[r_idx])
        c0 = table.cell(r_idx, 0)
        c1 = table.cell(r_idx, 1)
        c0.width = Cm(5.2)
        c1.width = Cm(10.4)
        _set_cell_bg(c0, COLOR_GRAY_LIGHT)  # F2F2F2，复刻样板
        c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.first_line_indent = Pt(0)
        run0 = p0.add_run(k)
        _set_run_font(run0, FONT_HEI, 12, bold=True)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p1.paragraph_format.first_line_indent = Pt(0)
        run1 = p1.add_run(v)
        _set_run_font(run1, FONT_FANG, 12, bold=False)

    # 封面后分页（同节内分页，非新节）
    document.add_page_break()


# ---------- 目录 ----------
def build_toc(document, spec):
    """目录内容（单节架构：与封面同节，封面后已分页）。"""
    # 目录标题
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run("目录")
    _set_run_font(r, FONT_HEI, 18, bold=True)

    # TOC 字段（复刻样板：TOC \o "1-3" \h \u）
    # toc_levels 可配置：默认 "1-3"，规格说明书/标准规范可 "1-4"，用户明确要求可 "1-5"/"1-6"
    toc_levels = spec.get("toc_levels", "1-3")
    # 校验格式：应为 "1-N" 且 N 在 3-6
    import re as _re
    if not _re.match(r"^1-[3-6]$", toc_levels):
        print(f"警告: toc_levels='{toc_levels}' 不合法（应为 1-3 ~ 1-6），回退为 '1-3'")
        toc_levels = "1-3"
    toc_p = document.add_paragraph()
    toc_p.paragraph_format.first_line_indent = Pt(0)
    _add_field(toc_p, f'TOC \\o "{toc_levels}" \\h \\u', placeholder="（目录将在打开后自动生成）")

    # 目录后分页
    document.add_page_break()


def _setup_footer_pagenum(footer):
    """页脚居中：第 X 页 共 Y 页（复刻样板 PAGE \\* MERGEFORMAT 字段格式）。"""
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    # 清空已有 run
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    r1 = p.add_run("第 ")
    _set_run_font(r1, FONT_SONG, 10.5)
    _add_field(p, " PAGE  \\* MERGEFORMAT ", placeholder="1")
    r2 = p.add_run(" 页 共 ")
    _set_run_font(r2, FONT_SONG, 10.5)
    _add_field(p, " NUMPAGES  \\* MERGEFORMAT ", placeholder="1")
    r3 = p.add_run(" 页")
    _set_run_font(r3, FONT_SONG, 10.5)


def _setup_header(header, project_name, doc_name):
    """页眉：项目名+文档名+Tab，两端对齐，下方细横线（复刻样板）。"""
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Pt(0)
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    run1 = p.add_run(project_name)
    _set_run_font(run1, FONT_SONG, 10.5)
    run2 = p.add_run(doc_name)
    _set_run_font(run2, FONT_SONG, 10.5)
    # Tab 到右侧
    run3 = p.add_run("\t")
    _set_run_font(run3, FONT_SONG, 10.5)
    _set_para_border_bottom(p)


# ---------- 文档控制 ----------
def build_doc_control(document, spec):
    """文档控制模块（目录之后，独立成页；目录后已分页）。"""
    # 文档控制标题
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("文档控制")
    _set_run_font(r, FONT_HEI, 16, bold=True)

    rev = spec.get("revision_history") or [
        {
            "version": spec.get("version", "V1.0"),
            "date": spec.get("date", ""),
            "content": "初版编制",
            "author": spec.get("author", spec.get("org", "")),
            "reviewer": "",
            "status": spec.get("status", "正式发布"),
        }
    ]
    headers = ["版本", "修订日期", "修订内容", "编制", "审核", "状态"]
    table = document.add_table(rows=1 + len(rev), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_fixed_layout(table)
    _set_table_borders(table)
    # 修订记录列宽：为“修订内容”保留主要空间，避免长文本被拆成大量短行
    total = 15.6
    width_ratios = [10, 15, 37, 12, 12, 14]
    col_widths = [Cm(total * ratio / 100.0) for ratio in width_ratios]
    _set_table_grid_widths(table, col_widths)
    # 表头
    for i, h in enumerate(headers):
        c = table.cell(0, i)
        c.width = col_widths[i]
        _set_cell_bg(c, COLOR_GRAY)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        run = p.add_run(h)
        _set_run_font(run, FONT_HEI, 10.5, bold=True)
    _set_repeat_header(table.rows[0])
    _set_row_cant_split(table.rows[0])
    # 数据行
    for ri, item in enumerate(rev, start=1):
        vals = [
            item.get("version", ""),
            item.get("date", ""),
            item.get("content", ""),
            item.get("author", ""),
            item.get("reviewer", ""),
            item.get("status", ""),
        ]
        for ci, v in enumerate(vals):
            c = table.cell(ri, ci)
            c.width = col_widths[ci]
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            # 短字段居中，长文本（修订内容）左对齐
            if ci == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            run = p.add_run(str(v))
            _set_run_font(run, FONT_FANG, 10.5)
        _set_row_cant_split(table.rows[ri])

    # 文档控制后分页
    document.add_page_break()


# ---------- 正文渲染 ----------
def render_body(document, spec, spec_dir, profile):
    """渲染 body blocks；图片相对路径以 spec 所在目录为基准。"""
    include_figures = spec.get("include_figures", False) is True
    blocks = spec.get("body", [])
    for block in blocks:
        _render_block(document, block, spec_dir, profile, include_figures)
    # 文档以表格结尾时，OOXML 要求表格后必须有一个段落；若不显式添加，Word/LibreOffice
    # 会隐式插入一个默认样式段落（正文 12pt+行距），可能溢出到下一页形成孤行空页。
    # 这里添加一个极紧凑尾段落（1pt 字号、零间距），满足结构要求又不占可见空间。
    if blocks and blocks[-1].get("type") == "table":
        p = document.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        r = p.add_run("")
        _set_run_font(r, FONT_SONG, 1, bold=False)


def _render_block(document, block, spec_dir, profile, include_figures):
    t = block.get("type")
    heading_types = {f"h{i}": i for i in range(1, 7)}
    if t in heading_types:
        level = heading_types[t]
        p = document.add_paragraph(style=f"Heading {level}")
        r = p.add_run(block["text"])
        _set_run_font(r, FONT_HEI, profile["h_sizes"][level - 1], bold=True)
    elif t == "p":
        p = document.add_paragraph(style="正文")
        _add_runs_with_bold(p, block["text"])
    elif t == "note":
        p = document.add_paragraph(style="小节标题")
        r = p.add_run(block["title"])
        _set_run_font(r, FONT_HEI, profile["h_sizes"][3], bold=True)
        if block.get("text"):
            p2 = document.add_paragraph(style="正文")
            _add_runs_with_bold(p2, block["text"])
    elif t == "list":
        _render_list(document, block, profile)
    elif t == "table":
        _render_table(document, block, profile)
    elif t == "figure":
        if include_figures:
            _render_figure(document, block, spec_dir, profile)
    elif t == "pagebreak":
        document.add_page_break()
    else:
        p = document.add_paragraph(style="正文")
        _add_runs_with_bold(p, str(block))


def _add_runs_with_bold(paragraph, text):
    """支持 **加粗** 内联语法的正文 run 写入。"""
    text = str(text)
    if "**" not in text:
        paragraph.add_run(text)
        return
    parts = text.split("**")
    for i, part in enumerate(parts):
        if part == "":
            continue
        r = paragraph.add_run(part)
        if i % 2 == 1:
            r.font.bold = True


def _render_list(document, block, profile):
    items = block.get("items", [])
    ordered = block.get("ordered", False)
    for i, item in enumerate(items):
        p = document.add_paragraph(style="列表正文")
        prefix = f"（{_cn_num(i + 1)}）" if ordered else "• "
        p.add_run(prefix + str(item))
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.left_indent = Pt(profile["first_indent"])


def _cn_num(n):
    """简单中文数字（1-20）。"""
    m = "零一二三四五六七八九十"
    if n <= 10:
        return m[n]
    if n < 20:
        return "十" + (m[n - 10] if n > 10 else "")
    if n == 20:
        return "二十"
    return str(n)


def _validate_table_block(block):
    headers = block.get("headers")
    rows = block.get("rows", [])
    if not isinstance(headers, list) or not headers:
        raise ValueError("表格 headers 必须是非空列表，列数不能为 0")
    if not isinstance(rows, list):
        raise ValueError("表格 rows 必须是列表")
    ncols = len(headers)
    for index, row in enumerate(rows, start=1):
        if not isinstance(row, (list, tuple)) or len(row) != ncols:
            actual = len(row) if isinstance(row, (list, tuple)) else "非列表"
            raise ValueError(f"表格第 {index} 行列数为 {actual}，应与表头列数 {ncols} 一致")
    widths = block.get("widths")
    if widths is not None:
        if not isinstance(widths, list) or len(widths) != ncols:
            raise ValueError(f"表格 widths 长度必须等于列数 {ncols}")
        try:
            widths = [float(value) for value in widths]
        except (TypeError, ValueError):
            raise ValueError("表格 widths 必须全部为数字")
        if any(value <= 0 for value in widths):
            raise ValueError("表格 widths 必须全部为正数")
    return headers, rows, widths


def _table_width_weights(headers, rows, explicit_widths=None):
    if explicit_widths:
        total = sum(explicit_widths)
        return [value / total for value in explicit_widths]
    scores = []
    for col, header in enumerate(headers):
        lengths = [len(str(header))] + [len(str(row[col] if row[col] is not None else "")) for row in rows]
        representative = max(4, min(30, max(lengths, default=4)))
        scores.append(representative ** 0.65)
    total = sum(scores)
    return [score / total for score in scores]


def _render_table(document, block, profile):
    headers, rows, widths = _validate_table_block(block)
    ncols = len(headers)
    use_landscape = block.get("landscape") is True or ncols >= 7
    previous_section = document.sections[-1]
    if use_landscape and previous_section.orientation != WD_ORIENT.LANDSCAPE:
        table_section = document.add_section(WD_SECTION.NEW_PAGE)
        setup_page(table_section, profile, landscape=True)
        _copy_header_footer_links(previous_section, table_section)
    else:
        table_section = previous_section

    if block.get("caption"):
        p = document.add_paragraph(style="表题")
        p.add_run(str(block["caption"]))

    available_cm = _usable_width_cm(table_section)
    weights = _table_width_weights(headers, rows, widths)
    col_widths = [Cm(available_cm * weight) for weight in weights]
    table = document.add_table(rows=1 + len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_fixed_layout(table)
    _set_table_borders(table)
    _set_table_grid_widths(table, col_widths)

    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.width = col_widths[i]
        _set_cell_bg(cell, COLOR_GRAY)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        paragraph = cell.paragraphs[0]
        paragraph.style = document.styles["表头"]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Pt(0)
        paragraph.add_run(str(header))
    _set_repeat_header(table.rows[0])
    _set_row_cant_split(table.rows[0])

    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            cell = table.cell(row_index, col_index)
            cell.width = col_widths[col_index]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            paragraph = cell.paragraphs[0]
            paragraph.style = document.styles["表格正文"]
            paragraph.paragraph_format.first_line_indent = Pt(0)
            text = str(value) if value is not None else ""
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if len(text) > 10 else WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run(text)
        _set_row_cant_split(table.rows[row_index])

    if use_landscape:
        portrait_section = document.add_section(WD_SECTION.NEW_PAGE)
        setup_page(portrait_section, profile, landscape=False)
        _copy_header_footer_links(table_section, portrait_section)


def _render_figure(document, block, spec_dir, profile):
    img = str(block.get("image", "")).strip()
    if not img:
        raise ValueError("include_figures=true 时 figure.image 不能为空")
    img_path = img if os.path.isabs(img) else os.path.join(spec_dir, img)
    if not os.path.isfile(img_path):
        raise FileNotFoundError(f"图片文件不存在: {img_path}")
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run()
    try:
        max_width = min(14.5, _usable_width_cm(document.sections[-1]))
        run.add_picture(img_path, width=Cm(max_width))
    except Exception as exc:
        raise ValueError(f"图片加载失败 {img_path}: {exc}") from exc
    if block.get("caption"):
        caption = document.add_paragraph(style="图题")
        caption.add_run(str(block["caption"]))


# ---------- 主函数 ----------
def main():
    if len(sys.argv) < 3:
        print("用法: python generate_docx.py <spec.json> <output.docx>")
        sys.exit(1)
    spec_path = sys.argv[1]
    out_path = sys.argv[2]

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    document = Document()
    profile_key, profile = _resolve_layout_profile(spec)

    chapter_fmt = spec.get("chapter_num_format", "chinese")
    if chapter_fmt not in ("chinese", "arabic"):
        print(f"警告: chapter_num_format='{chapter_fmt}' 不合法，回退为 'chinese'")
        chapter_fmt = "chinese"
    numbering_profile = spec.get("numbering_profile", "formal")
    if numbering_profile not in ("formal", "lightweight"):
        print(f"警告: numbering_profile='{numbering_profile}' 不合法（应为 formal/lightweight），回退为 'formal'")
        numbering_profile = "formal"
    # lightweight profile 下 chapter_num_format 无意义（一级固定"一、"），提示但不禁用
    if numbering_profile == "lightweight" and chapter_fmt == "arabic":
        print("提示: lightweight profile 一级固定显示'一、'，chapter_num_format=arabic 不生效")
    heading_align_left = spec.get("heading_align_left", False) is True
    setup_styles(document, chapter_num_format=chapter_fmt,
                 numbering_profile=numbering_profile, profile=profile,
                 heading_align_left=heading_align_left)

    # 默认单节；仅封面存在时启用首页不同。宽表可在正文中临时增设横向节。
    sec = document.sections[0]
    setup_page(sec, profile)
    has_cover = spec.get("cover", True) is True
    header_footer_enabled = spec.get("header_footer", True) is True
    sec.different_first_page_header_footer = has_cover and header_footer_enabled

    if header_footer_enabled:
        header_doc_name = spec.get("doc_name") or spec.get("document_title") or "未命名文档"
        _setup_header(sec.header, spec.get("project_name", ""), header_doc_name)
        _setup_footer_pagenum(sec.footer)

    if has_cover:
        build_cover(document, spec)

    # 目录
    if spec.get("toc", True):
        build_toc(document, spec)

    # 文档控制
    if spec.get("doc_control", True):
        build_doc_control(document, spec)

    # 正文
    spec_dir = os.path.dirname(os.path.abspath(spec_path))
    render_body(document, spec, spec_dir, profile)

    # 确保输出目录存在
    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    document.save(out_path)
    print(f"OK 生成: {out_path}")


if __name__ == "__main__":
    main()
