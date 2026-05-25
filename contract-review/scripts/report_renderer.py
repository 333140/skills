#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合同审核报告 DOCX 渲染器

将 Markdown 格式的审核报告转换为专业排版的 DOCX 文档。
特性：
- 自动替换 emoji 风险图标为彩色文字标签
- 专业排版（标题层级、字体、行距、页眉页脚）
- 支持 Markdown 基础语法解析
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

try:
    from scripts.utilities import DISCLAIMER_TEXT
except ImportError:
    from utilities import DISCLAIMER_TEXT

# ============ 配色方案 ============
COLOR_HIGH_RISK = RGBColor(0xC5, 0x30, 0x30)      # 深红 #C53030
COLOR_MEDIUM_RISK = RGBColor(0xDD, 0x6B, 0x20)    # 橙黄 #DD6B20
COLOR_LOW_RISK = RGBColor(0x2B, 0x6C, 0xB0)       # 蓝色 #2B6CB0
COLOR_TITLE = RGBColor(0x1A, 0x20, 0x2C)          # 深灰黑 #1A202C
COLOR_SUBTITLE = RGBColor(0x2D, 0x37, 0x41)       # 深灰 #2D3748
COLOR_BODY = RGBColor(0x2D, 0x37, 0x41)           # 正文深灰
COLOR_ACCENT = RGBColor(0x74, 0x00, 0x00)         # 暗红强调

# ============ 字体配置 ============
FONT_TITLE = "黑体"
FONT_HEADING = "黑体"
FONT_BODY = "宋体"
FONT_EMPHASIS = "楷体"


def _set_run_font(
    run,
    font_name: str,
    size: Optional[int] = None,
    bold: Optional[bool] = None,
    color: Optional[RGBColor] = None,
    italic: Optional[bool] = None,
) -> None:
    """设置 run 的字体属性"""
    run.font.name = font_name
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    if italic is not None:
        run.italic = italic

    # 设置东亚字体
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:eastAsia"), font_name)


def _set_cell_shading(cell, color_hex: str) -> None:
    """设置单元格背景色"""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color_hex)
    cell._tc.get_or_add_tcPr().append(shading)


def _apply_line_spacing(para, line_spacing: float = 1.5) -> None:
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    para.paragraph_format.line_spacing = line_spacing


def _replace_emoji_with_colored_runs(para, text: str, font_name: str, font_size: int = 12) -> None:
    """
    将文本中的 emoji 替换为带颜色的文字标签 run。
    支持：🔴 → 【高风险】红色, 🟡 → 【中风险】橙黄, 🔵 → 【低风险】蓝色
    """
    # emoji 到 (标签, 颜色) 的映射
    emoji_map = {
        "🔴": ("【高风险】", COLOR_HIGH_RISK),
        "🟡": ("【中风险】", COLOR_MEDIUM_RISK),
        "🔵": ("【低风险】", COLOR_LOW_RISK),
    }

    # 按顺序分割文本
    pattern = re.compile(r"(🔴|🟡|🔵)")
    parts = pattern.split(text)

    for part in parts:
        if part in emoji_map:
            label, color = emoji_map[part]
            run = para.add_run(label)
            _set_run_font(run, font_name, size=font_size, bold=True, color=color)
        else:
            if part:
                run = para.add_run(part)
                _set_run_font(run, font_name, size=font_size, color=COLOR_BODY)


def _add_title(doc: Document, text: str) -> None:
    """添加报告大标题"""
    para = doc.add_paragraph()
    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = para.add_run(text)
    _set_run_font(run, FONT_TITLE, size=22, bold=True, color=COLOR_TITLE)
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.space_before = Pt(12)


def _add_subtitle(doc: Document, text: str) -> None:
    """添加副标题（如合同名称）"""
    para = doc.add_paragraph()
    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = para.add_run(text)
    _set_run_font(run, FONT_HEADING, size=14, bold=True, color=COLOR_SUBTITLE)
    para.paragraph_format.space_after = Pt(12)


def _add_h1(doc: Document, text: str) -> None:
    """添加一级标题（如 ## 第一层：基础审核）"""
    # 移除 Markdown 标记
    clean = re.sub(r"^#{1,2}\s*", "", text).strip()
    para = doc.add_paragraph()
    run = para.add_run(clean)
    _set_run_font(run, FONT_HEADING, size=16, bold=True, color=COLOR_TITLE)
    para.paragraph_format.space_before = Pt(18)
    para.paragraph_format.space_after = Pt(8)
    _apply_line_spacing(para, 1.5)


def _add_h2(doc: Document, text: str) -> None:
    """添加二级标题（如 ### 🔴 高风险问题）"""
    clean = re.sub(r"^#{3,4}\s*", "", text).strip()
    para = doc.add_paragraph()
    _replace_emoji_with_colored_runs(para, clean, FONT_HEADING, font_size=14)
    para.paragraph_format.space_before = Pt(14)
    para.paragraph_format.space_after = Pt(6)
    _apply_line_spacing(para, 1.5)


def _add_bold_paragraph(doc: Document, text: str) -> None:
    """添加加粗段落（如 **1. 合同编号缺失**）"""
    clean = text.strip().lstrip("*").rstrip("*").strip()
    para = doc.add_paragraph()
    run = para.add_run(clean)
    _set_run_font(run, FONT_BODY, size=12, bold=True, color=COLOR_TITLE)
    para.paragraph_format.space_before = Pt(8)
    para.paragraph_format.space_after = Pt(4)
    _apply_line_spacing(para, 1.5)


def _add_list_item(doc: Document, text: str) -> None:
    """添加列表项（如 - 位置：合同首页）"""
    clean = re.sub(r"^[\-\*]\s*", "", text).strip()
    para = doc.add_paragraph()
    # 缩进
    para.paragraph_format.left_indent = Cm(0.8)
    para.paragraph_format.first_line_indent = Cm(-0.4)

    # 检测是否有加粗前缀（如 "位置："、"问题："、"风险："、"建议："）
    match = re.match(r"^(位置|问题|风险|建议|当事人|合同类型|审核日期|注册资本|合同名称|目标公司)：(.+)$", clean)
    if match:
        label = match.group(1) + "："
        content = match.group(2)
        run_label = para.add_run(label)
        _set_run_font(run_label, FONT_BODY, size=12, bold=True, color=COLOR_SUBTITLE)
        run_content = para.add_run(content)
        _set_run_font(run_content, FONT_BODY, size=12, color=COLOR_BODY)
    else:
        run = para.add_run(clean)
        _set_run_font(run, FONT_BODY, size=12, color=COLOR_BODY)

    para.paragraph_format.space_after = Pt(2)
    _apply_line_spacing(para, 1.5)


def _add_numbered_list_item(doc: Document, text: str) -> None:
    """添加编号列表项（如 1. **公司治理违法风险**）"""
    clean = re.sub(r"^\d+\.\s*", "", text).strip()
    # 移除可能的加粗标记
    clean = clean.lstrip("*").rstrip("*").strip()
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Cm(0.8)
    para.paragraph_format.first_line_indent = Cm(-0.4)
    run = para.add_run(clean)
    _set_run_font(run, FONT_BODY, size=12, bold=True, color=COLOR_TITLE)
    para.paragraph_format.space_after = Pt(4)
    _apply_line_spacing(para, 1.5)


def _add_normal_paragraph(doc: Document, text: str) -> None:
    """添加普通段落"""
    para = doc.add_paragraph()
    run = para.add_run(text.strip())
    _set_run_font(run, FONT_BODY, size=12, color=COLOR_BODY)
    para.paragraph_format.space_after = Pt(4)
    _apply_line_spacing(para, 1.5)


def _add_horizontal_line(doc: Document) -> None:
    """添加水平分隔线"""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(8)
    para.paragraph_format.space_after = Pt(8)
    # 使用底部边框模拟分隔线
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "B7791F")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_info_table(doc: Document, lines: List[str]) -> None:
    """将合同基本信息列表转为表格"""
    # 解析键值对
    rows = []
    for line in lines:
        line = line.strip()
        if line.startswith("-") or line.startswith("*"):
            clean = re.sub(r"^[\-\*]\s*", "", line).strip()
            if "：" in clean:
                key, value = clean.split("：", 1)
                rows.append((key + "：", value))
            elif ":" in clean:
                key, value = clean.split(":", 1)
                rows.append((key + ":", value))

    if not rows:
        return

    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    # 设置列宽
    for row in table.rows:
        row.cells[0].width = Cm(4)
        row.cells[1].width = Cm(12)

    for key, value in rows:
        cells = table.add_row().cells
        # 左列（标签）
        cells[0].text = ""
        para0 = cells[0].paragraphs[0]
        run0 = para0.add_run(key)
        _set_run_font(run0, FONT_BODY, size=12, bold=True, color=COLOR_SUBTITLE)
        _set_cell_shading(cells[0], "FFF5F5")

        # 右列（值）
        cells[1].text = ""
        para1 = cells[1].paragraphs[0]
        run1 = para1.add_run(value)
        _set_run_font(run1, FONT_BODY, size=12, color=COLOR_BODY)

    # 表格后留空
    doc.add_paragraph()


def _setup_page(doc: Document, title_text: str = "合同审核报告") -> None:
    """设置页面格式和页眉页脚"""
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    # 页眉
    header = section.header
    header.is_linked_to_previous = False
    para_header = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    para_header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run_header = para_header.add_run(title_text)
    _set_run_font(run_header, FONT_HEADING, size=9, color=COLOR_SUBTITLE)
    # 页眉底部横线
    pPr = para_header._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "CCCCCC")
    pBdr.append(bottom)
    pPr.append(pBdr)

    # 页脚 - 页码
    footer = section.footer
    footer.is_linked_to_previous = False
    para_footer = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para_footer.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # 添加 PAGE 域
    run = para_footer.add_run()
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    run._r.append(fldChar1)

    run2 = para_footer.add_run()
    instrText = OxmlElement("w:instrText")
    instrText.text = "PAGE"
    run2._r.append(instrText)

    run3 = para_footer.add_run()
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    run3._r.append(fldChar2)

    run4 = para_footer.add_run("1")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")
    run4._r.append(fldChar3)

    _set_run_font(run, FONT_BODY, size=9, color=COLOR_SUBTITLE)
    _set_run_font(run2, FONT_BODY, size=9, color=COLOR_SUBTITLE)
    _set_run_font(run3, FONT_BODY, size=9, color=COLOR_SUBTITLE)
    _set_run_font(run4, FONT_BODY, size=9, color=COLOR_SUBTITLE)


def _setup_document_defaults(doc: Document) -> None:
    """设置文档默认样式"""
    for style_name in ("Normal", "Heading 1", "Heading 2", "Heading 3"):
        if style_name not in doc.styles:
            continue
        style = doc.styles[style_name]
        style.font.name = FONT_BODY
        style.font.size = Pt(12)
        style.font.color.rgb = COLOR_BODY

        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.rFonts
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:ascii"), FONT_BODY)
        rfonts.set(qn("w:hAnsi"), FONT_BODY)
        rfonts.set(qn("w:eastAsia"), FONT_BODY)


def render_review_report_docx(
    markdown_text: str,
    output_path: Path,
    report_title: str = "合同审核报告",
    contract_name: Optional[str] = None,
) -> Path:
    """
    将 Markdown 格式的审核报告渲染为专业排版的 DOCX。

    Args:
        markdown_text: Markdown 格式的审核报告文本
        output_path: 输出 DOCX 路径
        report_title: 报告标题
        contract_name: 合同名称（用于副标题）

    Returns:
        Path: 输出文件路径
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = Document()
    _setup_document_defaults(doc)
    _setup_page(doc, report_title)

    # 标题
    _add_title(doc, report_title)
    if contract_name:
        _add_subtitle(doc, f"《{contract_name}》")

    lines = markdown_text.splitlines()
    i = 0
    info_lines = []
    in_info_block = False

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # 主标题（# 股权投资合作协议 审核报告）— 已处理，跳过
        if stripped.startswith("# ") and i == 0:
            i += 1
            continue

        # 分隔线
        if stripped == "---" or stripped.startswith("---"):
            if info_lines:
                _add_info_table(doc, info_lines)
                info_lines = []
                in_info_block = False
            _add_horizontal_line(doc)
            i += 1
            continue

        # 一级标题（## 合同基本信息 / ## 第一层：基础审核）
        if stripped.startswith("## "):
            if info_lines:
                _add_info_table(doc, info_lines)
                info_lines = []
                in_info_block = False
            _add_h1(doc, stripped)
            i += 1
            continue

        # 二级标题（### 🔴 高风险问题）
        if stripped.startswith("### "):
            if info_lines:
                _add_info_table(doc, info_lines)
                info_lines = []
                in_info_block = False
            _add_h2(doc, stripped)
            i += 1
            continue

        # 加粗段落（**1. 合同编号缺失**）
        if re.match(r"^\*\*\d+\.\s+.+\*\*$", stripped):
            if info_lines:
                _add_info_table(doc, info_lines)
                info_lines = []
                in_info_block = False
            _add_bold_paragraph(doc, stripped)
            i += 1
            continue

        # 加粗标题段落（**立即处理（签约前必须修改）：**）
        if re.match(r"^\*\*.+\*\*$", stripped):
            if info_lines:
                _add_info_table(doc, info_lines)
                info_lines = []
                in_info_block = False
            _add_bold_paragraph(doc, stripped)
            i += 1
            continue

        # 列表项（- 位置：合同首页）
        if stripped.startswith("-") or stripped.startswith("*"):
            # 检测是否是合同基本信息块
            if "合同名称" in stripped or "合同类型" in stripped or "当事人" in stripped or "目标公司" in stripped or "注册资本" in stripped or "审核日期" in stripped:
                info_lines.append(stripped)
                in_info_block = True
                i += 1
                continue
            if in_info_block and ("：" in stripped or ":" in stripped):
                info_lines.append(stripped)
                i += 1
                continue
            if info_lines:
                _add_info_table(doc, info_lines)
                info_lines = []
                in_info_block = False
            _add_list_item(doc, stripped)
            i += 1
            continue

        # 编号列表（1. **公司治理违法风险**）
        if re.match(r"^\d+\.\s+.+$", stripped):
            if info_lines:
                _add_info_table(doc, info_lines)
                info_lines = []
                in_info_block = False
            _add_numbered_list_item(doc, stripped)
            i += 1
            continue

        # 普通段落
        if info_lines:
            _add_info_table(doc, info_lines)
            info_lines = []
            in_info_block = False
        _add_normal_paragraph(doc, stripped)
        i += 1

    # 尾部信息（审核人、日期）
    if info_lines:
        _add_info_table(doc, info_lines)

    # 添加尾部分隔线和免责声明
    _add_horizontal_line(doc)
    para = doc.add_paragraph()
    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = para.add_run(DISCLAIMER_TEXT)
    _set_run_font(run, FONT_EMPHASIS, size=10, italic=True, color=COLOR_SUBTITLE)
    para.paragraph_format.space_before = Pt(12)

    doc.save(output_path)
    return output_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="将 Markdown 审核报告转换为 DOCX")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("-o", "--output", help="输出 DOCX 文件路径")
    parser.add_argument("-t", "--title", default="合同审核报告", help="报告标题")
    parser.add_argument("-c", "--contract", help="合同名称（副标题）")
    args = parser.parse_args()

    md_path = Path(args.input)
    md_text = md_path.read_text(encoding="utf-8")

    output = args.output or md_path.with_suffix(".docx")
    render_review_report_docx(md_text, output, args.title, args.contract)
    print(f"✓ 审核报告已生成: {output}")
