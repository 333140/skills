#!/usr/bin/env python3
"""
法条核验报告 DOCX 生成器 (generate_report.py)

将 JSON 格式的核验结果转换为专业排版的 DOCX 报告。

用法：
    python generate_report.py <JSON文件路径> [--output <输出路径>] [--title <报告标题>]

JSON schema 详见 SKILL.md 中的 "核验结果 JSON Schema" 章节。
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path


# ============================================================
# 颜色与样式常量
# ============================================================

# 主题色
DARK_BLUE = "1B3A5C"
RED = "C0392B"
GREEN = "27AE60"
ORANGE = "E67E22"
GOLD = "C49B3A"
LIGHT_GRAY = "F2F2F2"
MEDIUM_GRAY = "95A5A6"
DARK_GRAY = "2C3E50"
WHITE = "FFFFFF"

# 差异级别颜色映射
DIFF_COLORS = {
    "ID": GREEN,      # 完全一致
    "MN": GOLD,       # 轻微差异
    "SD": ORANGE,     # 实质差异
    "CD": RED,        # 严重差异
    "UV": MEDIUM_GRAY # 无法核实
}

# 差异级别中文标签
DIFF_LABELS = {
    "ID": "完全一致",
    "MN": "轻微差异",
    "SD": "实质差异",
    "CD": "严重差异",
    "UV": "无法核实"
}

# 状态中文标签
STATUS_LABELS = {
    "current": "现行有效",
    "revised": "部分修订",
    "repealed": "已废止",
    "expired": "已失效",
    "not_yet_effective": "尚未生效"
}

# 状态颜色映射
STATUS_COLORS = {
    "current": GREEN,
    "revised": ORANGE,
    "repealed": RED,
    "expired": RED,
    "not_yet_effective": GOLD
}

# 置信度中文标签
CONFIDENCE_LABELS = {
    "H": "高",
    "M": "中",
    "L": "低"
}


# ============================================================
# 字体检测
# ============================================================

def detect_chinese_font():
    """检测系统中可用的中文字体"""
    import platform
    system = platform.system()

    if system == "Windows":
        font_candidates = [
            "微软雅黑", "Microsoft YaHei",
            "宋体", "SimSun",
            "仿宋", "FangSong",
        ]
        heading_font = "微软雅黑"
        body_font = "仿宋"
    elif system == "Darwin":
        font_candidates = ["PingFang SC", "Hiragino Sans GB", "STSong"]
        heading_font = "PingFang SC"
        body_font = "PingFang SC"
    else:
        font_candidates = ["Noto Sans CJK SC", "WenQuanYi Micro Hei", "AR PL UMing CN"]
        heading_font = "Noto Sans CJK SC"
        body_font = "Noto Sans CJK SC"

    return heading_font, body_font


# ============================================================
# DOCX 文档构建器
# ============================================================

class ReportGenerator:
    """法条核验报告 DOCX 生成器"""

    def __init__(self, data: dict, title: str = None):
        self.data = data
        self.title = title or "法条核验报告"
        self.heading_font, self.body_font = detect_chinese_font()

        # 延迟导入
        try:
            from docx import Document
            from docx.shared import Inches, Pt, Cm, RGBColor, Emu
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            from docx.enum.table import WD_TABLE_ALIGNMENT
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement
            self.Document = Document
            self.Inches = Inches
            self.Pt = Pt
            self.Cm = Cm
            self.RGBColor = RGBColor
            self.WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH
            self.WD_TABLE_ALIGNMENT = WD_TABLE_ALIGNMENT
            self.qn = qn
            self.OxmlElement = OxmlElement
        except ImportError:
            raise ImportError("需要 python-docx 库，请运行：pip install python-docx")

    def generate(self) -> 'Document':
        """生成完整报告"""
        doc = self.Document()

        # 设置默认字体
        style = doc.styles['Normal']
        style.font.name = self.body_font
        style.font.size = self.Pt(10.5)
        style.element.rPr.rFonts.set(self.qn('w:eastAsia'), self.body_font)

        # 判断报告模式
        mode = self.data.get('mode', 'single')

        if mode == 'batch':
            self._add_batch_report(doc)
        else:
            self._add_single_report(doc)

        return doc

    # ============================================================
    # 单条核验报告
    # ============================================================

    def _add_single_report(self, doc):
        """添加单条核验报告内容"""

        # 标题
        self._add_title(doc, f"⚖️ {self.title}")

        # 核验对象信息
        self._add_verification_object(doc)

        # 核验来源
        self._add_sources(doc)

        # 官方原文
        self._add_official_text(doc)

        # 内容比对
        self._add_comparison(doc)

        # 有效性状态
        self._add_validity_status(doc)

        # 置信度评估
        self._add_confidence(doc)

        # 核验结论
        self._add_conclusion(doc)

        # 免责声明
        self._add_disclaimer(doc)

    def _add_verification_object(self, doc):
        """添加核验对象信息"""
        obj = self.data.get('verification_object', {})

        self._add_heading(doc, "📋 核验对象", level=2)

        table = self._create_table(doc, cols=2, rows=5, col_widths=[3, 12])

        fields = [
            ("法律名称", obj.get('law_name', '—')),
            ("条款", self._format_article_ref(obj)),
            ("颁布机关", obj.get('issuing_authority', '—')),
            ("版本年份", obj.get('version_year', '未指定')),
            ("用户提供内容", obj.get('user_text', '—')),
        ]

        for i, (label, value) in enumerate(fields):
            row = table.rows[i]
            self._set_cell(row.cells[0], label, bold=True, bg_color=DARK_BLUE, font_color=WHITE)
            self._set_cell(row.cells[1], str(value))

    def _add_sources(self, doc):
        """添加核验来源"""
        sources = self.data.get('sources', [])

        self._add_heading(doc, "🔍 核验来源", level=2)

        if not sources:
            self._add_paragraph(doc, "未记录核验来源", italic=True)
            return

        table = self._create_table(doc, cols=4, rows=len(sources) + 1,
                                   col_widths=[4, 5, 2, 4])

        # 表头
        header = table.rows[0]
        headers = ["来源名称", "网址", "状态", "备注"]
        for i, h in enumerate(headers):
            self._set_cell(header.cells[i], h, bold=True, bg_color=DARK_BLUE, font_color=WHITE,
                          align=self.WD_ALIGN_PARAGRAPH.CENTER)

        # 数据行
        for i, source in enumerate(sources):
            row = table.rows[i + 1]
            status = source.get('status', 'unknown')
            status_text = "✅ 已访问" if status == 'accessed' else "❌ 未访问" if status == 'failed' else "⚠️ 部分访问"
            bg = LIGHT_GRAY if i % 2 == 0 else WHITE

            self._set_cell(row.cells[0], source.get('name', '—'), bg_color=bg)
            self._set_cell(row.cells[1], source.get('url', '—'), bg_color=bg, size=self.Pt(8))
            self._set_cell(row.cells[2], status_text, bg_color=bg, align=self.WD_ALIGN_PARAGRAPH.CENTER)
            self._set_cell(row.cells[3], source.get('note', ''), bg_color=bg)

    def _add_official_text(self, doc):
        """添加官方原文"""
        official = self.data.get('official_text', {})

        self._add_heading(doc, "📝 官方原文", level=2)

        if not official.get('text'):
            self._add_paragraph(doc, "未能获取官方原文", italic=True, color=RED)
            return

        # 引用块样式
        para = doc.add_paragraph()
        para.paragraph_format.left_indent = self.Cm(1)
        para.paragraph_format.right_indent = self.Cm(1)

        run = para.add_run(official['text'])
        run.font.size = self.Pt(10)
        run.font.italic = True
        run.font.color.rgb = self.RGBColor.from_string(DARK_GRAY)

        # 来源标注
        source_url = official.get('source_url', '')
        if source_url:
            source_para = doc.add_paragraph()
            source_para.paragraph_format.alignment = self.WD_ALIGN_PARAGRAPH.RIGHT
            source_run = source_para.add_run(f"—— 来源：{source_url}")
            source_run.font.size = self.Pt(8)
            source_run.font.color.rgb = self.RGBColor.from_string(MEDIUM_GRAY)

    def _add_comparison(self, doc):
        """添加内容比对结果"""
        comparison = self.data.get('comparison', {})
        diff_level = comparison.get('diff_level', 'UV')
        differences = comparison.get('differences', [])

        self._add_heading(doc, "内容比对", level=2)

        # 比对结论
        label = DIFF_LABELS.get(diff_level, "未知")
        color = DIFF_COLORS.get(diff_level, MEDIUM_GRAY)

        conclusion_para = doc.add_paragraph()
        run = conclusion_para.add_run(f"比对结论：{label}")
        run.bold = True
        run.font.size = self.Pt(12)
        run.font.color.rgb = self.RGBColor.from_string(color)

        # 差异详情
        if differences:
            self._add_paragraph(doc, "")  # 空行

            table = self._create_table(doc, cols=5, rows=len(differences) + 1,
                                       col_widths=[2, 3.5, 3.5, 2, 4])

            header = table.rows[0]
            headers = ["位置", "用户版", "官方版", "差异级别", "影响说明"]
            for i, h in enumerate(headers):
                self._set_cell(header.cells[i], h, bold=True, bg_color=DARK_BLUE, font_color=WHITE,
                              align=self.WD_ALIGN_PARAGRAPH.CENTER)

            for i, diff in enumerate(differences):
                row = table.rows[i + 1]
                diff_type = diff.get('diff_type', 'UV')
                bg = LIGHT_GRAY if i % 2 == 0 else WHITE
                diff_color = DIFF_COLORS.get(diff_type, MEDIUM_GRAY)

                self._set_cell(row.cells[0], diff.get('location', '—'), bg_color=bg)
                self._set_cell(row.cells[1], diff.get('user_text', '—'), bg_color=bg)
                self._set_cell(row.cells[2], diff.get('official_text', '—'), bg_color=bg)

                # 差异级别带颜色
                label_cell = row.cells[3]
                self._set_cell(label_cell, DIFF_LABELS.get(diff_type, '—'),
                              bg_color=diff_color, font_color=WHITE,
                              align=self.WD_ALIGN_PARAGRAPH.CENTER)

                self._set_cell(row.cells[4], diff.get('impact', ''), bg_color=bg)

        # 正确法条
        correct = comparison.get('correct_text')
        if correct:
            self._add_paragraph(doc, "")
            correct_para = doc.add_paragraph()
            label_run = correct_para.add_run("✅ 正确法条：")
            label_run.bold = True
            label_run.font.color.rgb = self.RGBColor.from_string(GREEN)
            correct_para.add_run(correct)

    def _add_validity_status(self, doc):
        """添加有效性状态"""
        validity = self.data.get('validity', {})

        self._add_heading(doc, "📅 有效性状态", level=2)

        table = self._create_table(doc, cols=2, rows=5, col_widths=[4, 11])

        status = validity.get('status', 'unknown')
        status_label = STATUS_LABELS.get(status, '未知')
        status_color = STATUS_COLORS.get(status, MEDIUM_GRAY)

        fields = [
            ("法律状态", status_label),
            ("最新版本", validity.get('latest_version', '—')),
            ("施行日期", validity.get('effective_date', '—')),
            ("该条款状态", validity.get('clause_status', '—')),
            ("替代指引", validity.get('replacement', '无')),
        ]

        for i, (label, value) in enumerate(fields):
            row = table.rows[i]
            self._set_cell(row.cells[0], label, bold=True, bg_color=DARK_BLUE, font_color=WHITE)

            if i == 0:  # 状态行带颜色
                self._set_cell(row.cells[1], value, font_color=status_color, bold=True)
            else:
                self._set_cell(row.cells[1], str(value))

        # 替代法律详情
        replacement_detail = validity.get('replacement_detail')
        if replacement_detail:
            self._add_paragraph(doc, "")
            rep_para = doc.add_paragraph()
            label_run = rep_para.add_run("🔄 替代法律：")
            label_run.bold = True
            label_run.font.color.rgb = self.RGBColor.from_string(ORANGE)
            rep_para.add_run(replacement_detail)

    def _add_confidence(self, doc):
        """添加置信度评估"""
        confidence = self.data.get('confidence', {})
        level = confidence.get('confidence_level', 'L')
        factors = confidence.get('confidence_factors', {})
        note = confidence.get('confidence_note', '')

        self._add_heading(doc, "🎯 置信度评估", level=2)

        level_label = CONFIDENCE_LABELS.get(level, '未知')
        level_color = {
            'H': GREEN, 'M': ORANGE, 'L': RED
        }.get(level, MEDIUM_GRAY)

        para = doc.add_paragraph()
        run = para.add_run(f"置信度等级：{level_label}（{level}）")
        run.bold = True
        run.font.size = self.Pt(12)
        run.font.color.rgb = self.RGBColor.from_string(level_color)

        # 因子详情
        if factors:
            factor_labels = {
                'official_sources_count': '官方来源数量',
                'sources_agree': '来源内容一致',
                'current_status_confirmed': '现行状态已确认',
                'latest_version_confirmed': '最新版本已确认',
                'word_by_word_comparison_done': '逐字比对已完成',
                'proviso_checked': '但书/除外条款已检查',
            }

            table = self._create_table(doc, cols=2, rows=len(factors) + 1,
                                       col_widths=[6, 9])

            header = table.rows[0]
            self._set_cell(header.cells[0], "评估因子", bold=True, bg_color=DARK_BLUE, font_color=WHITE)
            self._set_cell(header.cells[1], "结果", bold=True, bg_color=DARK_BLUE, font_color=WHITE,
                          align=self.WD_ALIGN_PARAGRAPH.CENTER)

            for i, (key, value) in enumerate(factors.items()):
                row = table.rows[i + 1]
                bg = LIGHT_GRAY if i % 2 == 0 else WHITE
                label = factor_labels.get(key, key)

                self._set_cell(row.cells[0], label, bg_color=bg)

                if isinstance(value, bool):
                    result = "✅ 是" if value else "❌ 否"
                else:
                    result = str(value)
                self._set_cell(row.cells[1], result, bg_color=bg, align=self.WD_ALIGN_PARAGRAPH.CENTER)

        if note:
            note_para = doc.add_paragraph()
            note_run = note_para.add_run(f"备注：{note}")
            note_run.font.italic = True
            note_run.font.color.rgb = self.RGBColor.from_string(MEDIUM_GRAY)

        # 低置信度警告
        if level == 'L':
            self._add_paragraph(doc, "")
            warn_para = doc.add_paragraph()
            warn_run = warn_para.add_run("⚠️ 本核验置信度为低，建议通过官方渠道二次确认")
            warn_run.bold = True
            warn_run.font.color.rgb = self.RGBColor.from_string(RED)
        elif level == 'M':
            self._add_paragraph(doc, "")
            warn_para = doc.add_paragraph()
            warn_run = warn_para.add_run("⚠️ 本核验置信度为中等，建议通过官方渠道二次确认")
            warn_run.font.color.rgb = self.RGBColor.from_string(ORANGE)

    def _add_conclusion(self, doc):
        """添加核验结论"""
        conclusion = self.data.get('conclusion', '')

        self._add_heading(doc, "💡 核验结论", level=2)

        para = doc.add_paragraph()
        run = para.add_run(conclusion)
        run.font.size = self.Pt(11)

    # ============================================================
    # 批量核验报告
    # ============================================================

    def _add_batch_report(self, doc):
        """添加批量核验报告内容"""

        self._add_title(doc, f"⚖️ {self.title}")

        # 文件概况
        overview = self.data.get('overview', {})
        self._add_heading(doc, "📋 文件概况", level=2)

        table = self._create_table(doc, cols=2, rows=4, col_widths=[4, 11])
        fields = [
            ("文件名称", overview.get('file_name', '—')),
            ("提取的法律引用数量", str(overview.get('total_citations', 0))),
            ("核验完成", f"{overview.get('passed', 0)} 条通过 / "
                        f"{overview.get('issues', 0)} 条存在问题 / "
                        f"{overview.get('unverified', 0)} 条无法核实"),
            ("核验日期", datetime.now().strftime('%Y-%m-%d')),
        ]
        for i, (label, value) in enumerate(fields):
            row = table.rows[i]
            self._set_cell(row.cells[0], label, bold=True, bg_color=DARK_BLUE, font_color=WHITE)
            self._set_cell(row.cells[1], str(value))

        # 核验结果概览
        items = self.data.get('items', [])
        self._add_heading(doc, "📊 核验结果概览", level=2)

        if items:
            overview_table = self._create_table(
                doc, cols=6, rows=len(items) + 1,
                col_widths=[1, 4, 2, 2, 2, 4]
            )

            header = overview_table.rows[0]
            headers = ["序号", "法律名称", "条款", "内容核验", "有效性", "说明"]
            for i, h in enumerate(headers):
                self._set_cell(header.cells[i], h, bold=True, bg_color=DARK_BLUE, font_color=WHITE,
                              align=self.WD_ALIGN_PARAGRAPH.CENTER)

            for i, item in enumerate(items):
                row = overview_table.rows[i + 1]
                bg = LIGHT_GRAY if i % 2 == 0 else WHITE

                diff_level = item.get('diff_level', 'UV')
                validity = item.get('validity', 'unknown')

                diff_label = DIFF_LABELS.get(diff_level, '—')
                diff_color = DIFF_COLORS.get(diff_level, MEDIUM_GRAY)
                valid_label = STATUS_LABELS.get(validity, '—')
                valid_color = STATUS_COLORS.get(validity, MEDIUM_GRAY)

                self._set_cell(row.cells[0], str(i + 1), bg_color=bg, align=self.WD_ALIGN_PARAGRAPH.CENTER)
                self._set_cell(row.cells[1], item.get('law_name', '—'), bg_color=bg)
                self._set_cell(row.cells[2], item.get('article_ref', '—'), bg_color=bg,
                              align=self.WD_ALIGN_PARAGRAPH.CENTER)
                self._set_cell(row.cells[3], diff_label, bg_color=diff_color, font_color=WHITE,
                              align=self.WD_ALIGN_PARAGRAPH.CENTER)
                self._set_cell(row.cells[4], valid_label, bg_color=valid_color, font_color=WHITE,
                              align=self.WD_ALIGN_PARAGRAPH.CENTER)
                self._set_cell(row.cells[5], item.get('note', ''), bg_color=bg)

        # 详细核验信息
        detail_items = [item for item in items if item.get('diff_level', 'ID') != 'ID' or item.get('validity') != 'current']
        if detail_items:
            self._add_heading(doc, "🔍 详细核验信息", level=2)
            for item in detail_items:
                self._add_heading(doc, f"{item.get('law_name', '—')} {item.get('article_ref', '')}", level=3)

                # 差异详情
                if item.get('differences'):
                    for diff in item['differences']:
                        diff_para = doc.add_paragraph()
                        run = diff_para.add_run(f"• 差异：用户版「{diff.get('user_text', '')}」→ 官方版「{diff.get('official_text', '')}」")
                        run.font.size = self.Pt(10)

                # 有效性说明
                if item.get('validity_note'):
                    note_para = doc.add_paragraph()
                    note_run = note_para.add_run(f"• 有效性：{item['validity_note']}")
                    note_run.font.size = self.Pt(10)

        # 总结建议
        summary = self.data.get('summary', {})
        self._add_heading(doc, "💡 总结建议", level=2)

        categories = {
            'continue': ('✅ 可继续使用', GREEN),
            'needs_correction': ('❌ 需修正内容', RED),
            'needs_replacement': ('🔄 需更换法律', ORANGE),
        }

        for key, (label, color) in categories.items():
            items_list = summary.get(key, [])
            if items_list:
                para = doc.add_paragraph()
                label_run = para.add_run(f"{label}：")
                label_run.bold = True
                label_run.font.color.rgb = self.RGBColor.from_string(color)
                para.add_run('、'.join(str(x) for x in items_list))

        # 免责声明
        self._add_disclaimer(doc)

    # ============================================================
    # 通用组件
    # ============================================================

    def _add_title(self, doc, text: str):
        """添加报告标题"""
        para = doc.add_paragraph()
        para.alignment = self.WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        run.bold = True
        run.font.size = self.Pt(18)
        run.font.color.rgb = self.RGBColor.from_string(DARK_BLUE)
        run.font.name = self.heading_font
        run.element.rPr.rFonts.set(self.qn('w:eastAsia'), self.heading_font)

        # 核验日期
        date_para = doc.add_paragraph()
        date_para.alignment = self.WD_ALIGN_PARAGRAPH.CENTER
        date_run = date_para.add_run(f"核验日期：{datetime.now().strftime('%Y年%m月%d日')}")
        date_run.font.size = self.Pt(10)
        date_run.font.color.rgb = self.RGBColor.from_string(MEDIUM_GRAY)

    def _add_heading(self, doc, text: str, level: int = 2):
        """添加标题"""
        heading = doc.add_heading(text, level=level)
        for run in heading.runs:
            run.font.color.rgb = self.RGBColor.from_string(DARK_BLUE)
            run.font.name = self.heading_font
            run.element.rPr.rFonts.set(self.qn('w:eastAsia'), self.heading_font)

    def _add_paragraph(self, doc, text: str, bold: bool = False, italic: bool = False,
                       color: str = None, size=None):
        """添加段落"""
        para = doc.add_paragraph()
        if text:
            run = para.add_run(text)
            run.bold = bold
            run.italic = italic
            if color:
                run.font.color.rgb = self.RGBColor.from_string(color)
            if size:
                run.font.size = size
        return para

    def _add_disclaimer(self, doc):
        """添加免责声明"""
        self._add_paragraph(doc, "")
        self._add_paragraph(doc, "")

        # 分割线
        para = doc.add_paragraph()
        pBdr = self.OxmlElement('w:pBdr')
        bottom = self.OxmlElement('w:bottom')
        bottom.set(self.qn('w:val'), 'single')
        bottom.set(self.qn('w:sz'), '6')
        bottom.set(self.qn('w:space'), '1')
        bottom.set(self.qn('w:color'), GOLD)
        pBdr.append(bottom)
        para.paragraph_format.element.get_or_add_pPr().append(pBdr)

        disclaimer = doc.add_paragraph()
        run = disclaimer.add_run("⚠️ 免责声明：本核验基于公开信息，仅供参考，不构成法律意见。具体法律问题请咨询专业律师。")
        run.font.size = self.Pt(8)
        run.font.color.rgb = self.RGBColor.from_string(MEDIUM_GRAY)
        run.italic = True

    def _create_table(self, doc, cols: int, rows: int, col_widths: list = None):
        """创建表格并设置列宽"""
        table = doc.add_table(rows=rows, cols=cols)
        table.style = 'Table Grid'
        table.alignment = self.WD_TABLE_ALIGNMENT.CENTER

        if col_widths and len(col_widths) == cols:
            for i, width in enumerate(col_widths):
                for row in table.rows:
                    row.cells[i].width = self.Cm(width)

        return table

    def _set_cell(self, cell, text: str, bold: bool = False, bg_color: str = None,
                  font_color: str = None, align=None, size=None):
        """设置单元格内容和样式"""
        cell.text = ''
        para = cell.paragraphs[0]
        if align:
            para.alignment = align

        run = para.add_run(str(text))
        run.bold = bold
        run.font.name = self.body_font
        run.element.rPr.rFonts.set(self.qn('w:eastAsia'), self.body_font)

        if font_color:
            run.font.color.rgb = self.RGBColor.from_string(font_color)
        if size:
            run.font.size = size
        else:
            run.font.size = self.Pt(9)

        if bg_color:
            shading = self.OxmlElement('w:shd')
            shading.set(self.qn('w:fill'), bg_color)
            shading.set(self.qn('w:val'), 'clear')
            cell._tc.get_or_add_tcPr().append(shading)

    def _format_article_ref(self, obj: dict) -> str:
        """格式化条款引用"""
        parts = []
        if obj.get('article'):
            parts.append(f"第{obj['article']}条")
        if obj.get('paragraph'):
            parts.append(f"第{obj['paragraph']}款")
        if obj.get('item'):
            parts.append(f"第{obj['item']}项")
        return ''.join(parts) if parts else '—'


# ============================================================
# 主函数
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description='法条核验报告 DOCX 生成器',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  python generate_report.py result.json
  python generate_report.py result.json --output report.docx
  python generate_report.py result.json --title "《劳动合同法》第47条核验报告"
        """,
    )
    parser.add_argument('json_file', help='核验结果 JSON 文件路径')
    parser.add_argument('-o', '--output', help='输出 DOCX 文件路径（默认与 JSON 同名 .docx）')
    parser.add_argument('-t', '--title', help='报告标题', default='法条核验报告')

    args = parser.parse_args()

    # 读取 JSON
    json_path = Path(args.json_file)
    if not json_path.exists():
        print(f"错误：文件不存在：{args.json_file}", file=sys.stderr)
        sys.exit(1)

    try:
        data = json.loads(json_path.read_text(encoding='utf-8'))
    except json.JSONDecodeError as e:
        print(f"错误：JSON 解析失败：{e}", file=sys.stderr)
        sys.exit(1)

    # 生成报告
    try:
        generator = ReportGenerator(data, title=args.title)
        doc = generator.generate()
    except ImportError as e:
        print(f"错误：{e}", file=sys.stderr)
        sys.exit(1)

    # 输出路径
    output_path = args.output or str(json_path.with_suffix('.docx'))

    # 保存
    doc.save(output_path)
    print(f"报告已保存到：{output_path}", file=sys.stderr)


if __name__ == '__main__':
    main()
