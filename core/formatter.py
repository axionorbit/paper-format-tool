# -*- coding: utf-8 -*-
"""
格式化器模块 - 样式管理和格式应用
"""

import re
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

from utils.config import (
    PART_STYLE_NAMES, OUTLINE_LEVELS,
    FONT_SIZES, ALIGNMENTS, LINE_SPACINGS,
    FIRST_LINE_INDENT_MAP, OVERALL_INDENT_MAP
)


class StyleManager:
    """样式管理器"""

    def __init__(self, document):
        self.document = document
        self.styles = document.styles

    def get_style_name(self, part_type):
        return PART_STYLE_NAMES.get(part_type, part_type)

    def _set_style_alignment(self, style, alignment_name):
        alignment_value = ALIGNMENTS.get(alignment_name)
        if alignment_value is None:
            return

        style.paragraph_format.alignment = alignment_value

        if style.paragraph_format.alignment != alignment_value:
            style.element.get_or_add_pPr().jc_val = alignment_value

    def _set_style_outline_level(self, style, part_type):
        ppr = style.element.get_or_add_pPr()
        outline_level = OUTLINE_LEVELS.get(part_type)

        if outline_level is None:
            if ppr.outlineLvl is not None:
                ppr._remove_outlineLvl()
            return

        ppr.get_or_add_outlineLvl().val = outline_level

    def create_or_update_style(self, part_type, rule):
        """
        创建或更新样式
        part_type: 段落类型（如"heading1"）
        rule: 格式和其他规则字典
        """
        style_name = self.get_style_name(part_type)

        # 检查样式是否存在
        if style_name in self.styles:
            style = self.styles[style_name]
            print(f"更新样式: {style_name}")
        else:
            # 创建新样式（基于Normal样式）
            style = self.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            print(f"创建样式: {style_name}")

        # 设置段落格式
        self._set_paragraph_format(style, rule)
        self._set_style_outline_level(style, part_type)

        # 设置字体格式
        self._set_font_format(style, rule)

        return style

    def _set_paragraph_format(self, style, rule):
        """设置段落格式"""
        pf = style.paragraph_format

        # 对齐方式
        if rule.get("alignment"):
            self._set_style_alignment(style, rule["alignment"])

        # 行距
        if rule.get("line_spacing"):
            line_spacing_map = {
                "单倍行距": 1.0, "1.15倍行距": 1.15, "1.25倍行距": 1.25,
                "1.5倍行距": 1.5, "1.75倍行距": 1.75, "2倍行距": 2.0,
                "2.5倍行距": 2.5, "3倍行距": 3.0
            }
            if rule["line_spacing"] in line_spacing_map:
                pf.line_spacing = line_spacing_map[rule["line_spacing"]]

        # 段前间距
        if rule.get("space_before"):
            try:
                space_value = int(rule["space_before"].rstrip('磅'))
                pf.space_before = Pt(space_value)
            except:
                pass

        # 段后间距
        if rule.get("space_after"):
            try:
                space_value = int(rule["space_after"].rstrip('磅'))
                pf.space_after = Pt(space_value)
            except:
                pass

        # 整体缩进
        overall_indent = rule.get("overall_indent", "无")
        left_chars, right_chars = OVERALL_INDENT_MAP.get(overall_indent, (0, 0))
        pf.left_indent = Pt(left_chars * 12)
        pf.right_indent = Pt(right_chars * 12)

        # 首行缩进
        first_line_indent = rule.get("first_line_indent", "无")
        indent_chars = FIRST_LINE_INDENT_MAP.get(first_line_indent, 0)
        pf.first_line_indent = Pt(indent_chars * 12)  # 约12磅=1字符

    def _set_font_format(self, style, rule):
        """设置字体格式"""
        font = style.font

        # 中文字体
        if rule.get("chinese_font"):
            font.name = rule["chinese_font"]
            # 设置西文字体
            font._element.rPr.rFonts.set(qn('w:eastAsia'), rule["chinese_font"])

        # 英文字体
        if rule.get("english_font"):
            font.name = rule["english_font"]

        # 字号
        if rule.get("font_size"):
            if rule["font_size"] in FONT_SIZES:
                font.size = Pt(FONT_SIZES[rule["font_size"]])

        # 加粗
        if rule.get("bold") is not None:
            font.bold = rule["bold"]

        # 倾斜
        if rule.get("italic") is not None:
            font.italic = rule["italic"]

    def apply_style_to_paragraph(self, paragraph, part_type, rules):
        """
        将样式应用到段落
        策略：清除原有直接格式，应用样式
        """
        if part_type is None:
            # 如果part_type是None，默认为正文内容
            part_type = "body"

        style_name = self.get_style_name(part_type)

        # 优先复用已创建样式，避免每个段落重复创建/更新
        # 仅当样式缺失时再兜底创建一次
        if style_name in self.styles:
            target_style = self.styles[style_name]
        else:
            target_style = self.create_or_update_style(part_type, rules.get(part_type, {}))

        # 应用样式到段落
        paragraph.style = target_style

        # 清除原有的直接格式（可选策略）
        # 这里我们选择完全依赖样式，所以清除直接格式
        try:
            paragraph.paragraph_format.clear()
        except:
            pass

        # 确保对齐方式正确设置（因为clear()可能不会正确继承样式对齐）
        style_pf = target_style.paragraph_format
        if hasattr(style_pf, 'alignment') and style_pf.alignment is not None:
            paragraph.paragraph_format.alignment = style_pf.alignment

        # 谨慎清除run上的直接格式，避免破坏图片和公式
        for run in paragraph.runs:
            try:
                # 检查是否是图片或公式（这些节点不应被当作纯文本run清理）
                has_drawing = False
                has_formula = False
                if hasattr(run, '_element') and run._element is not None:
                    try:
                        drawings = run._element.xpath('.//w:drawing')
                        has_drawing = len(drawings) > 0
                        formulas = run._element.xpath('.//*[local-name()="oMath" or local-name()="oMathPara"]')
                        has_formula = len(formulas) > 0
                    except:
                        pass

                if has_drawing or has_formula:
                    # 对于包含图片或公式的run，只清除安全的字符级属性
                    run.font.bold = None
                    run.font.italic = None
                    run.font.underline = None
                    run.font.color = None
                else:
                    # 对于纯文本run，清除所有格式
                    run.font.name = None
                    run.font.size = None
                    run.font.bold = None
                    run.font.italic = None
                    run.font.underline = None
                    run.font.color = None
            except:
                pass

    def apply_style_to_table_cell(self, cell, part_type, rules):
        """将样式应用到表格单元格中的段落"""
        for paragraph in cell.paragraphs:
            self.apply_style_to_paragraph(paragraph, part_type, rules)


class FormatApplier:
    """格式应用器（新版 - 样式驱动）"""

    def __init__(self, format_rules, document=None):
        self.rules = format_rules
        self.document = document
        self.style_manager = StyleManager(document) if document else None

    def apply_to_paragraph(self, paragraph, part_type):
        """对段落应用样式"""
        # 使用样式管理器
        if self.style_manager:
            self.style_manager.apply_style_to_paragraph(paragraph, part_type, self.rules)
        else:
            # 回退到直接格式（如果样式管理器不可用）
            self._apply_direct_format(paragraph, part_type)

    def _apply_direct_format(self, paragraph, part_type):
        """直接格式应用（回退方法）"""
        # 确保至少应用正文格式，即使part_type不在规则中
        if part_type not in self.rules:
            if "body" in self.rules:
                rule = self.rules["body"]
            else:
                return
        else:
            rule = self.rules[part_type]

        if not rule:
            return

        # 获取段落格式
        pf = paragraph.paragraph_format

        # 设置对齐方式
        if rule.get("alignment"):
            pf.alignment = ALIGNMENTS.get(rule["alignment"])

        # 设置行距
        if rule.get("line_spacing"):
            line_spacing_value = LINE_SPACINGS.get(rule["line_spacing"])
            if line_spacing_value:
                # 检查是否为固定值（包含"磅"）
                if "倍" in rule["line_spacing"]:
                    # 倍数行距
                    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                    pf.line_spacing = line_spacing_value
                else:
                    # 固定值行距
                    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                    pf.line_spacing = Pt(line_spacing_value)

        # 设置段前间距
        if rule.get("space_before"):
            space_value = self._parse_spacing(rule["space_before"])
            if space_value is not None:
                pf.space_before = Pt(space_value)

        # 设置段后间距
        if rule.get("space_after"):
            space_value = self._parse_spacing(rule["space_after"])
            if space_value is not None:
                pf.space_after = Pt(space_value)

        # 设置整体缩进
        left_chars, right_chars = self._parse_overall_indent(rule.get("overall_indent"))
        pf.left_indent = Pt(left_chars * 12)
        pf.right_indent = Pt(right_chars * 12)

        # 设置首行缩进
        indent_chars = self._parse_first_line_indent(rule.get("first_line_indent"))
        pf.first_line_indent = Pt(indent_chars * 12)

        # 设置字体格式
        for run in paragraph.runs:
            self._apply_font(run, rule)

    def _apply_font(self, run, rule):
        """设置字体格式"""
        # 中文字体
        if rule.get("chinese_font"):
            run.font.name = rule["chinese_font"]
            run._element.rPr.rFonts.set(qn('w:eastAsia'), rule["chinese_font"])

        # 英文字体
        if rule.get("english_font"):
            run.font.name = rule["english_font"]

        # 字号
        if rule.get("font_size"):
            size_pt = FONT_SIZES.get(rule["font_size"])
            if size_pt:
                run.font.size = Pt(size_pt)

        # 加粗
        if rule.get("bold") is not None:
            run.font.bold = rule["bold"]

        # 倾斜
        if rule.get("italic") is not None:
            run.font.italic = rule["italic"]

    def _parse_spacing(self, spacing_str):
        """解析间距字符串，返回磅值"""
        if not spacing_str:
            return None
        match = re.match(r'(\d+)', spacing_str)
        if match:
            return int(match.group(1))
        return None

    def _parse_first_line_indent(self, indent_str):
        """解析首行缩进字符串，返回字符数"""
        if not indent_str:
            return 0
        return FIRST_LINE_INDENT_MAP.get(indent_str, 0)

    def _parse_overall_indent(self, indent_str):
        """解析整体缩进字符串，返回左右缩进字符数"""
        if not indent_str:
            return (0, 0)
        return OVERALL_INDENT_MAP.get(indent_str, (0, 0))

    def apply_to_table(self, table, rules):
        """对表格应用样式（增强版）"""
        if not self.style_manager:
            return

        # 获取PartIdentifier用于识别（避免循环导入）
        from core.rule_engine import PartIdentifier
        identifier = PartIdentifier()

        # 遍历表格的每一行和单元格
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_paragraphs = list(cell.paragraphs)  # 复制列表，避免修改迭代

                # 对单元格中的每个段落进行识别和应用
                for para_idx, paragraph in enumerate(cell_paragraphs):
                    # 重新获取文本，因为段落对象可能在处理过程中被修改
                    text = paragraph.text.strip()

                    # 跳过空段落
                    if not text:
                        continue

                    # 识别优先级（从高到低）：
                    # 1. 图片标题（如果单元格中有图片且匹配图标题模式）
                    # 2. 表格标题
                    # 3. 图表注释（资料来源、数据来源等）
                    # 4. 其他（表格内容）

                    # 检查是否为表格标题或图片标题
                    chart_result = self._is_chart_caption(text)
                    if chart_result and chart_result[0]:
                        caption_type = chart_result[1]
                        self.style_manager.apply_style_to_paragraph(
                            paragraph, caption_type, rules
                        )
                        continue

                    # 检查是否为图表注释（资料来源、数据来源等）
                    if (text.startswith("资料来源") or
                        text.startswith("数据来源") or
                        text.startswith("来源") or
                        text.startswith("注")):

                        self.style_manager.apply_style_to_paragraph(
                            paragraph, "table_note", rules
                        )
                        continue

                    # 使用段落识别器判断其他类型
                    part_type, is_title = identifier.identify(paragraph)

                    # 应用对应的样式
                    if part_type and part_type in ["figure_caption", "table_caption", "table_note"]:
                        self.style_manager.apply_style_to_paragraph(paragraph, part_type, rules)
                    else:
                        # 默认为表格内容
                        self.style_manager.apply_style_to_paragraph(paragraph, "table_content", rules)

    def _is_chart_caption(self, text):
        """检查是否为图表标题"""
        # 表标题模式（支持各种变体）
        table_patterns = [
            r'^表\s*[0-9]+[-\.][0-9]+.*$',      # 表5.1...
            r'^表\s*[0-9]+[-\.][0-9]+\s*[:：].*$', # 表5.1: ...
            r'^表\s*[0-9]+.*$',                 # 表5...
            r'^表[0-9]+[-\.][0-9]+.*$',         # 表5.1
            r'^表[0-9]+.*$',                    # 表5
            r'^表\s+[0-9]+.*$',                # 表 5.1
        ]

        # 图标题模式（支持各种变体）
        figure_patterns = [
            r'^图\s*[0-9]+[-\.][0-9]+.*$',      # 图1.1...
            r'^图\s*[0-9]+[-\.][0-9]+\s*[:：].*$', # 图1.1: ...
            r'^图\s*[0-9]+.*$',                 # 图1...
            r'^图[0-9]+[-\.][0-9]+.*$',         # 图1.1
            r'^图[0-9]+.*$',                    # 图1
            r'^图\s+[0-9]+.*$',                # 图 1.1
        ]

        # 图表注释模式（单独处理，不与图表标题冲突）
        annotation_patterns = [
            r'^资料来源.*$',                   # 资料来源
            r'^数据来.*$',                     # 数据来源
            r'^来源.*$',                       # 来源
            r'^注.*$',                         # 注
        ]

        import re
        for pattern in table_patterns:
            if re.match(pattern, text):
                return (True, "table_caption")

        for pattern in figure_patterns:
            if re.match(pattern, text):
                return (True, "figure_caption")

        # 检查是否为图表注释（单独处理）
        for pattern in annotation_patterns:
            if re.match(pattern, text):
                return (True, "table_note")

        return None
