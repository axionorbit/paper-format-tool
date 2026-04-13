# -*- coding: utf-8 -*-
"""
规则引擎模块 - 基于规则的论文部分识别
"""

import re


class PartIdentifier:
    """论文部分识别器"""

    # 一级标题：第X章（支持中文数字和阿拉伯数字）
    HEADING1_PATTERN = re.compile(r'^第[一二三四五六七八九十百0-9]+章\s*.*$')

    # 二级标题：第X节（支持中文数字和阿拉伯数字）
    HEADING2_PATTERN = re.compile(r'^第[一二三四五六七八九十百0-9]+节\s*.*$')

    # 三级标题：
    # - 1.1.1 ××× 格式（阿拉伯数字，三个数字用.分隔）
    # - 一、××× 格式（中文数字后跟、）
    HEADING3_PATTERN = re.compile(r'^(?:[0-9]+\.[0-9]+\.[0-9]+\s+|[一二三四五六七八九十百]+、)\s*.*$')

    # 四级标题：
    # - 1.1.1.1 ××× 格式（阿拉伯数字，四个数字用.分隔）
    # - （一）××× 格式（中文数字在括号内）
    HEADING4_PATTERN = re.compile(r'^(?:[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+\s+|（[一二三四五六七八九十百]+）)\s*.*$')

    # 阿拉伯数字编号的标题（如1. 2. 1.1 2.1等）
    ARABIC_HEADING1_PATTERN = re.compile(r'^[0-9]+\s+.*$')  # 1 人民币国际化
    ARABIC_HEADING1_DOT_PATTERN = re.compile(r'^[0-9]+\s*[、．．]\s*.*$')  # 1、或1．
    ARABIC_HEADING2_PATTERN = re.compile(r'^[0-9]+\.[0-9]+\s*.*$')  # 1.1
    ARABIC_HEADING3_PATTERN = re.compile(r'^[0-9]+\.[0-9]+\.[0-9]+\s*.*$')  # 1.1.1

    # 图片标题：图X-X、图X.X、图X等（宽松的匹配）
    FIGURE_PATTERN = re.compile(r'^图\s*[0-9]+([-\.][0-9]+)?.*$')

    # 表格标题：表X-X、表X.X、表X等（宽松的匹配）
    TABLE_PATTERN = re.compile(r'^表\s*[0-9]+([-\.][0-9]+)?.*$')

    # 表格注释：注：或 注:
    TABLE_NOTE_PATTERN = re.compile(r'^注[：:].*$')
    FORMULA_NUMBER_PATTERN = re.compile(r'^\s*[\(\uff08]?\s*\d+(?:[-\.]\d+)*\s*[\)\uff09]?\s*$')

    # 特殊标题（允许中间有空格）
    @staticmethod
    def _normalize(text):
        """去除所有空格用于匹配"""
        return text.replace(" ", "").replace("\u3000", "")

    @staticmethod
    def _has_formula_omml(paragraph):
        """检测段落中是否包含 Word 插入公式生成的 OMML 节点"""
        try:
            return bool(paragraph._element.xpath('.//*[local-name()="oMath" or local-name()="oMathPara"]'))
        except Exception:
            return False

    @staticmethod
    def _is_formula_paragraph(paragraph, text):
        """判断是否为独立公式段落"""
        if not PartIdentifier._has_formula_omml(paragraph):
            return False

        normalized = PartIdentifier._normalize(text)

        # Word 公式在 python-docx 中经常显示为空文本
        if not normalized:
            return True

        # 仅带公式编号的段落
        if PartIdentifier.FORMULA_NUMBER_PATTERN.match(normalized):
            return True

        # 兼容"公式 + 编号"这类短文本段落，避免把正文中的行内公式误判成整段公式
        non_formula_text = re.sub(r'[\s\(\)\uff08\uff090-9\-\.=+*/<>≤≥,，:：;；]+', '', normalized)
        return len(non_formula_text) <= 6

    @staticmethod
    def identify(paragraph, position_context=None):
        """
        识别段落类型
        返回: (part_type, is_title)
        """
        text = paragraph.text.strip()

        if PartIdentifier._is_formula_paragraph(paragraph, text):
            return "formula", False

        # 空段落
        if not text:
            return None, False

        normalized = PartIdentifier._normalize(text)

        # 改进的单独一行判断：更智能的标题识别逻辑
        has_end_punctuation = text.endswith(('。', '！', '？', '.', '!', '?', '；', ';', '：', ':'))
        # 更宽松的长度限制，因为中文标题可能较长
        is_short_text = len(text) < 150
        # 检查是否包含换行符或其他段落特征
        has_line_breaks = '\n' in text or '\r' in text
        # 检查是否像标题（包含标题关键词且不以句号结束）
        contains_title_keywords = any(keyword in text for keyword in ['第', '章', '节', '、', '（', '）', '图', '表'])

        # 标题字数限制：超过50个字符（不含空格）不算标题
        char_count_without_spaces = len(text.replace(" ", "").replace("\u3000", ""))
        is_too_long_for_title = char_count_without_spaces > 50

        is_single_line = (is_short_text and not has_end_punctuation and not has_line_breaks and not is_too_long_for_title) or \
                        (contains_title_keywords and not is_too_long_for_title)

        # 1. 特殊标题检查（支持更多变体）
        if normalized in ["摘要", "摘 要", "ABSTRACT", "Abstract"]:
            return "abstract_title", True
        if normalized in ["参考文献", "参 考 文 献", "REFERENCE", "Reference", "REFERENCES", "References", "οο"]:
            return "ref_title", True
        if normalized in ["致谢", "致 谢", "谢辞", "ACKNOWLEDGEMENT", "Acknowledgement", "ACKNOWLEDGEMENTS", "Acknowledgements"]:
            return "ack_title", True
        if normalized in ["附录", "附 录", "APPENDIX", "Appendix", "APPENDICES", "Appendices"]:
            return "appendix_title", True

        # 2. 标题级别检查（需要单独一行）
        if is_single_line:
            if PartIdentifier.HEADING1_PATTERN.match(text):
                return "heading1", True
            if PartIdentifier.HEADING2_PATTERN.match(text):
                return "heading2", True
            if PartIdentifier.HEADING3_PATTERN.match(text):
                return "heading3", True
            if PartIdentifier.HEADING4_PATTERN.match(text):
                return "heading4", True

            # 阿拉伯数字编号的标题
            if PartIdentifier.ARABIC_HEADING2_PATTERN.match(text) or PartIdentifier.ARABIC_HEADING3_PATTERN.match(text):
                return "heading2", True
            if PartIdentifier.ARABIC_HEADING1_PATTERN.match(text) or PartIdentifier.ARABIC_HEADING1_DOT_PATTERN.match(text):
                return "heading1", True

        # 3. 图表标题检查（需要单独一行）
        if is_single_line:
            if PartIdentifier.FIGURE_PATTERN.match(text):
                return "figure_caption", True
            if PartIdentifier.TABLE_PATTERN.match(text):
                return "table_caption", True

        # 4. 表格注释
        if PartIdentifier.TABLE_NOTE_PATTERN.match(text):
            return "table_note", True

        # 5. 根据上下文和内容判断区域
        if position_context:
            if position_context == "abstract":
                return "abstract_content", False
            elif position_context == "ref":
                return "ref_content", False
            elif position_context == "ack":
                return "ack_content", False
            elif position_context == "appendix":
                return "appendix_content", False

        # 5. 检查是否为参考文献内容（在参考文献标题后且不是其他标题）
        if position_context == "ref":
            return "ref_content", False

        # 6. 默认为正文
        return "body", False
