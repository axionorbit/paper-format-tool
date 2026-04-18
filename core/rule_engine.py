# -*- coding: utf-8 -*-
"""
规则引擎模块 - 基于规则的论文部分识别
"""

import re
from docx.oxml.ns import qn


class PartIdentifier:
    """论文部分识别器"""

    # 一级标题：第X章（支持中文数字和阿拉伯数字）
    HEADING1_PATTERN = re.compile(r'^第[一二三四五六七八九十百0-9]+章\s*.*$')

    # 二级标题：第X节（支持中文数字和阿拉伯数字）
    HEADING2_PATTERN = re.compile(r'^第[一二三四五六七八九十百0-9]+节\s*.*$')

    # 三级标题：
    # - 1.1.1 ××× 格式（阿拉伯数字，三个数字用.分隔）
    # - 一、××× 格式（中文数字后跟、）
    HEADING3_PATTERN = re.compile(r'^(?:[0-9]+\.[0-9]+\.[0-9]+(?!\.[0-9])\s*|[一二三四五六七八九十百]+、)\s*.*$')

    # 四级标题：
    # - 1.1.1.1 ××× 格式（阿拉伯数字，四个数字用.分隔）
    # - （一）××× 格式（中文数字在括号内）
    HEADING4_PATTERN = re.compile(r'^(?:[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+\s*|（[一二三四五六七八九十百]+）)\s*.*$')

    # 阿拉伯数字编号的标题（如1. 2. 1.1 2.1等）
    ARABIC_HEADING1_PATTERN = re.compile(r'^[0-9]+\s+.*$')  # 1 人民币国际化
    ARABIC_HEADING1_DOT_PATTERN = re.compile(r'^[0-9]+\s*[、．．]\s*.*$')  # 1、或1．
    # 使用负向前瞻避免前缀误匹配：
    # 例如 1.1.1 不应被 1.1 规则吞掉
    ARABIC_HEADING2_PATTERN = re.compile(r'^[0-9]+\.[0-9]+(?!\.[0-9])\s*.*$')  # 1.1
    ARABIC_HEADING3_PATTERN = re.compile(r'^[0-9]+\.[0-9]+\.[0-9]+(?!\.[0-9])\s*.*$')  # 1.1.1

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
    def _is_title_candidate(text):
        """判断是否为标题候选：短文本且末尾没有句号类标点，允许存在换行"""
        if not text:
            return False

        compact_text = text.strip()
        char_count_without_spaces = len(
            compact_text.replace(" ", "").replace("\u3000", "")
        )
        if char_count_without_spaces == 0 or char_count_without_spaces > 30:
            return False

        end_punctuation = ('。', '！', '？', '.', '!', '?')
        return not compact_text.endswith(end_punctuation)

    @staticmethod
    def _iter_style_chain(paragraph):
        """遍历段落样式链（当前样式 -> 基样式）"""
        style = getattr(paragraph, "style", None)
        visited = set()
        while style is not None and id(style) not in visited:
            yield style
            visited.add(id(style))
            style = getattr(style, "base_style", None)

    @staticmethod
    def _extract_heading_level_from_style_name(style_name):
        """从样式名中提取标题级别（Heading/标题/List Number）"""
        if not style_name:
            return None

        normalized = re.sub(r'\s+', '', str(style_name)).lower()

        # Heading 1 / 标题 1
        match = re.match(r'^(?:heading|标题)([1-9])$', normalized)
        if match:
            return int(match.group(1))

        # List Number 2 / 列表编号2 / 编号2
        match = re.match(r'^(?:listnumber|列表编号|编号)([1-9])$', normalized)
        if match:
            return int(match.group(1))

        return None

    @staticmethod
    def _find_numpr(paragraph):
        """
        获取段落编号定义：
        1) 段落自身 pPr.numPr
        2) 样式链 pPr.numPr
        """
        try:
            ppr = paragraph._element.pPr
            if ppr is not None and ppr.numPr is not None:
                return ppr.numPr
        except Exception:
            pass

        for style in PartIdentifier._iter_style_chain(paragraph):
            try:
                style_ppr = style.element.pPr
                if style_ppr is not None and style_ppr.numPr is not None:
                    return style_ppr.numPr
            except Exception:
                continue

        return None

    @staticmethod
    def _safe_numpr_val(numpr, attr_name):
        try:
            node = getattr(numpr, attr_name, None)
            if node is None:
                return None
            return node.val
        except Exception:
            return None

    @staticmethod
    def _extract_heading_level_from_numbering_definition(paragraph, num_id, ilvl):
        """从 numbering 定义中的 lvlText（如 %1.%2）推断标题层级"""
        if num_id is None:
            return None

        try:
            numbering_root = paragraph.part.numbering_part.element
        except Exception:
            return None

        try:
            target_abstract_num_id = None
            for num_node in numbering_root.findall(qn('w:num')):
                if num_node.get(qn('w:numId')) == str(num_id):
                    abstract_id_node = num_node.find(qn('w:abstractNumId'))
                    if abstract_id_node is not None:
                        target_abstract_num_id = abstract_id_node.get(qn('w:val'))
                    break

            if target_abstract_num_id is None:
                return None

            target_abstract = None
            for abstract_node in numbering_root.findall(qn('w:abstractNum')):
                if abstract_node.get(qn('w:abstractNumId')) == str(target_abstract_num_id):
                    target_abstract = abstract_node
                    break

            if target_abstract is None:
                return None

            lvl_node = None
            if ilvl is not None:
                for node in target_abstract.findall(qn('w:lvl')):
                    if node.get(qn('w:ilvl')) == str(ilvl):
                        lvl_node = node
                        break
            if lvl_node is None:
                lvl_nodes = target_abstract.findall(qn('w:lvl'))
                if lvl_nodes:
                    lvl_node = lvl_nodes[0]

            if lvl_node is None:
                return None

            lvl_text_node = lvl_node.find(qn('w:lvlText'))
            lvl_text = lvl_text_node.get(qn('w:val')) if lvl_text_node is not None else ""
            if not lvl_text:
                return None

            level_tokens = re.findall(r'%\d+', lvl_text)
            level = len(level_tokens)
            if 1 <= level <= 4:
                return level
        except Exception:
            return None

        return None

    @staticmethod
    def _infer_heading_level_by_style_or_numbering(paragraph, text, is_title_candidate, char_count_without_spaces):
        """
        兜底标题识别（用于 Word 自动编号等场景）
        """
        if not text or not is_title_candidate or char_count_without_spaces >= 30:
            return None

        style_level = None
        for style in PartIdentifier._iter_style_chain(paragraph):
            level = PartIdentifier._extract_heading_level_from_style_name(
                getattr(style, "name", "")
            )
            if level and 1 <= level <= 4:
                style_level = level
                break

        # 1) 编号元数据优先判定（numPr + ilvl / lvlText）
        numpr = PartIdentifier._find_numpr(paragraph)
        if numpr is not None:
            num_id = PartIdentifier._safe_numpr_val(numpr, "numId")
            ilvl = PartIdentifier._safe_numpr_val(numpr, "ilvl")

            level = PartIdentifier._extract_heading_level_from_numbering_definition(
                paragraph, num_id, ilvl
            )
            if level and 1 <= level <= 4:
                # 对于仅解析出一级且编号层信息缺失的样式，允许样式级别补偿
                if level == 1 and ilvl is None and style_level and style_level > 1:
                    return style_level
                return level

            # 2) 次级兜底：仅有 ilvl 时（常见于某些自动编号模板）
            if ilvl is not None:
                try:
                    ilvl_int = int(str(ilvl))
                    # ilvl=0 更可能是普通编号列表，这里保守忽略，避免把正文列表识别成一级标题
                    if ilvl_int >= 1:
                        level = ilvl_int + 1
                        if 1 <= level <= 4:
                            return level
                except Exception:
                    pass

        # 3) 样式名直判（Heading/标题/List Number 2 等）
        if style_level and 1 <= style_level <= 4:
            return style_level

        return None

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
    def _extract_numbering_prefix(paragraph):
        """
        Extract visible numbering prefix from Word numbering definition.
        Only used to augment text for regex matching.
        """
        numpr = PartIdentifier._find_numpr(paragraph)
        if numpr is None:
            return ""

        num_id = PartIdentifier._safe_numpr_val(numpr, "numId")
        ilvl = PartIdentifier._safe_numpr_val(numpr, "ilvl")
        ilvl_int = None
        if ilvl is not None:
            try:
                ilvl_int = int(str(ilvl))
            except Exception:
                ilvl_int = None

        lvl_text = ""
        if num_id is not None:
            try:
                numbering_root = paragraph.part.numbering_part.element
                target_abstract_num_id = None
                for num_node in numbering_root.findall(qn("w:num")):
                    if num_node.get(qn("w:numId")) == str(num_id):
                        abstract_id_node = num_node.find(qn("w:abstractNumId"))
                        if abstract_id_node is not None:
                            target_abstract_num_id = abstract_id_node.get(qn("w:val"))
                        break

                if target_abstract_num_id is not None:
                    target_abstract = None
                    for abstract_node in numbering_root.findall(qn("w:abstractNum")):
                        if abstract_node.get(qn("w:abstractNumId")) == str(target_abstract_num_id):
                            target_abstract = abstract_node
                            break

                    if target_abstract is not None:
                        lvl_node = None
                        if ilvl_int is not None:
                            for node in target_abstract.findall(qn("w:lvl")):
                                if node.get(qn("w:ilvl")) == str(ilvl_int):
                                    lvl_node = node
                                    break
                        if lvl_node is None:
                            all_lvls = target_abstract.findall(qn("w:lvl"))
                            if all_lvls:
                                lvl_node = all_lvls[0]

                        if lvl_node is not None:
                            lvl_text_node = lvl_node.find(qn("w:lvlText"))
                            if lvl_text_node is not None:
                                lvl_text = lvl_text_node.get(qn("w:val")) or ""
            except Exception:
                lvl_text = ""

        prefix = ""
        if lvl_text:
            prefix = re.sub(r"%\d+", "1", lvl_text).strip()
            prefix = re.sub(r"[\.\u3002\uFF0E、\s]+$", "", prefix)
        if not prefix and ilvl_int is not None and ilvl_int >= 0:
            prefix = ".".join(["1"] * (ilvl_int + 1))
        return prefix

    @staticmethod
    def _build_heading_match_text(paragraph, text):
        if not text or not PartIdentifier._is_title_candidate(text):
            return text

        prefix = PartIdentifier._extract_numbering_prefix(paragraph).strip()
        if not prefix or text.startswith(prefix):
            return text
        return f"{prefix} {text}"

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
        heading_match_text = PartIdentifier._build_heading_match_text(paragraph, text)
        is_title_candidate = PartIdentifier._is_title_candidate(heading_match_text)

        # 1. 特殊标题检查（支持更多变体）
        if normalized in ["摘要", "摘 要", "ABSTRACT", "Abstract"]:
            return "abstract_title", True
        if normalized in ["参考文献", "参 考 文 献", "REFERENCE", "Reference", "REFERENCES", "References", "οο"]:
            return "ref_title", True
        if normalized in ["致谢", "致 谢", "谢辞", "ACKNOWLEDGEMENT", "Acknowledgement", "ACKNOWLEDGEMENTS", "Acknowledgements"]:
            return "ack_title", True
        if normalized in ["附录", "附 录", "APPENDIX", "Appendix", "APPENDICES", "Appendices"]:
            return "appendix_title", True

        # 2. 标题级别检查（标题候选段落）
        if is_title_candidate:
            if PartIdentifier.HEADING1_PATTERN.match(heading_match_text):
                return "heading1", True
            if PartIdentifier.HEADING2_PATTERN.match(heading_match_text):
                return "heading2", True
            if PartIdentifier.HEADING4_PATTERN.match(heading_match_text):
                return "heading4", True
            if PartIdentifier.HEADING3_PATTERN.match(heading_match_text):
                return "heading3", True

            # 阿拉伯数字编号（先高层级后低层级，避免前缀吞并）
            if PartIdentifier.ARABIC_HEADING3_PATTERN.match(heading_match_text):
                return "heading3", True
            if PartIdentifier.ARABIC_HEADING2_PATTERN.match(heading_match_text):
                return "heading2", True
            if (
                PartIdentifier.ARABIC_HEADING1_PATTERN.match(heading_match_text)
                or PartIdentifier.ARABIC_HEADING1_DOT_PATTERN.match(heading_match_text)
            ):
                return "heading1", True

        # 3. 图表标题检查（标题候选段落）
        if is_title_candidate:
            if PartIdentifier.FIGURE_PATTERN.match(text):
                return "figure_caption", True
            if PartIdentifier.TABLE_PATTERN.match(text):
                return "table_caption", True

        # 4. 表格注释
        if PartIdentifier.TABLE_NOTE_PATTERN.match(text):
            return "table_note", True

        # 5. 样式/编号兜底识别（放在文本正则后、正文前）
        fallback_level = None
        if fallback_level:
            return f"heading{fallback_level}", True

        # 6. 根据上下文和内容判断区域
        if position_context:
            if position_context == "abstract":
                return "abstract_content", False
            elif position_context == "ref":
                return "ref_content", False
            elif position_context == "ack":
                return "ack_content", False
            elif position_context == "appendix":
                return "appendix_content", False

        # 7. 检查是否为参考文献内容（在参考文献标题后且不是其他标题）
        if position_context == "ref":
            return "ref_content", False

        # 8. 默认为正文
        return "body", False
