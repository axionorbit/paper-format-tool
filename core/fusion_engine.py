# -*- coding: utf-8 -*-
"""
融合引擎模块 - 结合规则识别和AI识别
"""

from typing import Optional, Tuple, List, Dict

from core.rule_engine import PartIdentifier
from core.ai_engine import AIIdentifier, get_ai_identifier, is_ai_enabled


class FusionIdentifier:
    """融合识别器 - 结合规则和AI的识别结果"""

    def __init__(self, ai_identifier: Optional[AIIdentifier] = None):
        """
        初始化融合识别器
        ai_identifier: AI识别器实例，如果为None则使用全局实例
        """
        self.ai_identifier = ai_identifier or get_ai_identifier()
        self.rule_identifier = PartIdentifier()
        self.last_heading_level = None  # 记录上一个标题级别
        self.has_body_since_last_heading = False  # 标记自上一个标题后是否有正文

    def identify_paragraph(
        self,
        paragraph,
        position_context: Optional[str] = None
    ) -> Tuple[Optional[str], bool]:
        """
        识别段落类型（规则 + AI融合）
        返回: (part_type, is_title)
        """
        # 1. 先使用规则识别
        part_type, is_title = self.rule_identifier.identify(
            paragraph, position_context
        )

        # 2. 更新状态（用于检测连续标题）
        self._update_state(part_type, is_title)

        # 3. 检查是否需要AI辅助识别
        if self._should_use_ai_recognition(paragraph, part_type, is_title):
            ai_part_type, ai_is_title = self._identify_with_ai(
                paragraph, position_context
            )
            # 如果AI识别成功，使用AI结果
            if ai_part_type:
                part_type = ai_part_type
                is_title = ai_is_title

        return part_type, is_title

    def _update_state(self, part_type: Optional[str], is_title: bool):
        """更新状态，用于检测连续标题等异常情况"""
        if is_title and part_type and part_type.startswith("heading"):
            self.last_heading_level = part_type
            self.has_body_since_last_heading = False
        elif part_type == "body" and self.last_heading_level:
            self.has_body_since_last_heading = True

    def _should_use_ai_recognition(
        self,
        paragraph,
        part_type: Optional[str],
        is_title: bool
    ) -> bool:
        """
        判断是否需要使用AI识别
        条件：
        1. AI识别已启用
        2. 被识别成正文
        3. 没有句号
        4. 字数少于30个字
        5. 同一级别的标题连续出现（中间没有正文）
        """
        if not is_ai_enabled():
            return False

        text = paragraph.text.strip()

        # 条件5：同级别标题连续出现（中间没有正文）
        # 检测情况：1.1 标题A → 1.2 标题B（之间没有正文）
        if is_title and part_type and part_type.startswith("heading"):
            if part_type == self.last_heading_level and not self.has_body_since_last_heading:
                # 同级别标题连续出现，中间没有正文 - 异常情况
                return True

        # 条件2-4：正文相关的检查
        if part_type != "body":
            return False

        # 检查是否有句号
        end_punctuation = ('。', '！', '？', '.', '!', '?')
        has_end_punctuation = text.endswith(end_punctuation)
        if has_end_punctuation:
            return False

        # 检查字数是否少于30个字
        char_count = len(text)
        if char_count >= 30:
            return False

        # 满足所有条件，需要AI识别
        return True

    def _identify_with_ai(
        self,
        paragraph,
        position_context: Optional[str] = None
    ) -> Tuple[Optional[str], bool]:
        """
        使用AI识别段落类型
        """
        if not self.ai_identifier:
            return None, False

        text = paragraph.text.strip()
        context_info = {
            "position_context": position_context,
            "text_length": len(text)
        }

        return self.ai_identifier.identify_paragraph(text, context_info)


class DocumentStructureValidator:
    """文档结构验证器"""

    def __init__(self, ai_identifier: Optional[AIIdentifier] = None):
        """
        初始化结构验证器
        """
        self.ai_identifier = ai_identifier or get_ai_identifier()

    def validate_document_structure(
        self,
        heading_sequence: List[str]
    ) -> Dict:
        """
        验证文档结构
        heading_sequence: 文档中所有标题的级别列表

        返回验证结果:
        {
            "is_valid": bool,           # 结构是否有效
            "missing_levels": List[str],   # 缺失的级别
            "has_h4_without_h3": bool,   # 是否有H4但没有H3
            "details": dict,              # 详细信息
            "ai_suggestions": List[dict],  # AI建议（如果有）
        }
        """
        result = {
            "is_valid": True,
            "missing_levels": [],
            "has_h4_without_h3": False,
            "details": {
                "has_h1": False,
                "has_h2": False,
                "has_h3": False,
                "has_h4": False,
                "total_headings": len(heading_sequence),
            },
            "ai_suggestions": []
        }

        # 1. 检查各个级别是否存在
        for heading in heading_sequence:
            if heading == "heading1":
                result["details"]["has_h1"] = True
            elif heading == "heading2":
                result["details"]["has_h2"] = True
            elif heading == "heading3":
                result["details"]["has_h3"] = True
            elif heading == "heading4":
                result["details"]["has_h4"] = True

        # 2. 验证规则
        if not result["details"]["has_h1"]:
            result["is_valid"] = False
            result["missing_levels"].append("heading1")
            result["ai_suggestions"].append({
                "type": "missing_heading",
                "level": "heading1",
                "message": "文档中缺少一级标题（如"第X章"或"1"格式）",
            })

        if not result["details"]["has_h2"]:
            result["is_valid"] = False
            result["missing_levels"].append("heading2")
            result["ai_suggestions"].append({
                "type": "missing_heading",
                "level": "heading2",
                "message": "文档中缺少二级标题（如"第X节"或"1.1"格式）",
            })

        if result["details"]["has_h4"] and not result["details"]["has_h3"]:
            result["is_valid"] = False
            result["has_h4_without_h3"] = True
            result["ai_suggestions"].append({
                "type": "invalid_structure",
                "message": "文档中存在四级标题（"1.1.1.1"格式）但没有三级标题",
            })

        # 3. 如果AI可用，可以生成更多建议
        if is_ai_enabled() and result["ai_suggestions"]:
            result["ai_suggestions"].append({
                "type": "ai_assistance_available",
                "message": "AI辅助识别已启用，可以帮助识别可能的标题",
            })

        return result


# 便捷函数

def create_fusion_identifier(api_key: Optional[str] = None) -> FusionIdentifier:
    """创建融合识别器（便捷函数）"""
    if api_key:
        from core.ai_engine import set_ai_api_key
        set_ai_api_key(api_key)
    return FusionIdentifier()


def validate_structure(
    heading_sequence: List[str],
    api_key: Optional[str] = None
) -> Dict:
    """验证文档结构（便捷函数）"""
    if api_key:
        from core.ai_engine import set_ai_api_key
        set_ai_api_key(api_key)
    validator = DocumentStructureValidator()
    return validator.validate_document_structure(heading_sequence)
