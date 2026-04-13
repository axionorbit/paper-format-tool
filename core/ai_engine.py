# -*- coding: utf-8 -*-
"""
AI引擎模块 - 使用AI辅助识别段落类型
"""

from typing import Optional, Tuple


class AIIdentifier:
    """AI识别器 - 使用AI辅助判断段落类型"""

    def __init__(self, api_key: Optional[str] = None):
        """
        初始化AI识别器
        api_key: Claude API密钥（可选，也可从配置文件读取）
        """
        self.api_key = api_key
        self.enabled = bool(api_key)  # 有API密钥时才启用

    def identify_paragraph(
        self,
        text: str,
        context: Optional[dict] = None
    ) -> Tuple[Optional[str], bool]:
        """
        使用AI识别段落类型
        返回: (part_type, is_title)
        """
        if not self.enabled:
            return None, False

        # TODO: 调用Claude API进行识别
        # 这里先返回None，表示AI未识别
        return None, False

    def validate_heading_structure(self, heading_sequence: list) -> dict:
        """
        验证标题层级结构是否正确
        heading_sequence: 文档中所有标题的级别列表，如 ["heading1", "heading2", "heading3"]

        返回验证结果:
        {
            "is_valid": bool,  # 结构是否有效
            "missing_levels": list,  # 缺失的级别
            "has_h4_without_h3": bool,  # 是否有H4但没有H3
            "has_h1": bool,
            "has_h2": bool,
            "has_h3": bool,
            "has_h4": bool,
        }
        """
        result = {
            "is_valid": True,
            "missing_levels": [],
            "has_h4_without_h3": False,
            "has_h1": False,
            "has_h2": False,
            "has_h3": False,
            "has_h4": False,
        }

        # 检查各个级别是否存在
        for heading in heading_sequence:
            if heading == "heading1":
                result["has_h1"] = True
            elif heading == "heading2":
                result["has_h2"] = True
            elif heading == "heading3":
                result["has_h3"] = True
            elif heading == "heading4":
                result["has_h4"] = True

        # 验证规则
        if not result["has_h1"]:
            result["is_valid"] = False
            result["missing_levels"].append("heading1")

        if not result["has_h2"]:
            result["is_valid"] = False
            result["missing_levels"].append("heading2")

        if result["has_h4"] and not result["has_h3"]:
            result["is_valid"] = False
            result["has_h4_without_h3"] = True

        return result


# 全局AI识别器实例（可配置）
_global_ai_identifier = None


def get_ai_identifier(api_key: Optional[str] = None) -> AIIdentifier:
    """获取全局AI识别器实例"""
    global _global_ai_identifier
    if _global_ai_identifier is None:
        _global_ai_identifier = AIIdentifier(api_key)
    return _global_ai_identifier


def set_ai_api_key(api_key: str):
    """设置AI API密钥"""
    global _global_ai_identifier
    _global_ai_identifier = AIIdentifier(api_key)


def is_ai_enabled() -> bool:
    """检查AI识别是否已启用"""
    return _global_ai_identifier is not None and _global_ai_identifier.enabled
