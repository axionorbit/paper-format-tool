# -*- coding: utf-8 -*-
"""
AI识别模块使用示例
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.fusion_engine import FusionIdentifier, DocumentStructureValidator, validate_structure


def example_basic_identification():
    """基础识别示例"""
    print("=== 示例1：基础识别（规则 + AI） ===\n")

    # 创建融合识别器（不设置API密钥，只使用规则识别）
    identifier = FusionIdentifier()

    # 模拟段落文本
    test_texts = [
        "第一章 绪论",           # 一级标题（规则可识别）
        "1.1 研究背景",        # 二级标题（规则可识别）
        "1.1.1 研究意义",      # 三级标题（规则可识别）
        "本文主要研究...",       # 正文（有句号，不需要AI）
        "重要结论",            # 可能的标题（无句号，少于30字，需要AI）
    ]

    for text in test_texts:
        # 这里创建模拟段落对象
        class MockParagraph:
            def __init__(self, text):
                self.text_value = text
            def text(self):
                return self.text_value

        paragraph = MockParagraph(text)

        part_type, is_title = identifier.identify_paragraph(paragraph)

        print(f"文本: {text}")
        print(f"  识别结果: {part_type or 'None'}, 是否标题: {is_title}")
        print()


def example_document_structure_validation():
    """文档结构验证示例"""
    print("=== 示例2：文档结构验证 ===\n")

    # 情况1：正常的文档结构
    print("情况1：正常结构 (H1 -> H2 -> H3)")
    normal_structure = ["heading1", "heading2", "heading3"]
    result = validate_structure(normal_structure)
    print(f"  是否有效: {result['is_valid']}")
    print(f"  缺失级别: {result['missing_levels']}")
    print(f"  H4无H3: {result['has_h4_without_h3']}")
    print(f"  建议: {[s['message'] for s in result['ai_suggestions']]}")
    print()

    # 情况2：缺少H1
    print("情况2：缺少H1")
    missing_h1_structure = ["heading2", "heading3"]
    result = validate_structure(missing_h1_structure)
    print(f"  是否有效: {result['is_valid']}")
    print(f"  缺失级别: {result['missing_levels']}")
    print(f"  建议: {[s['message'] for s in result['ai_suggestions']]}")
    print()

    # 情况3：缺少H2
    print("情况3：缺少H2")
    missing_h2_structure = ["heading1", "heading3"]
    result = validate_structure(missing_h2_structure)
    print(f"  是否有效: {result['is_valid']}")
    print(f"  缺失级别: {result['missing_levels']}")
    print(f"  建议: {[s['message'] for s in result['ai_suggestions']]}")
    print()

    # 情况4：有H4但无H3
    print("情况4：有H4但无H3")
    invalid_structure = ["heading1", "heading2", "heading4"]
    result = validate_structure(invalid_structure)
    print(f"  是否有效: {result['is_valid']}")
    print(f"  H4无H3: {result['has_h4_without_h3']}")
    print(f" 建议: {[s['message'] for s in result['ai_suggestions']]}")
    print()


def example_with_ai_key():
    """使用AI密钥的示例"""
    print("=== 示例3：使用AI密钥 ===\n")

    # 设置API密钥后，AI识别会被启用
    api_key = "your-claude-api-key"

    # 创建融合识别器（启用AI）
    identifier = FusionIdentifier(ai_key=None)  # 传入api_key或通过set_ai_api_key设置

    # 当AI启用时，满足条件的段落会被AI识别
    # 条件：被识别成正文 + 无句号 + 少于30字
    test_text = "研究结论与分析"

    class MockParagraph:
        def __init__(self, text):
            self.text_value = text
        def text(self):
            return self.text_value

    paragraph = MockParagraph(test_text)
    part_type, is_title = identifier.identify_paragraph(paragraph)

    print(f"文本: {test_text}")
    print(f"  识别结果: {part_type}, 是否标题: {is_title}")
    print(f"  （当AI启用时，会调用AI进行识别）")


if __name__ == "__main__":
    example_basic_identification()
    print("\n" + "="*60 + "\n")
    example_document_structure_validation()
    print("\n" + "="*60 + "\n")
    example_with_ai_key()
