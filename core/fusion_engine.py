# -*- coding: utf-8 -*-
"""
融合引擎模块 - 结合规则识别和AI识别
"""

from typing import Callable, Dict, List, Optional, Tuple

from core.parser import DocumentParser, ParagraphUnit
from core.rule_engine import PartIdentifier
from core.ai_engine import AIIdentifier, get_ai_identifier, is_ai_enabled
from utils.logger import default_logger


class FusionIdentifier:
    """融合识别器 - 文档级三阶段识别（规则 -> 异常检测 -> AI）"""

    TITLE_LABELS = {
        "heading1",
        "heading2",
        "heading3",
        "heading4",
        "figure_caption",
        "table_caption",
        "table_note",
        "abstract_title",
        "ref_title",
        "ack_title",
        "appendix_title",
    }

    def __init__(self, ai_identifier: Optional[AIIdentifier] = None):
        """
        初始化融合识别器
        ai_identifier: AI识别器实例，如果为None则使用全局实例
        """
        self.ai_identifier = ai_identifier or get_ai_identifier()
        self.rule_identifier = PartIdentifier()
        self.parser = DocumentParser()

        default_logger.info(
            "AI status: enabled=%s, model=%s, api_key_configured=%s",
            is_ai_enabled(self.ai_identifier),
            getattr(self.ai_identifier, "model", "unknown"),
            bool(getattr(self.ai_identifier, "api_key", "")),
        )

    def identify_paragraph(
        self,
        paragraph,
        position_context: Optional[str] = None
    ) -> Tuple[Optional[str], bool]:
        """
        单段识别（优先规则；满足条件时尝试AI）
        返回: (part_type, is_title)
        """
        part_type, is_title = self.rule_identifier.identify(
            paragraph, position_context
        )

        # 为保证流程一致性，单段识别不直接触发 AI。
        # AI 仅在 identify_document 的“规则识别 -> 异常检测 -> AI识别”流程中启动。
        return part_type, is_title

    def _identify_document_legacy(self, units: List[ParagraphUnit]) -> Dict:
        """
        文档级识别流程:
        1. 规则识别
        2. 异常检测（是否需要AI）
        3. AI识别候选段落并覆盖候选标签
        """
        return self._identify_document_with_stage(units, stage_callback=None)

    def _identify_document_with_stage(
        self,
        units: List[ParagraphUnit],
        stage_callback: Optional[Callable[[str, Dict], None]] = None,
    ) -> Dict:
        """
        文档级识别流程（带阶段回调）:
        1. 规则识别
        2. 异常检测（是否需要AI）
        3. AI识别候选段落并覆盖候选标签
        """
        self._emit_stage(stage_callback, "rule_start")
        rule_results = self._run_rule_stage(units)
        self._emit_stage(stage_callback, "rule_done", total=len(rule_results))

        self._emit_stage(stage_callback, "anomaly_start")
        anomalies = self._detect_anomalies(units, rule_results)
        self._emit_stage(
            stage_callback,
            "anomaly_done",
            need_ai=anomalies.get("need_ai", False),
            reasons=anomalies.get("reasons", []),
        )

        final_results = [dict(item) for item in rule_results]
        ai_used = False
        ai_label_map: Dict[int, str] = {}
        ai_candidates = self.parser.extract_ai_candidates(units, max_chars=30)

        should_call_ai = anomalies["need_ai"] and is_ai_enabled(self.ai_identifier)
        default_logger.info(
            "AI trigger: should_call_ai=%s, reasons=%s, candidates=%s",
            should_call_ai,
            anomalies.get("reasons", []),
            len(ai_candidates),
        )

        if should_call_ai and ai_candidates:
            self._emit_stage(
                stage_callback,
                "ai_start",
                candidates=len(ai_candidates),
                reasons=anomalies.get("reasons", []),
            )
            ai_used = True
            candidate_payload = [
                {
                    "id": unit.index,
                    "text": unit.ai_text,
                }
                for unit in ai_candidates
            ]
            ai_context = {
                "trigger_reasons": anomalies.get("reasons", []),
                "summary": anomalies.get("summary", {}),
            }
            ai_label_map = self.ai_identifier.identify_candidates(
                candidate_payload,
                context=ai_context,
            )

            default_logger.info(
                "AI labeled paragraphs: %s",
                [(idx, label) for idx, label in sorted(ai_label_map.items())],
            )
            if not ai_label_map:
                ai_status = {}
                try:
                    ai_status = self.ai_identifier.get_last_status()
                except Exception:
                    ai_status = {}
                default_logger.warning(
                    "AI produced no labels. last_error=%s, response_preview=%s",
                    ai_status.get("last_error", ""),
                    ai_status.get("last_response_preview", "")[:200],
                )

            for idx, label in ai_label_map.items():
                if idx < 0 or idx >= len(final_results):
                    continue
                final_results[idx]["part_type"] = label
                final_results[idx]["is_title"] = label in self.TITLE_LABELS
                final_results[idx]["source"] = "ai"
            self._emit_stage(
                stage_callback,
                "ai_done",
                labeled=len(ai_label_map),
            )
        elif anomalies.get("need_ai") and not is_ai_enabled(self.ai_identifier):
            self._emit_stage(stage_callback, "ai_skipped_disabled")
        elif anomalies.get("need_ai") and not ai_candidates:
            self._emit_stage(stage_callback, "ai_skipped_no_candidates")
        else:
            self._emit_stage(stage_callback, "ai_not_needed")

        final_labels = {
            row["index"]: row["part_type"]
            for row in final_results
            if row.get("part_type") is not None
        }

        result = {
            "rule_results": rule_results,
            "final_results": final_results,
            "final_labels": final_labels,
            "anomalies": anomalies,
            "ai_used": ai_used,
            "ai_candidate_count": len(ai_candidates),
            "ai_labeled_count": len(ai_label_map),
        }
        self._emit_stage(
            stage_callback,
            "identify_done",
            ai_used=ai_used,
            ai_candidate_count=len(ai_candidates),
            ai_labeled_count=len(ai_label_map),
        )
        return result

    def identify_document(
        self,
        units: List[ParagraphUnit],
        stage_callback: Optional[Callable[[str, Dict], None]] = None,
    ) -> Dict:
        return self._identify_document_with_stage(units, stage_callback=stage_callback)

    def _emit_stage(
        self,
        callback: Optional[Callable[[str, Dict], None]],
        stage: str,
        **payload,
    ):
        if callback:
            callback(stage, payload)

    def _run_rule_stage(self, units: List[ParagraphUnit]) -> List[Dict]:
        """第一阶段：规则识别"""
        results: List[Dict] = []
        current_context = None

        for unit in units:
            part_type, is_title = self.rule_identifier.identify(
                unit.paragraph,
                current_context,
            )
            results.append(
                {
                    "index": unit.index,
                    "part_type": part_type,
                    "is_title": is_title,
                    "source": "rule",
                }
            )
            current_context = self._next_context(current_context, part_type, is_title)

        return results

    def _next_context(
        self,
        current_context: Optional[str],
        part_type: Optional[str],
        is_title: bool,
    ) -> Optional[str]:
        if not is_title:
            return current_context

        if part_type == "abstract_title":
            return "abstract"
        if part_type == "ref_title":
            return "ref"
        if part_type == "ack_title":
            return "ack"
        if part_type == "appendix_title":
            return "appendix"
        if part_type in {"heading1", "heading2", "heading3", "heading4"}:
            return "body"
        return current_context

    def _detect_anomalies(
        self,
        units: List[ParagraphUnit],
        rule_results: List[Dict],
    ) -> Dict:
        """
        第二阶段：异常检测（是否触发AI）
        条件：
        1. 某段被识别为正文，且无句末标点，且字数 < 30
        2. 文档中不存在一级标题
        3. 文档中不存在二级标题
        4. 同一级别连续出现，且中间没有其他内容
        5. 存在四级标题但不存在三级标题
        """
        has_h1 = False
        has_h2 = False
        has_h3 = False
        has_h4 = False
        short_body_without_period_indices: List[int] = []
        consecutive_same_level_pairs: List[Tuple[int, int, str]] = []

        last_heading_level = None
        last_heading_index = None
        has_content_since_last_heading = False

        for unit, result in zip(units, rule_results):
            part_type = result.get("part_type")

            if part_type == "heading1":
                has_h1 = True
            elif part_type == "heading2":
                has_h2 = True
            elif part_type == "heading3":
                has_h3 = True
            elif part_type == "heading4":
                has_h4 = True

            if (
                part_type == "body"
                and not unit.is_empty
                and unit.char_count_without_spaces <= 30
                and not self.parser.has_sentence_ending(unit.text)
            ):
                short_body_without_period_indices.append(unit.index)

            if part_type in {"heading1", "heading2", "heading3", "heading4"}:
                if (
                    last_heading_level == part_type
                    and not has_content_since_last_heading
                    and last_heading_index is not None
                ):
                    consecutive_same_level_pairs.append(
                        (last_heading_index, unit.index, part_type)
                    )
                last_heading_level = part_type
                last_heading_index = unit.index
                has_content_since_last_heading = False
            elif part_type and not unit.is_empty:
                has_content_since_last_heading = True

        condition1 = bool(short_body_without_period_indices)
        condition2 = not has_h1
        condition3 = not has_h2
        condition4 = bool(consecutive_same_level_pairs)
        condition5 = has_h4 and not has_h3

        reasons: List[str] = []
        if condition1:
            reasons.append("body_short_without_period")
        if condition2:
            reasons.append("missing_heading1")
        if condition3:
            reasons.append("missing_heading2")
        if condition4:
            reasons.append("consecutive_same_heading_level")
        if condition5:
            reasons.append("has_heading4_without_heading3")

        return {
            "need_ai": bool(reasons),
            "condition1": condition1,
            "condition2": condition2,
            "condition3": condition3,
            "condition4": condition4,
            "condition5": condition5,
            "reasons": reasons,
            "summary": {
                "has_h1": has_h1,
                "has_h2": has_h2,
                "has_h3": has_h3,
                "has_h4": has_h4,
                "short_body_without_period_count": len(short_body_without_period_indices),
                "consecutive_same_level_count": len(consecutive_same_level_pairs),
            },
            "details": {
                "short_body_without_period_indices": short_body_without_period_indices,
                "consecutive_same_level_pairs": consecutive_same_level_pairs,
            },
        }

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
            result["ai_suggestions"].append(
                {
                    "type": "missing_heading",
                    "level": "heading1",
                    "message": '文档中缺少一级标题（如"第X章"或"1"格式）',
                }
            )

        if not result["details"]["has_h2"]:
            result["is_valid"] = False
            result["missing_levels"].append("heading2")
            result["ai_suggestions"].append(
                {
                    "type": "missing_heading",
                    "level": "heading2",
                    "message": '文档中缺少二级标题（如"第X节"或"1.1"格式）',
                }
            )

        if result["details"]["has_h4"] and not result["details"]["has_h3"]:
            result["is_valid"] = False
            result["has_h4_without_h3"] = True
            result["ai_suggestions"].append(
                {
                    "type": "invalid_structure",
                    "message": '文档中存在四级标题（"1.1.1.1"格式）但没有三级标题',
                }
            )

        # 3. 如果AI可用，可以生成更多建议
        # 使用当前验证器实例的 AI 状态，避免读取全局状态导致串扰
        if is_ai_enabled(self.ai_identifier) and result["ai_suggestions"]:
            result["ai_suggestions"].append({
                "type": "ai_assistance_available",
                "message": "AI辅助识别已启用，可以帮助识别可能的标题",
            })

        return result


# 便捷函数

def create_fusion_identifier(
    api_key: Optional[str] = None,
    model: Optional[str] = None,
) -> FusionIdentifier:
    """创建融合识别器（便捷函数）"""
    ai_identifier = get_ai_identifier(api_key, model=model)
    return FusionIdentifier(ai_identifier=ai_identifier)


def validate_structure(
    heading_sequence: List[str],
    api_key: Optional[str] = None,
    model: Optional[str] = None,
) -> Dict:
    """验证文档结构（便捷函数）"""
    ai_identifier = get_ai_identifier(api_key, model=model)
    validator = DocumentStructureValidator(ai_identifier=ai_identifier)
    return validator.validate_document_structure(heading_sequence)
