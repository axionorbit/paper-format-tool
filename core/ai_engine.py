# -*- coding: utf-8 -*-
"""
AI引擎模块 - 使用AI辅助识别段落类型
"""

import json
import re
import time
from typing import Dict, List, Optional, Tuple
from urllib import error, request

from utils.config import PARTS
from utils.logger import default_logger


class AIIdentifier:
    """AI识别器 - 使用AI辅助判断段落类型"""

    SUPPORTED_LABELS = {part_key for _, part_key in PARTS}

    TITLE_LABELS = {
        "heading1",
        "heading2",
        "heading3",
        "heading4",
        "abstract_title",
        "ref_title",
        "ack_title",
        "appendix_title",
        "figure_caption",
        "table_caption",
        "table_note",
    }

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: str = "GLM-4.7-Flash",
        timeout: int = 90,
        max_retries: int = 2,
        batch_size: int = 4,
    ):
        """
        初始化AI识别器
        api_key: 智谱AI API密钥
        """
        self.api_key = (api_key or "").strip()
        self.model = model
        self.timeout = timeout
        self.max_retries = max(0, int(max_retries))
        self.batch_size = max(1, int(batch_size))
        self.base_url = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
        self.enabled = bool(self.api_key)  # 有API密钥时才启用
        self.last_error: str = ""
        self.last_raw_content: str = ""
        self.last_response_preview: str = ""

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

        candidate = [{"id": 0, "text": (text or "").strip()}]
        label_map = self.identify_candidates(candidate, context=context)
        label = label_map.get(0)
        if not label:
            return None, False

        return label, label in self.TITLE_LABELS

    def identify_candidates(
        self,
        candidates: List[dict],
        context: Optional[dict] = None,
    ) -> Dict[int, str]:
        """
        批量识别候选段落，返回 {candidate_id: label}
        """
        if not self.enabled or not candidates:
            return {}

        cleaned_candidates = []
        for item in candidates:
            cid = item.get("id")
            text = (item.get("text") or "").strip()
            if cid is None or not text:
                continue
            candidate = {"id": cid, "text": text}
            for key in ("style_name", "numbering_level", "numbering_prefix"):
                if key in item and item.get(key) is not None:
                    candidate[key] = item.get(key)
            cleaned_candidates.append(candidate)

        if not cleaned_candidates:
            return {}

        result_map: Dict[int, str] = {}
        self.last_error = ""
        self.last_raw_content = ""

        expected_ids = [int(item["id"]) for item in cleaned_candidates]
        context_keys = sorted((context or {}).keys())
        total_chars = sum(len((item.get("text") or "")) for item in cleaned_candidates)
        default_logger.info(
            "AI request start. model=%s, candidates=%s, total_chars=%s, context_keys=%s",
            self.model,
            len(cleaned_candidates),
            total_chars,
            context_keys,
        )

        messages = self._build_messages(cleaned_candidates, context=context)
        content = self._call_chat_api(
            messages,
            request_meta={
                "candidate_count": len(cleaned_candidates),
                "total_chars": total_chars,
            },
        )
        self.last_raw_content = content or self.last_raw_content
        parsed = self._parse_ai_json(content)

        if parsed:
            raw_rows = self._extract_result_rows(parsed)
            default_logger.info(
                "AI parse success. rows=%s, candidate_count=%s",
                len(raw_rows),
                len(cleaned_candidates),
            )
            for row in raw_rows:
                cid = self._extract_row_id(row)
                if cid is None:
                    continue
                label = self._extract_row_label(row)
                if label in self.SUPPORTED_LABELS:
                    result_map[cid] = label
        else:
            default_logger.warning(
                "AI parse failed. last_error=%s, response_preview=%s",
                self.last_error or "none",
                (self.last_response_preview or "")[:200],
            )

        missing_ids = [cid for cid in expected_ids if cid not in result_map]
        if missing_ids:
            default_logger.warning(
                "AI label coverage incomplete. labeled=%s/%s, missing_ids_sample=%s",
                len(result_map),
                len(expected_ids),
                missing_ids[:10],
            )
        else:
            default_logger.info(
                "AI label coverage complete. labeled=%s/%s",
                len(result_map),
                len(expected_ids),
            )

        if not result_map:
            default_logger.warning(
                "AI returned no usable labels. candidates=%s, last_error=%s, response_preview=%s",
                len(cleaned_candidates),
                self.last_error or "none",
                (self.last_response_preview or "")[:200],
            )

        return result_map

    def _build_messages(
        self,
        candidates: List[dict],
        context: Optional[dict] = None,
    ) -> List[dict]:
        allowed = sorted(self.SUPPORTED_LABELS)
        payload = {
            "allowed_labels": allowed,
            "candidates": candidates,
        }
        if context:
            payload["context"] = context

        _legacy_system_prompt = (
            "你是论文结构识别助手。"
            "仅输出严格JSON，不要输出解释、前后缀或Markdown。"
            "JSON格式必须是: "
            '{"results":[{"id":1,"label":"heading2"}]}。'
            "label 只能从 allowed_labels 中选择。"
            "当 text 含有类似 1.1 / 1.1.1 / 第1章 的前缀时，应按标题级别判断。"
        )

        user_prompt = (
            "请识别每条候选段落的结构标签。\n"
            "输入数据如下：\n"
            f"{json.dumps(payload, ensure_ascii=False)}"
        )

        # Override with the new global-context prompt.
        system_prompt = (
            "你是论文结构标注助手。你将收到按文中出现顺序排列的候选段落列表（candidates）。"
            "你的任务是基于全部候选的全局上下文给每个候选标注结构标签。"
            "要求："
            "1. 不要先假设固定编号体系；允许任意学校/任意风格的标题写法。"
            "2. 必须结合全体候选的前后关系进行一致判断，不能逐条孤立判断。"
            "3. 优先识别图表相关：图题->figure_caption；表题->table_caption；注释/来源类->table_note。"
            "4. 对标题层级（heading1~heading4）做相对判断："
            "同一风格的标记应保持同一层级；"
            "更细分/更深的标记应对应更低层级（例如更长的编号链、明显子项格式）。"
            "5. 若某候选不应视为标题或图表注释，可标为 body。"
            "6. label 只能从 allowed_labels 中选择。"
            "7. 必须覆盖所有 candidate 的 id，每个 id 只能出现一次，不得漏标。"
            "输出要求（严格）："
            "只输出 JSON，不要输出解释、不要 markdown。"
            "格式必须为："
            "{\"results\":[{\"id\":1,\"label\":\"heading2\"}]}"
        )

        _legacy_user_prompt = (
            "请识别每条候选段落的结构标签。\n"
            "输入数据如下：\n"
            f"{json.dumps(payload, ensure_ascii=False)}"
        )

        user_prompt = (
            "请识别每条候选段落的结构标签。\n"
            "输入数据如下：\n"
            f"{json.dumps(payload, ensure_ascii=False)}"
        )

        # Final prompt (effective): includes all thesis labels.
        system_prompt = (
            "你是论文结构标注助手。你将收到按文中出现顺序排列的候选段落。"
            "请基于全局上下文为每条候选打标签。"
            "不要假设固定编号体系。"
            "同一风格标记保持同层级；更细分/更深标记对应更低层级。"
            "图题用 figure_caption，表题用 table_caption，注释/来源用 table_note。"
            "若不是标题/图表注释可标为 body。"
            "你还可以使用摘要、参考文献、致谢、附录、公式、表格内容等标签。"
            "label 必须从 allowed_labels 中选择，必须覆盖所有 id 且不重复。"
            "只输出 JSON，不要解释，不要 markdown。"
            "格式：{\"results\":[{\"id\":1,\"label\":\"heading2\"}]}"
        )
        user_prompt = (
            "请识别每条候选段落的结构标签。\n"
            "输入数据如下：\n"
            f"{json.dumps(payload, ensure_ascii=False)}"
        )

        return [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]

    def _call_chat_api(self, messages: List[dict], request_meta: Optional[Dict[str, int]] = None) -> str:
        body = {
            "model": self.model,
            "messages": messages,
            "temperature": 0.0,
            "stream": False,
        }
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.last_error = ""
        self.last_response_preview = ""
        payload_bytes = len(data)
        candidate_count = (request_meta or {}).get("candidate_count", -1)
        total_chars = (request_meta or {}).get("total_chars", -1)

        default_logger.info(
            "AI HTTP request prepared. model=%s, candidates=%s, total_chars=%s, payload_bytes=%s, timeout=%ss, max_retries=%s",
            self.model,
            candidate_count,
            total_chars,
            payload_bytes,
            self.timeout,
            self.max_retries,
        )

        for attempt in range(self.max_retries + 1):
            attempt_start = time.perf_counter()
            req = request.Request(
                self.base_url,
                method="POST",
                data=data,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self.api_key}",
                },
            )
            try:
                with request.urlopen(req, timeout=self.timeout) as resp:
                    raw = resp.read().decode("utf-8", errors="replace")
                    self.last_response_preview = raw[:500]
                elapsed = time.perf_counter() - attempt_start
                default_logger.info(
                    "AI HTTP success. attempt=%s/%s, elapsed=%.2fs, response_bytes=%s",
                    attempt + 1,
                    self.max_retries + 1,
                    elapsed,
                    len(raw.encode("utf-8", errors="ignore")),
                )
            except error.HTTPError as exc:
                body_preview = ""
                try:
                    body_preview = exc.read().decode("utf-8", errors="replace")[:500]
                except Exception:
                    body_preview = ""
                self.last_error = f"http_error:{exc.code}"
                elapsed = time.perf_counter() - attempt_start
                default_logger.error(
                    "AI HTTPError. status=%s, attempt=%s/%s, elapsed=%.2fs, body=%s",
                    exc.code,
                    attempt + 1,
                    self.max_retries + 1,
                    elapsed,
                    body_preview,
                )
                if attempt < self.max_retries:
                    wait_seconds = 1.5 ** attempt
                    default_logger.warning(
                        "AI retry scheduled after HTTPError. wait=%.2fs, next_attempt=%s/%s",
                        wait_seconds,
                        attempt + 2,
                        self.max_retries + 1,
                    )
                    time.sleep(wait_seconds)
                    continue
                return ""
            except error.URLError as exc:
                self.last_error = f"url_error:{exc}"
                elapsed = time.perf_counter() - attempt_start
                default_logger.error(
                    "AI URLError. attempt=%s/%s, elapsed=%.2fs, detail=%s",
                    attempt + 1,
                    self.max_retries + 1,
                    elapsed,
                    exc,
                )
                if attempt < self.max_retries:
                    wait_seconds = 1.5 ** attempt
                    default_logger.warning(
                        "AI retry scheduled after URLError. wait=%.2fs, next_attempt=%s/%s",
                        wait_seconds,
                        attempt + 2,
                        self.max_retries + 1,
                    )
                    time.sleep(wait_seconds)
                    continue
                return ""
            except Exception as exc:
                self.last_error = f"request_exception:{type(exc).__name__}:{exc}"
                elapsed = time.perf_counter() - attempt_start
                default_logger.error(
                    "AI request exception. attempt=%s/%s, elapsed=%.2fs, detail=%s",
                    attempt + 1,
                    self.max_retries + 1,
                    elapsed,
                    exc,
                )
                if attempt < self.max_retries:
                    wait_seconds = 1.5 ** attempt
                    default_logger.warning(
                        "AI retry scheduled after request exception. wait=%.2fs, next_attempt=%s/%s",
                        wait_seconds,
                        attempt + 2,
                        self.max_retries + 1,
                    )
                    time.sleep(wait_seconds)
                    continue
                return ""

            try:
                payload = json.loads(raw)
            except json.JSONDecodeError as exc:
                self.last_error = f"response_json_decode_error:{exc}"
                default_logger.error("AI response JSON decode error: %s", exc)
                return ""

            choices = payload.get("choices") or []
            if not choices:
                self.last_error = "no_choices_in_response"
                default_logger.error("AI response has no choices. payload=%s", str(payload)[:500])
                return ""
            message = choices[0].get("message") or {}
            return str(message.get("content") or "").strip()

        return ""

    def _parse_ai_json(self, content: str) -> Optional[dict]:
        if not content:
            self.last_error = self.last_error or "empty_content"
            return None

        json_text = self._extract_json_text(content)
        if not json_text:
            self.last_error = self.last_error or "json_block_not_found"
            return None

        try:
            parsed = json.loads(json_text)
        except json.JSONDecodeError:
            self.last_error = self.last_error or "json_decode_failed"
            return None

        if isinstance(parsed, list):
            return {"results": parsed}
        if not isinstance(parsed, dict):
            self.last_error = self.last_error or "parsed_not_dict"
            return None

        for key in ("results", "data", "items", "predictions", "labels"):
            value = parsed.get(key)
            if isinstance(value, list):
                return {"results": value}

        self.last_error = self.last_error or "results_list_missing"
        return None

    def _extract_json_text(self, content: str) -> str:
        fence_match = re.search(r"```(?:json)?\s*(\{[\s\S]*\})\s*```", content, flags=re.I)
        if fence_match:
            return fence_match.group(1).strip()
        fence_array_match = re.search(r"```(?:json)?\s*(\[[\s\S]*\])\s*```", content, flags=re.I)
        if fence_array_match:
            return fence_array_match.group(1).strip()

        first = content.find("{")
        last = content.rfind("}")
        if first != -1 and last != -1 and last > first:
            return content[first:last + 1].strip()

        first_arr = content.find("[")
        last_arr = content.rfind("]")
        if first_arr != -1 and last_arr != -1 and last_arr > first_arr:
            return content[first_arr:last_arr + 1].strip()

        return ""

    def _extract_result_rows(self, parsed: dict) -> List[dict]:
        rows = parsed.get("results", [])
        if not isinstance(rows, list):
            return []
        return [row for row in rows if isinstance(row, dict)]

    def _extract_row_id(self, row: dict) -> Optional[int]:
        for key in ("id", "index", "paragraph_id", "paragraphIndex"):
            if key in row:
                try:
                    return int(row.get(key))
                except (TypeError, ValueError):
                    return None
        return None

    def _extract_row_label(self, row: dict) -> Optional[str]:
        raw_label = None
        for key in ("label", "type", "category", "tag"):
            if key in row:
                raw_label = row.get(key)
                break
        if raw_label is None:
            return None
        return self._normalize_label(raw_label)

    def _normalize_label_legacy(self, label: Optional[str]) -> Optional[str]:
        if label is None:
            return None
        label = str(label).strip().lower()
        alias_map = {
            "h1": "heading1",
            "h2": "heading2",
            "h3": "heading3",
            "h4": "heading4",
            "heading_1": "heading1",
            "heading_2": "heading2",
            "heading_3": "heading3",
            "heading_4": "heading4",
            "一级标题": "heading1",
            "二级标题": "heading2",
            "三级标题": "heading3",
            "四级标题": "heading4",
            "figure": "figure_caption",
            "figurecaption": "figure_caption",
            "图片标题": "figure_caption",
            "table": "table_caption",
            "tablecaption": "table_caption",
            "表格标题": "table_caption",
            "tablenote": "table_note",
            "图表注释": "table_note",
            "正文": "body",
        }
        return alias_map.get(label, label)

    def _normalize_label(self, label: Optional[str]) -> Optional[str]:
        if label is None:
            return None

        label = str(label).strip().lower()
        alias_map = {
            # Headings
            "h1": "heading1",
            "h2": "heading2",
            "h3": "heading3",
            "h4": "heading4",
            "heading_1": "heading1",
            "heading_2": "heading2",
            "heading_3": "heading3",
            "heading_4": "heading4",
            "一级标题": "heading1",
            "二级标题": "heading2",
            "三级标题": "heading3",
            "四级标题": "heading4",

            # Body / abstract
            "body": "body",
            "正文": "body",
            "abstract": "abstract_title",
            "abstracttitle": "abstract_title",
            "abstractcontent": "abstract_content",
            "摘要标题": "abstract_title",
            "摘要内容": "abstract_content",

            # References
            "reference": "ref_title",
            "references": "ref_title",
            "reftitle": "ref_title",
            "refcontent": "ref_content",
            "参考文献标题": "ref_title",
            "参考文献内容": "ref_content",

            # Acknowledgement
            "acknowledgement": "ack_title",
            "acknowledgements": "ack_title",
            "acknowledgment": "ack_title",
            "acknowledgments": "ack_title",
            "acktitle": "ack_title",
            "ackcontent": "ack_content",
            "致谢标题": "ack_title",
            "致谢内容": "ack_content",

            # Appendix
            "appendix": "appendix_title",
            "appendixtitle": "appendix_title",
            "appendixcontent": "appendix_content",
            "附录标题": "appendix_title",
            "附录内容": "appendix_content",

            # Figures / tables
            "figure": "figure_caption",
            "figurecaption": "figure_caption",
            "图片标题": "figure_caption",
            "table": "table_caption",
            "tablecaption": "table_caption",
            "表格标题": "table_caption",
            "tablenote": "table_note",
            "table_note": "table_note",
            "图表注释": "table_note",
            "tablecontent": "table_content",
            "table_content": "table_content",
            "表格内容": "table_content",

            # Formula
            "formula": "formula",
            "公式": "formula",
        }
        return alias_map.get(label, label)

    def get_last_status(self) -> Dict[str, str]:
        return {
            "last_error": self.last_error or "",
            "last_response_preview": self.last_response_preview or "",
        }

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
    if api_key is not None:
        _global_ai_identifier = AIIdentifier(api_key)
    elif _global_ai_identifier is None:
        _global_ai_identifier = AIIdentifier(api_key)
    return _global_ai_identifier


def set_ai_api_key(api_key: str):
    """设置AI API密钥"""
    global _global_ai_identifier
    _global_ai_identifier = AIIdentifier(api_key)


def is_ai_enabled(ai_identifier: Optional[AIIdentifier] = None) -> bool:
    """检查AI识别是否已启用"""
    if ai_identifier is not None:
        return ai_identifier.enabled
    return _global_ai_identifier is not None and _global_ai_identifier.enabled
