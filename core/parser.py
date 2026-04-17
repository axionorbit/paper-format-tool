# -*- coding: utf-8 -*-
"""
文档解析与段落预处理
"""

from dataclasses import dataclass
import re
from typing import Dict, List, Optional

from docx.oxml.ns import qn


@dataclass
class ParagraphUnit:
    """预处理后的段落单元"""

    index: int
    paragraph: object
    text: str
    ai_text: str
    normalized_text: str
    char_count_without_spaces: int
    is_empty: bool
    style_name: str
    numbering_level: Optional[int]
    numbering_prefix: str


class DocumentParser:
    """将 Word 文档转换为段落单元列表"""

    @staticmethod
    def normalize_text(text: str) -> str:
        return (text or "").replace(" ", "").replace("\u3000", "").strip()

    @staticmethod
    def has_sentence_ending(text: str) -> bool:
        if not text:
            return False
        return text.strip().endswith(("。", ".", "！", "？", "!", "?"))

    @staticmethod
    def _iter_style_chain(paragraph):
        style = getattr(paragraph, "style", None)
        visited = set()
        while style is not None and id(style) not in visited:
            yield style
            visited.add(id(style))
            style = getattr(style, "base_style", None)

    @staticmethod
    def _find_numpr(paragraph):
        try:
            ppr = paragraph._element.pPr
            if ppr is not None and ppr.numPr is not None:
                return ppr.numPr
        except Exception:
            pass

        for style in DocumentParser._iter_style_chain(paragraph):
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
    def _extract_level_from_style_name(style_name: str) -> Optional[int]:
        if not style_name:
            return None
        normalized = re.sub(r"\s+", "", str(style_name)).lower()
        match = re.match(r"^(?:heading|标题|listnumber|列表编号|编号)([1-9])$", normalized)
        if not match:
            return None
        try:
            return int(match.group(1))
        except Exception:
            return None

    def _extract_numbering_metadata(self, paragraph) -> Dict[str, Optional[object]]:
        style_name = ""
        try:
            style_name = getattr(paragraph.style, "name", "") or ""
        except Exception:
            style_name = ""

        metadata: Dict[str, Optional[object]] = {
            "style_name": style_name,
            "numbering_level": None,
            "numbering_prefix": "",
        }
        style_level = self._extract_level_from_style_name(style_name)

        numpr = self._find_numpr(paragraph)
        if numpr is None:
            return metadata

        num_id = self._safe_numpr_val(numpr, "numId")
        ilvl = self._safe_numpr_val(numpr, "ilvl")

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

        level = None
        level_token_count = 0
        if lvl_text:
            level_tokens = re.findall(r"%\d+", lvl_text)
            level_token_count = len(level_tokens)
            if level_tokens:
                level = level_token_count
        if level is None and ilvl_int is not None:
            level = ilvl_int + 1
        if (level is None or (level == 1 and ilvl_int is None)) and style_level:
            level = style_level
        if level is not None and level > 0:
            metadata["numbering_level"] = level

        prefix = ""
        if lvl_text:
            prefix = re.sub(r"%\d+", "1", lvl_text).strip()
            # 清理结尾分隔符，避免 "1.1." 这类前缀影响句末判断
            prefix = re.sub(r"[\.\u3002\uFF0E、\s]+$", "", prefix)
        if level is not None and 1 <= level <= 9 and (not prefix or level_token_count < level):
            prefix = ".".join(["1"] * level)
        metadata["numbering_prefix"] = prefix

        return metadata

    def parse_document(self, document) -> List[ParagraphUnit]:
        units: List[ParagraphUnit] = []
        for idx, paragraph in enumerate(document.paragraphs):
            raw_text = paragraph.text or ""
            text = raw_text.strip()
            numbering_metadata = self._extract_numbering_metadata(paragraph)
            numbering_prefix = str(numbering_metadata.get("numbering_prefix") or "").strip()
            ai_text = text
            if numbering_prefix and text and not text.startswith(numbering_prefix):
                ai_text = f"{numbering_prefix} {text}"

            normalized = self.normalize_text(text)
            units.append(
                ParagraphUnit(
                    index=idx,
                    paragraph=paragraph,
                    text=text,
                    ai_text=ai_text,
                    normalized_text=normalized,
                    char_count_without_spaces=len(normalized),
                    is_empty=not bool(text),
                    style_name=str(numbering_metadata.get("style_name") or ""),
                    numbering_level=numbering_metadata.get("numbering_level"),
                    numbering_prefix=numbering_prefix,
                )
            )
        return units

    def extract_ai_candidates(
        self,
        units: List[ParagraphUnit],
        max_chars: int = 30,
    ) -> List[ParagraphUnit]:
        """
        提取全文中可能是标题的候选段落:
        - 字数 <= 30（去掉空格）
        - 不以句号/问号/感叹号等句末标点结束
        """
        candidates: List[ParagraphUnit] = []
        for unit in units:
            if unit.is_empty:
                continue

            # 自动编号标题候选优先保留（通用支持，不依赖具体模板）
            if unit.numbering_level is not None and 1 <= int(unit.numbering_level) <= 4:
                ai_chars = len(self.normalize_text(unit.ai_text))
                if ai_chars <= max_chars * 2:
                    candidates.append(unit)
                    continue

            ai_chars = len(self.normalize_text(unit.ai_text))
            if ai_chars > max_chars:
                continue
            if self.has_sentence_ending(unit.ai_text):
                continue
            candidates.append(unit)
        return candidates
