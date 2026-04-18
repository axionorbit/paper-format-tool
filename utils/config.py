# -*- coding: utf-8 -*-
"""
配置数据模块
"""

from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==================== 字体选项 ====================

# 中文字体选项
CHINESE_FONTS = [
    "宋体", "黑体", "楷体", "仿宋", "微软雅黑", "华文宋体", "华文黑体",
    "华文楷体", "华文仿宋", "方正宋体", "方正黑体"
]

# 英文字体选项
ENGLISH_FONTS = [
    "Times New Roman", "Arial", "Calibri", "Georgia",
    "Palatino Linotype", "Book Antiqua"
]

# 字号选项（名称 -> 磅值）
FONT_SIZES = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "七号": 5.5, "八号": 5
}
FONT_SIZE_NAMES = list(FONT_SIZES.keys())

# ==================== 段落格式选项 ====================

# 对齐方式
ALIGNMENTS = {
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY
}
ALIGNMENT_NAMES = list(ALIGNMENTS.keys())

# 行距选项（倍数和固定值）
LINE_SPACINGS = {
    "单倍行距": 1.0,
    "1.15倍行距": 1.15,
    "1.25倍行距": 1.25,
    "1.5倍行距": 1.5,
    "1.75倍行距": 1.75,
    "2倍行距": 2.0,
    "2.5倍行距": 2.5,
    "3倍行距": 3.0,
    "6磅": 6,
    "12磅": 12,
    "15磅": 15,
    "18磅": 18,
    "20磅": 20,
    "21磅": 21,
    "22磅": 22,
    "24磅": 24,
    "28磅": 28,
    "30磅": 30
}
LINE_SPACING_NAMES = list(LINE_SPACINGS.keys())

# 段前段后间距选项（磅）
SPACINGS = ["0磅", "6磅", "12磅", "18磅", "24磅", "30磅", "36磅", "48磅"]

# 首行缩进选项
FIRST_LINE_INDENTS = ["无", "1字符", "2字符", "3字符", "4字符"]

# 整体缩进选项
OVERALL_INDENTS = [
    "无",
    "左缩进1字符", "左缩进2字符", "左缩进3字符", "左缩进4字符",
    "右缩进1字符", "右缩进2字符",
    "左右各1字符", "左右各2字符"
]

# 缩进映射表
FIRST_LINE_INDENT_MAP = {
    "无": 0,
    "1字符": 1,
    "2字符": 2,
    "3字符": 3,
    "4字符": 4,
}

OVERALL_INDENT_MAP = {
    "无": (0, 0),
    "左缩进1字符": (1, 0),
    "左缩进2字符": (2, 0),
    "左缩进3字符": (3, 0),
    "左缩进4字符": (4, 0),
    "右缩进1字符": (0, 1),
    "右缩进2字符": (0, 2),
    "左右各1字符": (1, 1),
    "左右各2字符": (2, 2),
}

# ==================== 论文部分定义 ====================

DEFAULT_OUTPUT_DIR_HINT = "与原文件相同目录"
DEFAULT_FILENAME_HINT = "留空则使用默认名称"

# ==================== AI 模型配置 ====================

DEFAULT_AI_MODEL = "GLM-4.7-Flash"
AI_MODEL_OPTIONS = [
    "GLM-4.7-Flash",
    "GLM-5.1",
    "GLM-5",
    "GLM-5-Turbo",
    "GLM-4.7",
    "GLM-4.7-FlashX",
    "GLM-4.6",
    "GLM-4.5-Air",
    "GLM-4.5-AirX",
    "GLM-4-Long",
    "GLM-4-FlashX-250414",
]

# 论文部分列表
PARTS = [
    ("摘要标题", "abstract_title"),
    ("摘要内容", "abstract_content"),
    ("一级标题", "heading1"),
    ("二级标题", "heading2"),
    ("三级标题", "heading3"),
    ("四级标题", "heading4"),
    ("正文内容", "body"),
    ("图片标题", "figure_caption"),
    ("表格标题", "table_caption"),
    ("表格内容", "table_content"),
    ("图表注释", "table_note"),
    ("公式", "formula"),
    ("参考文献标题", "ref_title"),
    ("参考文献内容", "ref_content"),
    ("致谢标题", "ack_title"),
    ("致谢内容", "ack_content"),
    ("附录标题", "appendix_title"),
    ("附录内容", "appendix_content"),
]

PART_STYLE_NAMES = {part_key: part_name for part_name, part_key in PARTS}

# 大纲级别映射
OUTLINE_LEVELS = {
    "abstract_title": 0,
    "heading1": 0,
    "heading2": 1,
    "heading3": 2,
    "heading4": 3,
    "ref_title": 0,
    "appendix_title": 0,
}

# 表格列定义
TABLE_COLUMNS = [
    ("part_name", "", 150),
    ("chinese_font", "中文字体", 130),
    ("english_font", "英文字体", 150),
    ("font_size", "字号", 88),
    ("alignment", "对齐", 104),
    ("line_spacing", "行距", 118),
    ("space_before", "段前", 88),
    ("space_after", "段后", 88),
    ("first_line_indent", "首行缩进", 104),
    ("overall_indent", "整体缩进", 124),
    ("bold", "加粗", 78),
    ("italic", "倾斜", 78),
]

# ==================== 默认格式设置 ====================

DEFAULT_FORMATS = {
    "abstract_title": {
        "chinese_font": "黑体",
        "english_font": "Times New Roman",
        "font_size": "三号",
        "alignment": "居中",
        "line_spacing": "1.5倍行距",
        "space_before": "0磅",
        "space_after": "12磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": True,
        "italic": False
    },
    "abstract_content": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "小四",
        "alignment": "两端对齐",
        "line_spacing": "1.5倍行距",
        "space_before": "0磅",
        "space_after": "0磅",
        "first_line_indent": "2字符",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "heading1": {
        "chinese_font": "黑体",
        "english_font": "Times New Roman",
        "font_size": "三号",
        "alignment": "居中",
        "line_spacing": "1.5倍行距",
        "space_before": "24磅",
        "space_after": "12磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": True,
        "italic": False
    },
    "heading2": {
        "chinese_font": "黑体",
        "english_font": "Times New Roman",
        "font_size": "四号",
        "alignment": "左对齐",
        "line_spacing": "1.5倍行距",
        "space_before": "18磅",
        "space_after": "6磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": True,
        "italic": False
    },
    "heading3": {
        "chinese_font": "黑体",
        "english_font": "Times New Roman",
        "font_size": "小四",
        "alignment": "左对齐",
        "line_spacing": "1.5倍行距",
        "space_before": "12磅",
        "space_after": "6磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": True,
        "italic": False
    },
    "heading4": {
        "chinese_font": "黑体",
        "english_font": "Times New Roman",
        "font_size": "小四",
        "alignment": "左对齐",
        "line_spacing": "1.5倍行距",
        "space_before": "6磅",
        "space_after": "6磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": True,
        "italic": False
    },
    "body": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "小四",
        "alignment": "两端对齐",
        "line_spacing": "1.5倍行距",
        "space_before": "0磅",
        "space_after": "0磅",
        "first_line_indent": "2字符",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "figure_caption": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "五号",
        "alignment": "居中",
        "line_spacing": "单倍行距",
        "space_before": "6磅",
        "space_after": "6磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "table_caption": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "五号",
        "alignment": "居中",
        "line_spacing": "单倍行距",
        "space_before": "6磅",
        "space_after": "6磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "table_content": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "五号",
        "alignment": "居中",
        "line_spacing": "单倍行距",
        "space_before": "0磅",
        "space_after": "0磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "table_note": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "小五",
        "alignment": "左对齐",
        "line_spacing": "单倍行距",
        "space_before": "0磅",
        "space_after": "0磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "formula": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "小四",
        "alignment": "居中",
        "line_spacing": "1.5倍行距",
        "space_before": "6磅",
        "space_after": "6磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "ref_title": {
        "chinese_font": "黑体",
        "english_font": "Times New Roman",
        "font_size": "三号",
        "alignment": "居中",
        "line_spacing": "1.5倍行距",
        "space_before": "24磅",
        "space_after": "12磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": True,
        "italic": False
    },
    "ref_content": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "五号",
        "alignment": "两端对齐",
        "line_spacing": "单倍行距",
        "space_before": "0磅",
        "space_after": "0磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "ack_title": {
        "chinese_font": "黑体",
        "english_font": "Times New Roman",
        "font_size": "三号",
        "alignment": "居中",
        "line_spacing": "1.5倍行距",
        "space_before": "24磅",
        "space_after": "12磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": True,
        "italic": False
    },
    "ack_content": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "小四",
        "alignment": "两端对齐",
        "line_spacing": "1.5倍行距",
        "space_before": "0磅",
        "space_after": "0磅",
        "first_line_indent": "2字符",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
    "appendix_title": {
        "chinese_font": "黑体",
        "english_font": "Times New Roman",
        "font_size": "三号",
        "alignment": "居中",
        "line_spacing": "1.5倍行距",
        "space_before": "24磅",
        "space_after": "12磅",
        "first_line_indent": "无",
        "overall_indent": "无",
        "bold": True,
        "italic": False
    },
    "appendix_content": {
        "chinese_font": "宋体",
        "english_font": "Times New Roman",
        "font_size": "小四",
        "alignment": "两端对齐",
        "line_spacing": "1.5倍行距",
        "space_before": "0磅",
        "space_after": "0磅",
        "first_line_indent": "2字符",
        "overall_indent": "无",
        "bold": False,
        "italic": False
    },
}
