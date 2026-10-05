# -*- coding: utf-8 -*-
"""字体工具:PDF 字体名 → DOCX 字体映射,处理 CJK 与符号字体。

PDF 字体名通常带有子集前缀(如 ``ABCDEF+SimSun``)和样式后缀。
本模块负责:
- 去除子集前缀
- 识别常见中文字体并映射到 eastAsia 字体
- 识别粗体/斜体/等宽标记
- 回退到合理默认字体
"""
from __future__ import annotations
import re

# 常见中文字体名 → 标准中文字体
CJK_FONT_MAP = {
    # 宋体类
    "simsun": "宋体", "songti": "宋体", "song": "宋体",
    "stsong": "华文宋体", "stsong-light": "华文宋体",
    "nsimsun": "新宋体",
    # 黑体类
    "simhei": "黑体", "heiti": "黑体", "hei": "黑体",
    "stheiti": "华文黑体", "stheiti-light": "华文黑体",
    "microsoft yahei": "微软雅黑", "msyh": "微软雅黑",
    # 楷体类
    "kaiti": "楷体", "stkaiti": "华文楷体", "kai": "楷体",
    "simkai": "楷体",
    # 仿宋类
    "fangsong": "仿宋", "stfangsong": "华文仿宋", "fs": "仿宋",
    "simfang": "仿宋",
    # 其他
    "stxihei": "华文细黑", "stzhongsong": "华文中宋",
    "stkaiti": "华文楷体", "stcaiyun": "华文彩云",
    "li": "隶书", "suli": "隶书",
    "youyuan": "幼圆",
}

# 西文字体名标准化
LATIN_FONT_MAP = {
    "timesnewroman": "Times New Roman",
    "times": "Times New Roman",
    "arial": "Arial",
    "helvetica": "Arial",
    "courier": "Courier New",
    "couriernew": "Courier New",
    "calibri": "Calibri",
    "cambria": "Cambria",
    "verdana": "Verdana",
    "tahoma": "Tahoma",
    "georgia": "Georgia",
    "garamond": "Garamond",
}

# CJK Unicode 范围(用于判断字符是否为中日韩)
CJK_RANGES = [
    (0x4E00, 0x9FFF),    # CJK 统一汉字
    (0x3400, 0x4DBF),    # CJK 扩展 A
    (0x20000, 0x2A6DF),  # CJK 扩展 B
    (0x2A700, 0x2B73F),  # CJK 扩展 C
    (0x2B740, 0x2B81F),  # CJK 扩展 D
    (0x3000, 0x303F),    # CJK 标点
    (0xFF00, 0xFFEF),    # 全角字符
    (0x3040, 0x309F),    # 平假名
    (0x30A0, 0x30FF),    # 片假名
    (0xAC00, 0xD7AF),    # 韩文音节
]


def has_cjk(text: str) -> bool:
    """判断文本中是否包含 CJK 字符。"""
    if not text:
        return False
    for ch in text:
        cp = ord(ch)
        for lo, hi in CJK_RANGES:
            if lo <= cp <= hi:
                return True
    return False


def clean_font_name(name: str) -> str:
    """去除字体子集前缀(如 ``ABCDEF+SimSun`` → ``SimSun``)。"""
    if not name:
        return name
    # 子集前缀:6个大写字母+加号
    name = re.sub(r"^[A-Z]{6}\+", "", name)
    return name


def normalize_font(name: str) -> str:
    """规范化字体名:去前缀、去多余空格、统一大小写。"""
    if not name:
        return "Arial"
    name = clean_font_name(name)
    # 去除样式后缀(如 ",Bold"、"-BoldMT")
    name = re.split(r"[,\-]", name)[0].strip()
    key = name.lower().replace(" ", "")
    # 中文
    if key in CJK_FONT_MAP:
        return CJK_FONT_MAP[key]
    # 西文
    if key in LATIN_FONT_MAP:
        return LATIN_FONT_MAP[key]
    return name


def get_font_roles(name: str) -> tuple[str, str]:
    """返回 (latin_font, eastasia_font)。

    - 中文字体 → latin 用相同或 Arial,eastasia 用中文名
    - 西文字体 → latin 用原名,eastasia 用宋体(兜底)
    """
    if not name:
        return "Arial", "宋体"
    name = clean_font_name(name)
    key = name.lower().replace(" ", "")
    if key in CJK_FONT_MAP:
        cjk = CJK_FONT_MAP[key]
        return "Times New Roman", cjk
    latin = normalize_font(name)
    return latin, "宋体"


def font_flags_to_style(flags: int) -> dict:
    """把 PyMuPDF 的 font flags 转成样式字典。

    PyMuPDF flags 位定义:
    - bit 0 (1): superscripted
    - bit 1 (2): italic
    - bit 2 (4): serifed
    - bit 3 (8): monospaced
    - bit 4 (16): bold
    """
    return {
        "bold": bool(flags & 16),
        "italic": bool(flags & 2),
        "monospace": bool(flags & 8),
        "serif": bool(flags & 4),
    }
