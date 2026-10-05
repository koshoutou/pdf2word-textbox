# -*- coding: utf-8 -*-
"""字体工具:PDF 字体名 → DOCX 字体映射,处理 CJK 与符号字体。

PDF 字体名通常带有子集前缀(如 ``ABCDEF+SimSun``)和样式后缀。
本模块负责:
- 去除子集前缀
- 识别常见中文字体并映射到 eastAsia 字体
- 识别粗体/斜体/等宽标记
- 字号整数化(消除 PDF 浮点误差)
- 系统字体可用性检测(用于决定是否回退)
- 回退到合理默认字体

字体安装提示:本工具本身不含字体文件。为获得最佳渲染效果,建议安装
Windows 常用中文字体(仿宋/黑体/宋体/楷体/华文系列)。详见 README
"字体说明"章节。
"""
from __future__ import annotations
import re
import os
import subprocess
from functools import lru_cache

# 常见中文字体名 → 标准中文字体(docx eastAsia 字体名)
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
    # 方正系列(常见于公文)
    "fzfangsong": "方正仿宋_GBK", "fzkai": "方正楷体_GBK",
    "fzhei": "方正黑体_GBK", "fzxiaobiaosong": "方正小标宋_GBK",
    "fzdabiaosong": "方正大标宋简体",
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


def normalize_size(size: float, tolerance: float = 0.1) -> float:
    """字号整数化:消除 PDF 浮点误差。

    PDF 中字号常以矩阵缩放计算,可能产生 13.99 / 14.01 这类本应是 14.0 的值。
    本函数将接近整数的字号规整为整数,其余保留 1 位小数。

    #5.11 修复:用数学舍入(加 0.5 取 floor)替代 Python round 的银行家舍入,
    避免 round(13.5)=14、round(14.5)=14 的不直观行为。

    Args:
        size: 原始字号(pt)
        tolerance: 整数容差(默认 0.1pt)

    Returns:
        规整后的字号
    """
    if size is None:
        return 12.0
    # 数学舍入:加 0.5 取 floor(正数)
    import math
    rounded = math.floor(size + 0.5)
    if abs(size - rounded) <= tolerance:
        return float(rounded)
    # 非整数:保留 1 位小数(同样用数学舍入)
    return math.floor(size * 10 + 0.5) / 10.0


@lru_cache(maxsize=1)
def list_system_fonts() -> set[str]:
    """列出系统已安装的字体名(用 fc-list,带缓存)。

    Returns:
        字体名集合(小写)。若 fc-list 不可用则返回空集(不做可用性检测)。
    """
    names: set[str] = set()
    try:
        out = subprocess.run(
            ["fc-list", ":", "family"],
            capture_output=True, text=True, timeout=5,
        )
        for line in out.stdout.splitlines():
            fam = line.strip()
            if fam:
                names.add(fam.lower())
    except Exception:
        pass
    return names


@lru_cache(maxsize=128)
def is_font_available(font_name: str) -> bool:
    """检测指定字体名是否在系统已安装(用于决定是否回退)。

    匹配规则:字体名(小写、去空格)在系统字体集合中,或其任意前缀子串匹配。
    """
    if not font_name:
        return False
    sys_fonts = list_system_fonts()
    if not sys_fonts:
        return True  # 无法检测时假定可用,不强制回退
    key = font_name.lower().replace(" ", "")
    if key in sys_fonts:
        return True
    # 中文字体名直接匹配
    if font_name in sys_fonts or any(font_name in f for f in sys_fonts):
        return True
    return False


def get_font_roles_checked(name: str) -> tuple[str, str, bool]:
    """返回 (latin_font, eastasia_font, eastasia_available)。

    与 :func:`get_font_roles` 相同,但额外检测 eastAsia 字体是否在系统可用。
    若不可用,会回退到系统已有的中文衬线/无衬线字体,避免渲染时被替换。
    """
    latin, ea = get_font_roles(name)
    if is_font_available(ea):
        return latin, ea, True
    # 回退:优先用系统已有的常见中文字体
    fallbacks = ["宋体", "SimSun", "Noto Serif SC", "Noto Sans CJK SC",
                 "WenQuanYi Zen Hei", "仿宋", "FangSong"]
    sys_fonts = list_system_fonts()
    for fb in fallbacks:
        if fb.lower() in sys_fonts or fb in sys_fonts:
            return latin, fb, False
    return latin, ea, False
