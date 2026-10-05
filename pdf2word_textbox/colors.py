# -*- coding: utf-8 -*-
"""颜色工具:PDF 颜色空间 → DOCX 颜色(RGB hex)。"""
from __future__ import annotations
from typing import Sequence, Optional, Tuple


def to_hex_color(color) -> str:
    """把 PyMuPDF 返回的颜色(sRGB 0~1 浮点)转成 6 位 hex 字符串。

    PyMuPDF 的 ``shape['color']`` / ``shape['fill']`` / ``span['color']``
    返回值为 0~1 浮点(已做 sRGB 转换)或 None。

    Args:
        color: 颜色值,可能是 None / float(灰度) / (r,g,b) 三元组。

    Returns:
        6 位大写 hex,如 ``"FF0000"``。None 时返回黑色 ``"000000"``。
    """
    if color is None:
        return "000000"
    # 浮点灰度
    if isinstance(color, (int, float)):
        v = max(0, min(255, int(round(float(color) * 255))))
        return f"{v:02X}{v:02X}{v:02X}"
    # 序列 (r, g, b) 或 (c, m, y, k)
    if isinstance(color, (tuple, list)):
        if len(color) >= 3:
            r = max(0, min(255, int(round(float(color[0]) * 255))))
            g = max(0, min(255, int(round(float(color[1]) * 255))))
            b = max(0, min(255, int(round(float(color[2]) * 255))))
            return f"{r:02X}{g:02X}{b:02X}"
        if len(color) == 1:
            return to_hex_color(color[0])
    return "000000"


def is_white(color) -> bool:
    """判断颜色是否接近白色(用于跳过白色文本/线条)。"""
    if color is None:
        return False
    if isinstance(color, (int, float)):
        return float(color) > 0.95
    if isinstance(color, (tuple, list)) and len(color) >= 3:
        return all(float(c) > 0.95 for c in color[:3])
    return False


def is_black(color) -> bool:
    """判断颜色是否接近黑色。"""
    if color is None:
        return True  # PDF 默认黑色
    if isinstance(color, (int, float)):
        return float(color) < 0.05
    if isinstance(color, (tuple, list)) and len(color) >= 3:
        return all(float(c) < 0.05 for c in color[:3])
    return False
