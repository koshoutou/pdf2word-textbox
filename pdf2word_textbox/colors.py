# -*- coding: utf-8 -*-
"""颜色工具:PDF 颜色空间 → DOCX 颜色(RGB hex)。

PyMuPDF 的颜色返回值有两种来源,语义不同:

1. **文本 span 颜色** ``span["color"]``:打包整数 ``0xRRGGBB``
   (黑=0, 白=16777215, 红=16711680=0xFF0000, 蓝=255=0x0000FF, 绿=65280=0x00FF00)
2. **绘图颜色** ``drawing["color"]`` / ``drawing["fill"]``:``None`` /
   0~1 浮点(灰度) / ``(r, g, b)`` 三元组(已 sRGB 转换) /
   ``(c, m, y, k)`` 四元组

本模块用 ``to_hex_color`` 统一处理,通过启发式区分:
- ``int`` 且 > 255 → 视为 ``0xRRGGBB`` 打包整数
- ``int``/``float`` 且 0~255 → 视为灰度(0=黑,255=白);0~1 浮点也按灰度
- ``(r,g,b)`` / ``(r,g,b,a)`` → sRGB 0~1
- ``(c,m,y,k)`` → 先转 RGB
"""
from __future__ import annotations
from typing import Sequence, Optional, Tuple


def to_hex_color(color) -> str:
    """把 PyMuPDF 返回的颜色转成 6 位大写 hex 字符串。

    Args:
        color: 颜色值,可能是:
            - ``None`` → 黑色 ``000000``
            - 打包整数 ``0xRRGGBB``(span["color"],范围 0~16777215)
            - 0~1 浮点灰度(drawing 中的单值)
            - ``(r, g, b)`` 0~1 浮点三元组
            - ``(r, g, b, a)`` 0~1 四元组(忽略 alpha)
            - ``(c, m, y, k)`` 0~1 四元组(CMYK)

    Returns:
        6 位大写 hex,如 ``"FF0000"``。None 时返回黑色 ``"000000"``。
    """
    if color is None:
        return "000000"

    # 整数:区分打包 0xRRGGBB 与 0~255 灰度
    if isinstance(color, int) and not isinstance(color, bool):
        if color > 255:
            # 打包整数 0xRRGGBB(PyMuPDF span["color"])
            return f"{color:06X}"
        # 0~255 灰度
        v = max(0, min(255, color))
        return f"{v:02X}{v:02X}{v:02X}"

    # 浮点:0~1 灰度
    if isinstance(color, float):
        v = max(0, min(255, int(round(color * 255))))
        return f"{v:02X}{v:02X}{v:02X}"

    # 序列
    if isinstance(color, (tuple, list)):
        if len(color) >= 4:
            # 四元组:判断 CMYK 还是 RGBA
            # PyMuPDF 绘图返回 CMYK 时值为 0~1;RGBA 也 0~1
            # 约定:四元组视为 CMYK(PyMuPDF drawing 常见)
            c, m, y, k = (float(color[i]) for i in range(4))
            r = int(round((1 - c) * (1 - k) * 255))
            g = int(round((1 - m) * (1 - k) * 255))
            b = int(round((1 - y) * (1 - k) * 255))
            return f"{max(0,min(255,r)):02X}{max(0,min(255,g)):02X}{max(0,min(255,b)):02X}"
        if len(color) >= 3:
            r = max(0, min(255, int(round(float(color[0]) * 255))))
            g = max(0, min(255, int(round(float(color[1]) * 255))))
            b = max(0, min(255, int(round(float(color[2]) * 255))))
            return f"{r:02X}{g:02X}{b:02X}"
        if len(color) == 1:
            return to_hex_color(color[0])
    return "000000"


def span_color_to_hex(color) -> str:
    """专门处理 PyMuPDF span["color"](打包整数 0xRRGGBB)。

    PyMuPDF 的 span["color"] 永远是 0xRRGGBB 打包整数:
    - 黑 = 0
    - 白 = 16777215 (0xFFFFFF)
    - 红 = 16711680 (0xFF0000)
    - 绿 = 65280 (0x00FF00)
    - 蓝 = 255 (0x0000FF)

    因此整数一律按打包 RGB 处理(255 → 0000FF 蓝色,而非灰度白)。
    """
    if color is None:
        return "000000"
    if isinstance(color, int) and not isinstance(color, bool):
        # span 颜色永远是 0xRRGGBB 打包整数
        return f"{color & 0xFFFFFF:06X}"
    if isinstance(color, float):
        # 极少见:浮点(按打包处理)
        return f"{int(color) & 0xFFFFFF:06X}"
    return to_hex_color(color)


def is_white(color) -> bool:
    """判断颜色是否接近白色(用于跳过白色文本/线条)。"""
    if color is None:
        return False
    if isinstance(color, (int, float)) and not isinstance(color, bool):
        if float(color) > 255:
            # 打包整数:检查 R/G/B 三通道是否都接近 255
            v = int(color) & 0xFFFFFF
            r, g, b = (v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF
            return r > 240 and g > 240 and b > 240
        return float(color) > 0.95 or (isinstance(color, int) and color > 240)
    if isinstance(color, (tuple, list)) and len(color) >= 3:
        return all(float(c) > 0.95 for c in color[:3])
    return False


def is_black(color) -> bool:
    """判断颜色是否接近黑色。"""
    if color is None:
        return True  # PDF 默认黑色
    if isinstance(color, (int, float)) and not isinstance(color, bool):
        if float(color) > 255:
            v = int(color) & 0xFFFFFF
            r, g, b = (v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF
            return r < 15 and g < 15 and b < 15
        return float(color) < 0.05 or (isinstance(color, int) and color < 15)
    if isinstance(color, (tuple, list)) and len(color) >= 3:
        return all(float(c) < 0.05 for c in color[:3])
    return False
