# -*- coding: utf-8 -*-
"""页眉页脚检测器。

策略:
1. 统计每页顶部/底部区域内的文本 span,找出跨页重复出现的文本 → 页眉/页脚
2. 找出顶部/底部区域内的水平横线(绘图指令)→ 页眉/页脚分割线
3. 识别页码(纯数字、跨页递增的文本)
4. 计算 header/footer 边距(到正文区域的距离)

检测完成后,每页的元素会被标记为 'body' / 'header' / 'footer',
并记录页眉页脚区域与分割线信息。
"""
from __future__ import annotations
from dataclasses import dataclass, field
from collections import Counter
from typing import Optional
import re

from .extractor import PageElements, TextSpan, Drawing


@dataclass
class Divider:
    """页眉/页脚分割横线。"""
    bbox: tuple[float, float, float, float]
    color: str
    width: float
    region: str  # 'header' | 'footer'


@dataclass
class PageRegion:
    """单页的区域划分结果。"""
    page_index: int
    header_top: float = 0.0       # 页眉区域底边(到此 y 为止是页眉)
    footer_top: float = 0.0       # 页脚区域顶边(从此 y 开始是页脚)
    header_dividers: list[Divider] = field(default_factory=list)
    footer_dividers: list[Divider] = field(default_factory=list)
    is_page_number: dict = field(default_factory=dict)  # span_id -> bool


@dataclass
class HeaderFooterConfig:
    """文档级页眉页脚配置。"""
    # 页眉区域高度(pt,从页面顶部往下)
    header_height: float = 72.0
    # 页脚区域高度(pt,从页面底部往上)
    footer_height: float = 72.0
    # 是否启用页眉页脚
    has_header: bool = True
    has_footer: bool = True
    # 页眉页脚分割线
    header_dividers: list[Divider] = field(default_factory=list)
    footer_dividers: list[Divider] = field(default_factory=list)


def is_horizontal_line(dr: Drawing, tolerance: float = 1.0) -> bool:
    """判断绘图是否为水平线。"""
    x0, y0, x1, y1 = dr.bbox
    # 水平线:y方向尺寸很小,x方向较长
    if (y1 - y0) <= tolerance and (x1 - x0) > 5:
        return True
    # 也检查 items 中的 'l' 类型
    for it in dr.items:
        if it[0] == "l":
            p1, p2 = it[1], it[2]
            if abs(p1.y - p2.y) <= tolerance and abs(p2.x - p1.x) > 5:
                return True
    return False


def is_page_number_text(text: str) -> bool:
    """判断文本是否像页码(纯数字、罗马数字、"第X页"等)。"""
    t = text.strip()
    if not t:
        return False
    # 纯数字
    if re.fullmatch(r"-?\d+", t):
        return True
    # 罗马数字
    if re.fullmatch(r"[ivxIVX]+", t):
        return True
    # 第X页 / Page X / X/Y
    if re.fullmatch(r"第\s*\d+\s*页", t):
        return True
    if re.fullmatch(r"-\s*\d+\s*-", t):
        return True
    if re.fullmatch(r"\d+\s*/\s*\d+", t):
        return True
    if re.fullmatch(r"Page\s+\d+", t, re.I):
        return True
    return False


def detect_dividers_in_region(
    drawings: list[Drawing],
    y_min: float,
    y_max: float,
    region: str,
) -> list[Divider]:
    """在指定 y 范围内找水平分割线。"""
    result = []
    for dr in drawings:
        x0, y0, x1, y1 = dr.bbox
        cy = (y0 + y1) / 2
        if y_min <= cy <= y_max and is_horizontal_line(dr):
            result.append(Divider(
                bbox=dr.bbox,
                color=dr.stroke_color or "000000",
                width=max(dr.width, 0.5),
                region=region,
            ))
    return result


def detect_header_footer(
    pages: list[PageElements],
    header_ratio: float = 0.12,
    footer_ratio: float = 0.12,
) -> tuple[HeaderFooterConfig, list[PageRegion]]:
    """检测文档级页眉页脚配置 + 每页区域划分。

    Args:
        pages: 所有页元素
        header_ratio: 页眉区域占页面高度比例(默认顶部 12%)
        footer_ratio: 页脚区域占页面高度比例(默认底部 12%)

    Returns:
        (HeaderFooterConfig, [PageRegion, ...])
    """
    if not pages:
        return HeaderFooterConfig(has_header=False, has_footer=False), []

    # 用第一页确定默认区域高度
    p0 = pages[0]
    default_header_h = p0.height * header_ratio
    default_footer_h = p0.height * footer_ratio

    # 1. 统计页眉/页脚区域的重复文本
    header_texts = Counter()
    footer_texts = Counter()
    n_pages = len(pages)
    for pe in pages:
        h_thresh = pe.height * header_ratio
        f_thresh = pe.height * (1 - footer_ratio)
        for sp in pe.spans:
            cy = (sp.bbox[1] + sp.bbox[3]) / 2
            key = sp.text.strip()
            if not key:
                continue
            if cy <= h_thresh:
                header_texts[key] += 1
            elif cy >= f_thresh:
                footer_texts[key] += 1

    # 重复出现 >= 一半页面 → 真正的页眉/页脚文本
    min_repeat = max(2, n_pages // 3)
    has_header = any(v >= min_repeat for v in header_texts.values()) or n_pages < 5
    has_footer = any(v >= min_repeat for v in footer_texts.values()) or n_pages < 5

    # 2. 找页眉/页脚分割线(用第一页作代表,后续每页单独找)
    cfg = HeaderFooterConfig(
        header_height=default_header_h,
        footer_height=default_footer_h,
        has_header=has_header,
        has_footer=has_footer,
    )

    # 3. 每页区域划分
    regions = []
    for pe in pages:
        h_thresh = pe.height * header_ratio
        f_thresh = pe.height * (1 - footer_ratio)
        # 找该页页眉区域的分割线
        h_divs = detect_dividers_in_region(pe.drawings, 0, h_thresh + 10, "header")
        f_divs = detect_dividers_in_region(pe.drawings, f_thresh - 10, pe.height, "footer")
        # 页眉区域底边 = 页眉分割线的 y(若有),否则用默认
        header_top = h_thresh
        if h_divs:
            # 取最低的分割线作为页眉区域底边
            header_top = max(d.bbox[3] for d in h_divs) + 2
        # 页脚区域顶边 = 页脚分割线的 y(若有),否则用默认
        footer_top = f_thresh
        if f_divs:
            footer_top = min(d.bbox[1] for d in f_divs) - 2

        # 收集分割线到文档配置(用第一页的作代表)
        if pe.page_index == 0:
            cfg.header_dividers = h_divs
            cfg.footer_dividers = f_divs
            cfg.header_height = header_top
            cfg.footer_height = pe.height - footer_top

        # 识别页码
        is_pn = {}
        for i, sp in enumerate(pe.spans):
            cy = (sp.bbox[1] + sp.bbox[3]) / 2
            if cy >= f_thresh or cy <= h_thresh:
                if is_page_number_text(sp.text):
                    is_pn[i] = True

        regions.append(PageRegion(
            page_index=pe.page_index,
            header_top=header_top,
            footer_top=footer_top,
            header_dividers=h_divs,
            footer_dividers=f_divs,
            is_page_number=is_pn,
        ))

    return cfg, regions


def classify_span(
    span: TextSpan,
    region: PageRegion,
    pe: PageElements,
) -> str:
    """判断 span 属于 header / footer / body。

    Args:
        span: 文本片段
        region: 该页区域划分
        pe: 该页元素

    Returns:
        'header' | 'footer' | 'body'
    """
    cy = (span.bbox[1] + span.bbox[3]) / 2
    if region.header_top > 0 and cy <= region.header_top:
        return "header"
    if region.footer_top > 0 and cy >= region.footer_top:
        return "footer"
    return "body"
