# -*- coding: utf-8 -*-
"""PDF 元素提取器:用 PyMuPDF 精确提取每页所有元素的坐标与样式。

提取的元素类型:
- TextSpan: 文本片段(bbox、文本、字体、字号、颜色、粗斜体下划线)
- Drawing: 矢量图形(bbox、类型、描边颜色/宽度、填充颜色)
- Image: 图片(bbox、图像字节)

所有坐标单位为 PDF point(1/72 英寸),原点在页面左上角
(PyMuPDF 已转换为顶部原点,与 DOCX 一致)。
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional
import fitz  # PyMuPDF


@dataclass
class TextSpan:
    """文本片段。一个 span 内字体、字号、颜色一致。"""
    bbox: tuple[float, float, float, float]  # x0, y0, x1, y1
    text: str
    font: str
    size: float
    color: str  # hex
    flags: int  # PyMuPDF font flags
    ascender: float = 0.8
    descender: float = -0.2
    underline: bool = False  # 由后处理检测
    strike: bool = False

    @property
    def bold(self) -> bool:
        return bool(self.flags & 16) or "bold" in self.font.lower()

    @property
    def italic(self) -> bool:
        return bool(self.flags & 2) or "italic" in self.font.lower() or "oblique" in self.font.lower()


@dataclass
class Drawing:
    """矢量图形(直线/矩形/曲线等)。"""
    bbox: tuple[float, float, float, float]
    type: str  # 'line' | 'rect' | 'curve' | 'path'
    stroke_color: Optional[str]  # hex 或 None
    fill_color: Optional[str]    # hex 或 None
    width: float                 # 线宽(pt)
    # 原始路径项,用于精确复刻
    items: list = field(default_factory=list)
    closed: bool = False


@dataclass
class ImageItem:
    """图片。"""
    bbox: tuple[float, float, float, float]
    data: bytes  # 图像字节
    width: int
    height: int
    ext: str  # 扩展名 png/jpeg


@dataclass
class PageElements:
    """一页的所有元素。"""
    page_index: int
    width: float
    height: float
    spans: list[TextSpan] = field(default_factory=list)
    drawings: list[Drawing] = field(default_factory=list)
    images: list[ImageItem] = field(default_factory=list)


class PDFExtractor:
    """从 PDF 提取所有页面元素。"""

    def __init__(self, pdf_path: str, password: str = ""):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        if password:
            self.doc.authenticate(password)
        self.page_count = self.doc.page_count

    def close(self):
        if self.doc:
            self.doc.close()
            self.doc = None

    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()

    def extract_page(self, page_index: int) -> PageElements:
        """提取单页所有元素。"""
        page = self.doc[page_index]
        rect = page.rect
        pe = PageElements(
            page_index=page_index,
            width=rect.width,
            height=rect.height,
        )
        # 1. 文本
        self._extract_text(page, pe)
        # 2. 矢量图形
        self._extract_drawings(page, pe)
        # 3. 图片
        self._extract_images(page, pe)
        return pe

    def extract_all(self, start: int = 0, end: int | None = None) -> list[PageElements]:
        """提取所有页(或指定范围)。"""
        if end is None:
            end = self.page_count
        return [self.extract_page(i) for i in range(start, min(end, self.page_count))]

    # ---------- 文本 ----------
    def _extract_text(self, page, pe: PageElements):
        """用 'dict' 模式提取文本,保留 span 级别的精确信息。"""
        from .colors import to_hex_color
        from .fonts import has_cjk

        try:
            d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        except Exception:
            return
        for block in d.get("blocks", []):
            if block.get("type", 0) != 0:  # 0=文本,1=图片
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if text == "":
                        continue
                    # 跳过纯空白且无宽度
                    bbox = span.get("bbox", (0, 0, 0, 0))
                    if bbox[2] - bbox[0] < 0.1 and text.strip() == "":
                        continue
                    font = span.get("font", "Arial")
                    size = span.get("size", 12)
                    color = to_hex_color(span.get("color", 0))
                    flags = span.get("flags", 0)
                    # 处理下划线:PyMuPDF 不直接给,需要从字符标志位推断
                    # flags bit 0 = superscript;下划线无标准位,后续由 drawings 推断
                    pe.spans.append(TextSpan(
                        bbox=tuple(bbox),
                        text=text,
                        font=font,
                        size=round(size, 2),
                        color=color,
                        flags=flags,
                        ascender=span.get("ascender", 0.8),
                        descender=span.get("descender", -0.2),
                    ))

    # ---------- 矢量图形 ----------
    def _extract_drawings(self, page, pe: PageElements):
        """提取矢量绘图指令。"""
        from .colors import to_hex_color, is_white
        try:
            drawings = page.get_drawings()
        except Exception:
            drawings = []
        for dr in drawings:
            rect = dr.get("rect", None)
            if rect is None:
                continue
            x0, y0, x1, y1 = rect
            # 跳过 0 尺寸
            if x1 - x0 < 0.1 and y1 - y0 < 0.1:
                continue
            stroke = dr.get("color", None)
            fill = dr.get("fill", None)
            width = dr.get("width", 0.0) or 0.0
            items = dr.get("items", [])
            closed = False
            dtype = "path"
            # 判断类型:纯直线 / 矩形 / 曲线
            if len(items) == 1 and items[0][0] == "l":
                dtype = "line"
            elif len(items) == 5 and items[0][0] == "re":
                dtype = "rect"
                closed = True
            elif items and all(it[0] in ("l", "re") for it in items):
                dtype = "path"
                if items[-1][0] == "l":
                    closed = True
            # 跳过纯白色填充(页面背景)
            if fill is not None and is_white(fill) and (stroke is None):
                continue
            pe.drawings.append(Drawing(
                bbox=(x0, y0, x1, y1),
                type=dtype,
                stroke_color=to_hex_color(stroke) if stroke is not None else None,
                fill_color=to_hex_color(fill) if fill is not None else None,
                width=round(width, 2),
                items=items,
                closed=closed,
            ))

    # ---------- 图片 ----------
    def _extract_images(self, page, pe: PageElements):
        """提取图片(含 bbox 与字节数据)。"""
        try:
            info = page.get_image_info(xrefs=True)
        except Exception:
            info = []
        for im in info:
            bbox = im.get("bbox")
            if not bbox:
                continue
            xref = im.get("xref", 0)
            if not xref:
                continue
            try:
                base = self.doc.extract_image(xref)
            except Exception:
                continue
            if not base or not base.get("image"):
                continue
            pe.images.append(ImageItem(
                bbox=tuple(bbox),
                data=base["image"],
                width=base.get("width", 0),
                height=base.get("height", 0),
                ext=base.get("ext", "png"),
            ))
