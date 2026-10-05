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
    hyperlink: Optional[str] = None  # 超链接 URI(由后处理匹配)
    rotation: float = 0.0  # 旋转角度(度,顺时针);0=水平,90=竖排
    wmode: int = 0  # 书写模式:0=水平,1=竖排

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
    # #5.3 修复:标记为下划线(检测到后,该 drawing 不再作为形状绘制,避免双重线)
    is_underline: bool = False


@dataclass
class LinkItem:
    """超链接注释。"""
    bbox: tuple[float, float, float, float]
    uri: Optional[str]      # 外部 URI
    kind: str = "uri"       # uri | page | named
    target_page: Optional[int] = None  # 内部跳转目标页


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
    links: list[LinkItem] = field(default_factory=list)


class PDFExtractor:
    """从 PDF 提取所有页面元素。"""

    def __init__(self, pdf_path: str, password: str = ""):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        # #9 修复:检查加密状态并校验密码
        if self.doc.is_encrypted:
            if not password:
                raise ValueError(
                    "PDF 已加密,请提供 password 参数。"
                    "(Encrypted PDF: password required)"
                )
            if not self.doc.authenticate(password):
                raise ValueError(
                    "PDF 密码错误,无法解密。"
                    "(Wrong password: authentication failed)"
                )
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
        # 4. 超链接
        self._extract_links(page, pe)
        return pe

    def extract_all(self, start: int = 0, end: int | None = None) -> list[PageElements]:
        """提取所有页(或指定范围)。"""
        if end is None:
            end = self.page_count
        return [self.extract_page(i) for i in range(start, min(end, self.page_count))]

    # ---------- 文本 ----------
    def _extract_text(self, page, pe: PageElements):
        """用 'dict' 模式提取文本,保留 span 级别的精确信息。"""
        from .colors import span_color_to_hex
        from .fonts import has_cjk, normalize_size

        try:
            d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        except Exception:
            return
        for block in d.get("blocks", []):
            if block.get("type", 0) != 0:  # 0=文本,1=图片
                continue
            for line in block.get("lines", []):
                # 读取行级变换矩阵(用于旋转/竖排文本)
                line_dir = line.get("dir", (1.0, 0.0))  # (cos, sin)
                line_wmode = line.get("wmode", 0)  # 0=水平, 1=竖排
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if text == "":
                        continue
                    # 跳过纯空白且无宽度
                    bbox = span.get("bbox", (0, 0, 0, 0))
                    if bbox[2] - bbox[0] < 0.1 and text.strip() == "":
                        continue
                    font = span.get("font", "Arial")
                    # 字号整数化:消除 PDF 浮点误差(13.99→14.0)
                    size = normalize_size(span.get("size", 12))
                    # span["color"] 是打包整数 0xRRGGBB,用专门函数处理
                    color = span_color_to_hex(span.get("color", 0))
                    flags = span.get("flags", 0)
                    # 旋转角度(度):从行方向向量计算
                    import math
                    cos_a, sin_a = line_dir
                    angle = math.degrees(math.atan2(sin_a, cos_a))
                    pe.spans.append(TextSpan(
                        bbox=tuple(bbox),
                        text=text,
                        font=font,
                        size=size,
                        color=color,
                        flags=flags,
                        ascender=span.get("ascender", 0.8),
                        descender=span.get("descender", -0.2),
                        rotation=round(angle, 2),
                        wmode=line_wmode,
                    ))

    # ---------- 矢量图形 ----------
    def _extract_drawings(self, page, pe: PageElements):
        """提取矢量绘图指令。

        关键修复:过滤掉文字字形(汉字矢量轮廓)。
        PyMuPDF 会把 CID 字体的文字字形作为 drawing 提取(大量贝塞尔曲线,
        fill=黑色,bbox 与文字 span 重合)。这些不是真正的矢量图形,
        若作为 drawing 绘制会产生黑色矩形遮盖文字。
        判断:items 含大量 'c' 曲线(>10)且 bbox 与某 span 重合 → 是字形,跳过。
        """
        from .colors import to_hex_color, is_white
        try:
            drawings = page.get_drawings()
        except Exception:
            drawings = []
        # 预收集文字 span bbox(用于过滤字形)
        span_bboxes = [(sp.bbox, sp.text) for sp in pe.spans if sp.text.strip()]

        for dr in drawings:
            rect = dr.get("rect", None)
            if rect is None:
                continue
            x0, y0, x1, y1 = rect
            if x1 - x0 < 0.1 and y1 - y0 < 0.1:
                continue
            stroke = dr.get("color", None)
            fill = dr.get("fill", None)
            width = dr.get("width", 0.0) or 0.0
            items = dr.get("items", [])
            # 关键修复:过滤文字字形
            # 字形特征:items 含大量 'c' 曲线(>10),fill=黑色,bbox 与文字重合
            c_count = sum(1 for it in items if isinstance(it, tuple) and it[0] == "c")
            if c_count > 10 and fill is not None:
                # 检查是否与文字 span 重合
                is_glyph = False
                for sbbox, stext in span_bboxes:
                    sx0, sy0, sx1, sy1 = sbbox
                    # 重合:bbox 交集 > 50%
                    ix0, iy0 = max(x0, sx0), max(y0, sy0)
                    ix1, iy1 = min(x1, sx1), min(y1, sy1)
                    if ix0 < ix1 and iy0 < iy1:
                        overlap = (ix1 - ix0) * (iy1 - iy0)
                        dr_area = (x1 - x0) * (y1 - y0)
                        if dr_area > 0 and overlap / dr_area > 0.3:
                            is_glyph = True
                            break
                if is_glyph:
                    continue  # 跳过文字字形
            closed = False
            dtype = "path"
            if len(items) == 1 and items[0][0] == "l":
                dtype = "line"
            elif len(items) == 5 and items[0][0] == "re":
                dtype = "rect"
                closed = True
            elif items and all(it[0] in ("l", "re") for it in items):
                dtype = "path"
                if items[-1][0] == "l":
                    closed = True
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
        """提取图片(含 bbox 与字节数据)。

        #5.6 修复:
        - 用 xref 去重(同一图片跨页/同页复用只存一份,减小体积)
        - 尝试用 Pixmap 渲染含 SMask 的透明图片(避免黑底)
        """
        try:
            info = page.get_image_info(xrefs=True)
        except Exception:
            info = []
        seen_xrefs: set[int] = set()
        for im in info:
            bbox = im.get("bbox")
            if not bbox:
                continue
            xref = im.get("xref", 0)
            if not xref:
                continue
            # #5.6 去重:同 xref 跳过(已在别处存过)
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            try:
                base = self.doc.extract_image(xref)
            except Exception:
                continue
            if not base or not base.get("image"):
                continue
            data = base["image"]
            ext = base.get("ext", "png")
            # #5.6 修复:若图片有 SMask(软掩码/透明),合成到白底避免黑底
            smask = base.get("smask", 0)
            if smask:
                try:
                    pix = fitz.Pixmap(self.doc, xref)
                    if pix.alpha:
                        # 合成 alpha 到白底
                        pix = fitz.Pixmap(pix, 0) if pix.n >= 5 else pix
                        pix.set_dpi(72, 72)
                        data = pix.tobytes("png")
                        ext = "png"
                    pix = None
                except Exception:
                    pass
            pe.images.append(ImageItem(
                bbox=tuple(bbox),
                data=data,
                width=base.get("width", 0),
                height=base.get("height", 0),
                ext=ext,
            ))

    # ---------- 超链接 ----------
    def _extract_links(self, page, pe: PageElements):
        """提取超链接注释,并匹配到对应文本 span。

        PyMuPDF 的 ``page.get_links()`` 返回链接列表,每个链接含:
        - 'kind': LINK_URI(外部)、LINK_GOTO(内部跳转)等
        - 'from': 链接矩形 Rect
        - 'uri': 外部 URI(LINK_URI 时)
        - 'page': 目标页(LINK_GOTO 时)
        """
        try:
            links = page.get_links()
        except Exception:
            links = []
        for lk in links:
            frm = lk.get("from")
            if not frm:
                continue
            x0, y0, x1, y1 = frm
            # 跳过 0 尺寸
            if x1 - x0 < 1 or y1 - y0 < 1:
                continue
            kind = lk.get("kind", 0)
            uri = lk.get("uri")
            target_page = lk.get("page")  # 0-based
            li = LinkItem(
                bbox=(x0, y0, x1, y1),
                uri=uri,
                kind="uri" if kind == fitz.LINK_URI else ("page" if kind == fitz.LINK_GOTO else "named"),
                target_page=target_page,
            )
            pe.links.append(li)
            # 匹配到 span:span 的中心点在 link 矩形内,则标记 hyperlink
            for sp in pe.spans:
                if sp.hyperlink:
                    continue
                cx = (sp.bbox[0] + sp.bbox[2]) / 2
                cy = (sp.bbox[1] + sp.bbox[3]) / 2
                if x0 <= cx <= x1 and y0 <= cy <= y1:
                    if uri:
                        sp.hyperlink = uri
                    elif li.kind == "page" and target_page is not None:
                        sp.hyperlink = f"#page{target_page + 1}"
