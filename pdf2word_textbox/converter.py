# -*- coding: utf-8 -*-
"""主转换器:PDF → DOCX(文本框一比一复刻)。

流程:
1. 用 :class:`PDFExtractor` 提取每页元素(文本/图形/图片)
2. 用 :func:`detect_header_footer` 检测页眉页脚区域与分割线
3. 对每页:
   - 创建一个 section,页面尺寸 = PDF 页面尺寸(精确)
   - 边距:top=页眉高度, bottom=页脚高度, left/right=0
   - 页眉区:文本框放置页眉文本 + 页眉分割横线
   - 页脚区:文本框放置页脚文本(含页码) + 页脚分割横线
   - 正文区:一个空段落作为"画布",所有正文文本框/图形/图片绝对定位
4. 检测文本下划线(若文本下方有紧贴的横线,则标记为下划线)
"""
from __future__ import annotations
import io
import logging
from typing import Optional

from docx import Document
from docx.shared import Pt, Emu, RGBColor
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .extractor import PDFExtractor, PageElements, TextSpan, Drawing, ImageItem
from .header_footer import (
    detect_header_footer, classify_span, HeaderFooterConfig, PageRegion,
    is_horizontal_line,
)
from .fonts import get_font_roles, has_cjk
from . import docx_builder

log = logging.getLogger("pdf2word_textbox")


class Converter:
    """PDF → DOCX 文本框一比一复刻转换器。

    Args:
        pdf_file: PDF 文件路径
        password: PDF 密码(可选)
    """

    def __init__(self, pdf_file: str, password: str = ""):
        self.pdf_file = pdf_file
        self.password = password

    def convert(
        self,
        docx_file: str,
        start: int = 0,
        end: Optional[int] = None,
        header_ratio: float = 0.12,
        footer_ratio: float = 0.12,
        use_real_header_footer: bool = True,
        detect_underline: bool = True,
    ) -> str:
        """执行转换。

        Args:
            docx_file: 输出 docx 路径
            start: 起始页(0-based)
            end: 结束页(不含),None 表示到末尾
            header_ratio: 页眉区域占页面高度比例
            footer_ratio: 页脚区域占页面高度比例
            use_real_header_footer: 是否使用 docx 真实页眉页脚
                (True: 页眉页脚内容放 docx header/footer;
                 False: 全部放正文,纯文本框复刻)
            detect_underline: 是否检测文本下划线

        Returns:
            输出 docx 文件路径
        """
        logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
        with PDFExtractor(self.pdf_file, self.password) as ext:
            total = ext.page_count
            if end is None or end > total:
                end = total
            if start < 0:
                start = 0
            log.info("提取 PDF 元素: 第 %d~%d 页 / 共 %d 页", start + 1, end, total)
            pages = ext.extract_all(start, end)

            # 检测页眉页脚
            cfg, regions = detect_header_footer(pages, header_ratio, footer_ratio)
            log.info("页眉页脚检测: has_header=%s has_footer=%s header_h=%.1f footer_h=%.1f",
                     cfg.has_header, cfg.has_footer, cfg.header_height, cfg.footer_height)

            # 下划线检测
            if detect_underline:
                self._detect_underlines(pages)

            # 构建 docx
            doc = Document()
            self._setup_default_doc(doc)
            # 复制 PDF 元数据到 docx
            self._copy_metadata(ext, doc)
            # 超链接 URI → rel_id 缓存
            self._hyperlink_cache: dict[str, str] = {}
            for i, pe in enumerate(pages):
                region = regions[i]
                log.info("(%d/%d) 生成第 %d 页", i + 1, len(pages), pe.page_index + 1)
                self._make_page(doc, pe, region, cfg, use_real_header_footer, i == 0)

            doc.save(docx_file)
            log.info("已保存: %s", docx_file)
            return docx_file

    # ---------- 元数据 ----------
    def _copy_metadata(self, ext: PDFExtractor, doc: Document):
        """把 PDF 元数据(title/author/subject/keywords)复制到 docx core properties。"""
        try:
            meta = ext.doc.metadata or {}
            cp = doc.core_properties
            if meta.get("title"):
                cp.title = meta["title"]
            if meta.get("author"):
                cp.author = meta["author"]
            if meta.get("subject"):
                cp.subject = meta["subject"]
            if meta.get("keywords"):
                cp.keywords = meta["keywords"]
            if meta.get("producer"):
                cp.comments = f"PDF producer: {meta['producer']}"
            log.info("元数据已复制: title=%s author=%s",
                     meta.get("title", "")[:30], meta.get("author", "")[:30])
        except Exception as e:
            log.warning("元数据复制失败: %s", e)

    def _get_hyperlink_rel_id(self, doc: Document, uri: str) -> Optional[str]:
        """为 URI 创建(或复用)docx 外部超链接关系,返回 rel_id。"""
        if not uri:
            return None
        # 内部跳转(#pageN)暂不处理为可点击链接
        if uri.startswith("#"):
            return None
        if uri in self._hyperlink_cache:
            return self._hyperlink_cache[uri]
        try:
            rId = doc.part.relate_to(
                uri,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                is_external=True,
            )
            self._hyperlink_cache[uri] = rId
            return rId
        except Exception as e:
            log.warning("超链接关系创建失败(%s): %s", uri[:50], e)
            return None

    # ---------- docx 初始化 ----------
    def _setup_default_doc(self, doc: Document):
        """设置默认文档样式:无边距、无段落间距、默认字体。"""
        # 默认段落样式
        style = doc.styles["Normal"]
        pf = style.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        # 默认字体
        font = style.font
        font.size = Pt(12)
        font.name = "Arial"
        # 设置 eastAsia 字体
        rpr = style.element.get_or_add_rPr()
        rFonts = rpr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rpr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), "宋体")

    # ---------- 单页生成 ----------
    def _make_page(self, doc: Document, pe: PageElements, region: PageRegion,
                   cfg: HeaderFooterConfig, use_real_hf: bool, is_first: bool):
        """生成一页。"""
        if is_first:
            section = doc.sections[0]
        else:
            section = doc.add_section(WD_SECTION.NEW_PAGE)

        # 页面尺寸(精确)
        section.page_width = Pt(pe.width)
        section.page_height = Pt(pe.height)

        # 边距:top=页眉高度, bottom=页脚高度, left/right=0
        top_m = cfg.header_height if (use_real_hf and cfg.has_header) else 0
        bot_m = cfg.footer_height if (use_real_hf and cfg.has_footer) else 0
        section.top_margin = Pt(top_m)
        section.bottom_margin = Pt(bot_m)
        section.left_margin = Pt(0)
        section.right_margin = Pt(0)
        # 页眉页脚距离
        section.header_distance = Pt(0)
        section.footer_distance = Pt(0)

        # 取消"首页不同"/"奇偶页不同"(确保每页都有页眉页脚)
        section.different_first_page_header_footer = False
        # 设置标题页为 false
        sectPr = section._sectPr
        titlePg = sectPr.find(qn("w:titlePg"))
        if titlePg is not None:
            sectPr.remove(titlePg)

        # ---- 页眉页脚(真实 docx header/footer) ----
        if use_real_hf:
            self._make_header_footer(doc, section, pe, region, cfg)

        # ---- 正文画布 ----
        # 找出正文区的元素
        body_spans = []
        body_drawings = []
        for i, sp in enumerate(pe.spans):
            r = classify_span(sp, region, pe)
            if r == "body":
                body_spans.append((i, sp))
        # 正文区图形:排除页眉页脚分割线(已在 header/footer 处理)
        for dr in pe.drawings:
            cy = (dr.bbox[1] + dr.bbox[3]) / 2
            if region.header_top > 0 and cy <= region.header_top:
                continue
            if region.footer_top > 0 and cy >= region.footer_top:
                continue
            body_drawings.append(dr)

        # 添加一个空段落作为画布
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        # 段落固定行高,避免撑高页面
        pPr = p._p.get_or_add_pPr()
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        # 段落字体设为极小,减少空白
        for run in p.runs:
            run.font.size = Pt(1)

        # 1. 先画正文图形(置于底层)
        for dr in body_drawings:
            self._add_drawing(p, dr, behind=True, z=0)

        # 2. 画正文图片
        for im in pe.images:
            self._add_image(doc, p, im)

        # 3. 画正文文本框(置于上层)
        for idx, sp in body_spans:
            self._add_text_span_as_textbox(p, sp, pe, z=10, doc=doc)

    # ---------- 页眉页脚 ----------
    def _make_header_footer(self, doc: Document, section, pe: PageElements,
                            region: PageRegion, cfg: HeaderFooterConfig):
        """把页眉页脚内容写入 docx header/footer。"""
        # 启用页眉页脚
        header = section.header
        header.is_linked_to_previous = False
        footer = section.footer
        footer.is_linked_to_previous = False

        # 清空默认段落,用我们自己的画布段落
        if header.paragraphs:
            hp = header.paragraphs[0]
            # 清空 runs
            for r in list(hp.runs):
                r._r.getparent().remove(r._r)
        else:
            hp = header.add_paragraph()
        self._zero_paragraph(hp)

        if footer.paragraphs:
            fp = footer.paragraphs[0]
            for r in list(fp.runs):
                r._r.getparent().remove(r._r)
        else:
            fp = footer.add_paragraph()
        self._zero_paragraph(fp)

        # 页眉文本框 + 分割线
        for i, sp in enumerate(pe.spans):
            r = classify_span(sp, region, pe)
            if r == "header":
                self._add_text_span_as_textbox(hp, sp, pe, z=10, doc=doc)
        for dv in region.header_dividers:
            docx_builder.add_line(
                hp, dv.bbox[0], dv.bbox[1], dv.bbox[2], dv.bbox[3],
                color=dv.color, width=dv.width, behind=True, z=0,
            )

        # 页脚文本框 + 分割线
        for i, sp in enumerate(pe.spans):
            r = classify_span(sp, region, pe)
            if r == "footer":
                self._add_text_span_as_textbox(fp, sp, pe, z=10, doc=doc)
        for dv in region.footer_dividers:
            docx_builder.add_line(
                fp, dv.bbox[0], dv.bbox[1], dv.bbox[2], dv.bbox[3],
                color=dv.color, width=dv.width, behind=True, z=0,
            )

    def _zero_paragraph(self, p):
        """把段落行高/间距设为 0,避免撑高页眉页脚区。"""
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        pPr = p._p.get_or_add_pPr()
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)
        spacing.set(qn("w:line"), "20")
        spacing.set(qn("w:lineRule"), "exact")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")

    # ---------- 元素放置 ----------
    def _add_text_span_as_textbox(self, paragraph, sp: TextSpan,
                                  pe: PageElements, z: int = 10,
                                  doc: Optional[Document] = None):
        """把一个文本 span 作为文本框放置(精确坐标)。"""
        x0, y0, x1, y1 = sp.bbox
        w = max(x1 - x0, 1.0)
        h = max(y1 - y0, sp.size, 1.0)
        # 估算文本所需宽度,防止字体替换导致折行
        est_w = self._estimate_text_width(sp.text, sp.size)
        w = max(w, est_w * 1.4) + 4.0
        # 高度也略加 buffer,避免被裁剪
        h = max(h, sp.size * 1.2)
        font_latin, font_ea = get_font_roles(sp.font)
        # 超链接 rel_id
        hyperlink_rel_id = None
        if sp.hyperlink and doc is not None:
            hyperlink_rel_id = self._get_hyperlink_rel_id(doc, sp.hyperlink)
        run = {
            "text": sp.text,
            "font_latin": font_latin,
            "font_ea": font_ea,
            "size": sp.size,
            "color": sp.color,
            "bold": sp.bold,
            "italic": sp.italic,
            "underline": getattr(sp, "underline", False) or bool(sp.hyperlink),
            "hyperlink_rel_id": hyperlink_rel_id,
        }
        align = "left"
        docx_builder.add_textbox(
            paragraph, x0, y0, w, h,
            runs=[run], align=align, line_spacing=1.0,
            behind=False, z=z, vertical_align="top", no_wrap=True,
        )

    @staticmethod
    def _estimate_text_width(text: str, size: float) -> float:
        """估算文本渲染宽度(pt),用于防止折行。

        CJK 字符约 1em,拉丁字符约 0.55em,空格约 0.3em。
        """
        from .fonts import has_cjk
        w = 0.0
        for ch in text:
            cp = ord(ch)
            # CJK
            if (0x4E00 <= cp <= 0x9FFF) or (0x3000 <= cp <= 0x303F) or \
               (0xFF00 <= cp <= 0xFFEF) or (0x3400 <= cp <= 0x4DBF):
                w += size * 1.0
            elif ch == " ":
                w += size * 0.3
            elif ch in "iIl.,;:'|!":
                w += size * 0.3
            elif ch.isupper():
                w += size * 0.7
            else:
                w += size * 0.55
        return w

    def _add_drawing(self, paragraph, dr: Drawing, behind: bool = True, z: int = 0):
        """把矢量图形作为形状放置。"""
        x0, y0, x1, y1 = dr.bbox
        w = x1 - x0
        h = y1 - y0
        # 判断是否为水平/垂直线
        is_h = abs(h) <= 1.0 and abs(w) > 1.0
        is_v = abs(w) <= 1.0 and abs(h) > 1.0
        # 检查 items 是否有明确直线
        if dr.type == "line" and len(dr.items) >= 1:
            it = dr.items[0]
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                docx_builder.add_line(
                    paragraph, p1.x, p1.y, p2.x, p2.y,
                    color=dr.stroke_color or "000000",
                    width=max(dr.width, 0.5),
                    behind=behind, z=z,
                )
                return
        if is_h or is_v:
            docx_builder.add_line(
                paragraph, x0, (y0+y1)/2, x1, (y0+y1)/2 if is_h else y1,
                color=dr.stroke_color or "000000",
                width=max(dr.width, abs(h) if is_h else abs(w), 0.5),
                behind=behind, z=z,
            )
        else:
            # 矩形/路径:用矩形包围盒复刻
            docx_builder.add_rect(
                paragraph, x0, y0, max(w, 0.5), max(h, 0.5),
                stroke_color=dr.stroke_color,
                fill_color=dr.fill_color,
                stroke_width=max(dr.width, 0.5),
                behind=behind, z=z,
            )

    def _add_image(self, doc: Document, paragraph, im: ImageItem):
        """放置图片(绝对定位)。"""
        x0, y0, x1, y1 = im.bbox
        w = max(x1 - x0, 1.0)
        h = max(y1 - y0, 1.0)
        try:
            from docx.image.image import Image as DocxImage
            from docx.parts.image import ImagePart
            image = DocxImage.from_blob(im.data)
            image_part = ImagePart.from_image(image, doc.part.package)
            rId = doc.part.relate_to(
                image_part,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            )
            docx_builder.add_image(paragraph, x0, y0, w, h, im.data, rId, behind=False, z=5)
        except Exception as e:
            log.warning("图片放置失败: %s", e)

    # ---------- 下划线检测 ----------
    def _detect_underlines(self, pages: list[PageElements]):
        """检测文本下划线:若文本 span 下方有紧贴的横线,则标记为下划线。"""
        for pe in pages:
            for sp in pe.spans:
                if not sp.text.strip():
                    continue
                x0, y0, x1, y1 = sp.bbox
                baseline = y1
                for dr in pe.drawings:
                    if not is_horizontal_line(dr, tolerance=1.5):
                        continue
                    dx0, dy0, dx1, dy1 = dr.bbox
                    line_y = (dy0 + dy1) / 2
                    # 横线在文本基线下方 0~3pt,且 x 范围有重叠
                    if 0 <= line_y - baseline <= 3.5:
                        # x 重叠 > 50%
                        ov = min(x1, dx1) - max(x0, dx0)
                        if ov > 0 and ov > (x1 - x0) * 0.4:
                            sp.underline = True  # type: ignore[attr-defined]
                            break
