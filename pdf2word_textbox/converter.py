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
        """为 URI 创建(或复用)docx 外部超链接关系,返回 rel_id。

        #5.7 修复:内部跳转(#pageN)用 w:hyperlink w:anchor 而非外部关系,
        返回特殊标记 "__internal__" 让 docx_builder 用 anchor 模式。
        """
        if not uri:
            return None
        if uri.startswith("#"):
            # 内部跳转:返回特殊标记,docx_builder 用 w:anchor
            return "__internal__:" + uri[1:]
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
        """生成一页。

        关键修复:
        - #2: --no-real-header-footer 模式不再丢弃页眉页脚,改为全部放正文
        - #4: 真实页眉页脚模式下,top_margin=0,页眉用 page-relative 定位,
              避免双重偏移(页眉部件本身高度由 header_distance 控制)
        - #5: 每页独立用 region.header_top/footer_top,不再用全局 cfg
        - #17: 页眉页脚区域的图片进 header/footer 部件
        """
        if is_first:
            section = doc.sections[0]
        else:
            section = doc.add_section(WD_SECTION.NEW_PAGE)

        # 页面尺寸(精确)
        section.page_width = Pt(pe.width)
        section.page_height = Pt(pe.height)

        # 边距:#4 修复——真实页眉页脚模式下也用 0 边距,让文本框 page-relative
        # 定位生效(避免 top_margin + page-relative 双重偏移)。
        # 页眉页脚部件的内容仍可正常显示(它们独立于正文 margin)。
        section.top_margin = Pt(0)
        section.bottom_margin = Pt(0)
        section.left_margin = Pt(0)
        section.right_margin = Pt(0)
        section.header_distance = Pt(0)
        section.footer_distance = Pt(0)

        section.different_first_page_header_footer = False
        sectPr = section._sectPr
        titlePg = sectPr.find(qn("w:titlePg"))
        if titlePg is not None:
            sectPr.remove(titlePg)

        # 划分元素:header / footer / body
        header_spans = []
        footer_spans = []
        body_spans = []
        for i, sp in enumerate(pe.spans):
            r = classify_span(sp, region, pe)
            if r == "header":
                header_spans.append((i, sp))
            elif r == "footer":
                footer_spans.append((i, sp))
            else:
                body_spans.append((i, sp))

        # 图片划分:页眉页脚区域内的图片归对应部件
        header_imgs = []
        footer_imgs = []
        body_imgs = []
        for im in pe.images:
            cy = (im.bbox[1] + im.bbox[3]) / 2
            if region.header_top > 0 and cy <= region.header_top:
                header_imgs.append(im)
            elif region.footer_top < pe.height and cy >= region.footer_top:
                footer_imgs.append(im)
            else:
                body_imgs.append(im)

        # 图形划分:排除页眉页脚分割线(单独处理)+ 排除下划线(已作为样式,#5.3)
        body_drawings = []
        header_divider_set = set(id(d) for d in region.header_dividers)
        footer_divider_set = set(id(d) for d in region.footer_dividers)
        for dr in pe.drawings:
            if id(dr) in header_divider_set or id(dr) in footer_divider_set:
                continue
            if getattr(dr, "is_underline", False):
                continue  # #5.3 下划线已作为文本样式,不再画形状
            cy = (dr.bbox[1] + dr.bbox[3]) / 2
            if region.header_top > 0 and cy <= region.header_top:
                continue
            if region.footer_top < pe.height and cy >= region.footer_top:
                continue
            body_drawings.append(dr)

        # ---- 页眉页脚 ----
        extra_dividers = []
        if use_real_hf:
            self._make_header_footer(doc, section, pe, region, cfg,
                                     header_spans, footer_spans,
                                     header_imgs, footer_imgs)
        else:
            # #2 修复:no-real-header-footer 模式把页眉页脚内容放正文(不丢失)
            header_spans.extend(footer_spans)
            body_spans = body_spans + header_spans
            body_imgs = body_imgs + header_imgs + footer_imgs
            # divider 单独处理(Divider 对象无 type 属性,用 add_line)
            extra_dividers = list(region.header_dividers) + list(region.footer_dividers)

        # ---- 正文画布 ----
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pPr = p._p.get_or_add_pPr()
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        for run in p.runs:
            run.font.size = Pt(1)

        # 1. 正文图形(底层)
        for dr in body_drawings:
            self._add_drawing(p, dr, behind=True, z=0)
        # 分割线(no-real-header-footer 模式下,divider 用 add_line)
        for dv in extra_dividers:
            docx_builder.add_line(
                p, dv.bbox[0], dv.bbox[1], dv.bbox[2], dv.bbox[3],
                color=dv.color, width=dv.width, behind=True, z=0,
            )

        # 2. 正文图片
        for im in body_imgs:
            self._add_image(doc, p, im)

        # 3. 正文文本框(上层)+ 合并相邻同样式 span(#5.4 性能优化)
        merged = self._merge_adjacent_spans([sp for _, sp in body_spans])
        for sp in merged:
            self._add_text_span_as_textbox(p, sp, pe, z=10, doc=doc)

    # ---------- 页眉页脚 ----------
    def _make_header_footer(self, doc: Document, section, pe: PageElements,
                            region: PageRegion, cfg: HeaderFooterConfig,
                            header_spans, footer_spans,
                            header_imgs, footer_imgs):
        """把页眉页脚内容写入 docx header/footer。

        #4 修复:页眉页脚内 VML 用 page-relative 定位,header_distance=0,
        所以页眉文本框 y 坐标 = 原 PDF 坐标(无双重偏移)。
        #17 修复:页眉页脚区域的图片也放进对应部件。
        """
        header = section.header
        header.is_linked_to_previous = False
        footer = section.footer
        footer.is_linked_to_previous = False

        if header.paragraphs:
            hp = header.paragraphs[0]
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

        # 页眉文本框 + 图片 + 分割线
        for _, sp in header_spans:
            self._add_text_span_as_textbox(hp, sp, pe, z=10, doc=doc)
        for im in header_imgs:
            self._add_image(doc, hp, im)
        for dv in region.header_dividers:
            docx_builder.add_line(
                hp, dv.bbox[0], dv.bbox[1], dv.bbox[2], dv.bbox[3],
                color=dv.color, width=dv.width, behind=True, z=0,
            )

        # 页脚文本框 + 图片 + 分割线
        for _, sp in footer_spans:
            self._add_text_span_as_textbox(fp, sp, pe, z=10, doc=doc)
        for im in footer_imgs:
            self._add_image(doc, fp, im)
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
        from .fonts import get_font_roles_checked, has_cjk
        x0, y0, x1, y1 = sp.bbox
        actual_w = max(x1 - x0, 1.0)
        h = max(y1 - y0, sp.size, 1.0)
        # 关键修复:用 fitz.get_text_length 精确测宽(而非粗估)
        # 选内置字体:中文用 china-s,西文用 helv
        font_for_measure = "china-s" if has_cjk(sp.text) else "helv"
        try:
            import fitz
            precise_w = fitz.get_text_length(sp.text, fontname=font_for_measure, fontsize=sp.size)
        except Exception:
            precise_w = self._estimate_text_width(sp.text, sp.size)
        # 文本框宽度:取实际bbox、精确测量、估算 三者最大值 + 充足buffer(1.3倍)
        # 1.3倍 buffer 应对字体替换导致的宽度膨胀(LibreOffice无原字体时用替代字体)
        w = max(actual_w, precise_w * 1.3, self._estimate_text_width(sp.text, sp.size) * 1.3) + 4.0
        # 高度:允许 2 行(防止折行时被裁剪)
        h = max(h, sp.size * 1.6)
        # 字体:CJK 字体的西文部分保留原字体名
        font_latin, font_ea, ea_available = get_font_roles_checked(sp.font)
        if has_cjk(sp.font) and not has_cjk(sp.text):
            from .fonts import normalize_font
            font_latin = normalize_font(sp.font)
        # 超链接 rel_id
        hyperlink_rel_id = None
        if sp.hyperlink and doc is not None:
            hyperlink_rel_id = self._get_hyperlink_rel_id(doc, sp.hyperlink)
        # 处理 \xa0:保留为非断空格(表单填空用),但确保不影响渲染
        text = sp.text
        run = {
            "text": text,
            "font_latin": font_latin,
            "font_ea": font_ea,
            "size": sp.size,
            "color": sp.color,
            "bold": sp.bold,
            "italic": sp.italic,
            "underline": getattr(sp, "underline", False) or bool(sp.hyperlink),
            "hyperlink_rel_id": hyperlink_rel_id,
        }
        # 文本框内文字一律左对齐:文本框 left 已是文字 x0,
        # 居中/右对齐会导致文字相对框偏移(框宽度大于文字宽度时)。
        # 原PDF的居中/右对齐通过文本框位置本身已体现(框left=文字x0)。
        align = "left"
        # y 补偿:字体 ascender 导致文字 baseline 相对框 top 下移约 2pt,
        # 将文本框 top 上移 2pt 让文字实际位置对齐原PDF。
        box_y = y0 - 2.0
        # 关键修复:no_wrap=False(允许折行),避免"‹"裁剪符
        docx_builder.add_textbox(
            paragraph, x0, box_y, w, h,
            runs=[run], align=align, line_spacing=1.0,
            behind=False, z=z, vertical_align="top", no_wrap=False,
            rotation=getattr(sp, "rotation", 0.0),
        )

    @staticmethod
    def _infer_alignment(bbox: tuple, page_width: float) -> str:
        """根据 bbox 在页面的水平位置推断对齐方式。

        - bbox 中心在页面中心 ±5% → center
        - bbox 右边缘接近页面右边(距离 < 10% 页宽)→ right
        - 否则 → left
        """
        if page_width <= 0:
            return "left"
        x0, _, x1, _ = bbox
        cx = (x0 + x1) / 2
        page_cx = page_width / 2
        # 居中:中心在页面中心 ±5% 页宽
        if abs(cx - page_cx) < page_width * 0.05:
            return "center"
        # 右对齐:右边缘距页面右边 < 8% 页宽,且文本宽度 > 20pt
        if (page_width - x1) < page_width * 0.08 and (x1 - x0) > 20:
            return "right"
        return "left"

    @staticmethod
    def _merge_adjacent_spans(spans: list) -> list:
        """合并相邻同样式的 span(#5.4 性能优化)。

        合并条件:同行(y 接近)、同字体、同字号、同颜色、同粗斜体、
        x 间距 < 字号×0.3。合并后 text 拼接,bbox 取并集。
        """
        if not spans:
            return spans
        # 按 y 再按 x 排序
        spans = sorted(spans, key=lambda s: (round(s.bbox[1], 1), s.bbox[0]))
        merged = []
        for sp in spans:
            if not merged:
                merged.append(sp)
                continue
            prev = merged[-1]
            same_style = (
                prev.font == sp.font
                and prev.size == sp.size
                and prev.color == sp.color
                and prev.flags == sp.flags
                and prev.hyperlink == sp.hyperlink
            )
            # 同行:y 中心差 < 字号×0.5
            py = (prev.bbox[1] + prev.bbox[3]) / 2
            cy = (sp.bbox[1] + sp.bbox[3]) / 2
            same_line = abs(py - cy) < sp.size * 0.5
            # x 紧邻:间距 < 字号×0.5
            gap = sp.bbox[0] - prev.bbox[2]
            adjacent = -1 <= gap < sp.size * 0.5
            if same_style and same_line and adjacent:
                # 合并:拼接 text,bbox 并集
                from .extractor import TextSpan as TS
                new_bbox = (
                    min(prev.bbox[0], sp.bbox[0]),
                    min(prev.bbox[1], sp.bbox[1]),
                    max(prev.bbox[2], sp.bbox[2]),
                    max(prev.bbox[3], sp.bbox[3]),
                )
                # 替换最后一个
                merged[-1] = TS(
                    bbox=new_bbox,
                    text=prev.text + sp.text,
                    font=prev.font, size=prev.size, color=prev.color,
                    flags=prev.flags, ascender=prev.ascender, descender=prev.descender,
                    underline=prev.underline or sp.underline,
                    strike=prev.strike or sp.strike,
                    hyperlink=prev.hyperlink,
                    rotation=prev.rotation, wmode=prev.wmode,
                )
            else:
                merged.append(sp)
        return merged

    @staticmethod
    def _estimate_text_width(text: str, size: float) -> float:
        """估算文本渲染宽度(pt),用于防止折行。

        CJK 字符约 1em,拉丁字符约 0.55em,空格约 0.3em。
        """
        w = 0.0
        for ch in text:
            cp = ord(ch)
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
            # 关键修复:只对真正的矩形(type=='rect')应用填充。
            # 曲线路径(type=='path', items 含 'c' 曲线)即使 fill!=None,
            # 实际填充区域远小于 bbox,用 bbox 画填充矩形会产生错误的大黑块。
            is_real_rect = dr.type == "rect"
            # 检查 items 是否有 're'(矩形操作)
            has_re_item = any(it[0] == "re" for it in dr.items if isinstance(it, tuple) and len(it) > 0)
            apply_fill = (is_real_rect or has_re_item) and dr.fill_color is not None
            docx_builder.add_rect(
                paragraph, x0, y0, max(w, 0.5), max(h, 0.5),
                stroke_color=dr.stroke_color,
                fill_color=dr.fill_color if apply_fill else None,
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
        """检测文本下划线:若文本 span 下方有紧贴的横线,则标记为下划线。

        #5.3 修复:检测到下划线后,标记对应 Drawing.is_underline=True,
        _make_page 排除 is_underline 的 drawing,避免"真实线 + 下划线样式"双重绘制。
        """
        for pe in pages:
            for sp in pe.spans:
                if not sp.text.strip():
                    continue
                x0, y0, x1, y1 = sp.bbox
                baseline = y1
                for dr in pe.drawings:
                    if dr.is_underline:
                        continue
                    if not is_horizontal_line(dr, tolerance=1.5):
                        continue
                    dx0, dy0, dx1, dy1 = dr.bbox
                    line_y = (dy0 + dy1) / 2
                    # 横线在文本基线下方 0~3pt,且 x 范围有重叠
                    if 0 <= line_y - baseline <= 3.5:
                        ov = min(x1, dx1) - max(x0, dx0)
                        if ov > 0 and ov > (x1 - x0) * 0.4:
                            sp.underline = True
                            dr.is_underline = True  # 标记,后续不再作为形状绘制
                            break
