# -*- coding: utf-8 -*-
"""DOCX 文本框与形状构建器(VML 版本,兼容 Word/LibreOffice/WPS)。

使用 VML(``w:pict`` / ``v:shape``)生成绝对定位(相对页面)的文本框与形状。
VML 虽是较旧规范,但在 Word 2007+、LibreOffice、WPS 中都能正确渲染,
兼容性优于 DrawingML 的 ``wps:wsp`` 文本框。

单位:VML 默认像素,本模块统一用 ``pt`` 后缀(1 pt = 1/72 英寸),
与 PDF/PyMuPDF 坐标系一致。所有坐标原点在页面左上角。

通过 ``mso-position-horizontal-relative:page`` /
``mso-position-vertical-relative:page`` 实现页相对定位。
"""
from __future__ import annotations
from typing import Optional, Sequence
from docx.oxml.ns import nsmap
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from lxml import etree

# 注册所需命名空间(python-docx 默认未注册 wps/mc/v/o/w10)
_NS = {
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
}
for _k, _v in _NS.items():
    if _k not in nsmap:
        nsmap[_k] = _v

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
W10_NS = "urn:schemas-microsoft-com:office:word"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 全局 id 计数器
_id_counter = [100]


def _next_id() -> int:
    _id_counter[0] += 1
    return _id_counter[0]


def _pt(v: float) -> str:
    """格式化为 VML 长度字符串(pt)。"""
    return f"{v:.3f}pt"


def _make_pict(paragraph: Paragraph) -> etree._Element:
    """在段落中创建 ``<w:r><w:pict/></w:r>`` 并返回 ``w:pict`` 元素。

    OOXML schema 要求 ``w:pict`` 必须在 ``w:r``(run)内,不能直接在 ``w:p`` 下。
    """
    r = etree.SubElement(paragraph._p, "{%s}r" % W_NS)
    pict = etree.SubElement(r, "{%s}pict" % W_NS)
    return pict


def _make_run_props_xml(font_latin: str, font_ea: str, size: float,
                        color: str, bold: bool, italic: bool,
                        underline: bool, strike: bool = False) -> str:
    """生成 ``w:rPr`` XML 字符串。"""
    parts = [f'<w:rFonts w:ascii="{font_latin}" w:hAnsi="{font_latin}" '
             f'w:eastAsia="{font_ea}" w:cs="{font_latin}"/>',
             f'<w:color w:val="{color}"/>',
             f'<w:sz w:val="{int(round(size * 2))}"/>',
             f'<w:szCs w:val="{int(round(size * 2))}"/>']
    if bold:
        parts.append('<w:b/><w:bCs/>')
    if italic:
        parts.append('<w:i/><w:iCs/>')
    if underline:
        parts.append('<w:u w:val="single"/>')
    if strike:
        parts.append('<w:strike/>')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


def add_textbox(
    paragraph: Paragraph,
    x: float, y: float, w: float, h: float,
    runs: Sequence[dict],
    align: str = "left",
    line_spacing: float = 1.0,
    behind: bool = False,
    z: int = 1,
    fill_color: Optional[str] = None,
    border: bool = False,
    vertical_align: str = "top",
    no_wrap: bool = False,
    rotation: float = 0.0,
):
    """添加绝对定位(页相对)的文本框。

    Args:
        paragraph: 目标段落
        x, y: 左上角坐标(pt,相对页面)
        w, h: 宽高(pt)
        runs: 文本 run 列表
        align: left/center/right/justify
        behind: 是否置于文字下方(z-index 为负)
        z: 层级(用于覆盖顺序)
        fill_color: 填充色 hex
        border: 是否显示边框
        vertical_align: top/center/bottom(写入 v-text-anchor)
        no_wrap: 是否禁止折行(默认 False,允许折行避免"‹"裁剪符)
        rotation: 旋转角度(度,顺时针)
    """
    _id = _next_id()
    z_index = z if not behind else -z

    # 构建段落属性
    jc = "" if align in ("left",) else f'<w:jc w:val="{align}"/>'
    line_val = int(240 * line_spacing)
    # 关键修复:不再强制 wordWrap=0。
    # 之前 wordWrap=0 + mso-wrap:none 导致 LibreOffice 文本溢出时显示"‹"裁剪符。
    # 现在允许文本自然折行(若框不够宽),避免裁剪符。
    pPr = (f'<w:pPr>'
           f'<w:spacing w:line="{line_val}" w:lineRule="auto" '
           f'w:before="0" w:after="0" w:beforeLines="0" w:afterLines="0"/>'
           f'{jc}'
           f'<w:ind w:left="0" w:right="0" w:firstLine="0"/>'
           f'</w:pPr>')

    runs_xml = ""
    for run in runs:
        rPr = _make_run_props_xml(
            font_latin=run.get("font_latin", "Arial"),
            font_ea=run.get("font_ea", "宋体"),
            size=run.get("size", 12),
            color=run.get("color", "000000"),
            bold=run.get("bold", False),
            italic=run.get("italic", False),
            underline=run.get("underline", False),
            strike=run.get("strike", False),
        )
        text = (run.get("text", "")
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;"))
        rel_id = run.get("hyperlink_rel_id")
        if rel_id:
            if rel_id.startswith("__internal__:"):
                anchor = rel_id[len("__internal__:"):]
                runs_xml += (f'<w:hyperlink w:anchor="{anchor}" '
                            f'xmlns:w="{W_NS}">'
                            f'<w:r>{rPr}<w:t xml:space="preserve">{text}</w:t></w:r>'
                            f'</w:hyperlink>')
            else:
                runs_xml += (f'<w:hyperlink r:id="{rel_id}" '
                            f'xmlns:r="{R_NS}">'
                            f'<w:r>{rPr}<w:t xml:space="preserve">{text}</w:t></w:r>'
                            f'</w:hyperlink>')
        else:
            runs_xml += f'<w:r>{rPr}<w:t xml:space="preserve">{text}</w:t></w:r>'

    txbx_content = f'<w:txbxContent xmlns:w="{W_NS}"><w:p>{pPr}{runs_xml}</w:p></w:txbxContent>'

    fill_attr = "false" if not fill_color else "true"
    fill_xml = f'<v:fill color="#{fill_color}"/>' if fill_color else '<v:fill on="false"/>'
    stroke_attr = "true" if border else "false"
    stroke_xml = f'<v:stroke on="true" color="#000000" weight="0.75pt"/>' if border else '<v:stroke on="false"/>'

    v_anchor = {"top": "top", "center": "middle", "bottom": "bottom"}.get(vertical_align, "top")
    rot_attr = f' rotation="{-rotation:.2f}"' if abs(rotation) > 0.5 else ''

    # 关键修复:移除 mso-wrap:none,改为 mso-wrap:square(允许文本在框内折行)
    # 保留 page-relative 定位,但不禁止折行
    style = (f"position:absolute;"
             f"left:{_pt(x)};top:{_pt(y)};"
             f"width:{_pt(w)};height:{_pt(h)};"
             f"z-index:{z_index};"
             f"mso-position-horizontal:absolute;"
             f"mso-position-vertical:absolute;"
             f"mso-position-horizontal-relative:page;"
             f"mso-position-vertical-relative:page;"
             f"mso-wrap:square;")

    shape_xml = f'''<v:shape xmlns:v="{V_NS}" xmlns:o="{O_NS}" xmlns:w10="{W10_NS}" xmlns:w="{W_NS}" xmlns:r="{R_NS}"
 id="TextBox{_id}" type="#_x0000_t202"{rot_attr}
 style="{style}"
 filled="{fill_attr}" stroked="{stroke_attr}"
 v-text-anchor="{v_anchor}">
 {fill_xml}
 {stroke_xml}
 <v:textbox style="mso-fit-shape-to-text:false" inset="0pt,0pt,0pt,0pt">
 {txbx_content}
 </v:textbox>
</v:shape>'''

    pict = _make_pict(paragraph)
    shape = etree.fromstring(shape_xml)
    pict.append(shape)
    return pict


def add_line(
    paragraph: Paragraph,
    x1: float, y1: float, x2: float, y2: float,
    color: str = "000000",
    width: float = 0.75,
    behind: bool = True,
    z: int = 0,
):
    """添加直线。水平/垂直线用细长矩形(渲染最稳定),斜线用 v:line。"""
    z_index = z if not behind else -z
    dx = x2 - x1
    dy = y2 - y1
    is_horizontal = abs(dy) < 0.5 and abs(dx) > 0.5
    is_vertical = abs(dx) < 0.5 and abs(dy) > 0.5

    if is_horizontal or is_vertical:
        # 细长矩形
        x = min(x1, x2)
        y = min(y1, y2)
        w = abs(dx) if is_horizontal else width
        h = width if is_horizontal else abs(dy)
        style = (f"position:absolute;"
                 f"left:{_pt(x)};top:{_pt(y)};"
                 f"width:{_pt(max(w, 0.1))};height:{_pt(max(h, 0.1))};"
                 f"z-index:{z_index};"
                 f"mso-position-horizontal:absolute;"
                 f"mso-position-vertical:absolute;"
                 f"mso-position-horizontal-relative:page;"
                 f"mso-position-vertical-relative:page;"
                 f"mso-wrap:none;")
        shape_xml = f'''<v:rect xmlns:v="{V_NS}" xmlns:w="{W_NS}" xmlns:w10="{W10_NS}"
 id="Line{_next_id()}" style="{style}" filled="true" stroked="false">
 <v:fill color="#{color}"/>
 <v:stroke on="false"/>
</v:rect>'''
        pict = _make_pict(paragraph)
        pict.append(etree.fromstring(shape_xml))
        return pict
    else:
        # 斜线用 v:line(from/to 为像素,需转换:1pt = 4/3 px ≈ 1.3333px)
        px = lambda pt: pt * 4.0 / 3.0
        style = (f"position:absolute;"
                 f"left:0;top:0;"
                 f"z-index:{z_index};"
                 f"mso-position-horizontal:absolute;"
                 f"mso-position-vertical:absolute;"
                 f"mso-position-horizontal-relative:page;"
                 f"mso-position-vertical-relative:page;"
                 f"mso-wrap:none;")
        shape_xml = f'''<v:line xmlns:v="{V_NS}" xmlns:w="{W_NS}" xmlns:w10="{W10_NS}"
 id="Line{_next_id()}" style="{style}"
 from="{px(x1)},{px(y1)}" to="{px(x2)},{px(y2)}"
 strokecolor="#{color}" strokeweight="{_pt(width)}">
</v:line>'''
        pict = _make_pict(paragraph)
        pict.append(etree.fromstring(shape_xml))
        return pict


def add_rect(
    paragraph: Paragraph,
    x: float, y: float, w: float, h: float,
    stroke_color: Optional[str] = None,
    fill_color: Optional[str] = None,
    stroke_width: float = 0.75,
    behind: bool = True,
    z: int = 0,
):
    """添加矩形(可带描边/填充)。"""
    z_index = z if not behind else -z
    style = (f"position:absolute;"
             f"left:{_pt(x)};top:{_pt(y)};"
             f"width:{_pt(max(w, 0.1))};height:{_pt(max(h, 0.1))};"
             f"z-index:{z_index};"
             f"mso-position-horizontal:absolute;"
             f"mso-position-vertical:absolute;"
             f"mso-position-horizontal-relative:page;"
             f"mso-position-vertical-relative:page;"
             f"mso-wrap:none;")
    fill_xml = f'<v:fill color="#{fill_color}"/>' if fill_color else '<v:fill on="false"/>'
    stroke_xml = (f'<v:stroke on="true" color="#{stroke_color}" weight="{_pt(stroke_width)}"/>'
                  if stroke_color else '<v:stroke on="false"/>')
    shape_xml = f'''<v:rect xmlns:v="{V_NS}" xmlns:w="{W_NS}" xmlns:w10="{W10_NS}"
 id="Rect{_next_id()}" style="{style}"
 filled="{'true' if fill_color else 'false'}" stroked="{'true' if stroke_color else 'false'}">
 {fill_xml}
 {stroke_xml}
</v:rect>'''
    pict = _make_pict(paragraph)
    pict.append(etree.fromstring(shape_xml))
    return pict


def add_image(
    paragraph: Paragraph,
    x: float, y: float, w: float, h: float,
    image_data: bytes,
    rel_id: str,
    behind: bool = False,
    z: int = 1,
):
    """添加图片(绝对定位)。用 v:imagedata 引用关系。"""
    z_index = z if not behind else -z
    style = (f"position:absolute;"
             f"left:{_pt(x)};top:{_pt(y)};"
             f"width:{_pt(max(w, 0.1))};height:{_pt(max(h, 0.1))};"
             f"z-index:{z_index};"
             f"mso-position-horizontal:absolute;"
             f"mso-position-vertical:absolute;"
             f"mso-position-horizontal-relative:page;"
             f"mso-position-vertical-relative:page;"
             f"mso-wrap:none;")
    shape_xml = f'''<v:rect xmlns:v="{V_NS}" xmlns:r="{R_NS}" xmlns:w="{W_NS}" xmlns:w10="{W10_NS}"
 id="Pic{_next_id()}" style="{style}" filled="false" stroked="false">
 <v:fill on="false"/>
 <v:stroke on="false"/>
 <v:imagedata r:id="{rel_id}" o:title="image"/>
</v:rect>'''
    pict = _make_pict(paragraph)
    pict.append(etree.fromstring(shape_xml))
    return pict
