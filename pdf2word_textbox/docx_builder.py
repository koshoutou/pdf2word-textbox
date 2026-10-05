# -*- coding: utf-8 -*-
"""DOCX 文本框与形状构建器(mc:AlternateContent 版,兼容 Word/WPS/LibreOffice)。

每个文本框/形状用 ``mc:AlternateContent`` 包裹:
- ``mc:Choice Requires="wps"``:DrawingML(Word 2007+/WPS 原生支持,精确定位)
- ``mc:Fallback``:VML(老版本 Word / LibreOffice 回退)

DrawingML 用 ``wp:anchor`` + ``mso-position-*-relative:page`` 绝对定位,
``a:bodyPr`` 的 ``lIns/tIns/rIns/bIns=0`` 消除内部 padding,
``anchor="t"`` 顶对齐,确保文字相对文本框左上角无偏移。

单位:PDF point → EMU(1 pt = 12700 EMU)。
"""
from __future__ import annotations
from typing import Optional, Sequence
from docx.oxml.ns import nsmap
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from lxml import etree

# 注册命名空间
_NS = {
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
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
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

EMU_PER_PT = 12700
_id_counter = [100]


def _next_id() -> int:
    _id_counter[0] += 1
    return _id_counter[0]


def pt_to_emu(pt: float) -> int:
    return int(round(pt * EMU_PER_PT))


def _pt(v: float) -> str:
    return f"{v:.3f}pt"


def _el(tag: str):
    return OxmlElement(tag)


def _make_run_props_xml(font_latin: str, font_ea: str, size: float,
                        color: str, bold: bool, italic: bool,
                        underline: bool, strike: bool = False) -> str:
    parts = [f'<w:rFonts w:ascii="{font_latin}" w:hAnsi="{font_latin}" '
             f'w:eastAsia="{font_ea}" w:cs="{font_latin}"/>',
             f'<w:color w:val="{color}"/>',
             f'<w:sz w:val="{int(round(size * 2))}"/>',
             f'<w:szCs w:val="{int(round(size * 2))}"/>']
    if bold: parts.append('<w:b/><w:bCs/>')
    if italic: parts.append('<w:i/><w:iCs/>')
    if underline: parts.append('<w:u w:val="single"/>')
    if strike: parts.append('<w:strike/>')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


def _escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _make_runs_xml(runs: Sequence[dict]) -> str:
    out = ""
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
        text = _escape(run.get("text", ""))
        rel_id = run.get("hyperlink_rel_id")
        if rel_id:
            if rel_id.startswith("__internal__:"):
                anchor = rel_id[len("__internal__:"):]
                out += (f'<w:hyperlink w:anchor="{anchor}">'
                        f'<w:r>{rPr}<w:t xml:space="preserve">{text}</w:t></w:r>'
                        f'</w:hyperlink>')
            else:
                out += (f'<w:hyperlink r:id="{rel_id}">'
                        f'<w:r>{rPr}<w:t xml:space="preserve">{text}</w:t></w:r>'
                        f'</w:hyperlink>')
        else:
            out += f'<w:r>{rPr}<w:t xml:space="preserve">{text}</w:t></w:r>'
    return out


def _make_txbx_content(runs: Sequence[dict], align: str, line_spacing: float) -> str:
    jc = "" if align in ("left",) else f'<w:jc w:val="{align}"/>'
    line_val = int(240 * line_spacing)
    pPr = (f'<w:pPr><w:spacing w:line="{line_val}" w:lineRule="auto" '
           f'w:before="0" w:after="0" w:beforeLines="0" w:afterLines="0"/>'
           f'{jc}<w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>')
    return f'<w:txbxContent xmlns:w="{W_NS}"><w:p>{pPr}{_make_runs_xml(runs)}</w:p></w:txbxContent>'


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
    """添加绝对定位文本框(mc:AlternateContent:DrawingML + VML 回退)。

    DrawingML 用 a:bodyPr lIns/tIns/rIns/bIns=0 消除内部 padding,
    anchor="t" 顶对齐,确保文字相对文本框左上角无偏移(WPS/Word 精确定位)。
    VML 回退用 inset=0,供 LibreOffice 使用。
    """
    _id = _next_id()
    z_index = z if not behind else -z
    txbx = _make_txbx_content(runs, align, line_spacing)
    v_anchor = {"top": "t", "center": "ctr", "bottom": "b"}.get(vertical_align, "t")
    cx = pt_to_emu(max(w, 0.1))
    cy = pt_to_emu(max(h, 0.1))
    px = pt_to_emu(x)
    py = pt_to_emu(y)

    # ===== DrawingML (Choice) =====
    dml = f'''<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}" xmlns:r="{R_NS}">
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{z_index}" behindDoc="{'1' if behind else '0'}" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>{px}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>{py}</wp:posOffset></wp:positionV>
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="{_id}" name="TextBox{_id}"/>
<wp:cNvGraphicFramePr/>
<a:graphic>
<a:graphicData uri="{WPS_NS}">
<wps:wsp>
<wps:cNvSpPr txBox="1"/>
<wps:cNvCnPr/>
<wps:spPr>
<a:xfrm{" rotation=\"" + str(int(round(-rotation * 60000))) + "\"" if abs(rotation) > 0.5 else ""}>
<a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/>
</a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
{"<a:solidFill><a:srgbClr val=\"" + fill_color + "\"/></a:solidFill>" if fill_color else "<a:noFill/>"}
{"<a:ln w=\"" + str(pt_to_emu(0.75)) + "\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>" if border else "<a:ln><a:noFill/></a:ln>"}
</wps:spPr>
<wps:txbx>
{txbx}
</wps:txbx>
<wps:bodyPr rot="0" vert="horz" anchor="{v_anchor}" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/>
</wps:wsp>
</a:graphicData>
</a:graphic>
</wp:anchor>
</w:drawing>'''

    # ===== VML (Fallback) =====
    fill_attr = "false" if not fill_color else "true"
    fill_xml = f'<v:fill color="#{fill_color}"/>' if fill_color else '<v:fill on="false"/>'
    stroke_xml = f'<v:stroke on="true" color="#000000" weight="0.75pt"/>' if border else '<v:stroke on="false"/>'
    rot_attr = f' rotation="{-rotation:.2f}"' if abs(rotation) > 0.5 else ''
    style = (f"position:absolute;left:{_pt(x)};top:{_pt(y)};"
             f"width:{_pt(max(w,0.1))};height:{_pt(max(h,0.1))};"
             f"z-index:{z_index};"
             f"mso-position-horizontal:absolute;mso-position-vertical:absolute;"
             f"mso-position-horizontal-relative:page;mso-position-vertical-relative:page;"
             f"mso-wrap:square;")
    vml = f'''<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}" xmlns:o="{O_NS}" xmlns:w10="{W10_NS}" xmlns:r="{R_NS}">
<v:shape id="VML{_id}" type="#_x0000_t202"{rot_attr} style="{style}" filled="{fill_attr}" stroked="{'true' if border else 'false'}" v-text-anchor="{ {"top":"top","center":"middle","bottom":"bottom"}.get(vertical_align,"top") }">
{fill_xml}{stroke_xml}
<v:textbox style="mso-fit-shape-to-text:false" inset="0pt,0pt,0pt,0pt">
{txbx}
</v:textbox>
</v:shape>
</w:pict>'''

    # ===== mc:AlternateContent 包裹 =====
    ac_xml = f'''<mc:AlternateContent xmlns:mc="{MC_NS}">
<mc:Choice xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}" Requires="wps">
{dml}
</mc:Choice>
<mc:Fallback xmlns:w="{W_NS}">
{vml}
</mc:Fallback>
</mc:AlternateContent>'''
    r = etree.SubElement(paragraph._p, "{%s}r" % W_NS)
    r.append(etree.fromstring(ac_xml))
    return r


def add_line(
    paragraph: Paragraph,
    x1: float, y1: float, x2: float, y2: float,
    color: str = "000000",
    width: float = 0.75,
    behind: bool = True,
    z: int = 0,
):
    """添加直线(DrawingML + VML 回退)。水平/垂直线用细长矩形,斜线用 line。"""
    z_index = z if not behind else -z
    dx = x2 - x1
    dy = y2 - y1
    is_horizontal = abs(dy) < 0.5 and abs(dx) > 0.5
    is_vertical = abs(dx) < 0.5 and abs(dy) > 0.5
    _id = _next_id()

    if is_horizontal or is_vertical:
        x = min(x1, x2); y = min(y1, y2)
        w = abs(dx) if is_horizontal else width
        h = width if is_horizontal else abs(dy)
        cx = pt_to_emu(max(w, 0.1)); cy = pt_to_emu(max(h, 0.1))
        px = pt_to_emu(x); py = pt_to_emu(y)
        # DrawingML: 填充矩形无线条
        dml = f'''<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}" xmlns:r="{R_NS}">
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{z_index}" behindDoc="{'1' if behind else '0'}" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>{px}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>{py}</wp:posOffset></wp:positionV>
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="{_id}" name="Line{_id}"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="{WPS_NS}"><wps:wsp>
<wps:cNvSpPr/><wps:cNvCnPr/>
<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
<a:ln><a:noFill/></a:ln>
</wps:spPr>
<wps:txbx><w:txbxContent xmlns:w="{W_NS}"><w:p/></w:txbxContent></wps:txbx>
<wps:bodyPr anchor="t" lIns="0" tIns="0" rIns="0" bIns="0"/>
</wps:wsp></a:graphicData></a:graphic>
</wp:anchor></w:drawing>'''
        vml_style = (f"position:absolute;left:{_pt(x)};top:{_pt(y)};"
                     f"width:{_pt(max(w,0.1))};height:{_pt(max(h,0.1))};z-index:{z_index};"
                     f"mso-position-horizontal:absolute;mso-position-vertical:absolute;"
                     f"mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap:none;")
        vml = f'''<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}"><v:rect id="VL{_id}" style="{vml_style}" filled="true" stroked="false"><v:fill color="#{color}"/><v:stroke on="false"/></v:rect></w:pict>'''
    else:
        # 斜线:VML v:line(DrawingML 用 line prstGeom)
        cx = pt_to_emu(abs(dx)); cy = pt_to_emu(abs(dy))
        px = pt_to_emu(min(x1,x2)); py = pt_to_emu(min(y1,y2))
        import math
        angle = int(math.degrees(math.atan2(dy, dx)) * 60000)
        dml = f'''<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}" xmlns:r="{R_NS}">
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{z_index}" behindDoc="{'1' if behind else '0'}" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>{px}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>{py}</wp:posOffset></wp:positionV>
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="{_id}" name="Line{_id}"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="{WPS_NS}"><wps:wsp>
<wps:cNvSpPr/><wps:cNvCnPr/>
<wps:spPr><a:xfrm rotation="{angle}"><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{pt_to_emu(max(width,0.5))}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
<a:ln><a:noFill/></a:ln>
</wps:spPr>
<wps:txbx><w:txbxContent xmlns:w="{W_NS}"><w:p/></w:txbxContent></wps:txbx>
<wps:bodyPr anchor="t" lIns="0" tIns="0" rIns="0" bIns="0"/>
</wps:wsp></a:graphicData></a:graphic>
</wp:anchor></w:drawing>'''
        px_v = x1 * 4.0 / 3.0; py_v = y1 * 4.0 / 3.0
        qx_v = x2 * 4.0 / 3.0; qy_v = y2 * 4.0 / 3.0
        vml_style = (f"position:absolute;left:0;top:0;z-index:{z_index};"
                     f"mso-position-horizontal:absolute;mso-position-vertical:absolute;"
                     f"mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap:none;")
        vml = f'''<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}"><v:line id="VL{_id}" style="{vml_style}" from="{px_v:.2f},{py_v:.2f}" to="{qx_v:.2f},{qy_v:.2f}" strokecolor="#{color}" strokeweight="{_pt(width)}"/></w:pict>'''

    ac_xml = f'''<mc:AlternateContent xmlns:mc="{MC_NS}">
<mc:Choice xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}" Requires="wps">{dml}</mc:Choice>
<mc:Fallback xmlns:w="{W_NS}">{vml}</mc:Fallback>
</mc:AlternateContent>'''
    r = etree.SubElement(paragraph._p, "{%s}r" % W_NS)
    r.append(etree.fromstring(ac_xml))
    return r


def add_rect(
    paragraph: Paragraph,
    x: float, y: float, w: float, h: float,
    stroke_color: Optional[str] = None,
    fill_color: Optional[str] = None,
    stroke_width: float = 0.75,
    behind: bool = True,
    z: int = 0,
):
    """添加矩形(DrawingML + VML 回退)。"""
    z_index = z if not behind else -z
    _id = _next_id()
    cx = pt_to_emu(max(w, 0.1)); cy = pt_to_emu(max(h, 0.1))
    px = pt_to_emu(x); py = pt_to_emu(y)
    fill_xml_dml = f'<a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>' if fill_color else '<a:noFill/>'
    stroke_xml_dml = (f'<a:ln w="{pt_to_emu(max(stroke_width,0.1))}"><a:solidFill><a:srgbClr val="{stroke_color}"/></a:solidFill></a:ln>'
                      if stroke_color else '<a:ln><a:noFill/></a:ln>')
    dml = f'''<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}" xmlns:r="{R_NS}">
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{z_index}" behindDoc="{'1' if behind else '0'}" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>{px}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>{py}</wp:posOffset></wp:positionV>
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="{_id}" name="Rect{_id}"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="{WPS_NS}"><wps:wsp>
<wps:cNvSpPr/><wps:cNvCnPr/>
<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
{fill_xml_dml}{stroke_xml_dml}
</wps:spPr>
<wps:txbx><w:txbxContent xmlns:w="{W_NS}"><w:p/></w:txbxContent></wps:txbx>
<wps:bodyPr anchor="t" lIns="0" tIns="0" rIns="0" bIns="0"/>
</wps:wsp></a:graphicData></a:graphic>
</wp:anchor></w:drawing>'''
    fill_xml_vml = f'<v:fill color="#{fill_color}"/>' if fill_color else '<v:fill on="false"/>'
    stroke_xml_vml = f'<v:stroke on="true" color="#{stroke_color}" weight="{_pt(stroke_width)}"/>' if stroke_color else '<v:stroke on="false"/>'
    style = (f"position:absolute;left:{_pt(x)};top:{_pt(y)};"
             f"width:{_pt(max(w,0.1))};height:{_pt(max(h,0.1))};z-index:{z_index};"
             f"mso-position-horizontal:absolute;mso-position-vertical:absolute;"
             f"mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap:none;")
    vml = f'''<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}"><v:rect id="VR{_id}" style="{style}" filled="{'true' if fill_color else 'false'}" stroked="{'true' if stroke_color else 'false'}">{fill_xml_vml}{stroke_xml_vml}</v:rect></w:pict>'''
    ac_xml = f'''<mc:AlternateContent xmlns:mc="{MC_NS}">
<mc:Choice xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}" Requires="wps">{dml}</mc:Choice>
<mc:Fallback xmlns:w="{W_NS}">{vml}</mc:Fallback>
</mc:AlternateContent>'''
    r = etree.SubElement(paragraph._p, "{%s}r" % W_NS)
    r.append(etree.fromstring(ac_xml))
    return r


def add_image(
    paragraph: Paragraph,
    x: float, y: float, w: float, h: float,
    image_data: bytes,
    rel_id: str,
    behind: bool = False,
    z: int = 1,
):
    """添加图片(DrawingML + VML 回退)。"""
    z_index = z if not behind else -z
    _id = _next_id()
    cx = pt_to_emu(max(w, 0.1)); cy = pt_to_emu(max(h, 0.1))
    px = pt_to_emu(x); py = pt_to_emu(y)
    PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    dml = f'''<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}" xmlns:r="{R_NS}">
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{z_index}" behindDoc="{'1' if behind else '0'}" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>{px}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>{py}</wp:posOffset></wp:positionV>
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="{_id}" name="Pic{_id}"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="{PIC_NS}">
<pic:pic xmlns:pic="{PIC_NS}">
<pic:nvPicPr><pic:cNvPr id="{_id}" name="Pic{_id}"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic>
</a:graphicData></a:graphic>
</wp:anchor></w:drawing>'''
    style = (f"position:absolute;left:{_pt(x)};top:{_pt(y)};"
             f"width:{_pt(max(w,0.1))};height:{_pt(max(h,0.1))};z-index:{z_index};"
             f"mso-position-horizontal:absolute;mso-position-vertical:absolute;"
             f"mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap:none;")
    vml = f'''<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}" xmlns:r="{R_NS}"><v:rect id="VP{_id}" style="{style}" filled="false" stroked="false"><v:fill on="false"/><v:stroke on="false"/><v:imagedata r:embed="{rel_id}"/></v:rect></w:pict>'''
    ac_xml = f'''<mc:AlternateContent xmlns:mc="{MC_NS}">
<mc:Choice xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" Requires="wps">{dml}</mc:Choice>
<mc:Fallback xmlns:w="{W_NS}">{vml}</mc:Fallback>
</mc:AlternateContent>'''
    r = etree.SubElement(paragraph._p, "{%s}r" % W_NS)
    r.append(etree.fromstring(ac_xml))
    return r
