# -*- coding: utf-8 -*-
"""pdf2word_textbox: 基于文本框一比一精准复刻的 PDF 转 Word 工具。

通过在 DOCX 中使用绝对定位的文本框(text box)和形状(shape),
实现 PDF 页面元素的 1:1 位置复刻,包括:
- 文本(字体、字号、颜色、粗体/斜体/下划线)
- 矢量图形(直线、矩形、曲线)
- 图片
- 页眉页脚(含横线、对齐、分割线)
- 页码
- 表格(按单元格文本框复刻)
- 数学公式(按文本框复刻)
"""

from .converter import Converter

__version__ = "1.0.0"
__all__ = ["Converter"]
__author__ = "koshoutou (Inkcoo) <admin@inkcoo.com>"
