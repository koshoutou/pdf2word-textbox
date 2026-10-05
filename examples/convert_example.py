# -*- coding: utf-8 -*-
"""示例:用 pdf2word_textbox 把 PDF 转成 docx。

运行:
    python examples/convert_example.py
"""
from pdf2word_textbox import Converter


def main():
    pdf_file = "input.pdf"        # 替换为你的 PDF
    docx_file = "output.docx"

    cv = Converter(pdf_file)
    # 全量转换
    cv.convert(docx_file)
    print(f"已生成: {docx_file}")

    # 也可只转前 5 页做快速预览
    # cv.convert("preview.docx", start=0, end=5)


if __name__ == "__main__":
    main()
