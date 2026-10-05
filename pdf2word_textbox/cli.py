# -*- coding: utf-8 -*-
"""命令行入口。

用法::

    python -m pdf2word_textbox input.pdf output.docx
    python -m pdf2word_textbox input.pdf output.docx --start 0 --end 5
    python -m pdf2word_textbox input.pdf output.docx --no-real-header-footer
"""
from __future__ import annotations
import argparse
import sys
from .converter import Converter


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="pdf2word_textbox",
        description="PDF 转 Word 工具 - 基于文本框一比一精准复刻",
    )
    parser.add_argument("pdf", help="输入 PDF 文件路径")
    parser.add_argument("docx", help="输出 DOCX 文件路径")
    parser.add_argument("--password", default="", help="PDF 密码(若有)")
    parser.add_argument("--start", type=int, default=0, help="起始页(0-based,默认0)")
    parser.add_argument("--end", type=int, default=None, help="结束页(不含),默认到末尾")
    parser.add_argument("--header-ratio", type=float, default=0.12,
                        help="页眉区域占页面高度比例(默认0.12)")
    parser.add_argument("--footer-ratio", type=float, default=0.12,
                        help="页脚区域占页面高度比例(默认0.12)")
    parser.add_argument("--no-real-header-footer", action="store_true",
                        help="不使用 docx 真实页眉页脚(全部放正文文本框)")
    parser.add_argument("--no-detect-underline", action="store_true",
                        help="关闭下划线自动检测")
    args = parser.parse_args(argv)

    cv = Converter(args.pdf, password=args.password)
    cv.convert(
        args.docx,
        start=args.start,
        end=args.end,
        header_ratio=args.header_ratio,
        footer_ratio=args.footer_ratio,
        use_real_header_footer=not args.no_real_header_footer,
        detect_underline=not args.no_detect_underline,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
