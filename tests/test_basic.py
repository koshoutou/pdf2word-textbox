# -*- coding: utf-8 -*-
"""基础冒烟测试:验证转换器能正常运行。

运行:
    python -m pytest tests/ -v
或:
    python tests/test_basic.py
"""
import os
import sys
import tempfile

# 让 tests 能找到包
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _make_test_pdf(path: str):
    """生成一个最小测试 PDF。"""
    import fitz
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4
    # 页眉
    page.insert_text((72, 50), "页眉文本", fontsize=10, fontname="helv")
    # 页眉横线
    page.draw_line((72, 60), (523, 60), color=(0, 0, 0), width=0.5)
    # 正文
    page.insert_text((72, 100), "Hello pdf2word-textbox!", fontsize=14, fontname="helv")
    page.insert_text((72, 130), "第二行文本内容", fontsize=12, fontname="helv")
    # 矩形
    page.draw_rect((72, 160, 200, 200), color=(1, 0, 0), width=1)
    # 页脚页码
    page.insert_text((280, 800), "1", fontsize=10, fontname="helv")
    doc.save(path)
    doc.close()


def test_basic_conversion():
    """测试基本转换:页数一致、文件生成。"""
    from pdf2word_textbox import Converter

    with tempfile.TemporaryDirectory() as td:
        pdf = os.path.join(td, "test.pdf")
        docx = os.path.join(td, "test.docx")
        _make_test_pdf(pdf)

        cv = Converter(pdf)
        cv.convert(docx, use_real_header_footer=False)

        assert os.path.exists(docx), "docx 未生成"
        assert os.path.getsize(docx) > 0, "docx 为空"

        # 检查页数
        from docx import Document
        d = Document(docx)
        assert len(d.sections) == 1, f"页数应为1,实际{len(d.sections)}"
        print("✓ 基础转换测试通过")


def test_multi_page():
    """测试多页转换。"""
    import fitz
    from pdf2word_textbox import Converter

    with tempfile.TemporaryDirectory() as td:
        pdf = os.path.join(td, "multi.pdf")
        docx = os.path.join(td, "multi.docx")
        # 生成3页PDF
        doc = fitz.open()
        for i in range(3):
            p = doc.new_page()
            p.insert_text((72, 100), f"Page {i+1}", fontsize=14)
        doc.save(pdf)
        doc.close()

        cv = Converter(pdf)
        cv.convert(docx, use_real_header_footer=False)

        from docx import Document
        d = Document(docx)
        assert len(d.sections) == 3, f"页数应为3,实际{len(d.sections)}"
        print("✓ 多页转换测试通过")


if __name__ == "__main__":
    test_basic_conversion()
    test_multi_page()
    print("\n全部测试通过 ✓")
