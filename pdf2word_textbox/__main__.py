"""模块入口,支持 python -m pdf2word_textbox 调用。"""
from .cli import main
import sys
sys.exit(main())
