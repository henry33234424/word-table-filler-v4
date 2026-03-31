"""
Excel-to-Word Template Filler -- Word表格模板填报工具
"""

import sys
import os

# 确保可以导入 src 模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.app import App


def main():
    app = App()
    app.run()


if __name__ == '__main__':
    main()
