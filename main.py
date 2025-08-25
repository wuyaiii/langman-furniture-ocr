# -*- coding: utf-8 -*-
"""
屏幕截图OCR工具 - 主程序入口
"""

import sys
import os
from pathlib import Path

# 添加src目录到Python路径
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from src.ui import MainWindow
from src.utils import logger


def main():
    """主函数"""
    try:
        logger.info("=" * 50)
        logger.info("屏幕截图OCR工具启动")
        logger.info("=" * 50)
        
        # 创建并运行主界面
        app = MainWindow()
        app.run()
        
    except Exception as e:
        logger.error(f"程序启动失败: {e}")
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("启动错误", f"程序启动失败:\n{str(e)}")
        root.destroy()
        sys.exit(1)
    
    finally:
        logger.info("程序已退出")


if __name__ == "__main__":
    main()
