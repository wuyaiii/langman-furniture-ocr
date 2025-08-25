# -*- coding: utf-8 -*-
"""
DPI和显示缩放处理模块
"""

import ctypes
import ctypes.wintypes
import tkinter as tk
from ctypes import windll
from .logger import logger


class DPIHelper:
    """DPI和显示缩放处理器"""

    def __init__(self):
        self.dpi_scale = 1.0
        self.system_dpi = 96  # Windows默认DPI
        self.current_dpi = 96
        self._init_dpi_awareness()

    def _init_dpi_awareness(self):
        """初始化DPI感知"""
        try:
            # 设置DPI感知
            windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
            logger.info("已启用DPI感知")
        except Exception as e:
            logger.warning(f"设置DPI感知失败: {e}")
            try:
                # 备用方案
                windll.user32.SetProcessDPIAware()
                logger.info("已启用备用DPI感知")
            except Exception as e2:
                logger.error(f"备用DPI感知也失败: {e2}")

    def get_dpi_scale(self):
        """获取当前DPI缩放比例"""
        try:
            # 获取主显示器DPI
            hdc = windll.user32.GetDC(0)
            self.current_dpi = windll.gdi32.GetDeviceCaps(
                hdc, 88)  # LOGPIXELSX
            windll.user32.ReleaseDC(0, hdc)

            self.dpi_scale = self.current_dpi / self.system_dpi

            # 计算缩放百分比
            scale_percent = int(self.dpi_scale * 100)
            logger.info(
                f"检测到DPI: {self.current_dpi}, 缩放比例: {self.dpi_scale:.2f} ({scale_percent}%)")

            return self.dpi_scale
        except Exception as e:
            logger.error(f"获取DPI缩放比例失败: {e}")
            return 1.0

    def scale_coordinates(self, x, y, width, height):
        """根据DPI缩放调整坐标"""
        if self.dpi_scale == 1.0:
            return x, y, width, height

        # 对坐标进行缩放调整
        scaled_x = int(x / self.dpi_scale)
        scaled_y = int(y / self.dpi_scale)
        scaled_width = int(width / self.dpi_scale)
        scaled_height = int(height / self.dpi_scale)

        logger.info(
            f"坐标缩放: ({x}, {y}, {width}, {height}) -> ({scaled_x}, {scaled_y}, {scaled_width}, {scaled_height})")

        return scaled_x, scaled_y, scaled_width, scaled_height

    def get_screen_size(self):
        """获取真实屏幕尺寸（考虑DPI缩放）"""
        try:
            # 获取真实屏幕尺寸
            user32 = windll.user32
            screensize = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

            # 如果启用了DPI感知，这里得到的就是真实像素尺寸
            logger.info(f"屏幕尺寸: {screensize[0]}x{screensize[1]}")
            return screensize
        except Exception as e:
            logger.error(f"获取屏幕尺寸失败: {e}")
            # 备用方案：使用tkinter获取
            root = tk.Tk()
            width = root.winfo_screenwidth()
            height = root.winfo_screenheight()
            root.destroy()
            return width, height

    def get_display_info(self):
        """获取显示器详细信息"""
        info = {
            'dpi': self.current_dpi,
            'scale': self.dpi_scale,
            'screen_size': self.get_screen_size()
        }

        logger.info(f"显示器信息: {info}")
        return info


# 全局DPI助手实例
dpi_helper = DPIHelper()
