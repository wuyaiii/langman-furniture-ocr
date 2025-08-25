# -*- coding: utf-8 -*-
"""
屏幕截图和区域选择模块
"""

import tkinter as tk
from PIL import Image, ImageGrab
from datetime import datetime
from pathlib import Path
from ..utils import logger, config_manager, dpi_helper


class ScreenCapture:
    """屏幕截图和区域选择器"""

    def __init__(self, parent_window):
        self.parent_window = parent_window
        self.selection_window = None
        self.canvas = None
        self.start_x = None
        self.start_y = None
        self.end_x = None
        self.end_y = None
        self.captured_image = None
        self.selection_coords = self._load_selection_coordinates()

        # 创建调试图片目录
        self.debug_dir = Path(config_manager.get(
            'debug_image_dir', 'debug_images'))
        self.debug_dir.mkdir(exist_ok=True)

        # 获取DPI缩放信息
        self.dpi_scale = dpi_helper.get_dpi_scale()
        logger.info(f"屏幕截图模块初始化完成，DPI缩放: {self.dpi_scale}")

    def _load_selection_coordinates(self):
        """加载选框坐标"""
        return config_manager.get("selection_coordinates", {"x1": 0, "y1": 0, "x2": 0, "y2": 0})

    def _save_selection_coordinates(self, x1, y1, x2, y2):
        """保存选框坐标"""
        coords = {"x1": x1, "y1": y1, "x2": x2, "y2": y2}
        config_manager.set("selection_coordinates", coords)
        self.selection_coords = coords
        logger.info(f"保存选框坐标: {coords}")

    def has_valid_selection(self):
        """检查是否有有效的选框坐标"""
        coords = self.selection_coords
        return (coords["x1"] != coords["x2"] and coords["y1"] != coords["y2"] and
                coords["x1"] != 0 and coords["y1"] != 0)

    def start_screen_capture(self):
        """开始屏幕区域选择"""
        logger.info("开始屏幕区域选择")
        self.parent_window.withdraw()  # 隐藏主窗口
        self._create_selection_window()
        return True  # 选择窗口创建成功

    def _create_selection_window(self):
        """创建选择窗口"""
        self.selection_window = tk.Toplevel()
        self.selection_window.attributes('-fullscreen', True)
        self.selection_window.attributes('-alpha', 0.3)
        self.selection_window.configure(bg='black')
        self.selection_window.attributes('-topmost', True)

        # 创建画布
        self.canvas = tk.Canvas(self.selection_window,
                                highlightthickness=0, bg='black')
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 绑定事件
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.selection_window.bind("<Escape>", self._cancel_selection)
        self.selection_window.focus_set()

        # 添加提示文字
        screen_width = self.selection_window.winfo_screenwidth()
        self.canvas.create_text(screen_width//2, 50, text="拖拽鼠标选择区域，ESC键取消",
                                fill="white", font=("微软雅黑", 16))

    def _on_click(self, event):
        """鼠标点击事件"""
        self.start_x = event.x
        self.start_y = event.y
        logger.debug(f"鼠标点击: ({self.start_x}, {self.start_y})")

    def _on_drag(self, event):
        """鼠标拖拽事件"""
        if self.start_x is not None and self.start_y is not None:
            # 清除之前的选择框
            self.canvas.delete("selection")

            # 绘制新的选择框
            self.canvas.create_rectangle(
                self.start_x, self.start_y, event.x, event.y,
                outline="red", width=2, tags="selection"
            )

    def _on_release(self, event):
        """鼠标释放事件"""
        self.end_x = event.x
        self.end_y = event.y

        # 确保坐标正确
        x1 = min(self.start_x, self.end_x)
        y1 = min(self.start_y, self.end_y)
        x2 = max(self.start_x, self.end_x)
        y2 = max(self.start_y, self.end_y)

        logger.info(f"选择区域: ({x1}, {y1}) - ({x2}, {y2})")

        # 检查选择区域是否有效
        if abs(x2 - x1) > 10 and abs(y2 - y1) > 10:
            # 转换为屏幕坐标
            screen_x1 = self.selection_window.winfo_rootx() + x1
            screen_y1 = self.selection_window.winfo_rooty() + y1
            screen_x2 = self.selection_window.winfo_rootx() + x2
            screen_y2 = self.selection_window.winfo_rooty() + y2

            logger.info(
                f"选择区域: ({screen_x1}, {screen_y1}) - ({screen_x2}, {screen_y2})")

            self._save_selection_coordinates(
                screen_x1, screen_y1, screen_x2, screen_y2)
            self._capture_selected_area(
                screen_x1, screen_y1, screen_x2, screen_y2)

        else:
            logger.warning("选择区域太小")
            self._cancel_selection()

    def _capture_selected_area(self, x1, y1, x2, y2):
        """截图选中区域"""
        try:
            # 关闭选择窗口
            if self.selection_window:
                self.selection_window.destroy()
                self.selection_window = None

            # 截图
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            self.captured_image = screenshot

            # 保存调试图片
            if config_manager.get('save_debug_images', True):
                self._save_debug_image(screenshot)

            # 显示主窗口
            self.parent_window.deiconify()

            logger.info(f"截图完成，区域大小: {x2-x1}x{y2-y1} 像素")
            return True

        except Exception as e:
            logger.error(f"截图失败: {e}")
            self._cancel_selection()
            return False

    def _save_debug_image(self, image):
        """保存调试图片"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"capture_{timestamp}.png"
            filepath = self.debug_dir / filename
            image.save(filepath)
            logger.info(f"调试图片已保存: {filepath}")
        except Exception as e:
            logger.error(f"保存调试图片失败: {e}")

    def _cancel_selection(self, event=None):
        """取消选择"""
        if self.selection_window:
            self.selection_window.destroy()
            self.selection_window = None
        self.parent_window.deiconify()
        logger.info("取消屏幕选择")

    def capture_current_selection(self):
        """根据当前选框坐标重新截图"""
        if not self.has_valid_selection():
            logger.error("没有有效的选框区域")
            return None

        coords = self.selection_coords
        x1, y1, x2, y2 = coords["x1"], coords["y1"], coords["x2"], coords["y2"]

        # 确保坐标顺序正确
        left = min(x1, x2)
        top = min(y1, y2)
        right = max(x1, x2)
        bottom = max(y1, y2)

        try:
            # 截图选中区域
            screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
            self.captured_image = screenshot

            # 保存调试图片
            if config_manager.get('save_debug_images', True):
                self._save_debug_image(screenshot)

            logger.info(f"重新截图完成: ({left}, {top}) - ({right}, {bottom})")
            return screenshot

        except Exception as e:
            logger.error(f"重新截图失败: {e}")
            return None

    def get_captured_image(self):
        """获取截图图像"""
        return self.captured_image

    def get_selection_info(self):
        """获取选择区域信息"""
        if not self.has_valid_selection():
            return None

        coords = self.selection_coords
        x1, y1, x2, y2 = coords["x1"], coords["y1"], coords["x2"], coords["y2"]
        width = abs(x2 - x1)
        height = abs(y2 - y1)

        return {
            'coordinates': coords,
            'size': (width, height),
            'area': width * height
        }
