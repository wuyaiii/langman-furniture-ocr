# -*- coding: utf-8 -*-
"""
主界面模块
"""

import tkinter as tk
from tkinter import messagebox, ttk
from ..core import ScreenCapture, OCRProcessor, ExcelManager, DataSorter
from ..utils import logger, config_manager, dpi_helper


class MainWindow:
    """主界面窗口"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("屏幕截图OCR工具 v2.0")

        # 初始化DPI感知
        dpi_helper.get_dpi_scale()

        # 设置窗口大小并居中显示
        self._setup_window()

        # 初始化核心组件
        self.screen_capture = ScreenCapture(self.root)
        self.ocr_processor = OCRProcessor()
        self.excel_manager = ExcelManager()
        self.data_sorter = DataSorter(self.excel_manager)

        # OCR结果
        self.extracted_title = None
        self.extracted_item_names = []

        # 选框边框窗口
        self.selection_border_window = None

        # 创建界面
        self._create_widgets()

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        # 应用初始设置
        self._apply_initial_settings()

        logger.info("主界面初始化完成")

    def _setup_window(self):
        """设置窗口属性"""
        window_width = 423
        window_height = 557
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(
            f'{window_width}x{window_height}+{center_x}+{center_y}')

    def _create_widgets(self):
        """创建主界面控件"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 标题
        title_label = ttk.Label(
            main_frame, text="屏幕截图OCR工具", font=("微软雅黑", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # 说明文字
        info_label = ttk.Label(
            main_frame, text="点击下方按钮开始屏幕区域选择", font=("微软雅黑", 10))
        info_label.pack(pady=(0, 10))

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        # 截图选择按钮
        self.capture_btn = ttk.Button(
            button_frame, text="开始屏幕选择", command=self._start_screen_capture)
        self.capture_btn.pack(fill=tk.X, pady=5)

        # 识别按钮
        self.recognize_btn = ttk.Button(
            button_frame, text="识别选中区域", command=self._recognize_screen_area, state=tk.DISABLED)
        self.recognize_btn.pack(fill=tk.X, pady=5)

        # Excel操作按钮
        excel_frame = ttk.Frame(button_frame)
        excel_frame.pack(fill=tk.X, pady=5)

        ttk.Button(excel_frame, text="打开Excel文件", command=self._prepare_excel_file).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(excel_frame, text="Excel排序处理", command=self._sort_excel_data).pack(
            side=tk.LEFT, padx=(0, 10))

        # 选项设置
        self._create_option_widgets(button_frame)

        # 状态显示
        self.status_label = ttk.Label(
            main_frame, text="准备就绪", foreground="green")
        self.status_label.pack(pady=(20, 0))

        # 结果显示区域
        self._create_result_area(main_frame)

    def _create_option_widgets(self, parent):
        """创建选项控件"""
        # 自动准备Excel选项
        self.auto_open_excel = tk.BooleanVar(
            value=config_manager.get("auto_open_excel", False))
        auto_open_check = ttk.Checkbutton(parent, text="启动时自动打开Excel",
                                          variable=self.auto_open_excel,
                                          command=self._save_auto_open_setting)
        auto_open_check.pack(fill=tk.X, pady=2)

        # 窗口控制选项
        window_frame = ttk.Frame(parent)
        window_frame.pack(fill=tk.X, pady=5)

        # 置顶按钮
        self.topmost = tk.BooleanVar(
            value=config_manager.get("window_topmost", False))
        topmost_check = ttk.Checkbutton(window_frame, text="窗口置顶",
                                        variable=self.topmost,
                                        command=self._toggle_topmost)
        topmost_check.pack(side=tk.LEFT, padx=(0, 10))

        # 选框显示按钮
        self.show_selection_border = tk.BooleanVar(
            value=config_manager.get("show_selection_border", False))
        border_check = ttk.Checkbutton(window_frame, text="显示选框",
                                       variable=self.show_selection_border,
                                       command=self._toggle_selection_border)
        border_check.pack(side=tk.LEFT)

        # 确认选项
        self.show_confirmation = tk.BooleanVar(
            value=config_manager.get("show_confirmation", True))
        confirm_check = ttk.Checkbutton(parent, text="识别后显示确认对话框",
                                        variable=self.show_confirmation,
                                        command=self._save_confirmation_setting)
        confirm_check.pack(fill=tk.X, pady=2)

    def _create_result_area(self, parent):
        """创建结果显示区域"""
        result_frame = ttk.LabelFrame(parent, text="识别结果", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.result_text = tk.Text(
            result_frame, height=8, wrap=tk.WORD, font=("微软雅黑", 9))
        result_scrollbar = ttk.Scrollbar(
            result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=result_scrollbar.set)

        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _apply_initial_settings(self):
        """应用初始设置"""
        # 应用置顶设置
        if self.topmost.get():
            self.root.attributes('-topmost', True)

        # 验证配置
        valid, errors = config_manager.validate_env_config()
        if not valid:
            self.status_label.config(
                text=f"配置错误: {', '.join(errors)}", foreground="red")

            error_message = f"环境变量配置错误，请检查是否存在.env文件:\n\n" + \
                "\n".join(f"• {error}" for error in errors)
            error_message += f"\n\n请参考.env.example文件进行配置。\n点击确定后程序将关闭。"

            messagebox.showerror("配置错误", error_message)

            # 关闭程序
            logger.error("配置验证失败，程序退出")
            self.root.destroy()
            return
        else:
            self.status_label.config(text="配置验证通过", foreground="green")

        # 如果勾选了自动打开Excel，则自动准备Excel文件
        if self.auto_open_excel.get():
            self.root.after(1000, self._prepare_excel_file)

        # 如果启动时需要显示选框，则显示
        if self.show_selection_border.get() and self.screen_capture.has_valid_selection():
            self.root.after(1000, self._show_selection_border_window)

    def _start_screen_capture(self):
        """开始屏幕选择"""
        self.status_label.config(text="请在屏幕上拖拽选择区域...", foreground="blue")
        success = self.screen_capture.start_screen_capture()

        if success:
            self.recognize_btn.config(state=tk.NORMAL)
            selection_info = self.screen_capture.get_selection_info()
            if selection_info:
                size = selection_info['size']
                self.status_label.config(
                    text=f"已选择区域: {size[0]}x{size[1]} 像素", foreground="green")
                self._display_capture_info(selection_info)
        else:
            self.status_label.config(text="屏幕选择失败", foreground="red")

    def _display_capture_info(self, selection_info):
        """显示截图信息"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "✅ 已截取屏幕区域\n")
        size = selection_info['size']
        coords = selection_info['coordinates']
        self.result_text.insert(tk.END, f"区域大小: {size[0]} x {size[1]} 像素\n")
        self.result_text.insert(
            tk.END, f"位置: ({coords['x1']}, {coords['y1']}) - ({coords['x2']}, {coords['y2']})\n\n")
        self.result_text.insert(tk.END, "点击'识别选中区域'开始OCR识别\n")

        # 如果启用了选框显示，则显示边框
        if self.show_selection_border.get():
            self._show_selection_border_window()

    def _recognize_screen_area(self):
        """识别截图区域"""
        if not self.screen_capture.has_valid_selection():
            messagebox.showerror("错误", "没有有效的选框区域，请先选择区域")
            return

        try:
            self.status_label.config(text="正在识别...", foreground="blue")
            self.root.update()

            # 重新截图
            image = self.screen_capture.capture_current_selection()
            if not image:
                messagebox.showerror("错误", "截图失败")
                return

            # OCR识别
            ocr_result = self.ocr_processor.recognize_table(image)
            if not ocr_result:
                messagebox.showerror("错误", "OCR识别失败")
                return

            # 提取标题和物品名
            self.extracted_title, self.extracted_item_names = self.ocr_processor.extract_title_and_items(
                ocr_result)

            # 显示结果
            self._display_recognition_results()

            # 更新状态
            self.status_label.config(text="识别完成", foreground="green")

            # 根据配置决定是否显示确认对话框
            if self.show_confirmation.get() and (self.extracted_title or self.extracted_item_names):
                self._show_edit_confirmation_dialog()
            else:
                self._write_to_excel_direct()

        except Exception as e:
            logger.error(f"识别过程失败: {e}")
            messagebox.showerror("错误", f"识别失败: {str(e)}")
            self.status_label.config(text="识别失败", foreground="red")

    def _display_recognition_results(self):
        """显示识别结果"""
        self.result_text.delete(1.0, tk.END)

        self.result_text.insert(tk.END, "🎯 OCR识别结果\n")
        self.result_text.insert(tk.END, "=" * 30 + "\n\n")

        if self.extracted_title:
            self.result_text.insert(
                tk.END, f"📋 标题: {self.extracted_title}\n\n")
        else:
            self.result_text.insert(tk.END, "📋 标题: 未找到\n\n")

        if self.extracted_item_names:
            self.result_text.insert(
                tk.END, f"📦 物品名列表 (共{len(self.extracted_item_names)}项):\n")
            for i, item in enumerate(self.extracted_item_names, 1):
                self.result_text.insert(tk.END, f"  {i}. {item}\n")
        else:
            self.result_text.insert(tk.END, "📦 物品名列表: 未找到\n")

    def _show_edit_confirmation_dialog(self):
        """显示编辑确认对话框"""
        if not self.extracted_title and not self.extracted_item_names:
            messagebox.showwarning("警告", "没有识别到标题和物品名")
            return

        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("确认识别结果")
        dialog.geometry("286x428")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        # 居中显示
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() +
                        100, self.root.winfo_rooty() + 100))

        # 主框架
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 说明标签
        info_label = ttk.Label(
            main_frame, text="请确认或修改识别结果，然后选择操作：", font=("微软雅黑", 10))
        info_label.pack(anchor=tk.W, pady=(0, 10))

        # 标题部分
        title_frame = ttk.LabelFrame(main_frame, text="标题", padding=5)
        title_frame.pack(fill=tk.X, pady=(0, 5))

        self.title_var = tk.StringVar(value=self.extracted_title or "")
        title_entry = ttk.Entry(
            title_frame, textvariable=self.title_var, font=("微软雅黑", 9))
        title_entry.pack(fill=tk.X)

        # 物品名部分
        items_frame = ttk.LabelFrame(main_frame, text="物品名列表", padding=5)
        items_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        # 创建可编辑的文本框
        self.items_text = tk.Text(items_frame, font=(
            "微软雅黑", 9), height=12, wrap=tk.WORD)
        items_scrollbar = ttk.Scrollbar(
            items_frame, orient=tk.VERTICAL, command=self.items_text.yview)
        self.items_text.configure(yscrollcommand=items_scrollbar.set)

        # 填充物品名
        items_content = "\n".join(self.extracted_item_names)
        self.items_text.insert(tk.END, items_content)

        self.items_text.pack(side=tk.LEFT, fill=tk.BOTH,
                             expand=True, padx=(0, 2))
        items_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))

        # 按钮
        confirm_btn = ttk.Button(button_frame, text="写入Excel并继续",
                                 command=lambda: self._confirm_and_write_excel(dialog))
        confirm_btn.pack(side=tk.RIGHT, padx=(3, 0))

        cancel_btn = ttk.Button(button_frame, text="放弃",
                                command=dialog.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=(3, 0))

    def _confirm_and_write_excel(self, dialog):
        """确认并写入Excel"""
        # 获取修改后的数据
        modified_title = self.title_var.get().strip()

        # 从文本框获取物品名
        items_content = self.items_text.get("1.0", tk.END).strip()
        modified_items = [item.strip()
                          for item in items_content.split('\n') if item.strip()]

        # 更新提取的数据
        self.extracted_title = modified_title if modified_title else None
        self.extracted_item_names = modified_items

        # 关闭对话框
        dialog.destroy()

        # 写入Excel
        self._write_to_excel_direct()

    def _write_to_excel_direct(self):
        """直接写入Excel文件"""
        if not self.extracted_title and not self.extracted_item_names:
            messagebox.showwarning("警告", "没有数据可写入")
            return

        try:
            success = self.excel_manager.write_data(
                self.extracted_title, self.extracted_item_names)

            if success:
                self.status_label.config(
                    text="数据已写入Excel，可以继续选择新区域", foreground="green")
                logger.info("数据写入Excel成功")
            else:
                self.status_label.config(text="数据写入失败", foreground="red")
                messagebox.showerror("错误", "数据写入Excel失败")

        except Exception as e:
            logger.error(f"写入Excel失败: {e}")
            messagebox.showerror("错误", f"写入Excel失败: {str(e)}")

    def _prepare_excel_file(self):
        """准备Excel文件"""
        try:
            success = self.excel_manager.prepare_excel_file()
            if success:
                self.status_label.config(
                    text=f"Excel已打开: {self.excel_manager.excel_file_name}", foreground="green")
            else:
                self.status_label.config(text="Excel文件准备失败", foreground="red")
        except Exception as e:
            logger.error(f"准备Excel文件失败: {e}")
            messagebox.showerror("错误", f"准备Excel文件失败: {str(e)}")

    def _sort_excel_data(self):
        """Excel排序处理"""
        try:
            success = self.data_sorter.sort_excel_data()
            if success:
                self.status_label.config(
                    text="Excel排序处理完成", foreground="green")
                messagebox.showinfo("成功", "Excel数据排序处理完成！")
                # 重新打开Excel文件
                self.excel_manager._open_excel_file()
            else:
                self.status_label.config(text="排序处理失败", foreground="red")
                messagebox.showerror("错误", "Excel排序处理失败")
        except Exception as e:
            logger.error(f"排序处理失败: {e}")
            messagebox.showerror("错误", f"排序处理失败: {str(e)}")

    def _toggle_topmost(self):
        """切换窗口置顶状态"""
        self.root.attributes('-topmost', self.topmost.get())
        config_manager.set("window_topmost", self.topmost.get())

    def _toggle_selection_border(self):
        """切换选框显示状态"""
        config_manager.set("show_selection_border",
                           self.show_selection_border.get())

        if self.show_selection_border.get():
            if self.screen_capture.has_valid_selection():
                self._show_selection_border_window()
            else:
                messagebox.showwarning("警告", "没有有效的选框区域，请先选择区域")
                self.show_selection_border.set(False)
        else:
            self._hide_selection_border_window()

    def _show_selection_border_window(self):
        """显示选框边框窗口"""
        if not self.screen_capture.has_valid_selection():
            return

        # 如果已经有边框窗口，先关闭
        if self.selection_border_window:
            self._hide_selection_border_window()

        coords = self.screen_capture.selection_coords
        x1, y1, x2, y2 = coords["x1"], coords["y1"], coords["x2"], coords["y2"]

        # 确保坐标顺序正确
        left = min(x1, x2)
        top = min(y1, y2)
        right = max(x1, x2)
        bottom = max(y1, y2)
        width = right - left
        height = bottom - top

        # 创建透明的边框窗口
        self.selection_border_window = tk.Toplevel()
        self.selection_border_window.title("选框边框")
        self.selection_border_window.geometry(f"{width}x{height}+{left}+{top}")

        # 设置窗口属性
        self.selection_border_window.attributes('-topmost', True)
        self.selection_border_window.attributes('-transparentcolor', 'white')
        self.selection_border_window.overrideredirect(True)

        # 创建画布绘制边框
        canvas = tk.Canvas(self.selection_border_window, width=width, height=height,
                           bg='white', highlightthickness=0)
        canvas.pack()

        # 绘制红色边框
        border_width = 3
        canvas.create_rectangle(border_width//2, border_width//2,
                                width-border_width//2, height-border_width//2,
                                outline='red', width=border_width, fill='')

        logger.info(f"显示选框边框: ({left}, {top}) - ({right}, {bottom})")

    def _hide_selection_border_window(self):
        """隐藏选框边框窗口"""
        if self.selection_border_window:
            self.selection_border_window.destroy()
            self.selection_border_window = None
            logger.info("隐藏选框边框")

    def _save_auto_open_setting(self):
        """保存自动打开Excel设置"""
        config_manager.set("auto_open_excel", self.auto_open_excel.get())

    def _save_confirmation_setting(self):
        """保存确认对话框设置"""
        config_manager.set("show_confirmation", self.show_confirmation.get())

    def _on_closing(self):
        """程序关闭时的清理工作"""
        # 关闭选框边框窗口
        if self.selection_border_window:
            self._hide_selection_border_window()

        logger.info("程序正在关闭")
        self.root.destroy()

    def run(self):
        """运行主界面"""
        logger.info("启动主界面")
        self.root.mainloop()
