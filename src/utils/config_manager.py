# -*- coding: utf-8 -*-
"""
配置管理模块
"""

import os
import json
from pathlib import Path
from .logger import logger

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    logger.warning("dotenv未安装，将直接读取环境变量")


class ConfigManager:
    """配置管理器"""

    def __init__(self, config_file="screen_ocr_config.json"):
        self.config_file = Path(config_file)
        self.config = self._load_config()
        self.env_config = self._load_env_config()

    def _load_config(self):
        """加载JSON配置文件"""
        default_config = {
            "auto_open_excel": False,
            "show_confirmation": True,
            "window_topmost": False,
            "show_selection_border": False,
            "selection_coordinates": {"x1": 0, "y1": 0, "x2": 0, "y2": 0},
            "save_debug_images": True,
            "debug_image_dir": "debug_images"
        }

        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 合并默认配置和用户配置
                    return {**default_config, **config}
            else:
                logger.info(f"配置文件不存在，使用默认配置: {self.config_file}")
                return default_config
        except Exception as e:
            logger.error(f"加载配置文件失败: {e}")
            return default_config

    def _load_env_config(self):
        """加载环境变量配置"""
        return {
            # 腾讯云API配置
            'secret_id': os.getenv("TENCENTCLOUD_SECRET_ID"),
            'secret_key': os.getenv("TENCENTCLOUD_SECRET_KEY"),

            # Excel文件配置
            'excel_file_path': os.getenv("EXCEL_FILE_PATH"),
            'excel_file_name': os.getenv("EXCEL_FILE_NAME"),
            'excel_sheet_name': os.getenv("EXCEL_SHEET_NAME"),
            'excel_sorted_sheet_name': os.getenv("EXCEL_SORTED_SHEET_NAME", "排序结果"),

            # OCR过滤配置
            'filter_invalid_patterns': self._parse_list(os.getenv("FILTER_INVALID_ITEM_PATTERNS", "")),
            'filter_invalid_texts': self._parse_list(os.getenv("FILTER_INVALID_ITEM_TEXTS", "")),
            'filter_title_blacklist': self._parse_list(os.getenv("FILTER_TITLE_BLACKLIST", "")),
            'category_modifiers_to_remove': self._parse_list(os.getenv("CATEGORY_MODIFIERS_TO_REMOVE", "")),
            'category_priority_order': self._parse_list(os.getenv("CATEGORY_PRIORITY_ORDER", ""))
        }

    def _parse_list(self, value):
        """解析逗号分隔的字符串为列表"""
        if not value:
            return []
        return [item.strip() for item in value.split(",") if item.strip()]

    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            logger.info(f"配置已保存到: {self.config_file}")
        except Exception as e:
            logger.error(f"保存配置文件失败: {e}")

    def get(self, key, default=None):
        """获取配置值"""
        return self.config.get(key, default)

    def set(self, key, value):
        """设置配置值"""
        self.config[key] = value
        self.save_config()

    def get_env(self, key, default=None):
        """获取环境变量配置"""
        return self.env_config.get(key, default)

    def validate_env_config(self):
        """验证环境变量配置"""
        errors = []

        if not self.env_config['secret_id']:
            errors.append("TENCENTCLOUD_SECRET_ID未设置")

        if not self.env_config['secret_key']:
            errors.append("TENCENTCLOUD_SECRET_KEY未设置")

        if not self.env_config['excel_file_path']:
            errors.append("EXCEL_FILE_PATH未设置")

        if errors:
            logger.error(f"环境变量配置错误: {', '.join(errors)}")
            return False, errors

        logger.info("环境变量配置验证通过")
        return True, []

    def get_filter_config(self):
        """获取过滤配置"""
        return {
            "invalid_item_patterns": self.env_config['filter_invalid_patterns'],
            "invalid_item_texts": self.env_config['filter_invalid_texts'],
            "title_blacklist": self.env_config['filter_title_blacklist']
        }

    def get_category_config(self):
        """获取类别配置"""
        return {
            "modifiers_to_remove": self.env_config['category_modifiers_to_remove'],
            "priority_order": self.env_config['category_priority_order']
        }


# 全局配置管理器实例
config_manager = ConfigManager()
