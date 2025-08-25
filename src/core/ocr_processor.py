# -*- coding: utf-8 -*-
"""
OCR处理模块
"""

import json
import base64
import io
import re
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.ocr.v20181119 import ocr_client, models
from ..utils import logger, config_manager


class OCRProcessor:
    """OCR处理器"""

    def __init__(self):
        self.ocr_client = None
        self.filter_config = config_manager.get_filter_config()
        self._init_ocr_client()

    def _init_ocr_client(self):
        """初始化腾讯云OCR客户端"""
        try:
            secret_id = config_manager.get_env('secret_id')
            secret_key = config_manager.get_env('secret_key')

            if not secret_id or not secret_key:
                logger.error("腾讯云API密钥未设置")
                return False

            cred = credential.Credential(secret_id, secret_key)
            httpProfile = HttpProfile()
            httpProfile.endpoint = "ocr.tencentcloudapi.com"

            clientProfile = ClientProfile()
            clientProfile.httpProfile = httpProfile

            self.ocr_client = ocr_client.OcrClient(cred, "", clientProfile)
            logger.info("腾讯云OCR客户端初始化成功")
            return True

        except Exception as e:
            logger.error(f"初始化OCR客户端失败: {e}")
            return False

    def recognize_table(self, image):
        """识别表格图像"""
        if not self.ocr_client:
            logger.error("OCR客户端未初始化")
            return None

        try:
            # 转换图片为base64
            buffer = io.BytesIO()
            image.save(buffer, format='PNG')
            image_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')

            # 发送表格OCR请求
            req = models.RecognizeTableAccurateOCRRequest()
            params = {"ImageBase64": image_base64}
            req.from_json_string(json.dumps(params))

            # 调用表格识别接口
            resp = self.ocr_client.RecognizeTableAccurateOCR(req)
            result = json.loads(resp.to_json_string())

            logger.info("表格OCR识别完成")
            logger.info(result)
            return result

        except TencentCloudSDKException as e:
            logger.error(f"OCR识别失败: {e}")
            return None
        except Exception as e:
            logger.error(f"OCR处理失败: {e}")
            return None

    def extract_title_and_items(self, ocr_result):
        """从OCR结果中提取标题和物品名"""
        if not ocr_result or "TableDetections" not in ocr_result:
            logger.warning("OCR结果为空或格式不正确")
            return None, []

        extracted_title = None
        extracted_items = []

        try:
            # 查找标题
            extracted_title = self._extract_title(
                ocr_result["TableDetections"])

            # 提取物品名
            extracted_items = self._extract_items(
                ocr_result["TableDetections"])

            logger.info(
                f"提取完成 - 标题: {extracted_title}, 物品数量: {len(extracted_items)}")

        except Exception as e:
            logger.error(f"提取标题和物品失败: {e}")

        return extracted_title, extracted_items

    def _extract_title(self, table_detections):
        """提取标题"""
        for table in table_detections:
            # 查找所有位置为-1的单元格
            for cell in table.get("Cells", []):
                if not (cell.get("ColTl") == -1 and
                        cell.get("RowTl") == -1 and
                        cell.get("ColBr") == -1 and
                        cell.get("RowBr") == -1):
                    continue

                text = cell.get("Text", "").strip()
                if not text or text == "X":
                    continue

                # 匹配年份开头的标题
                year_pattern = r'^\d{4}.*$'
                if re.match(year_pattern, text):
                    logger.info(f"从年份格式文本中找到标题: {text}")
                    return self._clean_title(text)

        # 如果没有找到标题，尝试从数据表格中查找
        return self._extract_title_from_data_table(table_detections)

    def _extract_title_from_data_table(self, table_detections):
        """从数据表格中提取标题"""
        for table in table_detections:
            # 兼容不同类型的表格：Type 0（标题表格）、Type 1（数据表格）和 Type 2（其他表格类型）
            if table.get("Type") in [0, 1, 2]:
                cells = table.get("Cells", [])

                # 查找"图片"单元格所在行
                pic_row = None
                for cell in cells:
                    if cell.get("Text", "").strip() == "图片":
                        pic_row = cell.get("RowTl")
                        break

                if pic_row is not None:
                    # 查找图片行上方的文本作为标题候选
                    title_candidates = []
                    for cell in cells:
                        text = cell.get("Text", "").strip()
                        row = cell.get("RowTl")

                        if text and row is not None and row < pic_row:
                            # 过滤黑名单文本
                            if text not in self.filter_config.get('title_blacklist', []):
                                distance = pic_row - row
                                title_candidates.append({
                                    'text': text,
                                    'distance': distance
                                })

                    if title_candidates:
                        # 选择距离最近的非空文本作为标题
                        closest_text = min(
                            title_candidates, key=lambda x: x['distance'])
                        logger.info(f"从表格中找到标题: {closest_text['text']}")
                        return self._clean_title(closest_text['text'])

        return None

    def _extract_items(self, table_detections):
        """提取物品名"""
        for table in table_detections:
            # 兼容不同类型的表格：Type 1（数据表格）和 Type 2（其他表格类型）
            if table.get("Type") in [0, 1, 2]:
                cells = table.get("Cells", [])
                items = self._extract_items_from_table(cells)
                if items:  # 如果找到物品，直接返回
                    return items

        return []

    def _extract_items_from_table(self, cells):
        """从表格单元格中提取物品名"""
        # 查找"物品名"列索引
        item_col = None
        for cell in cells:
            if cell.get("Text", "").strip() == "物品名":
                item_col = cell.get("ColTl")
                logger.info(f"找到'物品名'列，列索引: {item_col}")
                break

        if item_col is None:
            logger.warning("未找到'物品名'列标题")
            return []

        # 收集所有物品名
        items = []
        for cell in cells:
            text = cell.get("Text", "").strip()
            if not text or text == "物品名":
                continue

            # 检查是否在物品名列
            if cell.get("ColTl") == item_col and cell.get("RowTl") > 0:  # 排除表头行
                if self._is_valid_item_name(text):
                    items.append({
                        'text': self._clean_text(text),
                        'row': cell.get("RowTl", 0)
                    })

        # 按行号排序
        items.sort(key=lambda x: x['row'])
        item_names = [item['text'] for item in items]

        logger.info(f"找到物品名: {item_names}")
        return item_names

    def _is_valid_item_name(self, text):
        """判断文本是否是有效的物品名"""
        if not text or not text.strip():
            return False

        text = text.strip()

        # 使用配置中的模式过滤
        for pattern in self.filter_config.get('invalid_item_patterns', []):
            if re.match(pattern, text):
                return False

        # 使用配置中的黑名单文本过滤
        if text in self.filter_config.get('invalid_item_texts', []):
            return False

        # 检查是否包含中文字符
        chinese_pattern = r'[\u4e00-\u9fff]'
        if re.search(chinese_pattern, text):
            return True

        # 如果没有中文但是长度大于1且不是纯符号，也可能是有效的物品名
        if len(text) > 1 and not re.match(r'^[\+\-\*\d\s]+$', text):
            return True

        return False

    def _clean_text(self, text):
        """清理文本中的特殊字符"""
        if not text:
            return text

        # 替换中文括号为英文括号
        cleaned_text = text.replace('（', '(').replace('）', ')')

        # 移除所有空格（包括全角和半角）
        cleaned_text = cleaned_text.replace(' ', '').replace('　', '')

        # 去除首尾空格
        cleaned_text = cleaned_text.strip()

        return cleaned_text

    def _clean_title(self, title):
        """专门清理标题的方法"""
        if not title:
            return title

        # 先进行基础文本清理
        cleaned_title = self._clean_text(title)

        # 如果标题包含"一"，去掉"一"及其后面的所有文字
        if '一' in cleaned_title:
            cleaned_title = cleaned_title.split('一')[0]
            logger.info("标题处理: 去掉'一'及后面的文字")

        if '-' in cleaned_title:
            cleaned_title = cleaned_title.split('-')[0]
            logger.info("标题处理: 去掉'-'及后面的文字")

        return cleaned_title.strip()
