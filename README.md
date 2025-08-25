# 屏幕截图OCR工具 v2.0

一个专业的屏幕截图OCR识别工具，专门用于提取游戏家具套装信息并自动整理到Excel表格中。

## ✨ 主要功能

- 🖼️ **智能屏幕截图**：支持拖拽选择任意屏幕区域
- 🔍 **高精度OCR识别**：基于腾讯云OCR API，准确识别表格内容
- 📊 **Excel自动化**：自动写入识别结果到Excel文件
- 🗂️ **智能数据排序**：按年份、类别、套装号自动分类排序
- 🎯 **DPI自适应**：完美支持高分辨率显示器和各种缩放比例
- 💾 **历史数据保护**：排序时自动保留用户填写的数量数据
- 🐛 **调试支持**：自动保存识别图片用于问题排查

## 🚀 快速开始

### 环境要求

- Python 3.7+
- Windows 10/11（推荐）
- Excel 2016+（可选，用于更好的集成）

### 安装步骤

1. **克隆项目**

```bash
git clone https://github.com/wuyaiii/langman-furniture-ocr.git
cd langman-furniture-ocr
```

2. **安装依赖**

```bash
pip install -r requirements_new.txt
```

3. **配置环境变量**
   复制 `.env.example` 为 `.env` 并填写配置：

```bash
cp .env.example .env
```

4. **运行程序**

```bash
python main.py
```

## ⚙️ 配置说明

### 环境变量配置 (.env文件)

#### 腾讯云API配置（必需）

```env
# 从腾讯云控制台获取: https://console.cloud.tencent.com/cam/capi
TENCENTCLOUD_SECRET_ID=你的SecretId
TENCENTCLOUD_SECRET_KEY=你的SecretKey
```

#### Excel文件配置（必需）

```env
# Excel文件完整路径
EXCEL_FILE_PATH=F:\path\to\your\excel\file.xlsx
# Excel文件名（用于xlwings识别）
EXCEL_FILE_NAME=your_file.xlsx
# 工作表名称（可选）
EXCEL_SHEET_NAME=识别结果
# 排序结果工作表名称
EXCEL_SORTED_SHEET_NAME=排序结果
```

#### OCR过滤配置（可选）

```env
# 无效物品名的正则表达式模式（逗号分隔）
FILTER_INVALID_ITEM_PATTERNS=^\\+\\d+$,^\\*\\d+$,^-\\d+$,^\\d+$
# 无效物品名的文本列表（逗号分隔）
FILTER_INVALID_ITEM_TEXTS=关闭,定,图片,附加,附加最大值
# 标题黑名单（逗号分隔）
FILTER_TITLE_BLACKLIST=附加最大值,物品名,图片,关闭
```

#### 类别排序配置（可选）

```env
# 需要移除的修饰词（逗号分隔）
CATEGORY_MODIFIERS_TO_REMOVE=精品,动态,年,叮当猫,节
# 类别优先级顺序（逗号分隔）
CATEGORY_PRIORITY_ORDER=春,情人,元宵,劳动,端午,七夕,中秋,国庆,万圣,圣诞
```

### 程序配置 (screen_ocr_config.json)

程序运行时会自动创建配置文件，包含以下选项：

```json
{
  "auto_open_excel": false,          // 启动时自动打开Excel
  "show_confirmation": true,         // 识别后显示确认对话框
  "window_topmost": false,          // 窗口置顶
  "show_selection_border": false,   // 显示选框边框
  "save_debug_images": true,        // 保存调试图片
  "debug_image_dir": "debug_images", // 调试图片目录
  "selection_coordinates": {        // 选框坐标（自动保存）
    "x1": 0, "y1": 0, "x2": 0, "y2": 0
  }
}
```

## 📖 使用指南

### 基本使用流程

1. **启动程序**：运行 `python main.py`
2. **选择区域**：点击"开始屏幕选择"，拖拽选择要识别的区域
3. **OCR识别**：点击"识别选中区域"开始识别
4. **确认结果**：在弹出的对话框中确认或修改识别结果
5. **自动写入**：数据自动写入到Excel文件中

### 高级功能

#### Excel排序处理

- 点击"Excel排序处理"按钮
- 程序会自动读取Excel中的数据
- 按年份、类别、套装号进行智能排序
- 自动保留用户填写的数量数据
- 生成格式化的排序结果表

#### DPI适配

程序自动检测并适配不同的显示器分辨率和缩放设置：

- 1920x1080 (100% 缩放)
- 2560x1440 (125% 缩放)
- 3840x2160 (150% 缩放)
- 其他分辨率和缩放比例

#### 调试功能

- 每次识别前自动保存截图到 `debug_images/` 目录
- 详细的日志记录保存到 `logs/` 目录
- 便于问题排查和结果验证

## 🏗️ 项目结构

```
langmanCV/
├── src/                    # 源代码目录
│   ├── core/              # 核心功能模块
│   │   ├── screen_capture.py    # 屏幕截图
│   │   ├── ocr_processor.py     # OCR处理
│   │   ├── excel_manager.py     # Excel操作
│   │   └── data_sorter.py       # 数据排序
│   ├── ui/                # 用户界面
│   │   └── main_window.py       # 主界面
│   └── utils/             # 工具模块
│       ├── logger.py            # 日志系统
│       ├── config_manager.py    # 配置管理
│       └── dpi_helper.py        # DPI处理
├── main.py                # 程序入口
├── requirements_new.txt   # 依赖包列表
├── .env.example          # 环境变量模板
└── README.md             # 说明文档
```

## 🔧 故障排除

### 常见问题

**Q: 识别结果不准确怎么办？**
A:

- 检查截图区域是否包含完整的表格
- 确保截图清晰，避免模糊或重叠
- 查看 `debug_images/` 目录中的调试图片
- 调整OCR过滤配置

**Q: 高分辨率显示器识别异常？**
A:

- 程序已内置DPI自适应功能
- 确保Windows显示缩放设置正确
- 调整分辨率或者文字缩放后重新运行程序，默认为1920x1080 缩放100%

**Q: Excel写入失败？**
A:

- 确保Excel文件路径正确且可写
- 检查Excel程序是否正在运行
- 验证xlwings是否正确安装

**Q: 排序时数量数据丢失？**
A:

- v2.0版本已修复此问题
- 程序会自动保留历史数量数据
- 查看日志确认数据恢复情况

### 日志查看

程序运行日志保存在 `logs/` 目录中：

- 文件名格式：`ScreenOCR_YYYYMMDD.log`
- 包含详细的操作记录和错误信息
- 用于问题诊断和性能分析

## 🔄 版本更新

### v2.0.0 (当前版本)

- ✅ 完全重构代码架构，模块化设计
- ✅ 修复高分辨率显示器兼容性问题
- ✅ 添加DPI自适应功能
- ✅ 实现专业日志系统
- ✅ 排序时保留历史数量数据
- ✅ 添加调试图片保存功能
- ✅ 优化依赖包，减小打包体积
- ✅ 完善错误处理和用户提示

### v1.x.x (旧版本)

- 基础OCR识别功能
- Excel数据写入
- 简单的数据排序

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🤝 贡献

欢迎提交Issue和Pull Request来帮助改进这个项目！

## 📞 支持

如果您在使用过程中遇到问题，请：

1. 查看本文档的故障排除部分
2. 检查 `logs/` 目录中的日志文件
3. 在GitHub上提交Issue，并附上相关日志信息

---

**注意**：使用本工具需要腾讯云OCR API密钥，请确保已正确配置相关凭证。
