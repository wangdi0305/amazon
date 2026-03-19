# Amazon Auto Parts Listing Generator

自动生成 Amazon 汽配 Listing 的 Python 工具，基于 DeepSeek AI。

## 功能特性

- 🚗 **汽配 OE 号解析** - 自动提取 OE 号对应的车型适配信息
- 📝 **Listing 自动生成** - 生成标准化的 Amazon 汽配产品描述
- 📊 **Excel 批量处理** - 支持 Excel 批量导入/导出
- 💾 **断点续传** - 支持 checkpoint 机制，中断后可继续
- 🌐 **Selenium 自动化** - 自动与 DeepSeek 网页端交互

## 使用方法

1. 安装依赖：
```bash
pip install openpyxl selenium
```

2. 配置 Excel 文件路径（修改脚本中的 `EXCEL_PATH` 和 `OUTPUT_PATH`）

3. 运行脚本：
```bash
python deepseek_listing_generator_final.py
```

4. 在浏览器中登录 DeepSeek，确认后脚本自动开始处理

## 输出字段

| 列 | 说明 |
|---|---|
| A | OE 号 |
| G | 产品标题/名称 |
| AD | Vehicle Fitment |
| AE-AI | 产品特性描述 |
| AJ-AN | 关键词 |

## 技术栈

- Python 3.x
- Selenium WebDriver (Edge)
- OpenPyXL
- DeepSeek AI

## 注意事项

- 需要先登录 DeepSeek 网页版
- 支持 Edge 浏览器
- 建议使用稳定的网络环境

## License

MIT
