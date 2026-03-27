# DocKit

[English](README.md) | **中文**

文档处理工具箱，支持 Word、PowerPoint、Excel 和 CSV 文件。

**字节进，字节出。** 纯处理逻辑，不绑定文件 I/O——可在命令行脚本、Web 应用或任意 Python 程序中使用。

[![在线试用](https://img.shields.io/badge/在线试用-dockit.tianlizeng.cloud-blue?style=for-the-badge)](https://dockit.tianlizeng.cloud)
[![License: MIT](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-yellow?style=for-the-badge)](https://python.org)

---

### 立即试用，无需安装

**https://dockit.tianlizeng.cloud**

上传文档，选择工具，下载结果。零配置，开箱即用。

![DocKit 首页](docs/screenshots/homepage.png)

---

## 能做什么？

| 工具 | 功能 | 输入 | 输出 |
|------|------|------|------|
| **Word 格式修复** | 修复引号配对、英文标点转中文、单位符号标准化 | Word 文件 | Word 文件 |
| **Word 文本提取** | 全文提取为 Markdown | Word 文件 | Markdown |
| **Word 样式清理** | 删除未使用样式、重命名/合并样式 | Word 文件 | Word 文件 |
| **格式转换** | XLSX、CSV、TXT 互转 | 表格文件 | 表格文件 |
| **PPT 标准化** | 统一字体、修复文本格式、设置表格样式 | PowerPoint | PowerPoint |
| **PPT 转 Markdown** | 提取幻灯片文本及演讲者备注 | PowerPoint | Markdown |
| **表格合并** | 按列名匹配合并多个表格 | 多个 Excel | Excel 文件 |
| **表格标准化** | 检查表名和引导段落、一键补全/重排序 | Excel 文件 | Excel 文件 |
| **图表生成** | 柱状图 / 甘特图 / 流程图 | JSON 配置 | PNG |
| **Markdown 转 Word** | Markdown 转带样式 Word 文档 | Markdown | Word 文件 |
| **格式检查** | 检查文档结构和样式 | Word 文件 | 报告 |
| **修订标记** | 为 Word 文档添加修订标记 | Word 文件 | Word 文件 |
| **图片题注** | 为 Word 文档中的图片添加题注 | Word 文件 | Word 文件 |

## 安装

```bash
# 从 GitHub 安装（推荐）
pip install git+https://github.com/zengtianli/dockit.git

# 本地开发
git clone https://github.com/zengtianli/dockit.git
cd dockit && pip install -e .
```

## 快速上手

```python
from dockit.docx import format_text

with open("input.docx", "rb") as f:
    doc_bytes = f.read()

result = format_text(doc_bytes, fix_quotes=True, fix_punctuation=True, fix_units=True)

with open("output.docx", "wb") as f:
    f.write(result.data)

print(result.stats)  # {"quotes": 5, "punctuation": 12, "units": 3}
```

## 功能模块

### 文本格式化 (`dockit.text`)
- 修复中文引号配对（智能匹配前后引号）
- 英文标点转中文标点（逗号、句号、冒号等）
- 中文单位名称转标准符号（如 平方米 → m²、立方米每秒 → m³/s）

### Word 文档处理 (`dockit.docx`)
- 一键修复 Word 文档中的文本格式
- 引号字体拆分（为引号字符设置独立字体）
- 处理段落、表格、页眉和页脚中的文本

### PowerPoint 处理 (`dockit.pptx`)
- 统一所有幻灯片和母版的字体
- 修复文本格式（引号、标点、单位）
- 设置表格样式选项（标题行、带状行、首列）
- 一键标准化（以上全部）

### Excel 处理 (`dockit.xlsx`)
- XLSX、CSV、TXT 格式互转
- 按工作表拆分工作簿
- 列名转小写
- 旧版 .xls 转 .xlsx

### CSV 处理 (`dockit.csv`)
- 自动检测分隔符
- CSV 与分隔文本互转
- 圈号数字转普通格式（如 ① → 1）
- 按参照列表重新排序行

## 自部署

用 Docker 或直接运行：

```bash
# Docker
docker build -t dockit .
docker run -p 8503:8503 dockit

# 或直接运行
pip install -e .[web]
streamlit run app/app.py
```

或者直接用在线版：**https://dockit.tianlizeng.cloud**

## 许可证

MIT
