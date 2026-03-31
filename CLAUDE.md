# dockit — 文档处理工具包（Word/PPT/Excel/CSV/Chart）

## Quick Reference

| 项目 | 路径/值 |
|------|---------|
| 库源码 | `src/dockit/` |
| Web 应用 | `app/app.py` (Streamlit) |
| CLI 入口 | `src/dockit/cli.py` / `python -m dockit` |
| 线上地址 | https://dockit.tianlizeng.cloud |
| Docker | `Dockerfile` → 容器化部署 |
| 测试 | `tests/` (pytest) |

## 常用命令

```bash
# 开发安装
pip install -e ".[excel,chart,web]"

# 本地启动 Web 应用
streamlit run app/app.py

# CLI 使用
python -m dockit --help

# 运行测试
pytest tests/

# 构建 Docker 镜像
docker build -t dockit .
docker run -p 8501:8501 dockit
```

## 库使用模式

dockit 是 **bytes in, bytes out** 设计，所有处理函数接收字节流，返回带 `.data` 和 `.stats` 的结果对象：

```python
from dockit.docx import format_text
from dockit.pptx import standardize_ppt
from dockit.xlsx import merge_tables

# 读取 → 处理 → 写出
result = format_text(open("in.docx","rb").read(), fix_quotes=True)
open("out.docx","wb").write(result.data)
print(result.stats)
```

## 项目结构

```
src/dockit/        # 核心库
  docx.py          # Word 处理（格式修正、样式、修订标记等）
  pptx.py          # PPT 处理（标准化、转 Markdown）
  xlsx.py          # Excel/CSV（合并、标准化、格式转换）
  text.py          # 文本规范化（引号、标点、单位）
  chart.py         # 图表生成（柱状图、甘特图、流程图）
  cli.py           # CLI 命令行接口
app/
  app.py           # Streamlit Web 入口
  pages/           # 各工具页面
tests/             # pytest 测试 + fixtures
```

## 功能速查

| 模块 | 主要函数 | 说明 |
|------|---------|------|
| `dockit.text` | `fix_quotes`, `fix_punctuation` | 引号/标点/单位规范化 |
| `dockit.docx` | `format_text`, `extract_text`, `cleanup_styles` | Word 处理 |
| `dockit.pptx` | `standardize_ppt`, `ppt_to_markdown` | PPT 处理 |
| `dockit.xlsx` | `merge_tables`, `standardize_table`, `convert_format` | 表格处理 |
| `dockit.chart` | `generate_chart` | 图表生成（需 `[chart]` extra） |

## 注意事项

- 可选依赖分组：`[excel]` pandas/xlrd，`[chart]` matplotlib/numpy，`[web]` streamlit
- 部署走 VPS + Nginx，线上已有 CF Access 保护
- 无需外部 API，不依赖 `~/.personal_env`
