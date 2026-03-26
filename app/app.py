"""DocKit Web — Document Processing Toolkit."""

import streamlit as st

st.set_page_config(
    page_title="DocKit",
    page_icon="📄",
    layout="wide",
)

st.title("📄 DocKit")
st.subheader("文档处理工具箱")

st.markdown("拖入文件，一键处理，下载结果。支持 Word / PowerPoint / Excel / CSV 批量处理。")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📝 Word 格式修复")
    st.markdown("""
    - 双引号自动配对（中文标准引号）
    - 英文标点 → 中文标点
    - 中文单位 → 标准符号（平方米 → m²）
    - 引号字体自动设置为宋体
    """)

    st.markdown("### 📊 格式转换")
    st.markdown("""
    - CSV / TXT / Excel 互转（8 种转换）
    - 自动检测分隔符
    - 多工作表拆分
    - XLS → XLSX 升级
    """)

with col2:
    st.markdown("### 🎨 PPT 标准化")
    st.markdown("""
    - 字体统一（微软雅黑）
    - 文本格式修复（引号/标点/单位）
    - 表格样式设置（标题行/镶边行/首列）
    - 一键全部标准化
    """)

    st.markdown("### 🔗 表格合并")
    st.markdown("""
    - 多个 Excel 纵向拼接
    - 按关键列横向合并
    - 数据预览
    """)

st.divider()

st.markdown(
    "**DocKit** is open source — "
    "[GitHub](https://github.com/zengtianli/dockit) · "
    "`pip install dockit`"
)
