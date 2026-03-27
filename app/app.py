"""DocKit Web — Document Processing Toolkit."""

import streamlit as st

st.set_page_config(
    page_title="DocKit",
    page_icon="📄",
    layout="wide",
)

st.title("📄 DocKit")
st.subheader("文档处理工具箱")

st.markdown(
    '<div style="text-align:center; color:gray; font-size:13px;">'
    '⭐ <a href="https://github.com/zengtianli/dockit">Star on GitHub</a> · '
    '👤 <a href="https://github.com/zengtianli">Tianli Zeng</a> · '
    '🌊 <a href="https://hydro.tianlizeng.cloud">Hydro Toolkit — 水利计算工具集</a>'
    '</div>',
    unsafe_allow_html=True,
)

st.markdown("拖入文件，一键处理，下载结果。支持 Word / PowerPoint / Excel / CSV 批量处理。")

st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("### 📝 Word 格式修复")
    st.markdown("""
    - 双引号自动配对（中文标准引号）
    - 英文标点 → 中文标点
    - 中文单位 → 标准符号（平方米 → m²）
    """)

    st.markdown("### 📄 Word 文本提取")
    st.markdown("""
    - 全文转 Markdown
    - 段落结构化提取
    - 按章节拆分
    """)

    st.markdown("### 🧹 Word 样式清理")
    st.markdown("""
    - 删除未使用样式
    - 重命名 / 合并样式
    """)

with col2:
    st.markdown("### 📊 格式转换")
    st.markdown("""
    - CSV / TXT / Excel 互转
    - 自动检测分隔符
    - 多工作表拆分
    """)

    st.markdown("### 🎨 PPT 标准化")
    st.markdown("""
    - 字体统一（微软雅黑）
    - 文本格式修复
    - 表格样式设置
    """)

    st.markdown("### 📑 PPT 转 Markdown")
    st.markdown("""
    - 提取幻灯片文本
    - 含演讲备注
    """)

with col3:
    st.markdown("### 🔗 表格合并")
    st.markdown("""
    - 多个 Excel 纵向拼接
    - 按关键列横向合并
    """)

    st.markdown("### 📋 表格标准化")
    st.markdown("""
    - 检查表名和引导段落
    - 一键补全 / 重排序
    """)

    st.markdown("### 📊 图表生成")
    st.markdown("""
    - 柱状图 / 甘特图 / 流程图
    - JSON 配置，PNG 输出
    """)

st.divider()
st.markdown("#### v0.3.0 新增")

col4, col5, col6, col7 = st.columns(4)

with col4:
    st.markdown("### 🔄 Markdown 转 Word")
    st.markdown("""
    - 支持模板样式
    - 标题/表格/列表
    """)

with col5:
    st.markdown("### 🔍 格式检查")
    st.markdown("""
    - 页面/样式/水印快照
    - 双文档格式对比
    """)

with col6:
    st.markdown("### ✏️ 修订标记")
    st.markdown("""
    - 读取修订和批注
    - 批量写入修订
    """)

with col7:
    st.markdown("### 🖼️ 图片题注")
    st.markdown("""
    - 图片样式统一
    - 题注自动应用
    """)

st.divider()
st.markdown("""
<div style="text-align:center; color:gray; font-size:13px;">
    <b>DocKit</b> is open source ·
    <a href="https://github.com/zengtianli/dockit">⭐ Star on GitHub</a> ·
    <code>pip install dockit</code><br/>
    Built by <a href="https://github.com/zengtianli">Tianli Zeng</a> ·
    More tools: <a href="https://hydro.tianlizeng.cloud">🌊 Hydro Toolkit</a>
</div>
""", unsafe_allow_html=True)
