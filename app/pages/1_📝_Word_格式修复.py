"""Word 文本格式修复 — 引号、标点、单位一键修复。"""

import streamlit as st

from dockit.docx import format_text

st.set_page_config(page_title="Word 格式修复", page_icon="📝")

st.title("📝 Word 格式修复")
st.markdown("上传 `.docx` 文件，自动修复文本中的引号、标点和单位格式。")

# -- Upload --
uploaded = st.file_uploader("上传 Word 文件", type=["docx"], accept_multiple_files=False)

if uploaded is not None:
    st.divider()

    # -- Options --
    st.markdown("#### 修复选项")
    col1, col2, col3 = st.columns(3)
    with col1:
        do_quotes = st.checkbox("修复引号", value=True, help="统一为中文标准引号 "" ")
    with col2:
        do_punct = st.checkbox("修复标点", value=True, help="英文标点 → 中文标点")
    with col3:
        do_units = st.checkbox("转换单位", value=True, help="中文单位 → 标准符号（平方米→m²）")

    col4, col5 = st.columns(2)
    with col4:
        quote_font = st.text_input("引号字体", value="宋体", help="引号字符使用的字体")
    with col5:
        header_footer = st.radio(
            "页眉页脚",
            ["处理", "跳过", "删除"],
            horizontal=True,
            help="对页眉页脚的处理方式",
        )

    # -- Process --
    if st.button("🚀 开始处理", type="primary", use_container_width=True):
        with st.spinner("正在处理文档..."):
            doc_bytes = uploaded.read()

            result = format_text(
                doc_bytes,
                fix_quotes=do_quotes,
                fix_punctuation=do_punct,
                fix_units=do_units,
                quote_font=quote_font,
                process_headers_footers=(header_footer == "处理"),
                strip_headers_footers=(header_footer == "删除"),
            )

        # -- Stats --
        st.divider()
        st.markdown("#### 处理结果")

        c1, c2, c3 = st.columns(3)
        c1.metric("引号修复", f"{result.stats['quotes']} 处")
        c2.metric("标点修复", f"{result.stats['punctuation']} 处")
        c3.metric("单位转换", f"{result.stats['units']} 处")

        total = sum(result.stats.values())
        if total > 0:
            st.success(f"共修复 {total} 处格式问题")
        else:
            st.info("未发现需要修复的格式问题")

        # -- Download --
        output_name = uploaded.name.replace(".docx", "_fixed.docx")
        st.download_button(
            label="📥 下载修复后的文件",
            data=result.data,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
else:
    st.info("👆 请上传一个 .docx 文件开始使用")
