"""PPT 转 Markdown — 提取 PowerPoint 幻灯片内容为 Markdown。"""

import streamlit as st

from dockit.pptx import to_markdown

st.set_page_config(page_title="PPT 转 Markdown", page_icon="📑")

st.title("📑 PPT 转 Markdown")
st.markdown("上传 `.pptx` 文件，提取所有幻灯片文本（含演讲备注）为 Markdown 格式。")

# -- Upload --
uploaded = st.file_uploader("上传 PowerPoint 文件", type=["pptx"], accept_multiple_files=False)

if uploaded is not None:
    file_stem = uploaded.name.rsplit(".", 1)[0]

    if st.button("开始转换", type="primary", use_container_width=True):
        with st.spinner("正在转换..."):
            pptx_bytes = uploaded.read()
            md_text = to_markdown(pptx_bytes)

        # -- Results --
        st.divider()
        st.markdown("#### 转换结果")

        slides = md_text.split("---\n")
        slide_count = len([s for s in slides if s.strip()])
        st.info(f"共提取 **{slide_count}** 张幻灯片")

        st.code(md_text, language="markdown")

        # -- Download --
        st.download_button(
            label="下载 Markdown",
            data=md_text.encode("utf-8"),
            file_name=f"{file_stem}.md",
            mime="text/markdown",
            use_container_width=True,
        )
else:
    st.info("👆 请上传一个 .pptx 文件开始使用")
