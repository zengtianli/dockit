"""图片题注样式 — 批量为图片和题注段落应用样式。"""

import streamlit as st

from dockit.docx import add_captions

st.set_page_config(page_title="图片题注", page_icon="🖼️")

st.title("🖼️ 图片题注样式")
st.markdown("上传 `.docx` 文件，自动为图片段落和题注（下一行）应用指定样式。")

# -- Upload --
uploaded = st.file_uploader("上传 Word 文件", type=["docx"], accept_multiple_files=False)

if uploaded is not None:
    st.divider()

    # -- Options --
    col1, col2 = st.columns(2)
    with col1:
        style_name = st.text_input("样式名称", value="ZDWP图名", help="应用到图片和题注的段落样式（支持模糊匹配）")
    with col2:
        add_blank = st.checkbox("题注后加空行", value=True, help="在题注段落后自动插入空行")

    # -- Process --
    if st.button("应用样式", type="primary", use_container_width=True):
        with st.spinner("正在处理..."):
            doc_bytes = uploaded.read()
            result = add_captions(doc_bytes, style_name=style_name, add_blank_after=add_blank)

        st.divider()
        st.markdown("#### 处理结果")

        c1, c2 = st.columns(2)
        c1.metric("图片段落", result.images_styled)
        c2.metric("题注段落", result.captions_styled)

        total = result.images_styled + result.captions_styled
        if total > 0:
            st.success(f"共处理 {total} 个段落")
        else:
            st.info("未找到图片段落")

        file_stem = uploaded.name.rsplit(".", 1)[0]
        st.download_button(
            label="下载处理后的文件",
            data=result.data,
            file_name=f"{file_stem}_captioned.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
else:
    st.info("👆 请上传一个 .docx 文件开始使用")
