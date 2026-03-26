"""Markdown 转 Word — 带模板样式的 Markdown 转 Docx。"""

import streamlit as st

from dockit.docx import md_to_docx

st.set_page_config(page_title="Markdown 转 Word", page_icon="🔄")

st.title("🔄 Markdown 转 Word")
st.markdown("输入 Markdown 文本，生成带样式的 Word 文档。可上传模板 docx 提取标题样式。")

# -- Input --
tab_paste, tab_upload = st.tabs(["粘贴 Markdown", "上传 .md 文件"])

with tab_paste:
    md_text = st.text_area("Markdown 内容", height=300, placeholder="# 标题\n\n正文内容...")

with tab_upload:
    md_file = st.file_uploader("上传 .md 文件", type=["md"], key="md_upload")
    if md_file:
        md_text = md_file.read().decode("utf-8")
        st.code(md_text[:500] + ("..." if len(md_text) > 500 else ""), language="markdown")

# -- Template --
st.divider()
template_file = st.file_uploader(
    "上传模板 .docx（可选）",
    type=["docx"],
    help="从模板中提取标题 1-4 和正文样式。不上传则使用默认样式。",
    key="template_upload",
)

# -- Convert --
if md_text and md_text.strip():
    if st.button("生成 Word 文档", type="primary", use_container_width=True):
        template_bytes = template_file.read() if template_file else None

        with st.spinner("正在转换..."):
            try:
                doc_bytes = md_to_docx(md_text, template_bytes=template_bytes)
            except Exception as e:
                st.error(f"转换失败: {e}")
                st.stop()

        st.divider()
        st.success(f"转换完成！文件大小：{len(doc_bytes) / 1024:.1f} KB")

        st.download_button(
            label="下载 Word 文档",
            data=doc_bytes,
            file_name="output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
else:
    st.info("👆 请输入 Markdown 内容或上传 .md 文件")
