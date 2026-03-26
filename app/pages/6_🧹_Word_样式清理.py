"""Word 样式清理 — 删除未使用样式、重命名、合并样式。"""

import streamlit as st

from dockit.docx import cleanup_styles

st.set_page_config(page_title="Word 样式清理", page_icon="🧹")

st.title("🧹 Word 样式清理")
st.markdown("上传 `.docx` 文件，清理冗余样式。支持删除未使用样式、重命名样式、合并样式。")

# -- Upload --
uploaded = st.file_uploader("上传 Word 文件", type=["docx"], accept_multiple_files=False)

if uploaded is not None:
    doc_bytes = uploaded.read()
    file_stem = uploaded.name.rsplit(".", 1)[0]

    st.divider()

    # -- Options --
    st.markdown("#### 清理选项")

    delete_unused = st.checkbox("删除未使用的样式", value=True, help="移除文档中未被任何段落引用的样式定义")

    st.markdown("#### 样式重命名（可选）")
    st.caption("每行一条，格式：`旧名称 = 新名称`")
    rename_text = st.text_area("重命名规则", placeholder="标题 1 = Heading 1\n正文 = Body Text", height=100)

    st.markdown("#### 样式合并（可选）")
    st.caption("每行一条，格式：`源样式ID = 目标样式ID`（将源样式的段落全部改为目标样式）")
    merge_text = st.text_area("合并规则", placeholder="CustomStyle1 = Normal", height=100)

    # -- Process --
    if st.button("开始清理", type="primary", use_container_width=True):
        # Parse renames
        renames = {}
        for line in rename_text.strip().split("\n"):
            if "=" in line:
                old, new = line.split("=", 1)
                renames[old.strip()] = new.strip()

        # Parse merges
        merges = {}
        for line in merge_text.strip().split("\n"):
            if "=" in line:
                src, dst = line.split("=", 1)
                merges[src.strip()] = dst.strip()

        with st.spinner("正在清理样式..."):
            result = cleanup_styles(
                doc_bytes,
                renames=renames or None,
                merges=merges or None,
                delete_unused=delete_unused,
            )

        # -- Results --
        st.divider()
        st.markdown("#### 清理结果")

        if result.log:
            for entry in result.log:
                st.markdown(f"- {entry}")
        else:
            st.info("未发现需要清理的样式")

        # -- Download --
        output_name = f"{file_stem}_cleaned.docx"
        st.download_button(
            label="下载清理后的文件",
            data=result.data,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
else:
    st.info("👆 请上传一个 .docx 文件开始使用")
