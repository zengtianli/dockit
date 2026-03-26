"""格式检查 — 提取 Word 文档的格式快照，或对比两个文档的格式差异。"""

import streamlit as st

from dockit.docx import check_format, compare_format, format_report

st.set_page_config(page_title="格式检查", page_icon="🔍")

st.title("🔍 Word 格式检查")
st.markdown("上传 Word 文档，生成格式报告（页面设置、样式、页眉页脚、水印等）。支持两文档对比。")

mode = st.radio("模式", ["单文档报告", "双文档对比"], horizontal=True)

if mode == "单文档报告":
    uploaded = st.file_uploader("上传 Word 文件", type=["docx"], key="single")

    if uploaded and st.button("生成格式报告", type="primary", use_container_width=True):
        with st.spinner("正在分析文档格式..."):
            doc_bytes = uploaded.read()
            snap = check_format(doc_bytes)
            report = format_report(snap)

        st.divider()

        # 关键指标
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("分节符", len(snap.page_setup))
        c2.metric("段落样式", len(snap.styles))
        c3.metric("嵌入图片", snap.images_count)
        c4.metric("直接覆盖", snap.direct_overrides_count)

        if snap.watermark:
            st.warning(f"检测到水印：**{snap.watermark}**")

        st.markdown(report)

        st.download_button(
            label="下载格式报告",
            data=report.encode("utf-8"),
            file_name=f"{uploaded.name.rsplit('.', 1)[0]}_格式报告.md",
            mime="text/markdown",
            use_container_width=True,
        )

else:
    col1, col2 = st.columns(2)
    with col1:
        before = st.file_uploader("原始文档", type=["docx"], key="before")
    with col2:
        after = st.file_uploader("修改后文档", type=["docx"], key="after")

    if before and after and st.button("对比格式", type="primary", use_container_width=True):
        with st.spinner("正在对比文档格式..."):
            result = compare_format(before.read(), after.read())

        st.divider()

        if result.all_ok:
            st.success("格式完整性通过 — 无非预期变化")
        else:
            st.error("发现非预期变化 — 请检查报告详情")

        st.markdown(result.report)

        st.download_button(
            label="下载对比报告",
            data=result.report.encode("utf-8"),
            file_name="格式对比报告.md",
            mime="text/markdown",
            use_container_width=True,
        )

if mode == "单文档报告" and not st.session_state.get("single"):
    st.info("👆 请上传一个 .docx 文件")
elif mode == "双文档对比" and (not st.session_state.get("before") or not st.session_state.get("after")):
    st.info("👆 请上传两个 .docx 文件进行对比")
