"""表格标准化检查 — 检查 Markdown 表格的表名和引导段落。"""

import streamlit as st

from dockit.md import check_tables, fix_table_names, reorder_table_names

st.set_page_config(page_title="表格标准化", page_icon="📋")

st.title("📋 表格标准化检查")
st.markdown("粘贴 Markdown 文本，检查表格是否缺少表名或引导段落，支持一键补全和重排序。")

# -- Input --
md_input = st.text_area(
    "粘贴 Markdown 内容",
    height=300,
    placeholder="在此粘贴包含表格的 Markdown 文本...",
)

if not md_input.strip():
    st.info("👆 请粘贴 Markdown 内容开始检查")
    st.stop()

st.divider()

# -- Options --
col1, col2 = st.columns(2)
with col1:
    min_chars = st.number_input("引导段落最少字数", value=80, min_value=0, step=10, help="低于此字数视为引导段落过短")
with col2:
    chapter_num = st.number_input("章节编号（用于补全表名）", value=1, min_value=1, step=1, help="如第 3 章则填 3，生成 表3-1, 表3-2...")

# -- Check --
if st.button("检查表格", type="primary", use_container_width=True):
    result = check_tables(md_input, min_intro_chars=min_chars)

    st.divider()
    st.markdown("#### 检查结果")
    st.metric("表格总数", result.tables)

    if not result.issues:
        st.success("所有表格均符合规范！")
    else:
        st.warning(f"发现 **{len(result.issues)}** 个问题")
        for issue in result.issues:
            if issue["type"] == "missing_name":
                st.markdown(f"- **第 {issue['line']} 行**：缺少表名")
            elif issue["type"] == "missing_intro":
                st.markdown(f"- **第 {issue['line']} 行**：缺少引导段落")
            elif issue["type"] == "short_intro":
                st.markdown(f"- **第 {issue['line']} 行**：引导段落过短（{issue['intro_len']}/{issue['min_chars']} 字）")

    # -- Fix actions --
    st.divider()
    st.markdown("#### 修复操作")

    fix_col1, fix_col2 = st.columns(2)

    with fix_col1:
        if st.button("补全缺失表名", use_container_width=True):
            fixed = fix_table_names(md_input, chapter_num, min_intro_chars=min_chars)
            st.session_state["fixed_md"] = fixed
            st.success("已补全表名（含占位符 `[待命名]`）")

    with fix_col2:
        if st.button("重排序表名位置", use_container_width=True):
            reordered, count = reorder_table_names(md_input)
            st.session_state["fixed_md"] = reordered
            if count > 0:
                st.success(f"已调整 {count} 个表名的位置")
            else:
                st.info("无需调整")

# -- Output --
if "fixed_md" in st.session_state:
    st.divider()
    st.markdown("#### 修复后的内容")
    st.code(st.session_state["fixed_md"], language="markdown")

    st.download_button(
        label="下载修复后的 Markdown",
        data=st.session_state["fixed_md"].encode("utf-8"),
        file_name="fixed_tables.md",
        mime="text/markdown",
        use_container_width=True,
    )
