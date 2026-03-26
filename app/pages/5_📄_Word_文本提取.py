"""Word 文本提取 — 提取 Word 文档内容为 Markdown 格式。"""

import streamlit as st

from dockit.docx import extract_chapters, extract_paragraphs, extract_text

st.set_page_config(page_title="Word 文本提取", page_icon="📄")

st.title("📄 Word 文本提取")
st.markdown("上传 `.docx` 文件，提取文本内容为 Markdown 格式。支持全文、段落结构、章节拆分三种模式。")

# -- Upload --
uploaded = st.file_uploader("上传 Word 文件", type=["docx"], accept_multiple_files=False)

if uploaded is not None:
    doc_bytes = uploaded.read()
    file_stem = uploaded.name.rsplit(".", 1)[0]

    st.divider()

    mode = st.radio(
        "提取模式",
        ["全文 Markdown", "段落结构", "章节拆分"],
        horizontal=True,
        help="全文：转为 Markdown 文本；段落结构：带样式信息的结构化数据；章节拆分：按一级标题拆分",
    )

    if st.button("提取", type="primary", use_container_width=True):
        with st.spinner("正在提取文档内容..."):
            if mode == "全文 Markdown":
                md_text = extract_text(doc_bytes)

                st.divider()
                st.markdown("#### 提取结果")
                lines = md_text.strip().split("\n")
                st.info(f"共 **{len(lines)}** 行")
                st.code(md_text, language="markdown")

                st.download_button(
                    label="下载 Markdown",
                    data=md_text.encode("utf-8"),
                    file_name=f"{file_stem}.md",
                    mime="text/markdown",
                    use_container_width=True,
                )

            elif mode == "段落结构":
                paragraphs = extract_paragraphs(doc_bytes)

                st.divider()
                st.markdown("#### 段落结构")
                st.info(f"共 **{len(paragraphs)}** 个段落")

                for i, p in enumerate(paragraphs):
                    level_tag = f"H{p['level']}" if p["level"] >= 1 else "正文"
                    st.markdown(f"**{i + 1}.** `[{level_tag}]` `{p['style']}` — {p['text'][:100]}")

                # 提供 JSON 下载
                import json

                json_str = json.dumps(paragraphs, ensure_ascii=False, indent=2)
                st.download_button(
                    label="下载 JSON",
                    data=json_str.encode("utf-8"),
                    file_name=f"{file_stem}_paragraphs.json",
                    mime="application/json",
                    use_container_width=True,
                )

            else:  # 章节拆分
                chapters = extract_chapters(doc_bytes)

                st.divider()
                st.markdown("#### 章节拆分结果")
                st.info(f"共 **{len(chapters)}** 个章节")

                for ch in chapters:
                    title = ch["title"] or "(无标题)"
                    with st.expander(f"{title}（{len(ch['paragraphs'])} 段）"):
                        st.code(ch["markdown"], language="markdown")

                # 合并所有章节的 markdown 下载
                all_md = "\n\n---\n\n".join(ch["markdown"] for ch in chapters)
                st.download_button(
                    label="下载全部章节 Markdown",
                    data=all_md.encode("utf-8"),
                    file_name=f"{file_stem}_chapters.md",
                    mime="text/markdown",
                    use_container_width=True,
                )
else:
    st.info("👆 请上传一个 .docx 文件开始使用")
