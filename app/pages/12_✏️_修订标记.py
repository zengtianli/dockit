"""修订标记 — 读取修订/批注，或批量写入修订标记。"""

import json

import streamlit as st

from dockit.docx import apply_review, read_changes

st.set_page_config(page_title="修订标记", page_icon="✏️")

st.title("✏️ Word 修订标记")
st.markdown("读取 Word 文档中的修订和批注，或批量写入查找替换修订。")

mode = st.radio("模式", ["读取修订", "写入修订"], horizontal=True)

if mode == "读取修订":
    uploaded = st.file_uploader("上传 Word 文件", type=["docx"], key="read")

    if uploaded and st.button("读取修订", type="primary", use_container_width=True):
        with st.spinner("正在读取修订..."):
            doc_bytes = uploaded.read()
            result = read_changes(doc_bytes)

        st.divider()
        changes = result["changes"]
        comments = result["comments"]

        c1, c2 = st.columns(2)
        c1.metric("修订", len(changes))
        c2.metric("批注", len(comments))

        if changes:
            st.markdown("#### 修订记录")
            for i, c in enumerate(changes, 1):
                icon = "插入" if c["type"] == "insert" else "删除"
                date_str = c["date"][:10] if c["date"] else ""
                st.markdown(f"{i}. **{icon}** | {c['author']} | {date_str}")
                st.markdown(f"> {c['text']}")

        if comments:
            st.markdown("#### 批注")
            for c in comments:
                date_str = c["date"][:10] if c["date"] else ""
                st.markdown(f"- **[{c['id']}]** {c['author']} ({date_str}): {c['text']}")

        if not changes and not comments:
            st.info("文档中没有修订或批注")

        # 下载 JSON
        st.download_button(
            label="下载 JSON",
            data=json.dumps(result, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name="revisions.json",
            mime="application/json",
            use_container_width=True,
        )

else:  # 写入修订
    uploaded = st.file_uploader("上传 Word 文件", type=["docx"], key="write")

    st.markdown("#### 替换规则")
    st.caption('JSON 格式：`[{"find": "旧文本", "replace": "新文本", "comment": "可选批注"}]`')

    rules_text = st.text_area(
        "规则 JSON",
        height=200,
        value='[\n  {"find": "旧文本", "replace": "新文本", "comment": "修改说明"}\n]',
    )

    author = st.text_input("作者", value="DocKit")

    if uploaded and st.button("写入修订", type="primary", use_container_width=True):
        try:
            rules = json.loads(rules_text)
        except json.JSONDecodeError as e:
            st.error(f"JSON 格式错误: {e}")
            st.stop()

        with st.spinner("正在写入修订标记..."):
            doc_bytes = uploaded.read()
            result = apply_review(doc_bytes, rules, author=author)

        st.divider()
        if result.count > 0:
            st.success(f"完成！共 {result.count} 处替换已写入修订标记")
        else:
            st.warning("未找到匹配的文本，无替换发生")

        file_stem = uploaded.name.rsplit(".", 1)[0]
        st.download_button(
            label="下载修订后的文件",
            data=result.data,
            file_name=f"{file_stem}_reviewed.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
