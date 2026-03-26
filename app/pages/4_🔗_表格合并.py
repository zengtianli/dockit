"""表格合并 — 上传多个 Excel 文件，纵向拼接或横向合并。"""

import streamlit as st

st.set_page_config(page_title="表格合并", page_icon="🔗")

st.title("🔗 表格合并")
st.caption("上传多个 Excel 文件，按行拼接或按关键列合并")

# ---------------------------------------------------------------------------
# 1. File upload
# ---------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "上传 Excel 文件（支持多选）",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("请上传至少两个 .xlsx 文件开始合并。")
    st.stop()

if len(uploaded_files) < 2:
    st.warning("至少需要上传两个文件才能合并。")
    st.stop()

# ---------------------------------------------------------------------------
# 2. Read all files into DataFrames
# ---------------------------------------------------------------------------
import pandas as pd
from io import BytesIO

dfs: list[pd.DataFrame] = []
file_names: list[str] = []

for f in uploaded_files:
    raw = f.read()
    try:
        df = pd.read_excel(BytesIO(raw))
    except Exception as e:
        st.error(f"读取 **{f.name}** 失败：{e}")
        st.stop()
    dfs.append(df)
    file_names.append(f.name)

# ---------------------------------------------------------------------------
# 3. Preview each uploaded file
# ---------------------------------------------------------------------------
st.subheader("文件预览")

tabs = st.tabs(file_names)
for tab, name, df in zip(tabs, file_names, dfs):
    with tab:
        st.markdown(f"**{name}**　—　{len(df)} 行 × {len(df.columns)} 列")
        st.dataframe(df.head(5), use_container_width=True)

# ---------------------------------------------------------------------------
# 4. Merge settings
# ---------------------------------------------------------------------------
st.divider()
st.subheader("合并设置")

mode = st.radio(
    "合并方式",
    options=["纵向拼接", "横向合并"],
    horizontal=True,
    help="纵向拼接：将所有表格按行堆叠。横向合并：按关键列做外连接。",
)

key_column: str | None = None

if mode == "横向合并":
    # Use headers from the first file as candidates
    first_cols = list(dfs[0].columns)
    if not first_cols:
        st.error("第一个文件没有列标题，无法进行横向合并。")
        st.stop()

    key_column = st.selectbox(
        "选择关键列（Key Column）",
        options=first_cols,
        help="所有文件将按此列做外连接（outer join）。请确保每个文件都包含该列。",
    )

    # Validate that every file contains the key column
    missing = [name for name, df in zip(file_names, dfs) if key_column not in df.columns]
    if missing:
        st.error(f"以下文件缺少关键列 **{key_column}**：{', '.join(missing)}")
        st.stop()

# ---------------------------------------------------------------------------
# 5. Process
# ---------------------------------------------------------------------------
if st.button("开始合并", type="primary", use_container_width=True):
    with st.spinner("合并中…"):
        if mode == "纵向拼接":
            result = pd.concat(dfs, ignore_index=True)
        else:
            result = dfs[0]
            for df in dfs[1:]:
                result = pd.merge(result, df, on=key_column, how="outer")

    # Store in session state so it persists across reruns
    st.session_state["merge_result"] = result

# ---------------------------------------------------------------------------
# 6. Result preview & download
# ---------------------------------------------------------------------------
if "merge_result" in st.session_state:
    result = st.session_state["merge_result"]

    st.divider()
    st.subheader("合并结果")

    # Stats
    col1, col2, col3 = st.columns(3)
    col1.metric("总行数", f"{len(result):,}")
    col2.metric("总列数", f"{len(result.columns):,}")
    col3.metric("源文件数", f"{len(uploaded_files)}")

    st.dataframe(result.head(50), use_container_width=True)

    if len(result) > 50:
        st.caption(f"仅显示前 50 行，完整数据共 {len(result):,} 行。")

    # Download
    buf = BytesIO()
    result.to_excel(buf, index=False, engine="openpyxl")
    output_bytes = buf.getvalue()

    st.download_button(
        label="下载合并结果（.xlsx）",
        data=output_bytes,
        file_name="合并结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
