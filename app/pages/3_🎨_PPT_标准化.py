"""PPT 标准化 — 字体统一、文本格式修复、表格样式."""

import streamlit as st

from dockit.pptx import format_text, set_font, set_table_style, standardize

st.set_page_config(page_title="PPT 标准化", page_icon="🎨")

st.title("🎨 PPT 标准化")
st.caption("字体统一 · 文本格式修复 · 表格样式")

# -- File upload ---------------------------------------------------------------

uploaded = st.file_uploader("上传 PowerPoint 文件", type=["pptx"])

# -- Mode selector -------------------------------------------------------------

MODE_OPTIONS = {
    "一键标准化": "all",
    "字体统一": "font",
    "文本格式修复": "format",
    "表格样式": "table",
}

mode_label = st.radio(
    "处理模式",
    list(MODE_OPTIONS.keys()),
    index=0,
    horizontal=True,
    help="推荐使用「一键标准化」，一次完成所有处理",
)
mode = MODE_OPTIONS[mode_label]

# -- Mode description ----------------------------------------------------------

MODE_INFO = {
    "all": "将依次执行：文本格式修复 → 字体统一 → 表格样式设置，一步到位完成所有标准化操作。",
    "font": "将 PPT 中所有文本（包括文本框、表格、母版）的字体统一为指定字体。",
    "format": "修复引号、标点、单位等文本格式问题（如中英文标点混用、单位格式不规范等）。",
    "table": "设置表格的标题行、首列和隔行样式，使表格风格统一。",
}

st.info(MODE_INFO[mode])

# -- Font name input (only for font/all modes) ---------------------------------

font_name = "Microsoft YaHei"
if mode in ("all", "font"):
    font_name = st.text_input("目标字体", value="Microsoft YaHei")

# -- Process button ------------------------------------------------------------

if uploaded and st.button("开始处理", type="primary", use_container_width=True):
    pptx_bytes = uploaded.read()

    with st.spinner("正在处理，请稍候..."):
        if mode == "all":
            result = standardize(pptx_bytes, font_name=font_name)
        elif mode == "font":
            result = set_font(pptx_bytes, font_name=font_name)
        elif mode == "format":
            result = format_text(pptx_bytes)
        else:  # table
            result = set_table_style(pptx_bytes)

    # -- Stats display ---------------------------------------------------------

    st.success("处理完成！")

    STAT_LABELS = {
        "runs": ("文本段", "处理的文本段数"),
        "shapes": ("形状", "处理的形状数"),
        "tables": ("表格单元", "处理的表格单元格数"),
        "quotes": ("引号修复", "修复的引号数量"),
        "punctuation": ("标点修复", "修复的标点数量"),
        "units": ("单位修复", "修复的单位格式数量"),
        "tables_styled": ("表格样式", "设置样式的表格数"),
    }

    stats = result.stats
    visible = {k: v for k, v in stats.items() if k in STAT_LABELS}

    if visible:
        cols = st.columns(len(visible))
        for col, (key, value) in zip(cols, visible.items()):
            label, help_text = STAT_LABELS[key]
            col.metric(label=label, value=value, help=help_text)

    # -- Download button -------------------------------------------------------

    out_name = uploaded.name.rsplit(".", 1)[0] + "_标准化.pptx"

    st.download_button(
        label="下载处理结果",
        data=result.data,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )
