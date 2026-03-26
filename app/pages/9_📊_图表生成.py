"""图表生成 — 通过 JSON 配置生成柱状图、甘特图、流程图。"""

import json

import streamlit as st

st.set_page_config(page_title="图表生成", page_icon="📊")

st.title("📊 图表生成")
st.markdown("选择图表类型，填写 JSON 配置，生成 PNG 图表。")

# -- Chart type --
chart_type = st.selectbox("图表类型", ["柱状图", "甘特图", "流程图"])

# -- Templates --
TEMPLATES = {
    "柱状图": json.dumps(
        {
            "type": "horizontal",
            "title": "示例柱状图",
            "unit": "万元",
            "items": [
                {"label": "项目A", "value": 120, "percent": 40},
                {"label": "项目B", "value": 90, "percent": 30},
                {"label": "项目C", "value": 60, "percent": 20},
                {"label": "项目D", "value": 30, "percent": 10},
            ],
        },
        ensure_ascii=False,
        indent=2,
    ),
    "甘特图": json.dumps(
        {
            "title": "项目进度",
            "phases": [
                {"name": "前期准备", "start": "2026-01", "end": "2026-03"},
                {"name": "实施阶段", "start": "2026-03", "end": "2026-08"},
                {"name": "验收总结", "start": "2026-08", "end": "2026-10"},
            ],
            "milestones": [
                {"name": "中期检查", "date": "2026-06-01"},
            ],
        },
        ensure_ascii=False,
        indent=2,
    ),
    "流程图": json.dumps(
        {
            "type": "layers",
            "title": "系统架构",
            "layers": [
                {"name": "用户层", "boxes": [{"text": "Web 端"}, {"text": "移动端"}]},
                {"name": "服务层", "boxes": [{"text": "API 网关"}, {"text": "业务逻辑"}, {"text": "认证服务"}]},
                {"name": "数据层", "boxes": [{"text": "PostgreSQL"}, {"text": "Redis"}]},
            ],
        },
        ensure_ascii=False,
        indent=2,
    ),
}

TYPE_HELP = {
    "柱状图": "type 可选：`horizontal`（水平）、`vertical`（垂直）、`grouped`（分组）",
    "甘特图": "phases 为阶段列表，milestones 为里程碑（可选）",
    "流程图": "type 可选：`layers`（分层架构图）、`flow`（流程步骤图）",
}

st.caption(TYPE_HELP[chart_type])

# -- Config input --
config_text = st.text_area(
    "JSON 配置",
    value=TEMPLATES[chart_type],
    height=300,
)

# -- Generate --
if st.button("生成图表", type="primary", use_container_width=True):
    try:
        config = json.loads(config_text)
    except json.JSONDecodeError as e:
        st.error(f"JSON 格式错误: {e}")
        st.stop()

    try:
        from dockit.chart import draw_bar, draw_flow, draw_gantt

        with st.spinner("正在生成图表..."):
            if chart_type == "柱状图":
                png_bytes = draw_bar(config)
            elif chart_type == "甘特图":
                png_bytes = draw_gantt(config)
            else:
                png_bytes = draw_flow(config)

        # -- Display --
        st.divider()
        st.markdown("#### 生成结果")
        st.image(png_bytes, use_container_width=True)

        # -- Download --
        st.download_button(
            label="下载 PNG",
            data=png_bytes,
            file_name=f"chart_{chart_type}.png",
            mime="image/png",
            use_container_width=True,
        )

    except ImportError:
        st.error("图表功能需要安装额外依赖：`pip install dockit[chart]`（matplotlib + numpy）")
    except Exception as e:
        st.error(f"生成失败: {e}")
