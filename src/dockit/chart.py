"""Chart generation — JSON config to PNG bytes.

Pure logic module: dict in, bytes out. Supports bar, gantt, and flow charts.

Requires matplotlib and numpy: ``pip install dockit[chart]``

Usage:
    from dockit.chart import draw_bar, draw_gantt, draw_flow

    png_bytes = draw_bar({"title": "Demo", "items": [{"label": "A", "value": 10}]})
    png_bytes = draw_gantt({"phases": [{"name": "P1", "start": "2026-01", "end": "2026-06"}]})
    png_bytes = draw_flow({"layers": [{"name": "Layer1", "boxes": [{"text": "Box"}]}]})
"""

from __future__ import annotations

from datetime import datetime, timedelta
from io import BytesIO

# -- Color scheme -------------------------------------------------------------

PHASE_COLORS = [
    "#2E86AB", "#A23B72", "#F18F01", "#C73E1D",
    "#3B7A57", "#6C5B7B", "#45B7D1", "#F5A623",
]
MILESTONE_COLOR = "#E74C3C"
MILESTONE_EDGE = "#333333"
BG_COLOR = "#FAFAFA"
GRID_COLOR = "#E0E0E0"


def _phase_color(index: int) -> str:
    return PHASE_COLORS[index % len(PHASE_COLORS)]


# -- Font setup ---------------------------------------------------------------


def _setup_fonts():
    """Configure matplotlib for CJK text rendering."""
    import matplotlib as mpl
    import matplotlib.pyplot as plt

    candidates = ["Arial Unicode MS", "PingFang SC", "Microsoft YaHei", "SimHei", "DejaVu Sans"]
    plt.rcParams.update({"font.sans-serif": candidates, "axes.unicode_minus": False, "font.family": "sans-serif"})
    mpl.rc("font", **{"sans-serif": candidates})


def _to_bytes(fig, dpi: int) -> bytes:
    """Save figure to PNG bytes and close it."""
    import matplotlib.pyplot as plt

    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white", edgecolor="none")
    plt.close(fig)
    return buf.getvalue()


# -- Bar charts ---------------------------------------------------------------


def _draw_horizontal_bar(config: dict, dpi: int) -> bytes:
    import matplotlib.pyplot as plt
    import numpy as np

    items = config["items"]
    title = config.get("title", "")
    unit = config.get("unit", "")
    figsize = tuple(config.get("figsize", [12, 6]))
    show_total = config.get("show_total", False)

    labels = [it["label"] for it in items]
    values = [it["value"] for it in items]
    percents = [it.get("percent") for it in items]
    colors = [it.get("color") or _phase_color(i) for i, it in enumerate(items)]

    fig, ax = plt.subplots(figsize=figsize)
    fig.set_facecolor("white")
    ax.set_facecolor(BG_COLOR)

    y_pos = np.arange(len(items))
    bars = ax.barh(y_pos, values, height=0.55, color=colors, edgecolor="#333333", linewidth=0.8, alpha=0.9)

    max_val = max(values)
    for i, (_, val) in enumerate(zip(bars, values, strict=False)):
        pct = percents[i]
        label_text = f"{val} {unit}（{pct}%）" if pct is not None else f"{val} {unit}"
        if val > max_val * 0.3:
            ax.text(val / 2, i, label_text, ha="center", va="center", fontsize=10, fontweight="bold", color="white")
        else:
            ax.text(val + max_val * 0.01, i, label_text, ha="left", va="center", fontsize=9, color="#333333")

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=11, fontweight="bold")
    ax.invert_yaxis()
    ax.set_xlabel(unit, fontsize=11)
    ax.set_title(title, fontsize=16, fontweight="bold", pad=15)

    if show_total:
        total_val = sum(values)
        total_label = config.get("total_label", "合计")
        ax.text(
            0.98, 0.02, f"{total_label}：{total_val:.2f} {unit}",
            transform=ax.transAxes, ha="right", va="bottom", fontsize=11, fontweight="bold", color="#666666",
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", edgecolor="#CCCCCC", alpha=0.9),
        )

    ax.grid(True, axis="x", color=GRID_COLOR, linestyle="-", linewidth=0.6, alpha=0.7)
    ax.grid(False, axis="y")
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    plt.tight_layout()
    return _to_bytes(fig, dpi)


def _draw_vertical_bar(config: dict, dpi: int) -> bytes:
    import matplotlib.pyplot as plt
    import numpy as np

    items = config["items"]
    title = config.get("title", "")
    unit = config.get("unit", "")
    figsize = tuple(config.get("figsize", [12, 7]))

    labels = [it["label"] for it in items]
    values = [it["value"] for it in items]
    colors = [it.get("color") or _phase_color(i) for i, it in enumerate(items)]

    fig, ax = plt.subplots(figsize=figsize)
    fig.set_facecolor("white")
    ax.set_facecolor(BG_COLOR)

    x_pos = np.arange(len(items))
    bars = ax.bar(x_pos, values, width=0.55, color=colors, edgecolor="#333333", linewidth=0.8, alpha=0.9)

    for bar, val in zip(bars, values, strict=False):
        pct_item = next((it for it in items if it["value"] == val), None)
        pct = pct_item.get("percent") if pct_item else None
        text = f"{val}\n({pct}%)" if pct is not None else f"{val}"
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), text,
                ha="center", va="bottom", fontsize=9, fontweight="bold", color="#333333")

    ax.set_xticks(x_pos)
    ax.set_xticklabels(labels, fontsize=10, fontweight="bold")
    ax.set_ylabel(unit, fontsize=11)
    ax.set_title(title, fontsize=16, fontweight="bold", pad=15)
    ax.grid(True, axis="y", color=GRID_COLOR, linestyle="-", linewidth=0.6, alpha=0.7)
    ax.grid(False, axis="x")
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    plt.tight_layout()
    return _to_bytes(fig, dpi)


def _draw_grouped_bar(config: dict, dpi: int) -> bytes:
    import matplotlib.pyplot as plt
    import numpy as np

    groups = config["groups"]
    series = config["series"]
    title = config.get("title", "")
    unit = config.get("unit", "")
    figsize = tuple(config.get("figsize", [14, 7]))

    n_groups = len(groups)
    n_series = len(series)
    bar_width = 0.8 / n_series
    x = np.arange(n_groups)

    fig, ax = plt.subplots(figsize=figsize)
    fig.set_facecolor("white")
    ax.set_facecolor(BG_COLOR)

    for i, s in enumerate(series):
        offset = (i - n_series / 2 + 0.5) * bar_width
        color = s.get("color") or _phase_color(i)
        bars = ax.bar(x + offset, s["values"], bar_width * 0.9, label=s["name"],
                      color=color, edgecolor="#333333", linewidth=0.6, alpha=0.9)
        for bar, val in zip(bars, s["values"], strict=False):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), f"{val}",
                    ha="center", va="bottom", fontsize=8)

    ax.set_xticks(x)
    ax.set_xticklabels(groups, fontsize=10, fontweight="bold")
    ax.set_ylabel(unit, fontsize=11)
    ax.set_title(title, fontsize=16, fontweight="bold", pad=15)
    ax.legend(fontsize=10, framealpha=0.9)
    ax.grid(True, axis="y", color=GRID_COLOR, linestyle="-", linewidth=0.6, alpha=0.7)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    plt.tight_layout()
    return _to_bytes(fig, dpi)


# -- Gantt chart --------------------------------------------------------------


def _parse_date(s: str) -> datetime:
    for fmt in ("%Y-%m-%d", "%Y-%m"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"Cannot parse date: {s}")


def draw_gantt(config: dict, *, dpi: int = 300) -> bytes:
    """Generate a Gantt chart from config dict.

    Args:
        config: Dict with "phases" list (each has name, start, end) and optional "milestones".
        dpi: Output resolution.

    Returns:
        PNG image as bytes.
    """
    import matplotlib.dates as mdates
    import matplotlib.patches as mpatches
    import matplotlib.pyplot as plt

    _setup_fonts()
    phases = config.get("phases", [])
    milestones = config.get("milestones", [])
    title = config.get("title", "")
    subtitle = config.get("subtitle", "")
    figsize = tuple(config.get("figsize", [16, 8]))

    if not phases:
        raise ValueError("Config must contain non-empty 'phases' list")

    for i, p in enumerate(phases):
        p["_start"] = _parse_date(p["start"])
        p["_end"] = _parse_date(p["end"])
        p["_color"] = p.get("color") or _phase_color(i)
        p["_duration"] = (p["_end"] - p["_start"]).days
    for m in milestones:
        m["_date"] = _parse_date(m["date"])

    all_dates = [p["_start"] for p in phases] + [p["_end"] for p in phases]
    if milestones:
        all_dates += [m["_date"] for m in milestones]
    date_min = min(all_dates) - timedelta(days=7)
    date_max = max(all_dates) + timedelta(days=7)

    fig, ax = plt.subplots(figsize=figsize)
    fig.set_facecolor("white")
    ax.set_facecolor(BG_COLOR)

    n_phases = len(phases)
    bar_height = 0.6

    for i, p in enumerate(phases):
        ax.barh(i, p["_duration"], left=p["_start"], height=bar_height, color=p["_color"],
                alpha=0.85, edgecolor="#333333", linewidth=0.8, zorder=2)
        short = p.get("short", p["name"])
        date_range = f"{p['_start'].strftime('%m.%d')}—{p['_end'].strftime('%m.%d')}"
        mid = p["_start"] + timedelta(days=p["_duration"] / 2)
        if p["_duration"] > 30:
            ax.text(mid, i, f"{short}\n{date_range}", ha="center", va="center",
                    fontsize=9, fontweight="bold", color="white", zorder=3)
        else:
            ax.text(p["_end"] + timedelta(days=2), i, f"{short} ({date_range})",
                    ha="left", va="center", fontsize=8, color="#333333", zorder=3)

    for m in milestones:
        y = -0.8
        ax.scatter(m["_date"], y, marker="D", s=120, color=MILESTONE_COLOR,
                   edgecolor=MILESTONE_EDGE, linewidth=1.5, zorder=4)
        label_text = m["name"]
        if m.get("label"):
            label_text += f"\n({m['label']})"
        ax.annotate(label_text, (m["_date"], y), textcoords="offset points", xytext=(0, -18),
                    ha="center", va="top", fontsize=7.5, fontweight="bold", color="#333333")
        ax.axvline(m["_date"], color=MILESTONE_COLOR, linestyle="--", alpha=0.3, linewidth=0.8, zorder=1)

    ax.set_yticks(range(n_phases))
    ax.set_yticklabels([p.get("short", p["name"]) for p in phases], fontsize=11, fontweight="bold")
    ax.invert_yaxis()
    ax.set_ylim(n_phases - 0.5, -1.5 if milestones else -0.5)
    ax.set_xlim(date_min, date_max)
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y年%m月"))
    ax.xaxis.set_minor_locator(mdates.WeekdayLocator(byweekday=0))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right", fontsize=10)

    ax.grid(True, axis="x", which="major", color=GRID_COLOR, linestyle="-", linewidth=0.8, alpha=0.7)
    ax.grid(True, axis="x", which="minor", color=GRID_COLOR, linestyle=":", linewidth=0.4, alpha=0.5)
    ax.grid(False, axis="y")

    title_text = f"{title}\n{subtitle}" if subtitle else title
    ax.set_title(title_text, fontsize=16, fontweight="bold", pad=15)

    legend_patches = [mpatches.Patch(color=p["_color"], label=p.get("short", p["name"])) for p in phases]
    if milestones:
        legend_patches.append(
            plt.scatter([], [], marker="D", s=80, color=MILESTONE_COLOR, edgecolor=MILESTONE_EDGE, label="Milestone"))
    ax.legend(handles=legend_patches, loc="upper right", fontsize=9, framealpha=0.9)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    plt.tight_layout()
    return _to_bytes(fig, dpi)


# -- Flow / layer charts ------------------------------------------------------


def _draw_layers(config: dict, dpi: int) -> bytes:
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch

    layers = config["layers"]
    title = config.get("title", "")
    figsize = tuple(config.get("figsize", [14, 12]))
    bottom_bar = config.get("bottom_bar")

    n_layers = len(layers)
    fig, ax = plt.subplots(figsize=figsize)
    fig.set_facecolor("white")
    ax.set_facecolor("white")

    layer_height = 1.0
    layer_gap = 0.5
    total_height = n_layers * layer_height + (n_layers - 1) * layer_gap
    if bottom_bar:
        total_height += layer_height + layer_gap

    canvas_width = 10.0
    margin_x = 0.5
    box_margin = 0.3

    ax.set_xlim(-margin_x, canvas_width + margin_x)
    ax.set_ylim(-0.5, total_height + 1.0)
    ax.set_aspect("equal")
    ax.axis("off")

    ax.text(canvas_width / 2, total_height + 0.6, title, ha="center", va="center", fontsize=16, fontweight="bold")

    for layer_idx, layer in enumerate(layers):
        y_top = total_height - layer_idx * (layer_height + layer_gap)
        y_bottom = y_top - layer_height
        y_center = (y_top + y_bottom) / 2
        color = layer.get("color") or _phase_color(layer_idx)

        layer_bg = FancyBboxPatch((0, y_bottom), canvas_width, layer_height,
                                  boxstyle="round,pad=0.05", facecolor=color, alpha=0.12,
                                  edgecolor=color, linewidth=1.5)
        ax.add_patch(layer_bg)
        ax.text(0.2, y_top - 0.12, layer["name"], ha="left", va="top", fontsize=10, fontweight="bold", color=color)

        boxes = layer.get("boxes", [])
        if boxes:
            n_boxes = len(boxes)
            box_width = (canvas_width - box_margin * (n_boxes + 1)) / n_boxes
            box_h = 0.5
            box_y = y_center - box_h / 2 - 0.05
            for bi, box in enumerate(boxes):
                box_x = box_margin + bi * (box_width + box_margin)
                box_color = box.get("color", color)
                rect = FancyBboxPatch((box_x, box_y), box_width, box_h,
                                      boxstyle="round,pad=0.08", facecolor="white",
                                      edgecolor=box_color, linewidth=1.2)
                ax.add_patch(rect)
                ax.text(box_x + box_width / 2, box_y + box_h / 2, box["text"],
                        ha="center", va="center", fontsize=9, fontweight="bold", color="#333333")

        output_text = layer.get("output")
        if output_text:
            ax.text(canvas_width - 0.2, y_bottom + 0.08, f"\u2192 {output_text}",
                    ha="right", va="bottom", fontsize=8, fontstyle="italic", color="#666666")

        if layer_idx < n_layers - 1:
            arrow_y_start = y_bottom - 0.05
            arrow_y_end = y_bottom - layer_gap + 0.05
            ax.annotate("", xy=(canvas_width / 2, arrow_y_end), xytext=(canvas_width / 2, arrow_y_start),
                        arrowprops=dict(arrowstyle="-|>", color="#666666", lw=2, mutation_scale=20))

    if bottom_bar:
        bar_y = -0.3
        bar_h = 0.6
        bar_bg = FancyBboxPatch((0, bar_y), canvas_width, bar_h, boxstyle="round,pad=0.05",
                                facecolor="#555555", alpha=0.15, edgecolor="#555555", linewidth=1.5)
        ax.add_patch(bar_bg)
        ax.text(canvas_width / 2, bar_y + bar_h / 2, bottom_bar,
                ha="center", va="center", fontsize=10, fontweight="bold", color="#555555")

    plt.tight_layout()
    return _to_bytes(fig, dpi)


def _draw_flow_steps(config: dict, dpi: int) -> bytes:
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch

    steps = config["steps"]
    title = config.get("title", "")
    figsize = tuple(config.get("figsize", [16, 5]))

    n_steps = len(steps)
    fig, ax = plt.subplots(figsize=figsize)
    fig.set_facecolor("white")
    ax.set_facecolor("white")
    ax.axis("off")

    box_w = 1.8
    box_h = 1.2
    gap = 0.8
    total_w = n_steps * box_w + (n_steps - 1) * gap
    start_x = (figsize[0] - total_w) / 2

    ax.set_xlim(0, figsize[0])
    ax.set_ylim(0, figsize[1])
    ax.text(figsize[0] / 2, figsize[1] - 0.5, title, ha="center", va="center", fontsize=16, fontweight="bold")

    y_center = figsize[1] / 2 - 0.2

    for i, step in enumerate(steps):
        x = start_x + i * (box_w + gap)
        color = step.get("color") or _phase_color(i)

        circle = plt.Circle((x + box_w / 2, y_center + box_h / 2 + 0.35), 0.25, color=color, zorder=3)
        ax.add_patch(circle)
        ax.text(x + box_w / 2, y_center + box_h / 2 + 0.35, str(i + 1),
                ha="center", va="center", fontsize=12, fontweight="bold", color="white", zorder=4)

        rect = FancyBboxPatch((x, y_center - box_h / 2), box_w, box_h, boxstyle="round,pad=0.1",
                              facecolor=color, alpha=0.15, edgecolor=color, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x + box_w / 2, y_center + 0.15, step["name"],
                ha="center", va="center", fontsize=10, fontweight="bold", color="#333333")

        desc = step.get("desc", "")
        if desc:
            ax.text(x + box_w / 2, y_center - 0.2, desc,
                    ha="center", va="center", fontsize=8, color="#666666", wrap=True)

        if i < n_steps - 1:
            ax.annotate("", xy=(x + box_w + gap - 0.1, y_center), xytext=(x + box_w + 0.1, y_center),
                        arrowprops=dict(arrowstyle="-|>", color="#888888", lw=2, mutation_scale=18))

    plt.tight_layout()
    return _to_bytes(fig, dpi)


# -- Public API ---------------------------------------------------------------


def draw_bar(config: dict, *, dpi: int = 300) -> bytes:
    """Generate a bar chart from config dict.

    Args:
        config: Dict with "items" list (each has label, value, optional percent/color).
            Set "type" to "horizontal" (default), "vertical", or "grouped".
            For "grouped": use "groups" list and "series" list instead of "items".
        dpi: Output resolution.

    Returns:
        PNG image as bytes.
    """
    _setup_fonts()
    chart_type = config.get("type", "horizontal")
    dispatch = {"horizontal": _draw_horizontal_bar, "vertical": _draw_vertical_bar, "grouped": _draw_grouped_bar}
    handler = dispatch.get(chart_type)
    if not handler:
        raise ValueError(f"Unknown bar type: {chart_type} (expected: {', '.join(dispatch)})")
    return handler(config, dpi)


def draw_flow(config: dict, *, dpi: int = 300) -> bytes:
    """Generate a flow or layer chart from config dict.

    Args:
        config: Dict with "type" set to "layers" or "flow".
            For "layers": use "layers" list (each has name, boxes).
            For "flow": use "steps" list (each has name, optional desc).
        dpi: Output resolution.

    Returns:
        PNG image as bytes.
    """
    _setup_fonts()
    chart_type = config.get("type", "layers")
    dispatch = {"layers": _draw_layers, "flow": _draw_flow_steps}
    handler = dispatch.get(chart_type)
    if not handler:
        raise ValueError(f"Unknown flow type: {chart_type} (expected: {', '.join(dispatch)})")
    return handler(config, dpi)
