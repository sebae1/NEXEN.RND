import numpy as np
import matplotlib.pyplot as plt

from typing import Literal
from itertools import cycle
from matplotlib.axes import Axes
from matplotlib.patches import Rectangle

from util import simplify_won, COLORMAP, Config, pastel_gradient

DPI = 80
TITLE_FONTSIZE = 12
PIE_OFFSET = -0.5
PIE_LEGEND_BBOX_TO_ANCHOR = (0.8, 0.5)
HORIZONTAL_BAR_HEIGHT = 0.6
LEGEND_KWARGS = {
    "frameon": False,
    "labelspacing": 1.5
}
VGAP = 40

def hide_axis(ax: Axes):
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(left=False, bottom=False)
    ax.set_xticklabels([])
    ax.set_yticklabels([])

def draw_stacked_single_bar(ax: Axes, data: dict[str, float] = {}, title: str = ""):
    legend_labels = list(data)
    legend_labels.sort()
    for label in legend_labels:
        data[label] = max(0, data[label]) # bar 색상 통일성을 위해 0 이하도 유지해야함
    # title = ax.get_title()
    ax.clear()
    # ax.set_title(title, fontsize=TITLE_FONTSIZE)
    hide_axis(ax)
    ax.set_xlim(-1, 1.5)
    if not np.nansum(list(data.values())):
        p = ax.bar(0, -1, bottom=0, color="gray")
    else:
        labels = list(data)
        values = np.array(list(data.values()))
        normalized_values = values/np.sum(values)
        y = 0
        colors = {}
        color_cycle = cycle(COLORMAP)
        for label in labels[::-1]:
            colors[label] = next(color_cycle)
        for idx in list(range(len(labels)))[::-1]:
            label = labels[idx]
            norm_value = normalized_values[idx]
            raw_value = values[idx]
        # for label, norm_value, raw_value in zip(labels, normalized_values, values):
            p = ax.bar(0, min(0, -norm_value), label=f"{label}\n{simplify_won(raw_value)}\n{norm_value*100:0.1f}%", bottom=y, fc=colors[label])
            y -= norm_value
        ax.legend(loc="center right", **LEGEND_KWARGS)
    ax.text(
        p[0].get_x() + p[0].get_width()/2,
        0.03,
        title,
        ha="center",
        va="bottom",
        fontsize=TITLE_FONTSIZE
    )

def draw_pie(
        ax: Axes, 
        data: dict[str, float] = {}, 
        title: str = "", 
        colors: list[str] | None = None, 
        start_angle: float = 0.0,
        sort_by: Literal["label", "value"] = "label",
        desc: bool = False
    ):
    for label in list(data):
        data[label] = max(0, data[label]) # bar 색상 통일성을 위해 0 이하도 유지해야함
    # title = ax.get_title()
    ax.clear()
    # ax.set_title(title, fontsize=TITLE_FONTSIZE)
    # hide_axis(ax)
    ax.text(
        PIE_OFFSET, 0, title,
        ha="center", va="center",
        fontsize=TITLE_FONTSIZE
    )
    if not np.nansum(list(data.values())):
        wedges, _ = ax.pie([1,], colors=["gray",]) # type: ignore
        for w in wedges:
            w.set_center((PIE_OFFSET, 0))
        return
    sort_index = 0 if sort_by == "label" else 1
    data = dict(sorted(data.items(), key=lambda x: x[sort_index], reverse=desc))
    labels = data.keys()
    values = np.array(list(data.values()))
    normalized_values = values/np.sum(values)
    wedges, _ = ax.pie(normalized_values, colors=colors, startangle=start_angle) # type: ignore
    for w in wedges:
        w.set_center((PIE_OFFSET, 0))
    ax.legend(
        wedges,
        [f"{label}\n{simplify_won(raw_value)}\n{norm_value*100:0.1f}%" for label, norm_value, raw_value in zip(labels, normalized_values, values)],
        loc="center left",
        bbox_to_anchor=PIE_LEGEND_BBOX_TO_ANCHOR,
        **LEGEND_KWARGS
    )

def draw_donut(ax: Axes, data: dict[str, float] = {}, title: str = ""):
    for label in list(data):
        data[label] = max(0, data[label]) # bar 색상 통일성을 위해 0 이하도 유지해야함
    # title = ax.get_title()
    ax.clear()
    # ax.set_title(title, fontsize=TITLE_FONTSIZE)
    # hide_axis(ax)
    ax.text(
        PIE_OFFSET, 0, title,
        ha="center", va="center",
        fontsize=TITLE_FONTSIZE
    )
    if not np.nansum(list(data.values())):
        wedges, _ = ax.pie([1,], colors=["gray",], wedgeprops=dict(width=0.5)) # type: ignore
        for w in wedges:
            w.set_center((PIE_OFFSET, 0))
        return
    data = dict(sorted(data.items(), key=lambda x: x[0]))
    labels = data.keys()
    values = np.array(list(data.values()))
    normalized_values = values/np.sum(values)
    wedges, _ = ax.pie(normalized_values, wedgeprops=dict(width=0.5)) # type: ignore
    for w in wedges:
        w.set_center((PIE_OFFSET, 0))
    ax.legend(
        wedges,
        [f"{label}\n{simplify_won(raw_value)}\n{norm_value*100:0.1f}%" for label, norm_value, raw_value in zip(labels, normalized_values, values)],
        loc="center left",
        bbox_to_anchor=PIE_LEGEND_BBOX_TO_ANCHOR,
        **LEGEND_KWARGS
    )

def draw_stacked_multiple_bar(
        ax: Axes,
        data: dict[str, np.ndarray] = {},
        x_labels: list[str,] = [],
        is_percentage: bool = False,
        show_summation_on_top: bool = False
    ):
    cleaned_data: dict[str, np.ndarray] = {}
    # value legend label로 정렬하고 음수를 clean
    legend_labels = list(data)
    legend_labels.sort()
    for label in legend_labels:
        arr = np.array(data[label])
        cleaned_data[label] = arr
    title = ax.get_title()
    ax.clear()
    ax.set_title(title, fontsize=TITLE_FONTSIZE)
    hide_axis(ax)
    total = np.zeros(len(x_labels))
    for val in cleaned_data.values():
        total += val
    if not x_labels or not np.nansum(total):
        for i in range(8):
            ax.bar(i, 1, bottom=0, color="gray")
        ax.set_xlim(-1, 1.1*8)
        return
    y = np.zeros(len(x_labels))
    handles = []
    labels = []
    bar_patches: list[Rectangle] = []
    label_texts: list[str] = []
    colors = {}
    color_cycle = cycle(COLORMAP)
    for cat in legend_labels:
        colors[cat] = next(color_cycle)
    for cat in legend_labels[::-1]:
        values = cleaned_data[cat]
        eff_values = np.array(values) / (total if is_percentage else 1)
        p = ax.bar(x_labels, eff_values, label=cat, bottom=y, fc=colors[cat])
        y += eff_values
        handles.append(p)
        labels.append(cat)
        for i, (rect, v) in enumerate(zip(p.patches, values)):
            bar_patches.append(rect)
            label_texts.append(f"{v/total[i]*100:0.1f}%" if is_percentage else simplify_won(v))

    MIN_H_PX = 14
    fig = ax.figure
    axes_frac_h = ax.get_position().height
    axes_px_h = axes_frac_h * fig.get_size_inches()[1] * fig.dpi  # type: ignore
    ymin, ymax = ax.get_ylim()
    yrange = ymax - ymin if ymax > ymin else 1.0

    for rect, text in zip(bar_patches, label_texts):
        h_data = abs(rect.get_height())
        h_px = (h_data / yrange) * axes_px_h
        if not np.isfinite(h_px) or h_px < MIN_H_PX or h_data == 0:
            continue

        # 중앙 좌표
        cx = rect.get_x() + rect.get_width()/2.0
        cy = rect.get_y() + rect.get_height()/2.0

        # facecolor 밝기 기반 텍스트 색 자동 선택
        r, g, b, a = rect.get_facecolor() # type: ignore
        L = 0.2126*r + 0.7152*g + 0.0722*b # type: ignore
        txt_color = "white" if L < 0.7 else "black"

        ax.text(
            cx, cy, text,
            ha="center", va="center",
            color=txt_color, fontsize=9, clip_on=True
        )

    # 합계 텍스트 표시
    if show_summation_on_top:
        max_y = np.max(y)
        for i, (value, y_pos) in enumerate(zip(total, y)):
            ax.text(
                i, y_pos+max_y*0.02,
                simplify_won(value),
                ha="center",
                va="bottom",
                fontsize=9,
                fontweight="bold"
            )

    ax.set_xlim(-1, 1.1*len(x_labels))
    ax.set_xticks(range(len(x_labels)))
    ax.set_xticklabels(x_labels, rotation=45)
    ax.legend(handles[::-1], labels[::-1], loc="center right", **LEGEND_KWARGS)

def draw_multiple_bar(ax: Axes, title: str|None = None, data: dict[str, float] = {}, color: str|None = None):
    for key in list(data):
        val = data[key]
        if val <= 0:
            del data[key]
    ax.clear()
    ax.set_title(title or "", fontsize=TITLE_FONTSIZE)
    hide_axis(ax)
    ax.set_xlim(-1, 1.5)
    if not np.nansum(list(data.values())):
        for i in range(8):
            ax.bar(i, 1, bottom=0, color="gray")
        ax.set_xlim(-1, 1.1*8)
        return
    x_labels = list(data)
    if not x_labels:
        for i in range(8):
            ax.bar(i, 1, bottom=0, color="gray")
        ax.set_xlim(-1, 1.1*8)
        return
    values = list(data.values())
    y = np.zeros(len(x_labels))
    if color:
        p = ax.bar(x_labels, values, bottom=y, color=pastel_gradient(color, len(values)))
    else:
        p = ax.bar(x_labels, values, bottom=y)
    ax.bar_label(
        p,
        label_type="edge",
        labels=[simplify_won(val) for val in values]
    )
    ax.set_xlim(-1, 1.1*len(x_labels))
    ax.set_xticks(range(len(x_labels)))
    ax.set_xticklabels(x_labels, rotation=45)

def draw_horizontal_overlapped_bar(ax: Axes, data: dict[str, tuple[float, float]] = {}):
    ax.clear()
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(left=False, bottom=False)
    ax.set_xticks([]); ax.set_yticks([])
    ax.set_xticklabels([]); ax.set_yticklabels([])

    if not data:
        for i in range(3):
            ax.barh(i, 1, left=0, color="gray", height=HORIZONTAL_BAR_HEIGHT)
        ax.set_ylim(-0.4, 2.4)
        ax.set_xlim(0, 1.05)
        return

    # 데이터
    labels = list(data.keys())
    values = np.array(list(data.values()), dtype=float)   # shape (N, 2)
    plan = values[:, 0]
    execu = values[:, 1]
    max_val = float(np.max(values)) if values.size else 1.0
    if max_val <= 0:
        max_val = 1
    plan_n = plan / max_val
    execu_n = execu / max_val

    y = np.arange(len(labels))
    left_pad = 0.01
    right_pad = 0.01
    ax.set_xlim(-left_pad, 1.0 + right_pad)
    ax.set_ylim(-0.6, len(labels) - 1 + 0.6)

    color1 = COLORMAP[0]
    color2 = COLORMAP[2]

    arr_plan_n = np.array(plan_n)
    arr_execu_n = np.array(execu_n)
    # 계획/집행 바
    bars1 = ax.barh(
        y, arr_plan_n, left=0, height=HORIZONTAL_BAR_HEIGHT,
        label='계획', color=color1
    )
    bars2 = ax.barh(
        y, arr_execu_n, left=0, height=HORIZONTAL_BAR_HEIGHT*0.6,
        label='집행', color=color2
    )

    def fmt_money(v: float) -> str:
        try:
            return simplify_won(v)  # 사용자가 정의한 포맷터 있으면 활용
        except Exception:
            return f"{v:,.0f}"

    with np.errstate(divide="ignore", invalid="ignore"):
        rates = np.where(plan > 0, execu / plan * 100.0, 0.0)

    # 좌측에 카테고리명만 표시
    for i, lbl in enumerate(labels):
        ax.text(
            -left_pad*0.95, y[i],
            lbl, ha="right", va="center", fontsize=12, fontweight="bold"
        )

    # 바 상단에 "계획: 값  집행: 값  집행률: 값  잔액: 값" 일렬 표시
    renderer = ax.figure.canvas.get_renderer() # type: ignore
    for i, (p, e, r) in enumerate(zip(plan, execu, rates)):
        y_top = y[i] + 0.38
        # 네 개의 항목 정의
        texts = [
            ("계획: ", dict(fontweight="bold")),
            (fmt_money(p) + "    ", dict()),
            ("실적: ", dict(fontweight="bold")),
            (fmt_money(e) + "    ", dict()),
            ("집행률: ", dict(fontweight="bold")),
            (f"{r:.1f}%    ", dict()),
            ("잔액: ", dict(fontweight="bold")),
            (fmt_money(p-e), dict()),
        ]
        x_cur = 0.0
        for t, style in texts:
            txt = ax.text(
                x_cur, y_top, t,
                ha="left", va="center", fontsize=10,
                **style
            )
            bb = txt.get_window_extent(renderer=renderer)
            dx_axes = bb.width / ax.bbox.width
            x_cur += dx_axes
