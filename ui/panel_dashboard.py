import os
import numpy as np
import matplotlib.pyplot as plt
import wx

from io import BytesIO
from datetime import datetime
from traceback import format_exc
from typing import Literal, Callable
from threading import Thread
from itertools import cycle
from collections import defaultdict

from PIL import Image
from wx.lib.scrolledpanel import ScrolledPanel
from wx.lib.newevent import NewCommandEvent
from matplotlib.axes import Axes
from matplotlib.patches import Rectangle

from util import simplify_won, COLORMAP, Config, pastel_gradient
from db import CostCategory, CostElement, CostCtr, LoadedData
from ai import (
    _CostCategory, _CostElement, _CostCtr, _BudgetByCtr, _BudgetByElement,
    get_prompts_for_ai, analyze_by_claude, analyze_by_gpt
)
from ui.component import PanelAspectRatio, PanelCanvas, \
    FONT_COLOR_HIGH_PORTION, FONT_COLOR_MID_PORTION, FONT_COLOR_LOW_PORTION, FONT_COLOR_NEGATIVE_VALUE, \
    OPENAI_MARK_SVG, CLAUDE_MARK_SVG
from ui.component.ai_analysis import DialogAIResult, DialogModels

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
        wedges, _ = ax.pie([1,], colors=["gray",])
        for w in wedges:
            w.set_center((PIE_OFFSET, 0))
        return
    sort_index = 0 if sort_by == "label" else 1
    data = dict(sorted(data.items(), key=lambda x: x[sort_index], reverse=desc))
    labels = data.keys()
    values = np.array(list(data.values()))
    normalized_values = values/np.sum(values)
    wedges, _ = ax.pie(normalized_values, colors=colors, startangle=start_angle)
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
        wedges, _ = ax.pie([1,], colors=["gray",], wedgeprops=dict(width=0.5))
        for w in wedges:
            w.set_center((PIE_OFFSET, 0))
        return
    data = dict(sorted(data.items(), key=lambda x: x[0]))
    labels = data.keys()
    values = np.array(list(data.values()))
    normalized_values = values/np.sum(values)
    wedges, _ = ax.pie(normalized_values, wedgeprops=dict(width=0.5))
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
        # arr[arr < 0] = 0 # TODO 음수 처리
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
    axes_px_h = axes_frac_h * fig.get_size_inches()[1] * fig.dpi
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
        r, g, b, a = rect.get_facecolor()
        L = 0.2126*r + 0.7152*g + 0.0722*b
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

def draw_pie_with_annotation(ax: Axes, data: dict[str, float] = {}):
    """
    data: {
        "개발비": {category (str): value: (int)},
        "고정비": {category (str): value: (int)},
        "인건비": {category (str): value: (int)}
    }
    """
    for key in list(data):
        val = data[key]
        if val <= 0:
            del data[key]
    colormaps = {
        "개발비": COLORMAP[0],
        "고정비": COLORMAP[1],
        "인건비": COLORMAP[2]
    }

    # outer_labels = []
    # outer_sizes = []
    labels = []
    sizes = []
    colors = []
    for bigcat, subcats in data.items():
        cmap = plt.get_cmap(colormaps[bigcat])
        sub_keys = list(subcats.keys())
        sub_vals = list(subcats.values())
        color = cmap(np.linspace(0.4, 0.9, len(sub_vals)))
        labels.extend(sub_keys)
        sizes.extend(sub_vals)
        colors.extend(color)

    title = ax.get_title()
    ax.clear()
    ax.set_title(title, fontsize=TITLE_FONTSIZE)
    # hide_axis(ax)
    total = 0
    for cat, items in data.items():
        total += np.nansum(list(items.values()))
    if not total:
        wedges, _ = ax.pie([1,], colors=["gray",])
        for w in wedges:
            w.set_center((PIE_OFFSET, 0))
        return

    wedges, _ = ax.pie(
        sizes,
        colors=colors,
        startangle=-40,
    )

    # 1단계: 모든 wedge의 각도와 위치 정보 수집
    wedge_info = []
    for i, p in enumerate(wedges):
        # 각 wedge의 중심 각도
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        # 각도를 라디안으로 변환
        ang_rad = np.deg2rad(ang)
        # wedge 중심의 x, y 좌표
        y = np.sin(ang_rad)
        x = np.cos(ang_rad)
        
        wedge_info.append({
            'index': i,
            'angle': ang,
            'angle_rad': ang_rad,
            'x': x,
            'y': y,
            'size': sizes[i],
            'label': labels[i]
        })

    # 2단계: 각도 순으로 정렬하여 인접한 wedge 파악
    wedge_info_sorted = sorted(wedge_info, key=lambda w: w['angle'])

    # 3단계: 텍스트 위치 계산 - radius를 조정하여 겹침 방지
    text_positions = []
    min_text_distance = 0.15  # 텍스트 간 최소 거리
    base_radius = 1.3  # 기본 반지름
    radius_increment = 0.2  # 반지름 증가량

    for idx, w in enumerate(wedge_info_sorted):
        x_base = w['x']
        y_base = w['y']
        
        # 초기 radius 설정
        radius = base_radius
        found_position = False
        max_attempts = 10  # 최대 시도 횟수
        
        # 적절한 위치를 찾을 때까지 radius 증가
        for attempt in range(max_attempts):
            # 현재 radius로 텍스트 위치 계산
            text_x = radius * x_base
            text_y = radius * y_base
            
            # 다른 텍스트들과 거리 확인
            collision = False
            for prev_pos in text_positions:
                # 유클리드 거리 계산
                distance = np.sqrt((text_x - prev_pos['x'])**2 + (text_y - prev_pos['y'])**2)
                if distance < min_text_distance:
                    collision = True
                    break
            
            if not collision:
                found_position = True
                break
            else:
                # 충돌이 있으면 radius 증가
                radius += radius_increment
        
        text_positions.append({
            'x': text_x,
            'y': text_y,
            'radius': radius,
            'index': w['index']
        })

    # 4단계: 연결선을 먼저 그리기 (낮은 z-order)
    for w in wedge_info:
        i = w['index']
        pos = text_positions[i]
        
        # 파이 차트 가장자리 점 (반지름 1)
        edge_x = w['x'] * 1.0
        edge_y = w['y'] * 1.0
        
        # 텍스트 위치
        text_x = pos['x']
        text_y = pos['y']
        
        # 연결선 그리기 - 파이 중심(0,0)을 지나는 직선
        # 실제로는 파이 가장자리부터 텍스트까지만 그림
        ax.plot([edge_x, text_x], [edge_y, text_y], 
                color='gray', linewidth=0.8, zorder=1)

    # 5단계: annotation 텍스트 박스 그리기 (높은 z-order)
    bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)

    for w in wedge_info:
        i = w['index']
        pos = text_positions[i]
        text_x = pos['x']
        text_y = pos['y']
        
        # 수평 정렬 설정
        if w['x'] > 0.1:
            horizontalalignment = "left"
        elif w['x'] < -0.1:
            horizontalalignment = "right"
        else:
            horizontalalignment = "center"
        
        label_text = labels[i]
        
        # 텍스트와 박스만 추가 (연결선 없이)
        ax.text(text_x, text_y, label_text,
                ha=horizontalalignment, va="center",
                bbox=bbox_props, zorder=10)  # 높은 z-order

    # 축 범위 조정 - 가장 먼 텍스트를 포함하도록
    max_radius = max([pos['radius'] for pos in text_positions]) + 0.3
    ax.set_xlim(-max_radius, max_radius)
    ax.set_ylim(-max_radius, max_radius)
    ax.set(aspect="equal")

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
    # TODO 음수값 처리
    # for key in list(data):
    #     val1, val2 = data[key]
    #     vals = [val1, val2]
    #     if val1 < 0:
    #         vals[0] = 0
    #     if val2 < 0:
    #         vals[1] = 0
    #     data[key] = (val1, val2)

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
    # arr_plan_n[arr_plan_n < 0] = 0 # TODO 음수값 처리
    arr_execu_n = np.array(execu_n)
    # arr_execu_n[arr_execu_n < 0] = 0 # TODO 음수값 처리
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
    renderer = ax.figure.canvas.get_renderer()
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

class PanelPieAndBar(PanelAspectRatio):
    def __init__(self, parent: wx.Window, pie_values: dict[str, float], bar_values: dict[str, float], color: str):
        super().__init__(parent, 4, True)
        pn_pie = PanelCanvas(self, save_fig_callback=self.save_image)
        pn_bar = PanelCanvas(self, save_fig_callback=self.save_image)

        # 파이
        ax: Axes = pn_pie.ax
        colors = pastel_gradient(color, max(1, len(pie_values)))
        draw_pie(ax, pie_values, colors=colors, start_angle=90.0, sort_by="value", desc=True)

        # 바
        ax: Axes = pn_bar.ax
        ax.set_yticks([])
        ax.set_xticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        if bar_values:
            colors = pastel_gradient(color, len(bar_values))
            labels = list(bar_values.keys())
            values = list(bar_values.values())
            y = np.arange(len(labels))
            ax.barh(y, values, height=0.6, color=colors, edgecolor="none")
            ax.set_yticks(y)
            ax.set_yticklabels(labels)
            ax.tick_params(axis='y', labelsize=10, left=False)
            max_v = max(values) if values else 1
            for yi, val in enumerate(values):
                ax.text(val + (max_v * 0.01), yi, simplify_won(val), ha='left', va='center')
            ax.set_ylim(len(bar_values) - 0.5, -0.5)
            ax.set_xlim(0, max_v * 1.3)

        sz_horz = wx.BoxSizer(wx.HORIZONTAL)
        sz_horz.AddMany((
            (pn_pie, 2, wx.EXPAND), ((40, -1), 0),
            (pn_bar, 3, wx.EXPAND)
        ))
        self.SetSizer(sz_horz)
        pn_pie.draw()
        pn_bar.draw()

        self.pn_pie = pn_pie
        self.pn_bar = pn_bar

    def save_image(self, filepath: str):
        buf_pie = BytesIO()
        self.pn_pie.fig.savefig(buf_pie, format="png", bbox_inches="tight")
        buf_pie.seek(0)
        pil_pie = Image.open(buf_pie)

        buf_bar = BytesIO()
        self.pn_bar.fig.savefig(buf_bar, format="png", bbox_inches="tight")
        buf_bar.seek(0)
        pil_bar = Image.open(buf_bar)

        merged = Image.new(
            "RGB",
            (pil_pie.width + pil_bar.width, pil_pie.height)
        )
        merged.paste(pil_pie, (0, 0))
        merged.paste(pil_bar, (pil_pie.width, 0))
        merged.save(filepath)

class PanelChart(ScrolledPanel):
    def __init__(self, parent: wx.Window):
        ScrolledPanel.__init__(self, parent)
        self.SetBackgroundColour(wx.WHITE)
        self.__set_layout()
        self.__bind_events()
        self.SetupScrolling(False, True)
        self.__cv_lv1.ax[0].set_title("R&D", fontsize=TITLE_FONTSIZE)
        self.__cv_lv1.ax[1].set_title("개발 비용 구성", fontsize=TITLE_FONTSIZE)
        self.__cv_lv1.ax[2].set_title("OE 비용 구성", fontsize=TITLE_FONTSIZE)
        self.__cv_pie.ax[0, 0].set_title("R&D", fontsize=TITLE_FONTSIZE)
        self.__cv_pie.ax[0, 1].set_title("국내", fontsize=TITLE_FONTSIZE)
        self.__cv_pie.ax[0, 2].set_title("해외", fontsize=TITLE_FONTSIZE)
        self.__cv_pie.ax[1, 0].set_title("NATC", fontsize=TITLE_FONTSIZE)
        self.__cv_pie.ax[1, 1].set_title("NETC", fontsize=TITLE_FONTSIZE)
        self.__cv_pie.ax[1, 2].set_title("NCTC", fontsize=TITLE_FONTSIZE)
        self.__cv_pie.fig.set_constrained_layout_pads(wspace=0.2)
        self.__pie_and_bars: list[PanelPieAndBar] = []
        self.draw_empty()

    def __set_layout(self):
        pn_inner = wx.Panel(self)

        pn_header_table = wx.Panel(pn_inner)
        st_label_title  = wx.StaticText(pn_header_table, label="전체 총계")
        st_label_plan   = wx.StaticText(pn_header_table, label="계획")
        st_label_actual = wx.StaticText(pn_header_table, label="실적")
        st_label_exe    = wx.StaticText(pn_header_table, label="집행률")
        st_label_rem    = wx.StaticText(pn_header_table, label="잔액")
        st_value_plan   = wx.StaticText(pn_header_table, label="")
        st_value_actual = wx.StaticText(pn_header_table, label="")
        st_value_exe    = wx.StaticText(pn_header_table, label="")
        st_value_rem    = wx.StaticText(pn_header_table, label="")
        font_regular = self.GetFont()
        font_regular.SetPointSize(12)
        font_bold = self.GetFont().Bold()
        font_bold.SetPointSize(24)
        st_label_title .SetFont(font_regular)
        st_label_plan  .SetFont(font_regular)
        st_label_actual.SetFont(font_regular)
        st_label_exe   .SetFont(font_regular)
        st_label_rem   .SetFont(font_regular)
        st_value_plan  .SetFont(font_bold)
        st_value_actual.SetFont(font_bold)
        st_value_exe   .SetFont(font_bold)
        st_value_rem   .SetFont(font_bold)

        sz_grid_text = wx.FlexGridSizer(2, 5, 10, 10)
        sz_grid_text.AddGrowableRow(0)
        sz_grid_text.AddGrowableCol(1)
        sz_grid_text.AddGrowableCol(2)
        sz_grid_text.AddGrowableCol(3)
        sz_grid_text.AddGrowableCol(4)
        sz_grid_text.AddMany((
            ((-1, -1), 0),
            (st_label_plan  , 0, wx.ALIGN_CENTER),
            (st_label_actual, 0, wx.ALIGN_CENTER),
            (st_label_exe   , 0, wx.ALIGN_CENTER),
            (st_label_rem   , 0, wx.ALIGN_CENTER),
            (st_label_title , 0, wx.ALIGN_CENTER),
            (st_value_plan  , 0, wx.ALIGN_CENTER),
            (st_value_actual, 0, wx.ALIGN_CENTER),
            (st_value_exe   , 0, wx.ALIGN_CENTER),
            (st_value_rem   , 0, wx.ALIGN_CENTER),
        ))

        pn_header_table.SetSizer(sz_grid_text)

        pn_exe_portion = PanelAspectRatio(pn_inner, 3, True) #집행률
        cv_exe_portion = PanelCanvas(pn_exe_portion, 1, 1, save_fig_callback=self.save_fig_exe_portion)
        sz_exe_portion = wx.BoxSizer(wx.HORIZONTAL)
        sz_exe_portion.Add(cv_exe_portion, 1, wx.EXPAND)
        pn_exe_portion.SetSizer(sz_exe_portion)

        pn_lv1 = PanelAspectRatio(pn_inner, 3, True)
        cv_lv1 = PanelCanvas(pn_lv1, 1, 3)
        sz_lv1 = wx.BoxSizer(wx.HORIZONTAL)
        sz_lv1.Add(cv_lv1, 1, wx.EXPAND)
        pn_lv1.SetSizer(sz_lv1)

        pn_pie = PanelAspectRatio(pn_inner, 4/2, True)
        cv_pie = PanelCanvas(pn_pie, 2, 3)
        sz_pie = wx.BoxSizer(wx.HORIZONTAL)
        sz_pie.Add(cv_pie, 1, wx.EXPAND)
        pn_pie.SetSizer(sz_pie)

        pn_bs = PanelAspectRatio(pn_inner, 2, True)
        cv_bs = PanelCanvas(pn_bs)
        sz_bs = wx.BoxSizer(wx.HORIZONTAL)
        sz_bs.Add(cv_bs, 1, wx.EXPAND)
        pn_bs.SetSizer(sz_bs)

        pn_dev = PanelAspectRatio(pn_inner, 2, True) # BS별 개발비
        cv_dev = PanelCanvas(pn_dev)
        sz_dev = wx.BoxSizer(wx.HORIZONTAL)
        sz_dev.Add(cv_dev, 1, wx.EXPAND)
        pn_dev.SetSizer(sz_dev)

        sz_pie_and_bars = wx.BoxSizer(wx.VERTICAL)

        sz_inner = wx.BoxSizer(wx.VERTICAL)
        sz_inner.AddMany((
            ((-1, int(VGAP/2)), 0),
            (pn_header_table, 0, wx.EXPAND), ((-1, VGAP), 0),
            (pn_exe_portion, 0, wx.EXPAND), ((-1, int(VGAP*1.5)), 0),
            (pn_lv1, 0, wx.EXPAND), ((-1, VGAP), 0),
            (pn_pie, 0, wx.EXPAND), ((-1, VGAP), 0),
            (pn_bs, 0, wx.EXPAND), ((-1, VGAP), 0),
            (pn_dev, 0, wx.EXPAND),
            (sz_pie_and_bars, 0, wx.EXPAND),
            ((-1, int(VGAP/2)), 0)
        ))
        pn_inner.SetSizer(sz_inner)

        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.AddStretchSpacer(1)
        sz.Add(pn_inner, 0, wx.EXPAND)
        sz.AddStretchSpacer(1)
        self.SetSizer(sz)

        self.__pn_header_table = pn_header_table
        self.__st_label_title  = st_label_title 
        self.__st_value_plan   = st_value_plan  
        self.__st_value_actual = st_value_actual
        self.__st_value_exe    = st_value_exe   
        self.__st_value_rem    = st_value_rem   
        self.__pn_inner = pn_inner
        self.__cv_exe_portion = cv_exe_portion
        self.__cv_lv1  = cv_lv1
        self.__cv_pie  = cv_pie
        self.__cv_bs   = cv_bs 
        self.__cv_dev  = cv_dev
        self.__sz_pie_and_bars = sz_pie_and_bars
    
    @property
    def cv_exe_portion(self): return self.__cv_exe_portion
    @property
    def cv_lv1        (self): return self.__cv_lv1        
    @property
    def cv_pie        (self): return self.__cv_pie        
    @property
    def cv_bs         (self): return self.__cv_bs         
    @property
    def cv_dev        (self): return self.__cv_dev        
    @property
    def pie_and_bars  (self): return self.__pie_and_bars  

    def __bind_events(self):
        self.Bind(wx.EVT_SIZE, self.__on_size)
    
    def __on_size(self, event):
        w, _ = self.GetClientSize()
        if w < 400:
            self.__pn_inner.SetMinSize((w-40, -1))
        else:
            self.__pn_inner.SetMinSize((min(1000, w), -1))
        event.Skip()

    def save_fig_exe_portion(self, filepath: str):
        panel = self.__pn_header_table
        width, height = panel.GetClientSize().Get()
        client_dc = wx.ClientDC(panel)
        bitmap = wx.Bitmap(width, height)
        memory_dc = wx.MemoryDC(bitmap)
        memory_dc.Blit(0, 0, width, height, client_dc, 0, 0)
        wx_image = bitmap.ConvertToImage()
        width = wx_image.GetWidth()
        height = wx_image.GetHeight()
        rgb_data = wx_image.GetData()
        pil_table = Image.frombytes("RGB", (width, height), rgb_data)

        buf_chart = BytesIO()
        self.__cv_exe_portion.fig.savefig(buf_chart, format="png", bbox_inches="tight")
        buf_chart.seek(0)
        pil_chart = Image.open(buf_chart)

        merged = Image.new(
            "RGB",
            (pil_table.width, pil_table.height+pil_chart.height)
        )
        merged.paste(pil_table, (0, 0))
        merged.paste(pil_chart, (0, pil_table.height))
        merged.save(filepath)

    def set_months(self, months: list[int,]):
        self.__months = months.copy()
        self.load_data()

    def draw_empty(self):
        draw_horizontal_overlapped_bar(self.__cv_exe_portion.ax)
        draw_stacked_single_bar(self.__cv_lv1.ax[0], title="R&D")
        draw_stacked_single_bar(self.__cv_lv1.ax[1], title="개발 비용 구성")
        draw_stacked_single_bar(self.__cv_lv1.ax[2], title="OE 비용 구성")
        draw_donut(self.__cv_pie.ax[0, 0], title="R&D")
        draw_donut(self.__cv_pie.ax[0, 1], title="국내")
        draw_donut(self.__cv_pie.ax[0, 2], title="해외")
        draw_donut(self.__cv_pie.ax[1, 0], title="NATC")
        draw_donut(self.__cv_pie.ax[1, 1], title="NETC")
        draw_donut(self.__cv_pie.ax[1, 2], title="NCTC")
        draw_stacked_multiple_bar(self.__cv_bs.ax)
        draw_stacked_multiple_bar(self.__cv_dev.ax)
        self.__sz_pie_and_bars.Clear(True)
        self.__pie_and_bars.clear()
        # for pn in self.__pie_and_bars:
        #     pn.Destroy()
        self.PostSizeEvent()

    def load_data(self):
        df = LoadedData.get_filtered_df()
        if df is None:
            self.draw_empty()
            return
        self.Freeze()
        months = [f"ConvActual({i})" for i in Config.get_months()]
        months_plan = [f"ConvPlan({i})" for i in Config.get_months()]
        cat_pk_vs_indice = defaultdict(list)
        first_cat_pk_vs_indice = defaultdict(list) # R&D
        rnd_vs_indice = defaultdict(list) # 개발 비중
        oe_vs_indice = defaultdict(list) # OE 비중
        ctr_code_vs_indice = defaultdict(list)
        bs_code_vs_indice = defaultdict(list)
        cat_pk_vs_elem_codes = defaultdict(list)
        for elem in LoadedData.cached_cost_element.values():
            if elem.category_pk in LoadedData.cached_cost_category:
                cat = LoadedData.cached_cost_category[elem.category_pk]
                cat_pk_vs_elem_codes[cat.pk].append(elem.code)

        for idx in df.index:
            sr = df.loc[idx]
            ctr_code = sr["Cost Center"]
            elem_code = sr["Cost Element"]
            currency_code = sr["Currency"]
            elem = LoadedData.cached_cost_element.get(elem_code)
            if elem:
                cat = LoadedData.cached_cost_category.get(elem.category_pk)
                if cat:
                    cat_pk_vs_indice[cat.pk].append(idx)
                first_category = LoadedData.get_first_category(cat)
                if first_category:
                    first_cat_pk_vs_indice[first_category.pk].append(idx)
            ctr = LoadedData.cached_cost_ctr.get(ctr_code)
            if ctr:
                rnd_vs_indice[ctr.rnd].append(idx)
                oe_vs_indice[ctr.oe].append(idx)
                ctr_code_vs_indice[ctr.code].append(idx)
                ctr_code_vs_indice[ctr.code].append(idx)
                bs = LoadedData.get_bs(ctr)
                if bs:
                    bs_code_vs_indice[bs.code].append(idx)

        self.__st_label_title.SetLabel(f"{Config.PERIOD} 총계")
        data_exe_portion = {
            LoadedData.cached_cost_category[cat_pk].name: (
                df.loc[indice, months_plan].fillna(0).sum().sum(),
                df.loc[indice, months].fillna(0).sum().sum()
            )
            for cat_pk, indice in first_cat_pk_vs_indice.items()
        }
        data_exe_portion = dict(sorted(data_exe_portion.items(), key=lambda item: item[1][0]))
        total_plan = np.sum([value[0] for value in data_exe_portion.values()])
        total_actual = np.sum([value[1] for value in data_exe_portion.values()])
        self.__st_value_plan.SetLabel(simplify_won(total_plan))
        self.__st_value_actual.SetLabel(simplify_won(total_actual))
        if total_plan:
            portion = total_actual/total_plan
            if portion >= 0.9:
                self.__st_value_exe.SetForegroundColour(wx.Colour(FONT_COLOR_HIGH_PORTION))
            elif portion >= 0.5:
                self.__st_value_exe.SetForegroundColour(wx.Colour(FONT_COLOR_MID_PORTION))
            else:
                self.__st_value_exe.SetForegroundColour(wx.Colour(FONT_COLOR_LOW_PORTION))
            self.__st_value_exe.SetLabel(f"{portion*100:0.1f}%")
        else:
            self.__st_value_exe.SetForegroundColour(wx.Colour(FONT_COLOR_LOW_PORTION))
            self.__st_value_exe.SetLabel("-")
        if total_plan-total_actual < 0:
            self.__st_value_rem.SetForegroundColour(wx.Colour(FONT_COLOR_NEGATIVE_VALUE))
        else:
            self.__st_value_rem.SetForegroundColour(wx.Colour(0, 0, 0))
        self.__st_value_rem.SetLabel(simplify_won(total_plan-total_actual))
        draw_horizontal_overlapped_bar(self.__cv_exe_portion.ax, data_exe_portion)

        draw_stacked_single_bar(self.__cv_lv1.ax[0], {LoadedData.cached_cost_category[cat_pk].name: df.loc[df.index.isin(indice), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "R&D")
        draw_stacked_single_bar(self.__cv_lv1.ax[1], {label: df.loc[df.index.isin(indice), months].fillna(0).sum().sum() for label, indice in rnd_vs_indice.items()}, "개발 비용 구성")
        draw_stacked_single_bar(self.__cv_lv1.ax[2], {label: df.loc[df.index.isin(indice), months].fillna(0).sum().sum() for label, indice in oe_vs_indice.items()}, "OE 비용 구성")
        draw_donut(self.__cv_pie.ax[0, 0], {LoadedData.cached_cost_category[cat_pk].name: df.loc[df.index.isin(indice), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "R&D")
        draw_donut(self.__cv_pie.ax[0, 1], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df.index.isin(indice)) & (df["Currency"] == "KRW"), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "국내")
        draw_donut(self.__cv_pie.ax[0, 2], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df.index.isin(indice)) & (df["Currency"] != "KRW"), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "해외")
        draw_donut(self.__cv_pie.ax[1, 0], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df["Currency"] == "USD") & (df.index.isin(indice)), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "NATC")
        draw_donut(self.__cv_pie.ax[1, 1], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df["Currency"] == "EUR") & (df.index.isin(indice)), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "NETC")
        draw_donut(self.__cv_pie.ax[1, 2], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df["Currency"] == "CNY") & (df.index.isin(indice)), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "NCTC")
        data = {}
        bs_code_vs_summation = {b: 0 for b in bs_code_vs_indice}
        for cat_pk, indice_cat in first_cat_pk_vs_indice.items():
            cat = LoadedData.cached_cost_category[cat_pk]
            name = cat.name
            data[name] = {}
            mask_cat = df.index.isin(indice_cat)
            for bs_code, indice_bs in bs_code_vs_indice.items():
                mask = mask_cat & df.index.isin(indice_bs)
                value = max(0, df.loc[mask, months].fillna(0).sum().sum()) # 음수는 0으로 처리
                data[name][bs_code] = value
                bs_code_vs_summation[bs_code] += value
        sorted_bs_codes = [
            bs for bs, sm in sorted(
                bs_code_vs_summation.items(),
                key=lambda x: -x[1]
            ) if sm > 0
        ]
        for cat_name in data:
            data[cat_name] = [data[cat_name][bs] for bs in sorted_bs_codes]
        xlabels = [LoadedData.cached_cost_ctr[bs].name for bs in sorted_bs_codes]
        draw_stacked_multiple_bar(self.__cv_bs.ax, data, xlabels, show_summation_on_top=True)

        # '직접개발비' 하위 카테고리에 대한 집행 비율
        category = CostCategory.get_direct_development_cost()
        if not category or not category.children:
            draw_stacked_multiple_bar(self.__cv_dev.ax)
        else:
            data = {}
            total = {}
            summation = {}
            bs_code_vs_summation = {}
            for cat in category.children:
                cat: CostCategory
                data[cat.pk] = {}
                mask_cat = df["Cost Element"].isin(cat_pk_vs_elem_codes[cat.pk])
                for bs_code, indice in bs_code_vs_indice.items():
                    mask = (df.index.isin(indice)) \
                        & (mask_cat)
                    value = df.loc[mask, months].fillna(0).sum().sum()
                    data[cat.pk][bs_code] = value
                    if bs_code not in bs_code_vs_summation:
                        bs_code_vs_summation[bs_code] = 0
                    bs_code_vs_summation[bs_code] += value
                total[cat.pk] = np.sum(list(data[cat.pk].values()))
            sorted_bs_code_vs_summation = dict(sorted(bs_code_vs_summation.items(), key=lambda items: -items[1]))
            for bs_code in list(sorted_bs_code_vs_summation):
                summation = sorted_bs_code_vs_summation[bs_code]
                if summation <= 0:
                    del sorted_bs_code_vs_summation[bs_code]
            data_assigned = {}
            for cat_pk, bs_code_vs_value in data.items():
                values = [bs_code_vs_value[bs_code] for bs_code in sorted_bs_code_vs_summation]
                values.insert(0, total[cat_pk])
                data_assigned[LoadedData.cached_cost_category[cat_pk].name] = values
            xlabels = ["전체",] + [LoadedData.cached_cost_ctr[bs_code].name for bs_code in sorted_bs_code_vs_summation]
            draw_stacked_multiple_bar(self.__cv_dev.ax, data_assigned, xlabels, True, True)

        data = []
        for cat_pk, indice in first_cat_pk_vs_indice.items():
            cat = LoadedData.cached_cost_category[cat_pk]
            lv2_data: dict[CostCategory, float] = {}
            lv3_data: dict[CostCategory, float] = {}
            for lv2_cat in cat.children:
                lv2_cat: CostCategory
                lv2_data[lv2_cat] = 0
                for lv3_cat in lv2_cat.children:
                    lv3_cat: CostCategory
                    indice = cat_pk_vs_indice[lv3_cat.pk]
                    lv3_data[lv3_cat] = df.loc[df.index.isin(indice), months].fillna(0).sum().sum()
                    lv2_data[lv2_cat] += lv3_data[lv3_cat]
            data.append({
                "lv1_cat": cat,
                "lv2_data": dict(sorted(lv2_data.items(), key=lambda item: -item[1])),
                "lv3_data": dict(sorted(lv3_data.items(), key=lambda item: -item[1]))
            })
        data.sort(key=lambda item: item["lv1_cat"].name)
        self.__sz_pie_and_bars.Clear(True)
        self.__pie_and_bars.clear()
        colors = cycle(COLORMAP)
        for dat in data:
            pn_canvas = PanelPieAndBar(
                self.__pn_inner,
                {cat.name: values for cat, values in dat["lv2_data"].items()},
                {cat.name: values for cat, values in dat["lv3_data"].items()},
                next(colors)
            )
            self.__sz_pie_and_bars.Add(pn_canvas, 0, wx.EXPAND|wx.TOP, VGAP)
            self.__pie_and_bars.append(pn_canvas)

        self.__cv_exe_portion.draw()
        self.__cv_lv1.draw()
        self.__cv_pie.draw()
        self.__cv_bs.draw()
        self.__cv_dev.draw()
        self.Layout()
        self.Thaw()
        self.PostSizeEvent()

class PanelDashboard(wx.Panel):
    def __init__(self, parent: wx.Window):
        wx.Panel.__init__(self, parent)
        self.__set_layout()
        self.__bind_events()

    def __set_layout(self):
        pn_menu = wx.Panel(self)
        periods = ["전체", "1Q", "2Q", "3Q", "4Q"]
        periods.extend([f"{i}월" for i in range(1, 13)])
        cb_month = wx.ComboBox(pn_menu, value="전체", choices=periods, style=wx.CB_READONLY)

        target = pn_menu.FromDIP(14)
        bt_chat_gpt = wx.Button(pn_menu, label="ChatGPT 분석")
        bundle = wx.BitmapBundle.FromSVG(OPENAI_MARK_SVG.encode("utf-8"), wx.Size(target, target))
        bt_chat_gpt.SetBitmap(bundle)
        bt_chat_gpt.SetBitmapPosition(wx.LEFT)
        bt_chat_gpt.SetBitmapMargins(pn_menu.FromDIP(8), 0)

        bt_claude = wx.Button(pn_menu, label="Claude 분석")
        bundle = wx.BitmapBundle.FromSVG(CLAUDE_MARK_SVG.encode("utf-8"), wx.Size(target, target))
        bt_claude.SetBitmap(bundle)
        bt_claude.SetBitmapPosition(wx.LEFT)
        bt_claude.SetBitmapMargins(pn_menu.FromDIP(8), 0)

        bt_save_all_images = wx.Button(pn_menu, label="이미지 일괄 저장")

        sz_horz = wx.BoxSizer(wx.HORIZONTAL)
        sz_horz.AddMany((
            (cb_month, 0, wx.ALIGN_CENTER_VERTICAL), ((10, -1), 0),
            (bt_chat_gpt, 0, wx.ALIGN_CENTER_VERTICAL), ((5, -1), 0),
            (bt_claude, 0, wx.ALIGN_CENTER_VERTICAL), ((30, -1), 1),
            (bt_save_all_images, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_menu = wx.BoxSizer(wx.HORIZONTAL)
        sz_menu.Add(sz_horz, 1, wx.ALL, 10)
        pn_menu.SetSizer(sz_menu)

        pn_chart = PanelChart(self)

        sz = wx.BoxSizer(wx.VERTICAL)
        sz.AddMany((
            (pn_menu, 0, wx.EXPAND),
            # (wx.StaticLine(self), 0, wx.EXPAND),
            (pn_chart, 1, wx.EXPAND)
        ))
        self.SetSizer(sz)

        self.__cb_month = cb_month
        self.__bt_chat_gpt = bt_chat_gpt
        self.__bt_claude = bt_claude
        self.__bt_save_all_images = bt_save_all_images
        self.__pn_chart = pn_chart

    def __bind_events(self):
        self.__cb_month.Bind(wx.EVT_COMBOBOX, self.__on_combo_month)
        self.__bt_chat_gpt.Bind(wx.EVT_BUTTON, self.__on_chat_gpt)
        self.__bt_claude.Bind(wx.EVT_BUTTON, self.__on_claude)
        self.__bt_save_all_images.Bind(wx.EVT_BUTTON, self.__on_save_all_images)
    
    def __on_combo_month(self, event):
        Config.PERIOD = self.__cb_month.GetValue()
        self.redraw_charts()
    
    def __on_chat_gpt(self, event):
        self.__on_ai("ChatGPT")
    
    def __on_claude(self, event):
        self.__on_ai("Claude")

    def __on_ai(self, ai_type: Literal["ChatGPT", "Claude"]):
        if LoadedData.df.empty:
            wx.MessageBox("분석할 데이터가 없습니다.\n먼저 데이터를 로드하세요.", "안내", parent=self)
            return
        match ai_type:
            case "ChatGPT":
                key = Config.OPENAI_API_KEY
                if not key:
                    wx.MessageBox("OpenAI API Key를 설정하세요.", "안내", parent=self)
                    return
                dlg = DialogModels(self, Config.GPT_MODELS, Config.LAST_USED_GPT_MODEL)
                ret = dlg.ShowModal()
                model = dlg.get_model()
                dlg.Destroy()
                if ret != wx.ID_OK:
                    return
                Config.LAST_USED_GPT_MODEL = model
                analyze = analyze_by_gpt
            case "Claude":
                key = Config.CLAUDE_API_KEY
                if not key:
                    wx.MessageBox("Claude API Key를 설정하세요.", "안내")
                    return
                dlg = DialogModels(self, Config.CLAUDE_MODELS, Config.LAST_USED_CLAUDE_MODEL)
                ret = dlg.ShowModal()
                model = dlg.get_model()
                dlg.Destroy()
                if ret != wx.ID_OK:
                    return
                Config.LAST_USED_CLAUDE_MODEL = model
                analyze = analyze_by_claude
            case _:
                raise RuntimeError
        
        dlgp = wx.ProgressDialog("안내", f"{ai_type}을 이용하여 분석 중입니다.", parent=self)
        dlgp.Pulse()

        def on_success(json_data: str, result: str):
            dlgp.Destroy()
            wx.Yield()
            dlg = DialogAIResult(self, ai_type, model, json_data, result)
            wx.MessageBox("분석을 완료했습니다.", "안내", parent=self)
            wx.Yield()
            dlg.ShowModal()
            dlg.Destroy()
        
        def on_fail(msg: str):
            dlgp.Destroy()
            wx.Yield()
            wx.MessageBox(msg, "안내", parent=self)

        def work():
            try:
                months = Config.get_months()
                all_categories = CostCategory.get_all()
                all_elements = CostElement.get_all()
                all_ctrs = CostCtr.get_all()

                category_list = [
                    _CostCategory(
                        pk=cat.pk,
                        parent_pk=cat.parent_pk,
                        name=cat.name
                    ) for cat in all_categories.values()
                ]
                element_list = [
                    _CostElement(
                        code=elem.code,
                        description=elem.description,
                        category_pk=elem.category_pk
                    ) for elem in all_elements.values()
                ]
                ctr_list = [
                    _CostCtr(
                        code=ctr.code,
                        parent_code=ctr.parent_code,
                        name=ctr.name,
                        rnd=["Research", "Develop"].index(ctr.rnd),
                        oe=["공통비", "RE", "OE"].index(ctr.oe)
                    ) for ctr in all_ctrs.values()
                ]

                df = LoadedData.df.fillna(0)
                cols = {name: i for i, name in enumerate(df.columns)}
                by_elem_code = {}
                by_ctr_code = {}
                for row in df.itertuples(index=False, name=None):
                    ctr_code = row[cols["Cost Center"]]
                    elem_code = row[cols["Cost Element"]]
                    if ctr_code not in all_ctrs \
                        or elem_code not in all_elements:
                        continue
                    if elem_code not in by_elem_code:
                        by_elem_code[elem_code] = {"code": elem_code, "planned": 0, "executed": 0}
                    if ctr_code not in by_ctr_code:
                        by_ctr_code[ctr_code] = {"code": ctr_code, "planned": 0, "executed": 0}
                    for month in months:
                        planned = row[cols[f"ConvPlan({month})"]]
                        executed = row[cols[f"ConvActual({month})"]]
                        by_elem_code[elem_code]["planned"] += planned
                        by_elem_code[elem_code]["executed"] += executed
                        by_ctr_code[ctr_code]["planned"] += planned
                        by_ctr_code[ctr_code]["executed"] += executed
                
                system_prompt, user_prompt, json_data = get_prompts_for_ai(
                    category_list,
                    element_list,
                    ctr_list,
                    [
                        _BudgetByElement(
                            cost_element_code=item["code"],
                            planned=item["planned"],
                            executed=item["executed"]
                        ) for item in by_elem_code.values()
                    ],
                    [
                        _BudgetByCtr(
                            cost_ctr_code=item["code"],
                            planned=item["planned"],
                            executed=item["executed"]
                        ) for item in by_ctr_code.values()
                    ]
                )

                res = analyze(system_prompt, user_prompt, key, model)

            except Exception as err:
                wx.CallAfter(on_fail, f"AI 분석 중 오류가 발생했습니니다.\n\n{err}")
            
            else:
                wx.CallAfter(on_success, json_data, res)

        Thread(target=work, daemon=True).start()

    def __on_save_all_images(self, event):
        dlg = wx.DirDialog(self, "이미지 일괄 저장", style=wx.DD_DEFAULT_STYLE)
        ret = dlg.ShowModal()
        dir_path = dlg.GetPath()
        dlg.Destroy()
        if ret != wx.ID_OK:
            return
        dlgp = wx.ProgressDialog("안내", "이미지 파일을 생성 중입니다.", parent=None)
        dlgp.Pulse()
        filehead = datetime.now().strftime("%y%m%d_%H%M%S")
        try:
            os.makedirs(dir_path, exist_ok=True)
            self.__pn_chart.save_fig_exe_portion(os.path.join(dir_path, f"{filehead}_01.png"))
            self.__pn_chart.cv_lv1        .fig.savefig(os.path.join(dir_path, f"{filehead}_02.png"), dpi=300)
            self.__pn_chart.cv_pie        .fig.savefig(os.path.join(dir_path, f"{filehead}_03.png"), dpi=300)
            self.__pn_chart.cv_bs         .fig.savefig(os.path.join(dir_path, f"{filehead}_04.png"), dpi=300)
            self.__pn_chart.cv_dev        .fig.savefig(os.path.join(dir_path, f"{filehead}_05.png"), dpi=300)
            ad = 1
            for pnb in self.__pn_chart.pie_and_bars:
                pnb.save_image(os.path.join(dir_path, f"{filehead}_{5+ad}.png"))
                ad += 1
        except:
            msg = f"이미지 생성 중 오류가 발생했습니다.\n\n{format_exc()}"
        else:
            msg = f"이미지들을 저장 하였습니다."
        finally:
            dlgp.Destroy()
            wx.Yield()
            wx.CallAfter(wx.MessageBox, msg, "안내", parent=self)

    def redraw_charts(self):
        self.__pn_chart.load_data()
