import os
import numpy as np
import wx

from datetime import datetime
from io import BytesIO
from itertools import cycle
from collections import defaultdict
from traceback import format_exc

from PIL import Image
from wx.lib.scrolledpanel import ScrolledPanel

from util import Config, simplify_won, pastel_gradient, COLORMAP
from util.chart import (
    draw_pie, draw_horizontal_overlapped_bar, draw_stacked_multiple_bar,
    VGAP,
)

from db.models import CostCtr, CostCategory
from db.loaded_data import LoadedData

from ui.component import PanelAspectRatio, PanelCanvas, \
    FONT_COLOR_HIGH_PORTION, FONT_COLOR_MID_PORTION, FONT_COLOR_LOW_PORTION, FONT_COLOR_NEGATIVE_VALUE


class PanelPieAndBar(PanelAspectRatio):
    def __init__(self, parent: wx.Window, pie_values: dict[str, float], bar_values: dict[str, float], color: str):
        super().__init__(parent, 4, True)
        pn_pie = PanelCanvas(self, save_fig_callback=self.save_image)
        pn_bar = PanelCanvas(self, save_fig_callback=self.save_image)

        # 파이
        ax: Axes = pn_pie.ax # type: ignore
        colors = pastel_gradient(color, max(1, len(pie_values)))
        draw_pie(ax, pie_values, colors=colors, start_angle=90.0, sort_by="value", desc=True)

        # 바
        ax: Axes = pn_bar.ax # type: ignore
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
        super().__init__(parent)
        self.SetBackgroundColour(wx.WHITE)
        self._set_layout()
        self._bind_events()
        self.SetupScrolling(False, True)
        self._pie_and_bars: list[PanelPieAndBar] = []
        self.draw_empty()

    def _set_layout(self):
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

        pn_team = PanelAspectRatio(pn_inner, 2, True)
        cv_team = PanelCanvas(pn_team)
        sz_team = wx.BoxSizer(wx.HORIZONTAL)
        sz_team.Add(cv_team, 1, wx.EXPAND)
        pn_team.SetSizer(sz_team)

        pn_dev = PanelAspectRatio(pn_inner, 2, True) # 팀별 개발비
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
            (pn_team, 0, wx.EXPAND), ((-1, VGAP), 0),
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
        self.__cv_team = cv_team
        self.__cv_dev = cv_dev
        self.__sz_pie_and_bars = sz_pie_and_bars

    @property
    def cv_team(self): return self.__cv_team
    @property
    def cv_dev(self): return self.__cv_dev
    @property
    def pie_and_bars(self): return self._pie_and_bars

    def _bind_events(self):
        return

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

    def draw_empty(self):
        draw_horizontal_overlapped_bar(self.__cv_exe_portion.ax) # type: ignore
        draw_stacked_multiple_bar(self.__cv_team.ax) # type: ignore
        draw_stacked_multiple_bar(self.__cv_dev.ax) # type: ignore
        self.__sz_pie_and_bars.Clear(True)
        self._pie_and_bars.clear()
        self.PostSizeEvent()

    def load_data(self, period: str, bs: CostCtr):
        month_list = Config.get_months(period)
        df = LoadedData.get_filtered_df()
        if df is None \
            or not month_list:
            self.draw_empty()
            return
        team_codes = []
        for ctr in LoadedData.cached_cost_ctr.values():
            if ctr.parent_code == bs.code:
                team_codes.append(ctr.code)
        bs_and_team_codes = [bs.code] + [code for code in team_codes]
        df = df.loc[df["Cost Center"].isin(bs_and_team_codes)]
        if df.empty:
            self.draw_empty()
            return
        self.Freeze()
        months = [f"ConvActual({i})" for i in month_list]
        months_plan = [f"ConvPlan({i})" for i in month_list]
        cat_pk_vs_indice = defaultdict(list)
        first_cat_pk_vs_indice = defaultdict(list) # R&D
        rnd_vs_indice = defaultdict(list) # 개발 비중
        oe_vs_indice = defaultdict(list) # OE 비중
        team_code_vs_indice = defaultdict(list)
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
                cat = LoadedData.cached_cost_category.get(elem.category_pk) # type: ignore
                if cat:
                    cat_pk_vs_indice[cat.pk].append(idx)
                first_category = LoadedData.get_first_category(cat)
                if first_category:
                    first_cat_pk_vs_indice[first_category.pk].append(idx)
            ctr = LoadedData.cached_cost_ctr.get(ctr_code)
            if ctr and ctr.code in team_codes:
                rnd_vs_indice[ctr.rnd].append(idx)
                oe_vs_indice[ctr.oe].append(idx)
                team_code_vs_indice[ctr.code].append(idx)
                team_code_vs_indice[ctr.code].append(idx)

        self.__st_label_title.SetLabel(f"{period} 총계")
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
        draw_horizontal_overlapped_bar(self.__cv_exe_portion.ax, data_exe_portion) # type: ignore

        data = {}
        bs_code_vs_summation = {b: 0 for b in team_code_vs_indice}
        for cat_pk, indice_cat in first_cat_pk_vs_indice.items():
            cat = LoadedData.cached_cost_category[cat_pk]
            name = cat.name
            data[name] = {}
            mask_cat = df.index.isin(indice_cat)
            for bs_code, indice_bs in team_code_vs_indice.items():
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
        draw_stacked_multiple_bar(self.__cv_team.ax, data, xlabels, show_summation_on_top=True) # type: ignore

        # '직접개발비' 하위 카테고리에 대한 집행 비율
        category = CostCategory.get_direct_development_cost()
        if not category or not category.children:
            draw_stacked_multiple_bar(self.__cv_dev.ax) # type: ignore
        else:
            data = {}
            total = {}
            summation = {}
            team_code_vs_summation = {}
            for cat in category.children:
                cat: CostCategory
                data[cat.pk] = {}
                mask_cat = df["Cost Element"].isin(cat_pk_vs_elem_codes[cat.pk])
                for team_code, indice in team_code_vs_indice.items():
                    mask = (df.index.isin(indice)) \
                        & (mask_cat)
                    value = df.loc[mask, months].fillna(0).sum().sum()
                    data[cat.pk][team_code] = value
                    if team_code not in team_code_vs_summation:
                        team_code_vs_summation[team_code] = 0
                    team_code_vs_summation[team_code] += value
                total[cat.pk] = np.sum(list(data[cat.pk].values()))
            sorted_team_code_vs_summation = dict(sorted(team_code_vs_summation.items(), key=lambda items: -items[1]))
            for team_code in list(sorted_team_code_vs_summation):
                summation = sorted_team_code_vs_summation[team_code]
                if summation <= 0:
                    del sorted_team_code_vs_summation[team_code]
            data_assigned = {}
            for cat_pk, team_code_vs_value in data.items():
                values = [team_code_vs_value[team_code] for team_code in sorted_team_code_vs_summation]
                values.insert(0, total[cat_pk])
                data_assigned[LoadedData.cached_cost_category[cat_pk].name] = values
            xlabels = ["전체",] + [LoadedData.cached_cost_ctr[team_code].name for team_code in sorted_team_code_vs_summation]
            draw_stacked_multiple_bar(self.__cv_dev.ax, data_assigned, xlabels, True, True) # type: ignore

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
        self._pie_and_bars.clear()
        colors = cycle(COLORMAP)
        for dat in data:
            pn_canvas = PanelPieAndBar(
                self.__pn_inner,
                {cat.name: values for cat, values in dat["lv2_data"].items()},
                {cat.name: values for cat, values in dat["lv3_data"].items()},
                next(colors)
            )
            self.__sz_pie_and_bars.Add(pn_canvas, 0, wx.EXPAND|wx.TOP, VGAP)
            self._pie_and_bars.append(pn_canvas)

        self.__cv_exe_portion.draw()
        self.__cv_team.draw()
        self.__cv_dev.draw()
        self.Layout()
        self.Thaw()
        self.PostSizeEvent()

class PanelBSChart(wx.Panel):
    def __init__(self, parent: wx.Window):
        super().__init__(parent)
        self._bs_list: list[CostCtr] = []
        self._set_layout()
        self._bind_events()
        self.load_bs_list()
        self.draw()

    def _set_layout(self):
        pn_menu = wx.Panel(self)
        periods = ["전체", "1Q", "2Q", "3Q", "4Q"]
        periods.extend([f"{i}월" for i in range(1, 13)])
        cb_month = wx.ComboBox(pn_menu, value="전체", choices=periods, style=wx.CB_READONLY)
        cb_bs = wx.ComboBox(pn_menu, size=wx.Size(200, -1), style=wx.CB_READONLY)
        bt_save_all_images = wx.Button(pn_menu, label="이미지 일괄 저장")
        sz_horz = wx.BoxSizer(wx.HORIZONTAL)
        sz_horz.AddMany((
            (cb_month, 0, wx.ALIGN_CENTER_VERTICAL), ((10, -1), 0),
            (cb_bs, 0, wx.ALIGN_CENTER_VERTICAL), ((30, -1), 1),
            (bt_save_all_images, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_menu = wx.BoxSizer(wx.HORIZONTAL)
        sz_menu.Add(sz_horz, 1, wx.ALL, 10)
        pn_menu.SetSizer(sz_menu)

        pn_chart = PanelChart(self)

        sz = wx.BoxSizer(wx.VERTICAL)
        sz.AddMany((
            (pn_menu, 0, wx.EXPAND),
            (pn_chart, 1, wx.EXPAND)
        ))
        self.SetSizer(sz)

        self._cb_month = cb_month
        self._cb_bs = cb_bs 
        self._bt_save_all_images = bt_save_all_images
        self._pn_chart = pn_chart

    def _bind_events(self):
        self._cb_month.Bind(wx.EVT_COMBOBOX, self._on_combo_month)
        self._cb_bs.Bind(wx.EVT_COMBOBOX, self._on_combo_bs)
        self._bt_save_all_images.Bind(wx.EVT_BUTTON, self._on_save_all_images)

    def _on_combo_month(self, evt):
        self.draw()

    def _on_combo_bs(self, evt):
        self.draw()
    
    def _on_save_all_images(self, evt):
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
            self._pn_chart.save_fig_exe_portion(os.path.join(dir_path, f"{filehead}_01.png"))
            self._pn_chart.cv_team.fig.savefig(os.path.join(dir_path, f"{filehead}_02.png"), dpi=300)
            self._pn_chart.cv_dev .fig.savefig(os.path.join(dir_path, f"{filehead}_03.png"), dpi=300)
            ad = 1
            for pnb in self._pn_chart.pie_and_bars:
                pnb.save_image(os.path.join(dir_path, f"{filehead}_{3+ad}.png"))
                ad += 1
        except:
            msg = f"이미지 생성 중 오류가 발생했습니다.\n\n{format_exc()}"
        else:
            msg = f"이미지들을 저장 하였습니다."
        finally:
            dlgp.Destroy()
            wx.Yield()
            wx.CallAfter(wx.MessageBox, msg, "안내", parent=self)

    def load_bs_list(self):
        """캐시로부터 BS 목록을 확인하여 ComboBox에 로드"""
        self._bs_list.clear()
        for ctr in LoadedData.cached_cost_ctr.values():
            level = LoadedData.get_level_of_ctr_from_cache(ctr)
            if level != 2:
                continue
            self._bs_list.append(ctr)
        self._bs_list.sort(key=lambda ctr: ctr.name)
        if not self._bs_list:
            self._cb_bs.Clear()
        else:
            self._cb_bs.SetItems([ctr.name for ctr in self._bs_list])
            self._cb_bs.SetValue(self._bs_list[0].name)
        self.draw()
    
    def draw(self):
        """선택한 기간, BS를 반영하여 차트 그림"""
        idx = self._cb_bs.GetSelection()
        if idx == wx.NOT_FOUND:
            self._pn_chart.draw_empty()
            return
        ctr = self._bs_list[idx]
        self._pn_chart.load_data(self._cb_month.GetValue(), ctr)

