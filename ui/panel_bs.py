import numpy as np
import pandas as pd
import wx

from collections import defaultdict
from wx.lib.scrolledpanel import ScrolledPanel

from util import Config, simplify_won
from util.chart import (
    draw_pie, draw_horizontal_overlapped_bar, draw_stacked_single_bar, draw_donut, draw_stacked_multiple_bar,
    TITLE_FONTSIZE, VGAP,
)

from db.models import CostCtr
from db.loaded_data import LoadedData

from ui.component import PanelAspectRatio, PanelCanvas, \
    FONT_COLOR_HIGH_PORTION, FONT_COLOR_MID_PORTION, FONT_COLOR_LOW_PORTION, FONT_COLOR_NEGATIVE_VALUE, \
    OPENAI_MARK_SVG, CLAUDE_MARK_SVG


class PanelChart(ScrolledPanel):
    def __init__(self, parent: wx.Window):
        super().__init__(parent)
        self.SetBackgroundColour(wx.WHITE)
        self._set_layout()
        self._bind_events()
        self.SetupScrolling(False, True)
        # self.__cv_lv1.ax[0].set_title("R&D", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_lv1.ax[1].set_title("개발 비용 구성", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_lv1.ax[2].set_title("OE 비용 구성", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_pie.ax[0, 0].set_title("R&D", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_pie.ax[0, 1].set_title("국내", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_pie.ax[0, 2].set_title("해외", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_pie.ax[1, 0].set_title("NATC", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_pie.ax[1, 1].set_title("NETC", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_pie.ax[1, 2].set_title("NCTC", fontsize=TITLE_FONTSIZE) # type: ignore
        # self.__cv_pie.fig.set_constrained_layout_pads(wspace=0.2) # type: ignore
        # self.__pie_and_bars: list[PanelPieAndBar] = []
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

        sz_inner = wx.BoxSizer(wx.VERTICAL)
        sz_inner.AddMany((
            ((-1, int(VGAP/2)), 0),
            (pn_header_table, 0, wx.EXPAND), ((-1, VGAP), 0),
            (pn_exe_portion, 0, wx.EXPAND), ((-1, int(VGAP*1.5)), 0),
            (pn_team, 0, wx.EXPAND), ((-1, VGAP), 0),
            # (pn_dev, 0, wx.EXPAND),
            # (sz_pie_and_bars, 0, wx.EXPAND),
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
        # self.__cv_dev  = cv_dev
        # self.__sz_pie_and_bars = sz_pie_and_bars

    def _bind_events(self):
        return

    def save_fig_exe_portion(self, filepath: str):
        # TODO
        return
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
        # draw_stacked_single_bar(self.__cv_lv1.ax[0], title="R&D") # type: ignore
        # draw_stacked_single_bar(self.__cv_lv1.ax[1], title="개발 비용 구성") # type: ignore
        # draw_stacked_single_bar(self.__cv_lv1.ax[2], title="OE 비용 구성") # type: ignore
        # draw_donut(self.__cv_pie.ax[0, 0], title="R&D") # type: ignore
        # draw_donut(self.__cv_pie.ax[0, 1], title="국내") # type: ignore
        # draw_donut(self.__cv_pie.ax[0, 2], title="해외") # type: ignore
        # draw_donut(self.__cv_pie.ax[1, 0], title="NATC") # type: ignore
        # draw_donut(self.__cv_pie.ax[1, 1], title="NETC") # type: ignore
        # draw_donut(self.__cv_pie.ax[1, 2], title="NCTC") # type: ignore
        draw_stacked_multiple_bar(self.__cv_team.ax) # type: ignore
        # draw_stacked_multiple_bar(self.__cv_dev.ax) # type: ignore
        # self.__sz_pie_and_bars.Clear(True)
        # self.__pie_and_bars.clear()
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

        # draw_stacked_single_bar(self.__cv_lv1.ax[0], {LoadedData.cached_cost_category[cat_pk].name: df.loc[df.index.isin(indice), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "R&D") # type: ignore
        # draw_stacked_single_bar(self.__cv_lv1.ax[1], {label: df.loc[df.index.isin(indice), months].fillna(0).sum().sum() for label, indice in rnd_vs_indice.items()}, "개발 비용 구성") # type: ignore
        # draw_stacked_single_bar(self.__cv_lv1.ax[2], {label: df.loc[df.index.isin(indice), months].fillna(0).sum().sum() for label, indice in oe_vs_indice.items()}, "OE 비용 구성") # type: ignore
        # draw_donut(self.__cv_pie.ax[0, 0], {LoadedData.cached_cost_category[cat_pk].name: df.loc[df.index.isin(indice), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "R&D") # type: ignore
        # draw_donut(self.__cv_pie.ax[0, 1], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df.index.isin(indice)) & (df["Currency"] == "KRW"), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "국내") # type: ignore
        # draw_donut(self.__cv_pie.ax[0, 2], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df.index.isin(indice)) & (df["Currency"] != "KRW"), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "해외") # type: ignore
        # draw_donut(self.__cv_pie.ax[1, 0], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df["Currency"] == "USD") & (df.index.isin(indice)), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "NATC") # type: ignore
        # draw_donut(self.__cv_pie.ax[1, 1], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df["Currency"] == "EUR") & (df.index.isin(indice)), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "NETC") # type: ignore
        # draw_donut(self.__cv_pie.ax[1, 2], {LoadedData.cached_cost_category[cat_pk].name: df.loc[(df["Currency"] == "CNY") & (df.index.isin(indice)), months].fillna(0).sum().sum() for cat_pk, indice in first_cat_pk_vs_indice.items()}, "NCTC") # type: ignore
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

        # # '직접개발비' 하위 카테고리에 대한 집행 비율
        # category = CostCategory.get_direct_development_cost()
        # if not category or not category.children:
        #     draw_stacked_multiple_bar(self.__cv_dev.ax) # type: ignore
        # else:
        #     data = {}
        #     total = {}
        #     summation = {}
        #     bs_code_vs_summation = {}
        #     for cat in category.children:
        #         cat: CostCategory
        #         data[cat.pk] = {}
        #         mask_cat = df["Cost Element"].isin(cat_pk_vs_elem_codes[cat.pk])
        #         for bs_code, indice in bs_code_vs_indice.items():
        #             mask = (df.index.isin(indice)) \
        #                 & (mask_cat)
        #             value = df.loc[mask, months].fillna(0).sum().sum()
        #             data[cat.pk][bs_code] = value
        #             if bs_code not in bs_code_vs_summation:
        #                 bs_code_vs_summation[bs_code] = 0
        #             bs_code_vs_summation[bs_code] += value
        #         total[cat.pk] = np.sum(list(data[cat.pk].values()))
        #     sorted_bs_code_vs_summation = dict(sorted(bs_code_vs_summation.items(), key=lambda items: -items[1]))
        #     for bs_code in list(sorted_bs_code_vs_summation):
        #         summation = sorted_bs_code_vs_summation[bs_code]
        #         if summation <= 0:
        #             del sorted_bs_code_vs_summation[bs_code]
        #     data_assigned = {}
        #     for cat_pk, bs_code_vs_value in data.items():
        #         values = [bs_code_vs_value[bs_code] for bs_code in sorted_bs_code_vs_summation]
        #         values.insert(0, total[cat_pk])
        #         data_assigned[LoadedData.cached_cost_category[cat_pk].name] = values
        #     xlabels = ["전체",] + [LoadedData.cached_cost_ctr[bs_code].name for bs_code in sorted_bs_code_vs_summation]
        #     draw_stacked_multiple_bar(self.__cv_dev.ax, data_assigned, xlabels, True, True) # type: ignore

        # data = []
        # for cat_pk, indice in first_cat_pk_vs_indice.items():
        #     cat = LoadedData.cached_cost_category[cat_pk]
        #     lv2_data: dict[CostCategory, float] = {}
        #     lv3_data: dict[CostCategory, float] = {}
        #     for lv2_cat in cat.children:
        #         lv2_cat: CostCategory
        #         lv2_data[lv2_cat] = 0
        #         for lv3_cat in lv2_cat.children:
        #             lv3_cat: CostCategory
        #             indice = cat_pk_vs_indice[lv3_cat.pk]
        #             lv3_data[lv3_cat] = df.loc[df.index.isin(indice), months].fillna(0).sum().sum()
        #             lv2_data[lv2_cat] += lv3_data[lv3_cat]
        #     data.append({
        #         "lv1_cat": cat,
        #         "lv2_data": dict(sorted(lv2_data.items(), key=lambda item: -item[1])),
        #         "lv3_data": dict(sorted(lv3_data.items(), key=lambda item: -item[1]))
        #     })
        # data.sort(key=lambda item: item["lv1_cat"].name)
        # self.__sz_pie_and_bars.Clear(True)
        # self.__pie_and_bars.clear()
        # colors = cycle(COLORMAP)
        # for dat in data:
        #     pn_canvas = PanelPieAndBar(
        #         self.__pn_inner,
        #         {cat.name: values for cat, values in dat["lv2_data"].items()},
        #         {cat.name: values for cat, values in dat["lv3_data"].items()},
        #         next(colors)
        #     )
        #     self.__sz_pie_and_bars.Add(pn_canvas, 0, wx.EXPAND|wx.TOP, VGAP)
        #     self.__pie_and_bars.append(pn_canvas)

        self.__cv_exe_portion.draw()
        self.__cv_team.draw()
        # self.__cv_dev.draw()
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
        # TODO
        return

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

