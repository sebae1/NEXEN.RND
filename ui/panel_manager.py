import math
import numpy as np
import pandas as pd
import wx
import wx.dataview as DV
from dataclasses import dataclass
from wx.lib.newevent import NewCommandEvent
from wx.lib.scrolledpanel import ScrolledPanel
from db import EnumOE, EnumRND, Currency, CostCtr, CostCategory, CostElement, MAXIMUM_DEPTH_OF_CATEGORY, LoadedData, Session
from util import get_error_message
from .component import TreeListCtrl, TreeListModelBase, EvtUpdate, TextEntryDialog

class DialogRootCtr(wx.Dialog):
    def __init__(self, parent: wx.Panel, root_ctr: CostCtr):
        wx.Dialog.__init__(self, parent, title="Root Ctr 수정")
        self.__ctr = root_ctr
        self.__set_layout()
        self.__bind_events()

    def __set_layout(self):
        st_bs   = wx.StaticText(self, label="Root Ctr")
        st_code = wx.StaticText(self, label="코드")
        st_rnd  = wx.StaticText(self, label="개발 비중")
        st_oe   = wx.StaticText(self, label="OE 비중")
        tc_name = wx.TextCtrl(self, size=(250, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER, value=self.__ctr.name)
        tc_code = wx.TextCtrl(self, size=(250, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER, value=self.__ctr.code)
        rb_rnd_dev = wx.RadioButton(self, label="Develop", style=wx.RB_GROUP)
        rb_rnd_res = wx.RadioButton(self, label="Research")
        rb_oe_com = wx.RadioButton(self, label="공통비", style=wx.RB_GROUP)
        rb_oe_oe  = wx.RadioButton(self, label="OE")
        rb_oe_re  = wx.RadioButton(self, label="RE")

        if self.__ctr.rnd == EnumRND.DEVELOP:
            rb_rnd_dev.SetValue(True)
        else:
            rb_rnd_res.SetValue(True)
        if self.__ctr.oe == EnumOE.COMMON:
            rb_oe_com.SetValue(True)
        elif self.__ctr.oe == EnumOE.OE:
            rb_oe_oe.SetValue(True)
        elif self.__ctr.oe == EnumOE.RE:
            rb_oe_re.SetValue(True)

        sz_rnd = wx.BoxSizer(wx.HORIZONTAL)
        sz_rnd.AddMany((
            (rb_rnd_dev, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_rnd_res, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_oe = wx.BoxSizer(wx.HORIZONTAL)
        sz_oe.AddMany((
            (rb_oe_com, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_oe_oe , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_oe_re , 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_grid = wx.FlexGridSizer(4, 2, 5, 5)
        sz_grid.AddMany((
            (st_bs  , 0, wx.ALIGN_CENTER_VERTICAL), (tc_name, 0, wx.ALIGN_CENTER_VERTICAL),
            (st_code, 0, wx.ALIGN_CENTER_VERTICAL), (tc_code, 0, wx.ALIGN_CENTER_VERTICAL),
            (st_rnd , 0, wx.ALIGN_CENTER_VERTICAL), (sz_rnd , 0, wx.ALIGN_CENTER_VERTICAL),
            (st_oe  , 0, wx.ALIGN_CENTER_VERTICAL), (sz_oe  , 0, wx.ALIGN_CENTER_VERTICAL)
        ))

        bt_confirm = wx.Button(self, label="수정")
        bt_cancel = wx.Button(self, label="취소")
        sz_bt = wx.BoxSizer(wx.HORIZONTAL)
        sz_bt.AddMany((
            (bt_confirm, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_cancel, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (sz_grid, 0), ((-1, 10), 0),
            (wx.StaticLine(self), 0, wx.EXPAND), ((-1, 10), 0),
            (sz_bt, 0, wx.ALIGN_RIGHT)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 20)
        self.SetSizerAndFit(sz)
        self.CenterOnParent()

        self.__tc_name    = tc_name   
        self.__tc_code    = tc_code   
        self.__rb_rnd_dev = rb_rnd_dev
        self.__rb_rnd_res = rb_rnd_res
        self.__rb_oe_com  = rb_oe_com 
        self.__rb_oe_oe   = rb_oe_oe  
        self.__rb_oe_re   = rb_oe_re  
        self.__bt_confirm = bt_confirm
        self.__bt_cancel  = bt_cancel 

    def __bind_events(self):
        self.__bt_confirm.Bind(wx.EVT_BUTTON, self.__on_confirm)
        self.__bt_cancel .Bind(wx.EVT_BUTTON, self.__on_cancel )
        self.__tc_name   .Bind(wx.EVT_TEXT_ENTER, self.__on_enter)
        self.__tc_code.Bind(wx.EVT_TEXT_ENTER, self.__on_enter)

    def __on_enter(self, event):
        self.__on_confirm(None)

    def __on_confirm(self, event):
        name = self.__tc_name.GetValue().strip()
        code = self.__tc_code.GetValue().strip()
        rnd = EnumRND.DEVELOP if self.__rb_rnd_dev.GetValue() else EnumRND.RESEARCH
        oe = EnumOE.COMMON
        if self.__rb_oe_oe.GetValue():
            oe = EnumOE.OE
        elif self.__rb_oe_re.GetValue():
            oe = EnumOE.RE
        if not name:
            wx.MessageBox("이름을 입력하세요.", "안내")
            return
        if not code:
            wx.MessageBox("코드를 입력하세요.", "안내")
            return
        try:
            if (self.__ctr.code == code
                and self.__ctr.name == name
                and self.__ctr.rnd == rnd
                and self.__ctr.oe == oe):
                self.EndModal(wx.ID_CANCEL)
                return
            self.__ctr = self.__ctr.update(code, name, rnd, oe, self.__ctr.parent_code)
        except Exception as err:
            wx.MessageBox(get_error_message(err), "안내")
            return
        self.EndModal(wx.ID_OK)

    def __on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)

    def GetCTR(self) -> CostCtr:
        return self.__ctr

class DialogCostBS(wx.Dialog):
    def __init__(self, parent: wx.Panel, bs: CostCtr|None):
        title = "Cost BS 수정" if bs else "Cost BS 추가"
        wx.Dialog.__init__(self, parent, title=title)
        self.__bs = bs
        self.__set_layout()
        self.__bind_events()

    def __set_layout(self):
        st_bs   = wx.StaticText(self, label="Cost BS")
        st_code = wx.StaticText(self, label="코드")
        st_rnd  = wx.StaticText(self, label="개발 비중")
        st_oe   = wx.StaticText(self, label="OE 비중")
        tc_bs   = wx.TextCtrl(self, size=(250, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER, value=self.__bs.name if self.__bs else "")
        tc_code = wx.TextCtrl(self, size=(250, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER, value=self.__bs.code if self.__bs else "")
        rb_rnd_dev = wx.RadioButton(self, label="Develop", style=wx.RB_GROUP)
        rb_rnd_res = wx.RadioButton(self, label="Research")
        rb_oe_com = wx.RadioButton(self, label="공통비", style=wx.RB_GROUP)
        rb_oe_oe  = wx.RadioButton(self, label="OE")
        rb_oe_re  = wx.RadioButton(self, label="RE")

        if self.__bs:
            if self.__bs.rnd == EnumRND.DEVELOP:
                rb_rnd_dev.SetValue(True)
            else:
                rb_rnd_res.SetValue(True)
            if self.__bs.oe == EnumOE.COMMON:
                rb_oe_com.SetValue(True)
            elif self.__bs.oe == EnumOE.OE:
                rb_oe_oe.SetValue(True)
            elif self.__bs.oe == EnumOE.RE:
                rb_oe_re.SetValue(True)

        sz_rnd = wx.BoxSizer(wx.HORIZONTAL)
        sz_rnd.AddMany((
            (rb_rnd_dev, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_rnd_res, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_oe = wx.BoxSizer(wx.HORIZONTAL)
        sz_oe.AddMany((
            (rb_oe_com, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_oe_oe , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_oe_re , 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_grid = wx.FlexGridSizer(4, 2, 5, 5)
        sz_grid.AddMany((
            (st_bs  , 0, wx.ALIGN_CENTER_VERTICAL), (tc_bs  , 0, wx.ALIGN_CENTER_VERTICAL),
            (st_code, 0, wx.ALIGN_CENTER_VERTICAL), (tc_code, 0, wx.ALIGN_CENTER_VERTICAL),
            (st_rnd , 0, wx.ALIGN_CENTER_VERTICAL), (sz_rnd , 0, wx.ALIGN_CENTER_VERTICAL),
            (st_oe  , 0, wx.ALIGN_CENTER_VERTICAL), (sz_oe  , 0, wx.ALIGN_CENTER_VERTICAL)
        ))

        bt_confirm = wx.Button(self, label="수정" if self.__bs else "추가")
        bt_cancel = wx.Button(self, label="취소")
        sz_bt = wx.BoxSizer(wx.HORIZONTAL)
        sz_bt.AddMany((
            (bt_confirm, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_cancel, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (sz_grid, 0), ((-1, 10), 0),
            (wx.StaticLine(self), 0, wx.EXPAND), ((-1, 10), 0),
            (sz_bt, 0, wx.ALIGN_RIGHT)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 20)
        self.SetSizerAndFit(sz)
        self.CenterOnParent()

        self.__tc_bs      = tc_bs     
        self.__tc_code    = tc_code   
        self.__rb_rnd_dev = rb_rnd_dev
        self.__rb_rnd_res = rb_rnd_res
        self.__rb_oe_com  = rb_oe_com 
        self.__rb_oe_oe   = rb_oe_oe  
        self.__rb_oe_re   = rb_oe_re  
        self.__bt_confirm = bt_confirm
        self.__bt_cancel  = bt_cancel 

    def __bind_events(self):
        self.__bt_confirm.Bind(wx.EVT_BUTTON, self.__on_confirm)
        self.__bt_cancel .Bind(wx.EVT_BUTTON, self.__on_cancel )
        self.__tc_bs  .Bind(wx.EVT_TEXT_ENTER, self.__on_enter)
        self.__tc_code.Bind(wx.EVT_TEXT_ENTER, self.__on_enter)

    def __on_enter(self, event):
        self.__on_confirm(None)

    def __on_confirm(self, event):
        bs_name = self.__tc_bs.GetValue().strip()
        code = self.__tc_code.GetValue().strip()
        rnd = EnumRND.DEVELOP if self.__rb_rnd_dev.GetValue() else EnumRND.RESEARCH
        oe = EnumOE.COMMON
        if self.__rb_oe_oe.GetValue():
            oe = EnumOE.OE
        elif self.__rb_oe_re.GetValue():
            oe = EnumOE.RE
        if not bs_name:
            wx.MessageBox("BS 이름을 입력하세요.", "안내")
            return
        if not code:
            wx.MessageBox("BS 코드를 입력하세요.", "안내")
            return
        try:
            if self.__bs:
                if (self.__bs.code == code
                    and self.__bs.name == bs_name
                    and self.__bs.rnd == rnd
                    and self.__bs.oe == oe):
                    self.EndModal(wx.ID_CANCEL)
                    return
                self.__bs = self.__bs.update(code, bs_name, rnd, oe, self.__bs.parent_code)
            else:
                root_ctr = CostCtr.get_root_ctr()
                self.__bs = CostCtr.add(code, bs_name, rnd, oe, root_ctr.code)
        except Exception as err:
            wx.MessageBox(get_error_message(err), "안내")
            return
        self.EndModal(wx.ID_OK)

    def __on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)

    def GetBS(self) -> CostCtr|None:
        return self.__bs

class DialogCostTeam(wx.Dialog):
    def __init__(self, parent: wx.Panel, bs: CostCtr, team: CostCtr|None):
        title = "Cost Team 수정" if team else "Cost Team 추가"
        wx.Dialog.__init__(self, parent, title=title)
        self.__bs = bs
        self.__team = team
        self.__set_layout()
        self.__bind_events()

    def __set_layout(self):
        st_bs   = wx.StaticText(self, label="Cost BS")
        st_team   = wx.StaticText(self, label="Cost Team")
        st_code = wx.StaticText(self, label="코드")
        st_rnd  = wx.StaticText(self, label="개발 비중")
        st_oe   = wx.StaticText(self, label="OE 비중")
        tc_bs   = wx.TextCtrl(self, size=(250, -1), style=wx.TE_CENTER|wx.TE_READONLY, value=f"({self.__bs.code}) {self.__bs.name}")
        tc_team = wx.TextCtrl(self, size=(250, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER, value=self.__team.name if self.__team else "")
        tc_code = wx.TextCtrl(self, size=(250, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER, value=self.__team.code if self.__team else "")
        rb_rnd_dev = wx.RadioButton(self, label="Develop", style=wx.RB_GROUP)
        rb_rnd_res = wx.RadioButton(self, label="Research")
        rb_oe_com = wx.RadioButton(self, label="공통비", style=wx.RB_GROUP)
        rb_oe_oe  = wx.RadioButton(self, label="OE")
        rb_oe_re  = wx.RadioButton(self, label="RE")

        if self.__team:
            if self.__team.rnd == EnumRND.DEVELOP:
                rb_rnd_dev.SetValue(True)
            else:
                rb_rnd_res.SetValue(True)
            if self.__team.oe == EnumOE.COMMON:
                rb_oe_com.SetValue(True)
            elif self.__team.oe == EnumOE.OE:
                rb_oe_oe.SetValue(True)
            elif self.__team.oe == EnumOE.RE:
                rb_oe_re.SetValue(True)

        sz_rnd = wx.BoxSizer(wx.HORIZONTAL)
        sz_rnd.AddMany((
            (rb_rnd_dev, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_rnd_res, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_oe = wx.BoxSizer(wx.HORIZONTAL)
        sz_oe.AddMany((
            (rb_oe_com, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_oe_oe , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (rb_oe_re , 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_grid = wx.FlexGridSizer(5, 2, 5, 5)
        sz_grid.AddMany((
            (st_bs  , 0, wx.ALIGN_CENTER_VERTICAL), (tc_bs  , 0, wx.ALIGN_CENTER_VERTICAL),
            (st_team, 0, wx.ALIGN_CENTER_VERTICAL), (tc_team, 0, wx.ALIGN_CENTER_VERTICAL),
            (st_code, 0, wx.ALIGN_CENTER_VERTICAL), (tc_code, 0, wx.ALIGN_CENTER_VERTICAL),
            (st_rnd , 0, wx.ALIGN_CENTER_VERTICAL), (sz_rnd , 0, wx.ALIGN_CENTER_VERTICAL),
            (st_oe  , 0, wx.ALIGN_CENTER_VERTICAL), (sz_oe  , 0, wx.ALIGN_CENTER_VERTICAL)
        ))

        bt_confirm = wx.Button(self, label="수정" if self.__team else "추가")
        bt_cancel = wx.Button(self, label="취소")
        sz_bt = wx.BoxSizer(wx.HORIZONTAL)
        sz_bt.AddMany((
            (bt_confirm, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_cancel, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (sz_grid, 0), ((-1, 10), 0),
            (wx.StaticLine(self), 0, wx.EXPAND), ((-1, 10), 0),
            (sz_bt, 0, wx.ALIGN_RIGHT)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 20)
        self.SetSizerAndFit(sz)
        self.CenterOnParent()

        self.__tc_bs      = tc_bs     
        self.__tc_team    = tc_team   
        self.__tc_code    = tc_code   
        self.__rb_rnd_dev = rb_rnd_dev
        self.__rb_rnd_res = rb_rnd_res
        self.__rb_oe_com  = rb_oe_com 
        self.__rb_oe_oe   = rb_oe_oe  
        self.__rb_oe_re   = rb_oe_re  
        self.__bt_confirm = bt_confirm
        self.__bt_cancel  = bt_cancel 

    def __bind_events(self):
        self.__bt_confirm.Bind(wx.EVT_BUTTON, self.__on_confirm)
        self.__bt_cancel .Bind(wx.EVT_BUTTON, self.__on_cancel )
        self.__tc_team.Bind(wx.EVT_TEXT_ENTER, self.__on_enter)
        self.__tc_code.Bind(wx.EVT_TEXT_ENTER, self.__on_enter)

    def __on_enter(self, event):
        self.__on_confirm(None)

    def __on_confirm(self, event):
        team_name = self.__tc_team.GetValue().strip()
        code = self.__tc_code.GetValue().strip()
        rnd = EnumRND.DEVELOP if self.__rb_rnd_dev.GetValue() else EnumRND.RESEARCH
        oe = EnumOE.COMMON
        if self.__rb_oe_oe.GetValue():
            oe = EnumOE.OE
        elif self.__rb_oe_re.GetValue():
            oe = EnumOE.RE
        if not team_name:
            wx.MessageBox("Team 이름을 입력하세요.", "안내")
            return
        if not code:
            wx.MessageBox("Team 코드를 입력하세요.", "안내")
            return
        try:
            if self.__team:
                if (self.__team.code == code
                    and self.__team.name == team_name
                    and self.__team.rnd == rnd
                    and self.__team.oe == oe):
                    self.EndModal(wx.ID_CANCEL)
                    return
                self.__team = self.__team.update(code, team_name, rnd, oe, self.__team.parent_code)
            else:
                self.__team = CostCtr.add(code, team_name, rnd, oe, self.__bs.code)
        except Exception as err:
            wx.MessageBox(get_error_message(err), "안내")
            return
        self.EndModal(wx.ID_OK)

    def __on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)

    def GetTeam(self) -> CostCtr|None:
        return self.__team

class DialogCurrency(wx.Dialog):
    def __init__(self, parent: wx.Panel, currency: Currency|None):
        title = "환율 수정" if currency else "환율 추가"
        wx.Dialog.__init__(self, parent, title=title)
        self.__currency = currency
        self.__set_layout()
        self.__bind_events()
    
    def __set_layout(self):
        st_code = wx.StaticText(self, label="통화코드")
        st_unit = wx.StaticText(self, label="단위")
        st_1q   = wx.StaticText(self, label="1Q")
        st_2q   = wx.StaticText(self, label="2Q")
        st_3q   = wx.StaticText(self, label="3Q")
        st_4q   = wx.StaticText(self, label="4Q")
        if self.__currency:
            initial_code = self.__currency.code
            initial_unit = self.__currency.unit
            initial_1q   = self.__currency.q1
            initial_2q   = self.__currency.q2
            initial_3q   = self.__currency.q3
            initial_4q   = self.__currency.q4
        else:
            initial_code = ""
            initial_unit = None
            initial_1q   = None
            initial_2q   = None
            initial_3q   = None
            initial_4q   = None
        tc_code = wx.TextCtrl(self, size=(100, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER, value=initial_code)
        sc_unit = wx.SpinCtrl(self, size=(80, -1), value=str(int(initial_unit)) if initial_unit else "100", initial=100 if initial_unit is None else int(initial_unit), min=1, max=99999)
        sc_1q   = wx.SpinCtrlDouble(self, size=(80, -1), value=str(initial_1q  ) if initial_1q   else "100", initial=initial_1q   or 100, min=1, max=99999, inc=0.1)
        sc_2q   = wx.SpinCtrlDouble(self, size=(80, -1), value=str(initial_2q  ) if initial_2q   else "100", initial=initial_2q   or 100, min=1, max=99999, inc=0.1)
        sc_3q   = wx.SpinCtrlDouble(self, size=(80, -1), value=str(initial_3q  ) if initial_3q   else "100", initial=initial_3q   or 100, min=1, max=99999, inc=0.1)
        sc_4q   = wx.SpinCtrlDouble(self, size=(80, -1), value=str(initial_4q  ) if initial_4q   else "100", initial=initial_4q   or 100, min=1, max=99999, inc=0.1)
        sz_grid = wx.FlexGridSizer(6, 2, 5, 5)
        sz_grid.AddMany((
            (st_code, 0, wx.ALIGN_CENTER_VERTICAL), (tc_code, 0, wx.ALIGN_CENTER_VERTICAL),
            (st_unit, 0, wx.ALIGN_CENTER_VERTICAL), (sc_unit, 0, wx.ALIGN_CENTER_VERTICAL),
            (st_1q  , 0, wx.ALIGN_CENTER_VERTICAL), (sc_1q  , 0, wx.ALIGN_CENTER_VERTICAL),
            (st_2q  , 0, wx.ALIGN_CENTER_VERTICAL), (sc_2q  , 0, wx.ALIGN_CENTER_VERTICAL),
            (st_3q  , 0, wx.ALIGN_CENTER_VERTICAL), (sc_3q  , 0, wx.ALIGN_CENTER_VERTICAL),
            (st_4q  , 0, wx.ALIGN_CENTER_VERTICAL), (sc_4q  , 0, wx.ALIGN_CENTER_VERTICAL)
        ))

        bt_confirm = wx.Button(self, label="수정" if self.__currency else "추가")
        bt_cancel = wx.Button(self, label="취소")
        sz_bt = wx.BoxSizer(wx.HORIZONTAL)
        sz_bt.AddMany((
            (bt_confirm, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_cancel, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
    
        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (sz_grid, 0), ((-1, 10), 0),
            (wx.StaticLine(self), 0, wx.EXPAND), ((-1, 10), 0),
            (sz_bt, 0, wx.ALIGN_RIGHT)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 20)
        self.SetSizerAndFit(sz)
        self.CenterOnParent()

        self.__tc_code = tc_code
        self.__sc_unit = sc_unit
        self.__sc_1q   = sc_1q  
        self.__sc_2q   = sc_2q  
        self.__sc_3q   = sc_3q  
        self.__sc_4q   = sc_4q  
        self.__bt_confirm = bt_confirm
        self.__bt_cancel  = bt_cancel 

    def __bind_events(self):
        self.__bt_confirm.Bind(wx.EVT_BUTTON, self.__on_confirm)
        self.__bt_cancel .Bind(wx.EVT_BUTTON, self.__on_cancel )
        self.__tc_code.Bind(wx.EVT_TEXT_ENTER, self.__on_enter)

    def __on_enter(self, event):
        self.__on_confirm(None)

    def __on_confirm(self, event):
        code = self.__tc_code.GetValue().strip()
        unit = self.__sc_unit.GetValue()
        q1   = self.__sc_1q  .GetValue()
        q2   = self.__sc_2q  .GetValue()
        q3   = self.__sc_3q  .GetValue()
        q4   = self.__sc_4q  .GetValue()

        if not code:
            wx.MessageBox("통화코드를 입력하세요.", "안내")
            return

        if self.__currency \
            and self.__currency.code != code \
            and Currency.has_code(code):
            wx.MessageBox("이미 존재하는 통화코드 입니다.", "안내")
            return

        try:
            if self.__currency:
                if (self.__currency.code == code
                    and self.__currency.unit == unit
                    and self.__currency.q1 == q1
                    and self.__currency.q2 == q2
                    and self.__currency.q3 == q3
                    and self.__currency.q4 == q4):
                    self.EndModal(wx.ID_CANCEL)
                    return
                if self.__currency.code != code \
                    and self.__currency.code in LoadedData.get_all_currencies():
                    wx.MessageBox(
                        f"통화코드 '{self.__currency.code}'가 현재 데이터에서 사용되고 있습니다.\n" \
                        "다른 통화코드로 변경할 수 없습니다.",
                        "안내"
                    )
                    return
                self.__currency = self.__currency.update(code, unit, q1, q2, q3, q4)
            else:
                self.__currency = Currency.add(code, unit, q1, q2, q3, q4)
        except Exception as err:
            msg = get_error_message(err)
            wx.MessageBox(msg, "안내")
            return
        self.EndModal(wx.ID_OK)

    def __on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)

    def GetCurrency(self) -> Currency|None:
        return self.__currency

class ModelCategory(TreeListModelBase):
    def __init__(self):
        TreeListModelBase.__init__(self, 1)

    def GetValue(self, item, col):
        if not item:
            return ""
        
        node_id = int(item.GetID())
        if node_id not in self.nodes:
            return ""
        node = self.nodes[node_id]
        node_item: CostCategory = node.item
        match col:
            case 0:
                return node_item.name
        return ""

class ModelElement(TreeListModelBase):
    def __init__(self):
        TreeListModelBase.__init__(self, 3)

    def GetValue(self, item, col):
        if not item:
            return ""
        
        node_id = int(item.GetID())
        if node_id not in self.nodes:
            return ""
        node = self.nodes[node_id]
        node_item: CostElement = node.item
        match col:
            case 0:
                return node_item.code
            case 1:
                return " > ".join(node_item.category_tree[1:])
            case 2:
                return node_item.description
        return ""

class ModelCtr(TreeListModelBase):
    def __init__(self):
        TreeListModelBase.__init__(self, 4)

    def GetValue(self, item, col):
        if not item:
            return ""
        
        node_id = int(item.GetID())
        if node_id not in self.nodes:
            return ""
        node = self.nodes[node_id]
        node_item: CostCtr = node.item
        match col:
            case 0:
                return node_item.name
            case 1:
                return node_item.code
            case 2:
                return node_item.rnd
            case 3:
                return node_item.oe

class ModelCurrency(TreeListModelBase):
    def __init__(self):
        TreeListModelBase.__init__(self, 6)

    def GetValue(self, item, col):
        if not item:
            return ""
        
        node_id = int(item.GetID())
        if node_id not in self.nodes:
            return ""
        node = self.nodes[node_id]
        node_item: Currency = node.item
        match col:
            case 0:
                return node_item.code
            case 1:
                return f"{node_item.unit:,}"
            case 2:
                return f"₩ {node_item.q1:,}"
            case 3:
                return f"₩ {node_item.q2:,}"
            case 4:
                return f"₩ {node_item.q3:,}"
            case 5:
                return f"₩ {node_item.q4:,}"

class ModelData(TreeListModelBase):
    def __init__(self):
        TreeListModelBase.__init__(self, 35)
        self._sort_orders_fixed = tuple(range(11, 34))

    def Compare(self, item1, item2, column, ascending):
        # 0: 미집계
        # 1, 2: 대계정 - 계정항목
        # 6, 7, 3, 4, 5: BS - Team - Ctr Code - 개발 - OE
        # 9, 8: Category - Element
        # 10: CUR
        match column:
            case 1: sort_orders_variable = [1, 2,   6, 7, 3, 4, 5,   9, 8,   0,   10]
            case 2: sort_orders_variable = [2, 1,   6, 7, 3, 4, 5,   9, 8,   0,   10]

            case 3: sort_orders_variable = [3, 6, 7, 4, 5,   9, 8,   1, 2,   0,   10]
            case 4: sort_orders_variable = [4, 3, 6, 7, 5,   9, 8,   1, 2,   0,   10]
            case 5: sort_orders_variable = [5, 3, 6, 7, 4,   9, 8,   1, 2,   0,   10]
            case 6: sort_orders_variable = [6, 3, 7, 4, 5,   9, 8,   1, 2,   0,   10]
            case 7: sort_orders_variable = [7, 3, 6, 4, 5,   9, 8,   1, 2,   0,   10]

            case 8: sort_orders_variable = [8, 9,   6, 7, 3, 4, 5,   1, 2,   0,   10]
            case 9: sort_orders_variable = [9, 8,   6, 7, 3, 4, 5,   1, 2,   0,   10]

            case 10: sort_orders_variable = [10,   6, 7, 3, 4, 5,   9, 8,   1, 2,   0]

            case _: sort_orders_variable = [0,   6, 7, 3, 4, 5,   9, 8,   1, 2,   10]

        ret = +1 if ascending else -1

        columns = sort_orders_variable + list(self._sort_orders_fixed)
        for col in columns:
            if 0 <= col <= 10:
                val1 = self.GetValue(item1, col) or ""
                val2 = self.GetValue(item2, col) or ""
            elif 11 <= col <= 34:
                val1 = self.GetValue(item1, col) or ""
                val2 = self.GetValue(item2, col) or ""
                try:
                    val1 = int(str(val1).replace(",", ""))
                except:
                    val1 = -np.inf
                try:
                    val2 = int(str(val2).replace(",", ""))
                except:
                    val2 = -np.inf
            else:
                continue
            if val1 > val2:
                return ret
            elif val1 < val2:
                return -ret
        return 0

    def GetValue(self, item, col):
        if not item:
            return ""
        
        node_id = int(item.GetID())
        if node_id not in self.nodes:
            return ""
        node = self.nodes[node_id]
        sr: pd.Series = node.item
        ctr_code = sr["Cost Center"]
        elem_code = sr["Cost Element"]
        curr_code = sr["Currency"]
        ctr = LoadedData.cached_cost_ctr.get(ctr_code)
        elem = LoadedData.cached_cost_element.get(elem_code)
        match col:
            case 0: # 미집계
                flag = ctr_code in LoadedData.cached_cost_ctr \
                    and elem_code in LoadedData.cached_cost_element \
                    and curr_code in LoadedData.cached_currency
                return "" if flag else "●"
            case 1: # 대계정
                return sr["대계정"]
            case 2: # 계정항목
                return sr["계정항목"]
            case 3: # Cost Ctr
                return ctr_code
            case 4: # 개발 비중
                return "" if ctr is None else ctr.rnd
            case 5: # OE 비중
                return "" if ctr is None else ctr.oe
            case 6|7: # BS, Team
                if ctr is None:
                    return ""
                ctr_level = LoadedData.get_level_of_ctr_from_cache(ctr)
                if col == 6:
                    if ctr_level == 3:
                        return LoadedData.cached_cost_ctr[ctr.parent_code].name
                    elif ctr_level == 2:
                        return ctr.name
                    return ""
                else:
                    if ctr_level == 3:
                        return ctr.name
                    return ""
            case 8: # Cost Element
                return elem_code
            case 9: # Cost Category
                if elem is None:
                    return ""
                category = LoadedData.cached_cost_category.get(elem.category_pk)
                if category is None:
                    return ""
                return LoadedData.get_category_path_from_cache(category)
            case 10: # 통화코드
                return sr["Currency"]
            case 11|12|13|14|15|16|17|18|19| \
                20|21|22|23|24|25|26|27|28|29| \
                30|31|32|33|34:
                label = "Plan" if col%2 else "Actual"
                month = math.floor((col-11)/2 + 1)
                value = sr[f"Conv{label}({month})"]
                if np.isnan(value):
                    return ""
                return f"{math.floor(value):,}"

    def GetAttr(self, item, col, attr):
        if not item:
            return False
        node_id = int(item.GetID())
        if node_id not in self.nodes:
            return ""
        node = self.nodes[node_id]
        sr: pd.Series = node.item
        match col:
            case 0: # 미집계
                attr.SetColour(wx.Colour(255, 0, 0))
                return True
            case 11|12|13|14|15|16|17|18|19| \
                20|21|22|23|24|25|26|27|28|29| \
                30|31|32|33|34|35:
                label = "Plan" if col%2 else "Actual"
                month = math.floor((col-11)/2 + 1)
                value = sr[f"Conv{label}({month})"]
                if np.isnan(value) \
                    or value >= 0:
                    return True
                attr.SetColour(wx.Colour(255, 0, 0))
                return True
            case _:
                return False

@dataclass
class DropCategory:
    level: int
    category: CostCategory

class TreeCategory(TreeListCtrl):
    EvtCategoryDragged, EVT_CATEGORY_DRAGGED = NewCommandEvent()

    def __init__(self, parent: wx.Panel):
        TreeListCtrl.__init__(
            self,
            parent,
            ModelCategory(),
            {"Cost Category": 100}
        )

        self.__drop_category: DropCategory|None = None
        self.__data_format = wx.DataFormat(wx.DF_UNICODETEXT)
        self.EnableDragSource(self.__data_format)
        self.EnableDropTarget(self.__data_format)
        self.Bind(DV.EVT_DATAVIEW_ITEM_BEGIN_DRAG, self.__on_begin_drag)
        self.Bind(DV.EVT_DATAVIEW_ITEM_DROP_POSSIBLE, self.__on_drop_possible)
        self.Bind(DV.EVT_DATAVIEW_ITEM_DROP, self.__on_drop)

    def __on_begin_drag(self, event):
        """드랍 아이템 확인
        레벨 3 또는 4만 허용
        """
        node_id = int(event.GetItem().GetID())
        node = self.model.nodes[node_id]
        category = node.item
        level = LoadedData.get_level_of_category_from_cache(category)
        if level != 3 and level != 4:
            event.Veto()
            return
        self.__drop_category = DropCategory(level, category)
        data_object = wx.TextDataObject(str(id(node)))
        event.SetDataObject(data_object)
        event.SetDragFlags(wx.Drag_AllowMove)
    
    def __on_drop_possible(self, event):
        """동일 레벨로의 이동만 허용"""
        if event.GetDataFormat() != self.__data_format:
            event.Veto()
            return
        node_id = int(event.GetItem().GetID())
        node = self.model.nodes[node_id]
        category = node.item
        level = LoadedData.get_level_of_category_from_cache(category)
        if level != self.__drop_category.level-1:
            event.Veto()
            return

    def __on_drop(self, event):
        if event.GetDataFormat() != self.__data_format \
            or not self.__drop_category:
            self.__drop_category = None
            event.Veto()
            return
        node_id = int(event.GetItem().GetID())
        node = self.model.nodes[node_id]
        category = node.item
        level = LoadedData.get_level_of_category_from_cache(category)
        try:
            assert level == self.__drop_category.level-1
            assert category.pk != self.__drop_category.category.parent_pk
        except:
            event.Veto()
        else:
            wx.PostEvent(
                self,
                __class__.EvtCategoryDragged(self.Id, category=self.__drop_category.category, new_parent=category)
            )
        finally:
            self.__drop_category = None

class TreeElement(TreeListCtrl):
    def __init__(self, parent: wx.Panel):
        TreeListCtrl.__init__(
            self,
            parent,
            ModelElement(),
            {
                "Cost Element": 100,
                "Cost Category": 250,
                "설명": 380
            }
        )

    def reload_db(self):
        """기존 노드들을 모두 삭제하고 DB의 내용을 불러와서 노드 생성"""
        # for node in self.model.logical_root.children:
        #     self.delete_node(node)
        self.clear_nodes()
        elems = CostElement.get_all()
        for elem in elems.values():
            self.add_node(
                None,
                elem.code,
                elem
            )
        self.Refresh()

    def update_values(self):
        """기존 노드를 삭제하지는 않고 값만 업데이트
        카테고리 정보가 변경된 경우 경로 재계산

        TODO 기타 정보
        """
        with Session() as session:
            for node in self.model.nodes.values():
                if not node.key:
                    continue
                elem = node.item
                elem._category_tree = elem._compute_category_trees(session, elem.category_pk)[elem.category_pk]
        self.Refresh()

@dataclass
class DropTeam:
    bs: CostCtr
    team: CostCtr

class TreeCtr(TreeListCtrl):
    EvtTeamDragged, EVT_TEAM_DRAGGED = NewCommandEvent()

    def __init__(self, parent: wx.Panel):
        TreeListCtrl.__init__(
            self,
            parent,
            ModelCtr(),
            {
                "Cost Ctr": 130,
                "코드": 80,
                "개발 비중": 80,
                "OE 비중": 80
            }
        )
        self.__drop_team: DropTeam|None = None
        self.__data_format = wx.DataFormat(wx.DF_UNICODETEXT)
        self.EnableDragSource(self.__data_format)
        self.EnableDropTarget(self.__data_format)
        self.Bind(DV.EVT_DATAVIEW_ITEM_BEGIN_DRAG, self.__on_begin_drag)
        self.Bind(DV.EVT_DATAVIEW_ITEM_DROP_POSSIBLE, self.__on_drop_possible)
        self.Bind(DV.EVT_DATAVIEW_ITEM_DROP, self.__on_drop)

    def __on_begin_drag(self, event):
        node_id = int(event.GetItem().GetID())
        node = self.model.nodes[node_id]
        ctr = node.item
        if LoadedData.get_level_of_ctr_from_cache(ctr) < 3:
            event.Veto()
            return
        self.__drop_team = DropTeam(LoadedData.get_bs(ctr), ctr)
        data_object = wx.TextDataObject(str(id(node)))
        event.SetDataObject(data_object)
        event.SetDragFlags(wx.Drag_AllowMove)
    
    def __on_drop_possible(self, event):
        if event.GetDataFormat() != self.__data_format:
            event.Veto()
            return
        node_id = int(event.GetItem().GetID())
        node = self.model.nodes[node_id]
        ctr = node.item
        if LoadedData.get_level_of_ctr_from_cache(ctr) != 2:
            event.Veto()
            return

    def __on_drop(self, event):
        if event.GetDataFormat() != self.__data_format \
            or not self.__drop_team:
            self.__drop_team = None
            event.Veto()
            return
        node_id = int(event.GetItem().GetID())
        node = self.model.nodes[node_id]
        ctr = node.item
        if LoadedData.get_level_of_ctr_from_cache(ctr) != 2:
            event.Veto()
            return
        bs = LoadedData.get_bs(ctr)
        try:
            assert bs.code != self.__drop_team.bs.code
        except:
            event.Veto()
        else:
            wx.PostEvent(self, __class__.EvtTeamDragged(self.Id, team=self.__drop_team.team, new_bs=ctr))
        finally:
            self.__drop_team = None

class TreeCurrency(TreeListCtrl):
    def __init__(self, parent: wx.Panel):
        TreeListCtrl.__init__(
            self,
            parent,
            ModelCurrency(),
            {
                "통화코드": 50,
                "단위": 70,
                "1분기": 80,
                "2분기": 80,
                "3분기": 80,
                "4분기": 80
            }
        )

class TreeData(TreeListCtrl):
    def __init__(self, parent: wx.Panel):
        TreeListCtrl.__init__(
            self,
            parent,
            ModelData(),
            {
                "미집계": 50,
                "대계정": 100,
                "계정항목": 100,
                "Cost Ctr": 100,
                "개발 비중": 80,
                "OE 비중": 80,
                "BS": 80,
                "팀": 80,
                "Cost Element": 100,
                "Cost Category": 200,
                "통화코드": 80,
                "계획(1월)" : 80, "실적(1월)" : 80,
                "계획(2월)" : 80, "실적(2월)" : 80,
                "계획(3월)" : 80, "실적(3월)" : 80,
                "계획(4월)" : 80, "실적(4월)" : 80,
                "계획(5월)" : 80, "실적(5월)" : 80,
                "계획(6월)" : 80, "실적(6월)" : 80,
                "계획(7월)" : 80, "실적(7월)" : 80,
                "계획(8월)" : 80, "실적(8월)" : 80,
                "계획(9월)" : 80, "실적(9월)" : 80,
                "계획(10월)": 80, "실적(10월)": 80,
                "계획(11월)": 80, "실적(11월)": 80,
                "계획(12월)": 80, "실적(12월)": 80,
            }
        )

class PanelManager(wx.SplitterWindow):
    def __init__(self, parent: wx.Panel):
        wx.SplitterWindow.__init__(self, parent)
        self.__ctr_expanded = True
        self.__category_expanded = True
        self.__set_layout()
        self.__bind_events()
        self.load_db_values()

    def __set_layout(self):
        pn_top = ScrolledPanel(self)
        pn_top.SetupScrolling(True, False)

        tr_category        = TreeCategory(pn_top)
        bt_category_add    = wx.Button(pn_top, label="추가", style=wx.BU_EXACTFIT)
        bt_category_rename = wx.Button(pn_top, label="이름 수정", style=wx.BU_EXACTFIT)
        bt_category_del    = wx.Button(pn_top, label="삭제", style=wx.BU_EXACTFIT)
        bt_category_expand = wx.Button(pn_top, label="펼치기/접기", style=wx.BU_EXACTFIT)
        sz_category_header = wx.BoxSizer(wx.HORIZONTAL)
        sz_category_header.AddMany((
            (bt_category_add   , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_category_rename, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_category_del   , 0, wx.ALIGN_CENTER_VERTICAL), ((15, -1), 1),
            (bt_category_expand, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_category = wx.BoxSizer(wx.VERTICAL)
        sz_category.AddMany((
            (sz_category_header, 0, wx.EXPAND), ((-1, 3), 0),
            (tr_category, 1, wx.EXPAND)
        ))

        tr_element        = TreeElement (pn_top)
        bt_element_add    = wx.Button(pn_top, label="추가", style=wx.BU_EXACTFIT)
        bt_element_recode = wx.Button(pn_top, label="코드 수정", style=wx.BU_EXACTFIT)
        bt_element_desc   = wx.Button(pn_top, label="설명 수정", style=wx.BU_EXACTFIT)
        bt_element_assign = wx.Button(pn_top, label="카테고리 할당", style=wx.BU_EXACTFIT)
        bt_element_del    = wx.Button(pn_top, label="삭제", style=wx.BU_EXACTFIT)
        tc_element_search = wx.TextCtrl(pn_top, size=(100, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER)
        tc_element_search.SetHint("코드 검색")
        bt_element_search = wx.Button(pn_top, label="검색", style=wx.BU_EXACTFIT)
        sz_element_header = wx.BoxSizer(wx.HORIZONTAL)
        sz_element_header.AddMany((
            (bt_element_add   , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_element_recode, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_element_desc  , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_element_assign, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_element_del   , 0, wx.ALIGN_CENTER_VERTICAL), ((15, -1), 1),
            (tc_element_search, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_element_search, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_element = wx.BoxSizer(wx.VERTICAL)
        sz_element.AddMany((
            (sz_element_header, 0, wx.EXPAND), ((-1, 3), 0),
            (tr_element, 1, wx.EXPAND)
        ))

        tr_ctr      = TreeCtr(pn_top)
        bt_ctr_add  = wx.Button(pn_top, label="추가", style=wx.BU_EXACTFIT)
        bt_ctr_edit = wx.Button(pn_top, label="수정", style=wx.BU_EXACTFIT)
        bt_ctr_del  = wx.Button(pn_top, label="삭제", style=wx.BU_EXACTFIT)
        tc_ctr_search = wx.TextCtrl(pn_top, size=(100, -1), style=wx.TE_CENTER|wx.TE_PROCESS_ENTER)
        tc_ctr_search.SetHint("코드 검색")
        bt_ctr_search = wx.Button(pn_top, label="검색", style=wx.BU_EXACTFIT)
        bt_ctr_expand = wx.Button(pn_top, label="펼치기/접기", style=wx.BU_EXACTFIT)
        sz_ctr_header = wx.BoxSizer(wx.HORIZONTAL)
        sz_ctr_header.AddMany((
            (bt_ctr_add   , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_ctr_edit  , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_ctr_del   , 0, wx.ALIGN_CENTER_VERTICAL), ((15, -1), 1),
            (tc_ctr_search, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_ctr_search, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_ctr_expand, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_ctr = wx.BoxSizer(wx.VERTICAL)
        sz_ctr.AddMany((
            (sz_ctr_header, 0, wx.EXPAND), ((-1, 3), 0),
            (tr_ctr, 1, wx.EXPAND)
        ))

        tr_currency      = TreeCurrency(pn_top)
        bt_currency_add  = wx.Button(pn_top, label="추가", style=wx.BU_EXACTFIT)
        bt_currency_edit = wx.Button(pn_top, label="수정", style=wx.BU_EXACTFIT)
        bt_currency_del  = wx.Button(pn_top, label="삭제", style=wx.BU_EXACTFIT)
        sz_currency_header = wx.BoxSizer(wx.HORIZONTAL)
        sz_currency_header.AddMany((
            (bt_currency_add , 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_currency_edit, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_currency_del , 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_currency = wx.BoxSizer(wx.VERTICAL)
        sz_currency.AddMany((
            (sz_currency_header, 0, wx.EXPAND), ((-1, 3), 0),
            (tr_currency, 1, wx.EXPAND)
        ))

        sz_horz = wx.BoxSizer(wx.HORIZONTAL)
        sz_horz.AddMany((
            (sz_category, 0, wx.EXPAND), ((10, -1), 0),
            (sz_element , 0, wx.EXPAND), ((10, -1), 0),
            (sz_ctr     , 0, wx.EXPAND), ((10, -1), 0),
            (sz_currency, 0, wx.EXPAND)
        ))
        sz_top = wx.BoxSizer(wx.HORIZONTAL)
        sz_top.Add(sz_horz, 1, wx.EXPAND|wx.ALL, 10)
        pn_top.SetSizer(sz_top)

        pn_bot = wx.Panel(self)
        bt_data_assign_category = wx.Button(pn_bot, label="Cost Category 할당")
        st_total_label = wx.StaticText(pn_bot, size=(150, -1), style=wx.ALIGN_CENTER, label="합계")
        st_avail_label = wx.StaticText(pn_bot, size=(150, -1), style=wx.ALIGN_CENTER, label="집계")
        st_na_label    = wx.StaticText(pn_bot, size=(150, -1), style=wx.ALIGN_CENTER, label="미집계")
        st_plan_label  = wx.StaticText(pn_bot, style=wx.ALIGN_CENTER, label="계획")
        st_act_label   = wx.StaticText(pn_bot, style=wx.ALIGN_CENTER, label="실적")
        tc_plan_total  = wx.TextCtrl(pn_bot, value="-", style=wx.TE_READONLY|wx.TE_CENTER)
        tc_plan_avail  = wx.TextCtrl(pn_bot, value="-", style=wx.TE_READONLY|wx.TE_CENTER)
        tc_plan_na     = wx.TextCtrl(pn_bot, value="-", style=wx.TE_READONLY|wx.TE_CENTER)
        tc_act_total   = wx.TextCtrl(pn_bot, value="-", style=wx.TE_READONLY|wx.TE_CENTER)
        tc_act_avail   = wx.TextCtrl(pn_bot, value="-", style=wx.TE_READONLY|wx.TE_CENTER)
        tc_act_na      = wx.TextCtrl(pn_bot, value="-", style=wx.TE_READONLY|wx.TE_CENTER)
        sz_summary = wx.FlexGridSizer(3, 4, 3, 5)
        sz_summary.AddGrowableCol(0)
        sz_summary.AddMany((
            ((-1, -1), 0), (st_total_label, 0, wx.ALIGN_CENTER), (st_avail_label, 0, wx.ALIGN_CENTER), (st_na_label, 0, wx.ALIGN_CENTER),
            (st_plan_label, 0, wx.ALIGN_CENTER), (tc_plan_total, 0, wx.ALIGN_CENTER_VERTICAL|wx.EXPAND), (tc_plan_avail, 0, wx.ALIGN_CENTER_VERTICAL|wx.EXPAND), (tc_plan_na, 0, wx.ALIGN_CENTER_VERTICAL|wx.EXPAND),
            (st_act_label, 0, wx.ALIGN_CENTER), (tc_act_total, 0, wx.ALIGN_CENTER_VERTICAL|wx.EXPAND), (tc_act_avail, 0, wx.ALIGN_CENTER_VERTICAL|wx.EXPAND), (tc_act_na, 0, wx.ALIGN_CENTER_VERTICAL|wx.EXPAND)
        ))
        sz_data_header = wx.BoxSizer(wx.HORIZONTAL)
        sz_data_header.AddMany((
            (bt_data_assign_category, 0, wx.ALIGN_BOTTOM), ((20, -1), 1),
            (sz_summary, 0, wx.ALIGN_BOTTOM)
        ))
        tr_data = TreeData(pn_bot)

        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (sz_data_header, 0, wx.EXPAND), ((-1, 5), 0),
            (tr_data, 1, wx.EXPAND)
        ))
        sz_bot = wx.BoxSizer(wx.HORIZONTAL)
        sz_bot.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 10)
        pn_bot.SetSizer(sz_bot)

        self.SplitHorizontally(pn_top, pn_bot)
        self.SetMinimumPaneSize(300)

        self.__tr_category         = tr_category        
        self.__bt_category_add     = bt_category_add    
        self.__bt_category_rename  = bt_category_rename 
        self.__bt_category_del     = bt_category_del    
        self.__bt_category_expand  = bt_category_expand 

        self.__tr_element        = tr_element       
        self.__bt_element_add    = bt_element_add   
        self.__bt_element_recode = bt_element_recode
        self.__bt_element_desc   = bt_element_desc  
        self.__bt_element_assign = bt_element_assign
        self.__bt_element_del    = bt_element_del   
        self.__tc_element_search = tc_element_search
        self.__bt_element_search = bt_element_search

        self.__tr_ctr        = tr_ctr       
        self.__bt_ctr_add    = bt_ctr_add   
        self.__bt_ctr_edit   = bt_ctr_edit  
        self.__bt_ctr_del    = bt_ctr_del   
        self.__tc_ctr_search = tc_ctr_search
        self.__bt_ctr_search = bt_ctr_search
        self.__bt_ctr_expand = bt_ctr_expand

        self.__tr_currency      = tr_currency     
        self.__bt_currency_add  = bt_currency_add 
        self.__bt_currency_edit = bt_currency_edit
        self.__bt_currency_del  = bt_currency_del 

        self.__bt_data_assign_category = bt_data_assign_category
        self.__tc_plan_total = tc_plan_total
        self.__tc_plan_avail = tc_plan_avail
        self.__tc_plan_na    = tc_plan_na   
        self.__tc_act_total  = tc_act_total 
        self.__tc_act_avail  = tc_act_avail 
        self.__tc_act_na     = tc_act_na    
        self.__tr_data = tr_data

    def __bind_events(self):
        self.__tr_category       .Bind(TreeCategory.EVT_CATEGORY_DRAGGED, self.__on_category_dragged)
        self.__bt_category_add   .Bind(wx.EVT_BUTTON, self.__on_category_add   )
        self.__bt_category_rename.Bind(wx.EVT_BUTTON, self.__on_category_rename)
        self.__bt_category_del   .Bind(wx.EVT_BUTTON, self.__on_category_del   )
        self.__bt_category_expand.Bind(wx.EVT_BUTTON, self.__on_category_expand)

        self.__bt_element_add   .Bind(wx.EVT_BUTTON, self.__on_element_add   )
        self.__bt_element_recode.Bind(wx.EVT_BUTTON, self.__on_element_recode)
        self.__bt_element_desc  .Bind(wx.EVT_BUTTON, self.__on_element_desc  )
        self.__bt_element_assign.Bind(wx.EVT_BUTTON, self.__on_element_assign)
        self.__bt_element_del   .Bind(wx.EVT_BUTTON, self.__on_element_del   )
        self.__bt_element_search.Bind(wx.EVT_BUTTON, self.__on_element_search)
        self.__tc_element_search.Bind(wx.EVT_TEXT_ENTER, self.__on_element_search)

        self.__tr_ctr       .Bind(TreeCtr.EVT_TEAM_DRAGGED, self.__on_team_dragged)
        self.__bt_ctr_add   .Bind(wx.EVT_BUTTON, self.__on_ctr_add   )
        self.__bt_ctr_edit  .Bind(wx.EVT_BUTTON, self.__on_ctr_edit  )
        self.__bt_ctr_del   .Bind(wx.EVT_BUTTON, self.__on_ctr_del   )
        self.__bt_ctr_expand.Bind(wx.EVT_BUTTON, self.__on_ctr_expand)
        self.__bt_ctr_search.Bind(wx.EVT_BUTTON, self.__on_ctr_search)
        self.__tc_ctr_search.Bind(wx.EVT_TEXT_ENTER, self.__on_ctr_search)

        self.__bt_currency_add .Bind(wx.EVT_BUTTON, self.__on_currency_add )
        self.__bt_currency_edit.Bind(wx.EVT_BUTTON, self.__on_currency_edit)
        self.__bt_currency_del .Bind(wx.EVT_BUTTON, self.__on_currency_del )

        self.__bt_data_assign_category.Bind(wx.EVT_BUTTON, self.__on_data_assign_category)

    def __on_category_dragged(self, event):
        category: CostCategory = event.category
        new_parent: CostCategory = event.new_parent
        old_path = LoadedData.get_category_path_from_cache(category)
        new_path = LoadedData.get_category_path_from_cache(new_parent) + f" > {category.name}"
        dlg = wx.MessageDialog(self, f"카테고리를 아래와 같이 수정할까요?\n수정 전: {old_path}\n수정 후: {new_path}", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        new_category = category.update(category.name, new_parent.pk)
        tr = self.__tr_category
        old_node = tr.model.nodes[tr.model.key_vs_id[category.pk]]
        parent_node = tr.model.nodes[tr.model.key_vs_id[new_parent.pk]]
        tr.delete_node(old_node)
        new_node = tr.add_node(parent_node, new_category.pk, new_category)
        for child in new_category.children:
            tr.add_node(new_node, child.pk, child)
        tr.expand_node(new_node, True)
        tr.reveal_and_select(new_node)
        LoadedData.cache_category()
        self.__tr_element.update_values()
        wx.MessageBox("카테고리를 수정했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_category_add(self, event):
        tr = self.__tr_category
        node = tr.get_selected_node()
        if node is None:
            wx.MessageBox("카테고리를 추가할 상위 카테고리를 선택하세요.", "안내")
            return
        if node.get_level() == MAXIMUM_DEPTH_OF_CATEGORY+1:
            wx.MessageBox("최하위 카테고리입니다.\n더 이상 추가할 수 없습니다.", "안내")
            return
        parent_cat = CostCategory.get(node.key)
        sibling_names = [child.name for child in parent_cat.children if child.name != node.item.name]
        value = ""
        while True:
            dlg = wx.TextEntryDialog(self, "카테고리 이름을 입력하세요.", "안내", value)
            ret = dlg.ShowModal()
            value = dlg.GetValue().strip()
            dlg.Destroy()
            if ret != wx.ID_OK:
                return
            if not value:
                continue
            if value in sibling_names:
                wx.MessageBox("카테고리 이름이 중복됩니다.", "안내")
                continue
            category = CostCategory.add(value, parent_cat.pk)
            tr.add_node(node, category.pk, category)
            LoadedData.cache_category()
            wx.MessageBox("Cost Category를 추가했습니다.", "안내")
            return

    def __on_category_rename(self, event):
        tr = self.__tr_category
        node = tr.get_selected_node()
        if node is None \
            or node.get_level() == 1:
            wx.MessageBox("이름을 수정하고자 하는 카테고리를 선택하세요.", "안내")
            return
        parent_cat = CostCategory.get(node.parent.key)
        sibling_names = [child.name for child in parent_cat.children if child.name != node.item.name]
        category: CostCategory = node.item
        value = category.name
        while True:
            dlg = wx.TextEntryDialog(self, "카테고리 이름을 입력하세요.", "안내", value)
            ret = dlg.ShowModal()
            value = dlg.GetValue().strip()
            dlg.Destroy()
            if ret != wx.ID_OK:
                return
            if value == category.name:
                return
            if not value:
                continue
            if value in sibling_names:
                wx.MessageBox("카테고리 이름이 중복됩니다.", "안내")
                continue
            node.item = category.update(value, category.parent_pk)
            tr.update_node(node)
            LoadedData.cache_category()
            self.__tr_element.update_values()
            wx.MessageBox("Cost Category 이름을 수정했습니다.", "안내")
            wx.PostEvent(self, EvtUpdate(self.Id))
            return

    def __on_category_del(self, event):
        tr = self.__tr_category
        node = tr.get_selected_node()
        if node is None \
            or node.get_level() == 1:
            wx.MessageBox("삭제하고자 하는 카테고리를 선택하세요.", "안내")
            return
        dlg = wx.MessageDialog(self, "모든 하위 카테고리가 삭제됩니다.\n카테고리를 삭제할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        CostCategory.delete(node.key)
        tr.delete_node(node)
        LoadedData.cache_category()
        self.__tr_element.update_values()
        wx.MessageBox("Cost Category를 삭제했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_category_expand(self, event):
        tr = self.__tr_category
        self.__category_expanded = not self.__category_expanded
        logical_root = tr.get_logical_root()
        if not self.__category_expanded:
            for root in logical_root.children:
                tr.expand_node(root, True)
                for child in root.children:
                    tr.expand_node(child, False)
        else:
            for node in tr.model.nodes.values():
                tr.expand_node(node, True)

    def __on_element_add(self, event):
        tr = self.__tr_element
        value = ""
        while True:
            dlg = wx.TextEntryDialog(self, "코드를 입력하세요.", "안내", value)
            ret = dlg.ShowModal()
            value = dlg.GetValue().strip()
            dlg.Destroy()
            if ret != wx.ID_OK:
                return
            if not value:
                continue
            if CostElement.has_code(value):
                wx.MessageBox("이미 존재하는 코드 입니다.", "안내")
                continue
            element = CostElement.add(value, None)
            node = tr.add_node(None, element.code, element)
            tr.reveal_and_select(node)
            LoadedData.cache_element()
            self.__tr_data.Refresh()
            self.update_summary()
            wx.MessageBox("Cost Element를 추가했습니다.", "안내")
            return

    def __on_element_recode(self, event):
        tr = self.__tr_element
        node = tr.get_selected_node()
        if node is None:
            wx.MessageBox("코드를 수정할 Cost Element를 선택하세요.", "안내")
            return
        element: CostElement = node.item
        value = node.key
        while True:
            dlg = wx.TextEntryDialog(self, "코드를 입력하세요.", "안내", value)
            ret = dlg.ShowModal()
            value = dlg.GetValue().strip()
            dlg.Destroy()
            if ret != wx.ID_OK:
                return
            if not value:
                continue
            if value == node.key:
                return
            if CostElement.has_code(value):
                wx.MessageBox("이미 존재하는 코드 입니다.", "안내")
                continue
            element = element.update(value, element.category_pk)
            node.key = value
            node.item = element
            tr.update_node(node)
            LoadedData.cache_element()
            self.__tr_data.Refresh()
            self.update_summary()
            wx.PostEvent(self, EvtUpdate(self.Id))
            wx.MessageBox("Cost Element 코드를 수정했습니다.", "안내")
            return

    def __on_element_desc(self, event):
        tr = self.__tr_element
        node = tr.get_selected_node()
        if node is None:
            wx.MessageBox("설명을 수정할 Cost Element를 선택하세요.", "안내")
            return
        element: CostElement = node.item
        value = element.description
        while True:
            dlg = TextEntryDialog(self, "안내", "설명을 입력하세요.", value, wx.Size(350, -1))
            ret = dlg.ShowModal()
            value = dlg.GetValue().strip()
            dlg.Destroy()
            if ret != wx.ID_OK:
                return
            if not value:
                continue
            if value == element.description:
                return
            if len(value) > 100:
                wx.MessageBox("최대 100자 이하로 입력하세요.", "안내")
                continue
            element = element.update(element.code, element.category_pk, value)
            node.item = element
            tr.update_node(node)
            LoadedData.cache_element()
            self.__tr_data.Refresh()
            self.update_summary()
            wx.PostEvent(self, EvtUpdate(self.Id))
            wx.MessageBox("Cost Element 설명을 수정했습니다.", "안내")
            return

    def __on_element_assign(self, event):
        tr = self.__tr_element
        node = tr.get_selected_node()
        if node is None:
            wx.MessageBox("Cost Element를 선택하세요.", "안내")
            return
        
        node_cat = self.__tr_category.get_selected_node()
        if node_cat is None:
            wx.MessageBox("Cost Category를 선택하세요.", "안내")
            return
        if node_cat.get_level() != MAXIMUM_DEPTH_OF_CATEGORY+1:
            wx.MessageBox("최하위 Cost Category를 선택하세요.", "안내")
            return
        category_tree = []
        curr = node_cat
        while curr.item:
            category_tree.append(curr.item.name)
            curr = curr.parent
        
        dlg = wx.MessageDialog(
            self,
            f"Cost Element({node.key})의 카테고리를 아래로 설정할까요?\n" \
                f"[{' > '.join(category_tree[::-1][1:])}]",
            "안내",
            style=wx.YES_NO|wx.NO_DEFAULT
        )
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        element: CostElement = node.item
        element = element.update(element.code, node_cat.key)
        node.item = element
        tr.update_node(node)
        LoadedData.cache_element()
        self.__tr_data.Refresh()
        self.update_summary()
        wx.PostEvent(self, EvtUpdate(self.Id))
        wx.MessageBox("Cost Element에 카테고리를 할당했습니다.", "안내")
       
    def __on_element_del(self, event):
        tr = self.__tr_element
        node = tr.get_selected_node()
        if node is None:
            wx.MessageBox("삭제할 Cost Element를 선택하세요.", "안내")
            return
        dlg = wx.MessageDialog(self, "Cost Element를 삭제할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        CostElement.delete(node.key)
        tr.delete_node(node)
        LoadedData.cache_element()
        self.__tr_data.Refresh()
        self.update_summary()
        wx.MessageBox("Cost Element를 삭제했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_element_search(self, event):
        code = self.__tc_element_search.GetValue().strip()
        if not code:
            return
        tr = self.__tr_element
        for node in tr.model.nodes.values():
            if not node.key:
                continue
            if code not in node.key:
                continue
            tr.reveal_and_select(node)
            return

    def __on_team_dragged(self, event):
        team: CostCtr = event.team
        new_bs: CostCtr = event.new_bs
        dlg = wx.MessageDialog(self, f"[{team.name}] 팀을 [{new_bs.name}] BS로 이동할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        ctr = team.update(team.code, team.name, team.rnd, team.oe, new_bs.code)
        tr = self.__tr_ctr
        old_node = tr.model.nodes[tr.model.key_vs_id[team.code]]
        bs_node = tr.model.nodes[tr.model.key_vs_id[new_bs.code]]
        tr.delete_node(old_node)
        new_node = tr.add_node(bs_node, ctr.code, ctr)
        tr.reveal_and_select(new_node)
        LoadedData.cache_ctr()
        self.__tr_data.Refresh()
        self.update_summary()
        wx.MessageBox("팀을 이동했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_ctr_add(self, event):
        root_ctr = CostCtr.get_root_ctr()
        tr = self.__tr_ctr
        node = tr.get_selected_node()
        if node is None \
            or node.get_level() == 3:
            wx.MessageBox(f"BS를 추가하려면 '{root_ctr.name}'을 선택,\n팀을 추가하려면 BS를 선택하세요.", "안내")
            return
        level = node.get_level()
        match level:
            case 1:
                dlg = DialogCostBS(self, None)
                ret = dlg.ShowModal()
                ctr = dlg.GetBS()
                dlg.Destroy()
            case 2:
                bs = CostCtr.get(node.key)
                dlg = DialogCostTeam(self, bs, None)
                ret = dlg.ShowModal()
                ctr = dlg.GetTeam()
                dlg.Destroy()
            case _:
                wx.MessageBox(f"Invalid level: {level}", "안내")
                return
        if ret != wx.ID_OK:
            return
        updated_node = tr.add_node(node, ctr.code, ctr)
        tr.reveal_and_select(updated_node)
        LoadedData.cache_ctr()
        self.update_summary()
        wx.MessageBox("Cost Ctr를 추가했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_ctr_edit(self, event):
        tr = self.__tr_ctr
        node = tr.get_selected_node()
        if node is None:
            wx.MessageBox(f"수정할 Cost Ctr를 선택하세요.", "안내")
            return
        level = node.get_level()
        match level:
            case 1:
                dlg = DialogRootCtr(self, CostCtr.get(node.key))
                ret = dlg.ShowModal()
                ctr = dlg.GetCTR()
                dlg.Destroy()
            case 2:
                dlg = DialogCostBS(self, CostCtr.get(node.key))
                ret = dlg.ShowModal()
                ctr = dlg.GetBS()
                dlg.Destroy()
            case 3:
                bs = CostCtr.get(node.parent.key)
                team = CostCtr.get(node.key)
                dlg = DialogCostTeam(self, bs, team)
                ret = dlg.ShowModal()
                ctr = dlg.GetTeam()
                dlg.Destroy()
            case _:
                wx.MessageBox(f"Invalid level: {level}", "안내")
                return
        if ret != wx.ID_OK:
            return
        node.item = ctr
        node.key = ctr.code
        tr.update_node(node)
        tr.reveal_and_select(node)
        LoadedData.cache_ctr()
        self.__tr_data.Refresh()
        self.update_summary()
        wx.MessageBox("Cost Ctr 정보를 수정했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_ctr_del(self, event):
        tr = self.__tr_ctr
        node = tr.get_selected_node()
        if node is None \
            or node.get_level() == 1:
            wx.MessageBox(f"삭제할 BS 또는 팀을 선택하세요.", "안내")
            return
        level = node.get_level()
        match level:
            case 2:
                dlg = wx.MessageDialog(self, "모든 하위 팀 정보가 삭제됩니다.\n선택한 BS를 삭제할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
                ret = dlg.ShowModal()
                dlg.Destroy()
            case 3:
                dlg = wx.MessageDialog(self, "선택한 팀을 삭제할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
                ret = dlg.ShowModal()
                dlg.Destroy()
            case _:
                wx.MessageBox(f"Invalid level: {level}", "안내")
                return
        if ret != wx.ID_YES:
            return
        CostCtr.delete(node.key)
        tr.delete_node(node)
        LoadedData.cache_ctr()
        self.__tr_data.Refresh()
        self.update_summary()
        wx.Yield()
        wx.MessageBox("Cost Ctr를 삭제했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_ctr_expand(self, event):
        self.__ctr_expanded = not self.__ctr_expanded
        logical_root = self.__tr_ctr.get_logical_root()
        for root in logical_root.children:
            self.__tr_ctr.expand_node(root, True)
            for bs in root.children:
                self.__tr_ctr.expand_node(bs, self.__ctr_expanded)

    def __on_ctr_search(self, event):
        code = self.__tc_ctr_search.GetValue().strip()
        if not code:
            return
        tr = self.__tr_ctr
        for node in tr.model.nodes.values():
            if not node.key:
                continue
            if code not in node.key:
                continue
            tr.reveal_and_select(node)
            return

    def __on_currency_add(self, event):
        dlg = DialogCurrency(self, None)
        ret = dlg.ShowModal()
        currency = dlg.GetCurrency()
        dlg.Destroy()
        if ret != wx.ID_OK:
            return
        self.__tr_currency.add_node(None, currency.code, currency)
        LoadedData.cache_currency()
        LoadedData.update_currency()
        self.redraw_data_tree()
        wx.MessageBox("환율 정보를 추가했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_currency_edit(self, event):
        node = self.__tr_currency.get_selected_node()
        if not node:
            wx.MessageBox("수정할 환율 정보를 선택하세요.", "안내")
            return
        code = node.key
        currency = Currency.get(code)
        if not currency:
            wx.MessageBox(f"환율 정보를 찾을 수 없습니다. ({code})", "안내")
            return
        dlg = DialogCurrency(self, currency)
        ret = dlg.ShowModal()
        currency = dlg.GetCurrency()
        dlg.Destroy()
        if ret != wx.ID_OK:
            return
        node.key = currency.code
        node.item = currency
        self.__tr_currency.update_node(node)
        LoadedData.cache_currency()
        LoadedData.update_currency()
        self.redraw_data_tree()
        wx.MessageBox("환율 정보를 수정했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_currency_del(self, event):
        node = self.__tr_currency.get_selected_node()
        if not node:
            wx.MessageBox("삭제할 환율 정보를 선택하세요.", "안내")
            return
        code = node.key
        dlg = wx.MessageDialog(self, f"통화코드 '{code}'에 대한 환율 정보를 삭제할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        Currency.delete(code)
        self.__tr_currency.delete_node(node)
        LoadedData.cache_currency()
        LoadedData.update_currency()
        self.redraw_data_tree()
        wx.MessageBox("환율 정보를 삭제했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def __on_data_assign_category(self, event):
        selected_data_node = self.__tr_data.get_selected_node()
        if not selected_data_node:
            wx.MessageBox("Cost Category를 할당할 데이터 행을 선택하세요.", "안내")
            return
        target_element_code = selected_data_node.item["Cost Element"]
        selected_category_node = self.__tr_category.get_selected_node()
        if not selected_category_node:
            wx.MessageBox(f"'{target_element_code}'에 할당할 Cost Category를 선택하세요.", "안내")
            return
        category: CostCategory = selected_category_node.item
        if LoadedData.get_level_of_category_from_cache(category) < 4:
            wx.MessageBox("최하위 Cost Category를 선택하세요.", "안내")
            return
        dlg = wx.MessageDialog(
            self,
            f"Cost Element({target_element_code})의 Cost Category를 아래와 같이 설정할까요?\n{LoadedData.get_category_path_from_cache(category)}",
            "안내",
            style=wx.YES_NO|wx.NO_DEFAULT
        )
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        tr = self.__tr_element
        if target_element_code in LoadedData.cached_cost_element:
            element = LoadedData.cached_cost_element[target_element_code]
            element = element.update(target_element_code, category.pk)
            node = tr.get_node_by_key(element.code)
            node.item = element
            tr.update_node(node)
        else:
            element = CostElement.add(target_element_code, category.pk)
            node = tr.add_node(None, element.code, element)
        LoadedData.cache_category()
        LoadedData.cache_element()
        tr.reveal_and_select(node)
        self.__tr_data.Refresh()
        self.update_summary()
        wx.MessageBox("Cost Category를 할당했습니다.", "안내")
        wx.PostEvent(self, EvtUpdate(self.Id))

    def load_db_values(self):
        # 환율
        currencies = Currency.get_all()
        self.__tr_currency.clear_nodes()
        for curr in currencies.values():
            self.__tr_currency.add_node(None, curr.code, curr)
        self.__tr_currency.Refresh()

        # Cost Ctr
        ctrs = CostCtr.get_all()
        self.__tr_ctr.clear_nodes()
        for ctr in ctrs.values():
            self.__tr_ctr.add_node(
                ctr.parent_code and self.__tr_ctr.get_node_by_key(ctr.parent_code),
                ctr.code,
                ctr
            )
        self.__tr_ctr.Refresh()

        # Cost Category
        cats = CostCategory.get_all()
        self.__tr_category.clear_nodes()
        for cat in cats.values():
            self.__tr_category.add_node(
                self.__tr_category.get_node_by_key(cat.parent_pk) if cat.parent_pk is not None else None,
                cat.pk,
                cat
            )
        self.__tr_category.Refresh()

        # Cost Element
        self.__tr_element.reload_db()

    def redraw_data_tree(self):
        tr = self.__tr_data
        tr.clear_nodes()
        df = LoadedData.df
        for key in df.index:
            tr.add_node(None, key, df.loc[key])
        self.update_summary()

    def update_summary(self):
        """우측 하단의 집계/미집계 합계 금액 업데이트"""
        df = LoadedData.df
        if df.empty:
            self.__tc_plan_total.SetValue("-")
            self.__tc_plan_avail.SetValue("-")
            self.__tc_plan_na   .SetValue("-")
            self.__tc_act_total .SetValue("-")
            self.__tc_act_avail .SetValue("-")
            self.__tc_act_na    .SetValue("-")
            return
        mask = LoadedData.get_available_mask()
        self.__tc_plan_total.SetValue(f'{int(df       [[f"ConvPlan({i})"   for i in range(1, 13)]].fillna(0).sum().sum()):,}')
        self.__tc_plan_avail.SetValue(f'{int(df[ mask][[f"ConvPlan({i})"   for i in range(1, 13)]].fillna(0).sum().sum()):,}')
        self.__tc_plan_na   .SetValue(f'{int(df[~mask][[f"ConvPlan({i})"   for i in range(1, 13)]].fillna(0).sum().sum()):,}')
        self.__tc_act_total .SetValue(f'{int(df       [[f"ConvActual({i})" for i in range(1, 13)]].fillna(0).sum().sum()):,}')
        self.__tc_act_avail .SetValue(f'{int(df[ mask][[f"ConvActual({i})" for i in range(1, 13)]].fillna(0).sum().sum()):,}')
        self.__tc_act_na    .SetValue(f'{int(df[~mask][[f"ConvActual({i})" for i in range(1, 13)]].fillna(0).sum().sum()):,}')