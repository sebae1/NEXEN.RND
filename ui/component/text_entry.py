import wx

class TextEntryDialog(wx.Dialog):
    def __init__(self, parent: wx.Window, title: str, description: str, initial_value: str = "", entry_size: wx.Size = wx.DefaultSize):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER)
        st = wx.StaticText(self, label=description)
        tc = wx.TextCtrl(self, value=initial_value, size=entry_size)
        bt_confirm = wx.Button(self, label="확인")
        bt_cancel = wx.Button(self, label="취소")
        sz_bt = wx.BoxSizer(wx.HORIZONTAL)
        sz_bt.AddMany((
            (bt_confirm, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_cancel, 0, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (st, 0), ((-1, 5), 0),
            (tc, 1, wx.EXPAND), ((-1, 5), 0),
            (sz_bt, 0, wx.ALIGN_RIGHT)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 20)
        self.SetSizerAndFit(sz)
        self.CenterOnParent()
        self.SetMinSize(self.GetSize())
        self.SetMaxSize((-1, self.GetSize()[1]))

        self.__tc = tc

        bt_confirm.Bind(wx.EVT_BUTTON, self.__on_confirm)
        bt_cancel.Bind(wx.EVT_BUTTON, self.__on_cancel)
    
    def __on_confirm(self, event):
        self.EndModal(wx.ID_OK)
    
    def __on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)
    
    def GetValue(self) -> str:
        return self.__tc.GetValue()

