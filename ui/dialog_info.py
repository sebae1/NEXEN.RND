import wx
from util import APP_NAME, VERSION
from .component import NEXEN_LOGO_FULL_SVG

class DialogInfo(wx.Dialog):
    def __init__(self, parent: wx.Window):
        super().__init__(parent, title="정보")

        w, h = self.FromDIP(130), self.FromDIP(90)
        size = wx.Size(w, h)
        bundle = wx.BitmapBundle.FromSVG(NEXEN_LOGO_FULL_SVG.encode("utf-8"), size)
        bmp = bundle.GetBitmap(size)
        sb = wx.StaticBitmap(self, bitmap=bmp)

        font = self.GetFont().Bold()
        font.SetPointSize(12)
        st_title = wx.StaticText(self, label=APP_NAME)
        st_title.SetFont(font)

        font = self.GetFont().Bold()
        font.SetPointSize(10)
        st_version = wx.StaticText(self, label=f"< {VERSION} >")
        st_version.SetFont(font)
        st_version.SetForegroundColour(wx.Colour(100, 100, 100))

        st_hypeware = wx.StaticText(self, label="2025 Hypeware\ncontact@hypeware.co.kr", style=wx.ALIGN_CENTER)

        bt_close = wx.Button(self, label="닫기")

        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (sb, 0, wx.ALIGN_CENTER_HORIZONTAL), ((-1, 15), 0),
            (st_title, 0, wx.ALIGN_CENTER_HORIZONTAL), ((-1, 5), 0),
            (st_version, 0, wx.ALIGN_CENTER_HORIZONTAL), ((-1, 30), 0),
            (st_hypeware, 0, wx.ALIGN_CENTER_HORIZONTAL), ((-1, 20), 0),
            (bt_close, 0, wx.EXPAND)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 40)
        self.SetSizerAndFit(sz)
        self.CenterOnParent()

        bt_close.Bind(wx.EVT_BUTTON, self.__on_close)
    
    def __on_close(self, event):
        self.EndModal(wx.ID_CLOSE)
