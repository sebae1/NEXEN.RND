import wx

class PanelAspectRatio(wx.Panel):
    def __init__(self, parent, ar: float, growable_height: bool, *args, **kw):
        """ar = w/h"""
        super().__init__(parent, *args, **kw)
        self.__ar = ar
        self.__growable_height = growable_height
        self.Bind(wx.EVT_SIZE, self.__on_size)

    def __on_size(self, event):
        w, h = self.GetClientSize()
        if self.__growable_height:
            h = int(w/self.__ar)
            self.SetMinSize((-1, h))
            self.SetMaxSize((-1, h))
        else:
            w = int(self.__ar*w)
            self.SetMinSize((w, -1))
            self.SetMaxSize((w, -1))
        event.Skip()
