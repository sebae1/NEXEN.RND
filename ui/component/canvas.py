import wx
import numpy as np
import matplotlib.pyplot as plt
from typing import Callable
from matplotlib.figure import Figure
from matplotlib.axes import Axes
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas

class PanelCanvas(wx.Panel):
    def __init__(
            self,
            parent: wx.Window,
            nrows: int = 1,
            ncols: int = 1,
            constrained_layout: bool = True,
            save_fig_callback: Callable[[str], None] | None = None,
            **kwargs
        ):
        wx.Panel.__init__(self, parent, **kwargs)
        self._save_fig_callback = save_fig_callback
        self._fig, self._ax = plt.subplots(nrows, ncols, constrained_layout=constrained_layout)
        self._cv = FigureCanvas(self, -1, self._fig)
        self._cv.SetMinSize((10, 10))
        self._cv.Bind(wx.EVT_RIGHT_DOWN, self._on_right_click)
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(self._cv, 1, wx.EXPAND)
        self.SetSizer(sz)
        self.Bind(wx.EVT_WINDOW_DESTROY, self._on_destroy)

    def _on_destroy(self, event):
        plt.close(self._fig)
        event.Skip()

    def _on_right_click(self, event):
        menu = wx.Menu()
        save_item = menu.Append(wx.ID_ANY, "이미지 저장")
        self.Bind(wx.EVT_MENU, self._on_save_chart, save_item)
        self.PopupMenu(menu)
        menu.Destroy()

    def _on_save_chart(self, event):
        dlg = wx.FileDialog(None, "이미지 저장", wildcard="PNG 이미지 (*.png)|*.png", style=wx.FD_SAVE|wx.FD_OVERWRITE_PROMPT)
        res = dlg.ShowModal()
        filepath = dlg.GetPath()
        dlg.Destroy()
        if res != wx.ID_OK:
            return
        if self._save_fig_callback:
            self._save_fig_callback(filepath)
        else:
            self._fig.savefig(filepath, dpi=300)
        wx.MessageBox("이미지를 저장하였습니다.", "안내", wx.OK|wx.ICON_INFORMATION)

    @property
    def fig(self) -> Figure:
        return self._fig
    
    @property
    def ax(self) -> Axes|np.ndarray[Axes,]:
        return self._ax

    @property
    def canvas(self):
        return self._cv

    def draw(self):
        self._cv.draw()
