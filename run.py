import wx

from traceback import format_exc
from ui import FrameMain
from db import initialize_db, LoadedData
from util import initialize_matplotlib

def main():
    app = wx.App()
    try:
        initialize_db()
        initialize_matplotlib()
        LoadedData.cache_all()
    except:
        wx.MessageBox(format_exc(), '오류')
    else:
        FrameMain().Show()
    app.MainLoop()

if __name__ == "__main__":
    main()
