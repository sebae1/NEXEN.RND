import markdown
import json
import base64
import wx
import wx.html2 as webview

class DialogAIResult(wx.Dialog):
    def __init__(self, parent: wx.Window, ai_type: str, model: str, json_data: str, mark_down_result: str):
        super().__init__(parent, title="AI 분석 결과", style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER)
        st_ai_type = wx.StaticText(self, label="AI")
        st_model   = wx.StaticText(self, label="모델")
        tc_ai_type = wx.TextCtrl(self, size=(150, -1), style=wx.TE_READONLY|wx.TE_CENTER, value=ai_type)
        tc_model   = wx.TextCtrl(self, size=(150, -1), style=wx.TE_READONLY|wx.TE_CENTER, value=model)
        sz_grid = wx.FlexGridSizer(2, 2, 5, 5)
        sz_grid.AddGrowableCol(0)
        sz_grid.AddMany((
            (st_ai_type, 0, wx.ALIGN_CENTER_VERTICAL), (tc_ai_type, 0, wx.ALIGN_CENTER_VERTICAL),
            (st_model  , 0, wx.ALIGN_CENTER_VERTICAL), (tc_model  , 0, wx.ALIGN_CENTER_VERTICAL)
        ))

        notebook = wx.Notebook(self)
        view = webview.WebView.New(notebook)
        html = markdown.markdown(mark_down_result, extensions=["fenced_code", "tables"])
        styled_html = f"""
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{
                    font-family: 'Segoe UI', sans-serif;
                    margin: 20px;
                    line-height: 1.6;
                }}
                pre, code {{
                    background-color: #f5f5f5;
                    padding: 4px 8px;
                    border-radius: 4px;
                    font-family: Consolas, monospace;
                }}
                h1, h2, h3 {{
                    color: #333;
                    border-bottom: 1px solid #ddd;
                    padding-bottom: 4px;
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                }}
                th, td {{
                    border: 1px solid #ccc;
                    padding: 4px 8px;
                    text-align: left;
                }}
            </style>
        </head>
        <body>{html}</body></html>
        """
        view.SetPage(styled_html, "")

        pn_data = wx.Panel(notebook)
        tc_data = wx.TextCtrl(
            pn_data,
            value=json.dumps(json.loads(json_data), ensure_ascii=False, indent=2),
            style=wx.TE_MULTILINE
        )
        sz_data = wx.BoxSizer(wx.HORIZONTAL)
        sz_data.Add(tc_data, 1, wx.EXPAND|wx.ALL, 20)
        pn_data.SetSizer(sz_data)

        notebook.AddPage(view, "분석 결과")
        notebook.AddPage(pn_data, "분석 데이터")

        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (sz_grid, 0), ((-1, 10), 0),
            (notebook, 1, wx.EXPAND)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 20)
        self.SetSizer(sz)
        display_size = wx.GetDisplaySize()
        self.SetSize((min(900, display_size[0]), int(display_size[1]*0.9)))
        self.SetMinSize(self.GetSize())
        self.CenterOnScreen()

class DialogModels(wx.Dialog):
    def __init__(self, parent: wx.Window, models: list[str], initial_model: str):
        super().__init__(parent, title="모델 선택")
        self.__cb = wx.ComboBox(self, value=initial_model, choices=models, style=wx.CB_READONLY)
        bt_confirm = wx.Button(self, label="확인")
        bt_cancel = wx.Button(self, label="취소")
        sz_bt = wx.BoxSizer(wx.HORIZONTAL)
        sz_bt.AddMany((
            (bt_confirm, 1, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_cancel, 1, wx.ALIGN_CENTER_VERTICAL)
        ))
        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (self.__cb, 0, wx.EXPAND), ((-1, 5), 0),
            (sz_bt, 0, wx.EXPAND)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 30)
        self.SetSizerAndFit(sz)
        self.CenterOnParent()

        bt_confirm.Bind(wx.EVT_BUTTON, self.__on_confirm)
        bt_cancel.Bind(wx.EVT_BUTTON, self.__on_cancel)
    
    def __on_confirm(self, event):
        self.EndModal(wx.ID_OK)
    
    def __on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)
    
    def get_model(self) -> str:
        return self.__cb.GetValue()

class DialogPDF(wx.Dialog):
    def __init__(self, parent: wx.Window, pdf_bytes: bytes):
        wx.Dialog.__init__(self, parent, title="AI 분석 결과", size=(900, 800), style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER)
        self.view = webview.WebView.New(self)
        b64 = base64.b64encode(pdf_bytes).decode("ascii")
        self.view.SetPage(
            f'<iframe src="data:application/pdf;base64,{b64}" '
            'style="width:100%;height:100%;border:0;"></iframe>', ""
        )
        self.CenterOnScreen()
