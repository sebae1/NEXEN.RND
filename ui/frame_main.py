import os
import shutil
import wx

from traceback import format_exc
from threading import Thread
from sqlalchemy import delete

from ai import get_gpt_models, get_claude_models
from db import LoadedData, EXT, DATABASE_PATH, Session, get_engine, validate_db
from db.models import read_ctr_excel, read_element_excel, CostCategory, CostCtr, CostElement
from util import APP_NAME, get_error_message, Config
from ui.component import EVT_UPDATE, NEXEN_LOGO_SVG, WARNING_MARK_SVG
from ui.panel_dashboard import PanelDashboard
from ui.panel_viewer import PanelViewer
from ui.panel_manager import PanelManager
from ui.panel_bs import PanelBSChart
from ui.dialog_info import DialogInfo
from ui.dialog_licence import DialogOSSL

class DialogManageRawData(wx.Dialog):
    def __init__(self, parent: wx.Window):
        wx.Dialog.__init__(self, parent, title="Raw Data 관리")
        self.__set_layout()
        self.__bind_events()
        self.__updated = False
    
    def __set_layout(self):
        lc = wx.ListCtrl(self, style=wx.LC_REPORT)
        lc.AppendColumn("", wx.LIST_FORMAT_CENTER, 0)
        lc.AppendColumn("SHA256", wx.LIST_FORMAT_CENTER, 160)
        lc.AppendColumn("파일명", wx.LIST_FORMAT_CENTER, 200)
        for i, (sha256, filepath) in enumerate(LoadedData.file_hash.items()):
            lc.InsertItem(i, "")
            lc.SetItem(i, 1, sha256)
            lc.SetItem(i, 2, os.path.split(filepath)[-1])
        bt_upload = wx.Button(self, label="업로드")
        bt_remove = wx.Button(self, label="삭제")
        sz_bt = wx.BoxSizer(wx.HORIZONTAL)
        sz_bt.AddMany((
            (bt_upload, 0, wx.ALIGN_CENTER_VERTICAL), ((3, -1), 0),
            (bt_remove, 0, wx.ALIGN_CENTER_VERTICAL)
        ))

        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (lc, 1, wx.EXPAND), ((-1, 5), 0),
            (sz_bt, 0, wx.ALIGN_RIGHT)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 20)

        self.SetSizer(sz)
        self.SetSize((440, 300))
        self.CenterOnParent()

        self.__lc = lc
        self.__bt_upload = bt_upload
        self.__bt_remove = bt_remove
    
    def __bind_events(self):
        self.__bt_upload.Bind(wx.EVT_BUTTON, self.__on_upload)
        self.__bt_remove.Bind(wx.EVT_BUTTON, self.__on_remove)

    def __on_upload(self, event):
        dlg = wx.FileDialog(self, "Raw Data 업로드", wildcard="Raw Data (*.xlsx)|*.xlsx", style=wx.FD_OPEN|wx.FD_MULTIPLE)
        ret = dlg.ShowModal()
        filepaths = dlg.GetPaths()
        dlg.Destroy()
        if ret != wx.ID_OK:
            return
        dlg = wx.ProgressDialog("안내", "데이터를 업로드 중입니다.")
        dlg.Pulse()
        try:
            LoadedData.load_raw_file(filepaths)
        except Exception as err:
            dlg.Destroy()
            wx.Yield()
            wx.MessageBox(get_error_message(err), "안내")
            return
        self.__sync_list_items()
        dlg.Destroy()
        wx.Yield()
        wx.MessageBox("데이터를 업로드했습니다.", "안내")
        self.__updated = True

    def __on_remove(self, event):
        selected_sha256s = []
        lc = self.__lc
        for i in range(lc.GetItemCount()):
            if lc.IsSelected(i):
                selected_sha256s.append(lc.GetItemText(i, 1))
        if not selected_sha256s:
            return
        dlg = wx.MessageDialog(self, "선택한 데이터들을 삭제할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        try:
            for sha256 in selected_sha256s:
                LoadedData.remove_raw_data(sha256)
        except Exception as err:
            wx.MessageBox(get_error_message(err), "안내")
        else:
            wx.MessageBox("데이터를 삭제했습니다.", "안내")
            self.__updated = True
        finally:
            self.__sync_list_items()

    def __sync_list_items(self):
        lc = self.__lc
        lc.DeleteAllItems()
        for sha256, filepath in LoadedData.file_hash.items():
            i = lc.GetItemCount()
            lc.InsertItem(i, "")
            lc.SetItem(i, 1, sha256)
            lc.SetItem(i, 2, os.path.split(filepath)[-1])
    
    def is_updated(self) -> bool:
        return self.__updated

class DialogLoadDB(wx.Dialog):
    def __init__(self, parent: wx.Window):
        super().__init__(parent, title="경고")

        w, h = self.FromDIP(80), self.FromDIP(73)
        size = wx.Size(w, h)
        bundle = wx.BitmapBundle.FromSVG(WARNING_MARK_SVG.encode("utf-8"), size)
        bmp = bundle.GetBitmap(size)
        sb = wx.StaticBitmap(self, bitmap=bmp)

        st = wx.StaticText(
            self, 
            label="다른 DB 파일을 불러오면 현재 프로그램의 기본 DB에 덮어씌워집니다.\n" \
                "기본 DB를 유지하려면 먼저 '다른 이름으로 저장' 하세요.\n\n" \
                "다른 DB 파일을 불러올까요?",
            style=wx.ALIGN_CENTER
            )
        bt_confirm = wx.Button(self, label="확인")
        bt_cancel = wx.Button(self, label="취소")
        bt_cancel.SetFocus()

        sz_vert = wx.BoxSizer(wx.VERTICAL)
        sz_vert.AddMany((
            (sb, 0, wx.ALIGN_CENTER_HORIZONTAL), ((-1, 15), 0),
            (st, 0, wx.ALIGN_CENTER_HORIZONTAL), ((-1, 15), 0),
            (bt_confirm, 0, wx.EXPAND), ((-1, 5), 0),
            (bt_cancel, 0, wx.EXPAND)
        ))
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(sz_vert, 1, wx.EXPAND|wx.ALL, 30)
        self.SetSizerAndFit(sz)
        self.CenterOnParent()

        bt_confirm.Bind(wx.EVT_BUTTON, self.__on_confirm)
        bt_cancel.Bind(wx.EVT_BUTTON, self.__on_cancel)
    
    def __on_confirm(self, event):
        self.EndModal(wx.ID_YES)
    
    def __on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)

class FrameMain(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title=APP_NAME)
        self.__set_icon()
        self.__set_layout()
        self.__set_menubar()
        self.__bind_events()
        self.SetSize((1000, 800))
        self.SetMinSize(self.GetSize())
        self.CenterOnScreen()

    def __set_menubar(self):
        menubar = wx.MenuBar()
        self.SetMenuBar(menubar)

        menu = wx.Menu()
        mi_manage_data = wx.MenuItem(menu, -1, "Raw Data 관리")
        mi_set_openai_key = wx.MenuItem(menu, -1, "OpenAI API 키 설정")
        mi_set_claude_key = wx.MenuItem(menu, -1, "Anthropic API 키 설정")
        mi_load_db = wx.MenuItem(menu, -1, "DB 불러오기")
        mi_save_db = wx.MenuItem(menu, -1, "DB 다른 이름으로 저장")
        mi_load_ctr = wx.MenuItem(menu, -1, "Cost Ctr 불러오기")
        mi_load_element = wx.MenuItem(menu, -1, "Cost Element / Category 불러오기")
        mi_quit = wx.MenuItem(menu, -1, "종료")
        menu.Append(mi_manage_data)
        menu.Append(mi_set_openai_key)
        menu.Append(mi_set_claude_key)
        menu.AppendSeparator()
        # menu.Append(mi_load_db)
        # menu.Append(mi_save_db)
        # menu.AppendSeparator()
        menu.Append(mi_load_ctr)
        menu.Append(mi_load_element)
        menu.AppendSeparator()
        menu.Append(mi_quit)
        menubar.Append(menu, "메뉴")

        menu = wx.Menu()
        mi_license = wx.MenuItem(menu, -1, "OSS 라이센스")
        mi_info = wx.MenuItem(menu, -1, "정보")
        menu.Append(mi_license)
        menu.Append(mi_info)
        menubar.Append(menu, "정보")

        self.__mi_manage_data = mi_manage_data
        self.__mi_set_openai_key = mi_set_openai_key
        self.__mi_set_claude_key = mi_set_claude_key
        self.__mi_load_db = mi_load_db
        self.__mi_save_db = mi_save_db
        self.__mi_load_ctr = mi_load_ctr
        self.__mi_load_element = mi_load_element
        self.__mi_quit = mi_quit
        self.__mi_license = mi_license
        self.__mi_info = mi_info

    def __set_layout(self):
        pn = wx.Panel(self)
        nb = wx.Notebook(pn)
        pn_dashboard = PanelDashboard(nb)
        pn_viewer    = PanelViewer(nb)
        pn_manager   = PanelManager(nb)
        pn_bs_chart  = PanelBSChart(nb)
        nb.AddPage(pn_dashboard, "대시보드")
        nb.AddPage(pn_viewer   , "뷰어")
        nb.AddPage(pn_manager  , "관리")
        nb.AddPage(pn_bs_chart , "BS별 차트")
        sz = wx.BoxSizer(wx.HORIZONTAL)
        sz.Add(nb, 1, wx.EXPAND|wx.ALL, 10)
        pn.SetSizer(sz)

        self.__pn_dashboard = pn_dashboard
        self.__pn_viewer    = pn_viewer   
        self.__pn_manager   = pn_manager  
        self.__pn_bs_chart  = pn_bs_chart 

    def __set_icon(self):
        base = self.FromDIP(32)
        bb = wx.BitmapBundle.FromSVG(NEXEN_LOGO_SVG.encode("utf-8"), wx.Size(base, base))
        sizes = [16, 24, 32, 48, 64, 128]
        ib = wx.IconBundle()
        for dp in sizes:
            px = self.FromDIP(dp)
            bmp = bb.GetBitmap(wx.Size(px, px))
            ico = wx.Icon()
            ico.CopyFromBitmap(bmp)
            ib.AddIcon(ico)
        self.SetIcons(ib)

    def __bind_events(self):
        self.Bind(wx.EVT_MENU, self.__on_manage_data, self.__mi_manage_data)
        self.Bind(wx.EVT_MENU, self.__on_set_openai_key, self.__mi_set_openai_key)
        self.Bind(wx.EVT_MENU, self.__on_set_claude_key, self.__mi_set_claude_key)
        self.Bind(wx.EVT_MENU, self.__on_load_db, self.__mi_load_db)
        self.Bind(wx.EVT_MENU, self.__on_save_db, self.__mi_save_db)
        self.Bind(wx.EVT_MENU, self.__on_load_ctr, self.__mi_load_ctr)
        self.Bind(wx.EVT_MENU, self.__on_load_element, self.__mi_load_element)
        self.Bind(wx.EVT_MENU, self.__on_quit, self.__mi_quit)
        self.Bind(wx.EVT_MENU, self.__on_licence, self.__mi_license)
        self.Bind(wx.EVT_MENU, self.__on_info, self.__mi_info)
        self.Bind(EVT_UPDATE, self.__on_data_updated)

    def __on_manage_data(self, event):
        dlg = DialogManageRawData(self)
        dlg.ShowModal()
        updated = dlg.is_updated()
        dlg.Destroy()
        if updated:
            self.__on_data_updated(None)
            self.__pn_manager.redraw_data_tree()

    def __on_set_openai_key(self, event):
        val = Config.OPENAI_API_KEY
        while True:
            dlg = wx.TextEntryDialog(self, "OpenAI API 키를 입력하세요.", "안내", value=val)
            ret = dlg.ShowModal()
            val = dlg.GetValue().strip()
            dlg.Destroy()
            if ret != wx.ID_OK:
                return
            if val == Config.OPENAI_API_KEY:
                return
            if not val:
                continue
            break

        dlgp = wx.ProgressDialog("안내", "API 키를 확인 중입니다.", parent=self)
        dlgp.Pulse()

        def on_done(msg: str):
            dlgp.Destroy()
            wx.Yield()
            wx.MessageBox(msg, "안내", parent=self)

        def work():
            try:
                models = get_gpt_models(val)
            except Exception as err:
                msg = f"연결이 불가합니다.\nAPI 키를 확인하세요.\n\n{err}"
            else:
                Config.OPENAI_API_KEY = val
                Config.GPT_MODELS = models
                Config.LAST_USED_GPT_MODEL = models[0]
                msg = "API 키를 설정 했습니다."
            finally:
                wx.CallAfter(on_done, msg)

        Thread(target=work, daemon=True).start()

    def __on_set_claude_key(self, event):
        val = Config.CLAUDE_API_KEY
        while True:
            dlg = wx.TextEntryDialog(self, "Anthropic API 키를 입력하세요.", "안내", value=val)
            ret = dlg.ShowModal()
            val = dlg.GetValue().strip()
            dlg.Destroy()
            if ret != wx.ID_OK:
                return
            if val == Config.CLAUDE_API_KEY:
                return
            if not val:
                continue
            break

        dlgp = wx.ProgressDialog("안내", "API 키를 확인 중입니다.", parent=self)
        dlgp.Pulse()

        def on_done(msg: str):
            dlgp.Destroy()
            wx.Yield()
            wx.MessageBox(msg, "안내", parent=self)

        def work():
            try:
                models = get_claude_models(val)
            except Exception as err:
                msg = f"연결이 불가합니다.\nAPI 키를 확인하세요.\n\n{err}"
            else:
                Config.CLAUDE_MODELS = models
                Config.LAST_USED_CLAUDE_MODEL = models[0]
                Config.CLAUDE_API_KEY = val
                msg = "API 키를 설정 했습니다."
            finally:
                wx.CallAfter(on_done, msg)

        Thread(target=work, daemon=True).start()

    def __on_load_db(self, event):
        dlg = DialogLoadDB(self)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        dlg = wx.FileDialog(self, "DB 파일 불러오기", wildcard=f"DB 파일 (*.{EXT})|*.{EXT}", style=wx.FD_OPEN|wx.FD_FILE_MUST_EXIST)
        ret = dlg.ShowModal()
        load_file_path = dlg.GetPath()
        dlg.Destroy()
        if ret != wx.ID_OK:
            return
        try:
            validate_db(load_file_path)
            shutil.copy(load_file_path, DATABASE_PATH)
            Session.configure(bind=get_engine())
        except AssertionError as err:
            msg = f"DB 파일을 불러오던 중 오류가 발생했습니다.\n\n{err}"
        except:
            msg = f"DB 파일을 불러오던 중 오류가 발생했습니다.\n\n{format_exc()}"
        else:
            msg = "DB 파일을 불러왔습니다."
            self.__pn_manager.load_db_values()
            self.__pn_manager.redraw_data_tree()
            self.__on_data_updated(None)
        finally:
            wx.MessageBox(msg, "안내")

    def __on_save_db(self, event):
        dlg = wx.FileDialog(self, "DB 파일 저장", wildcard=f"DB 파일 (*.{EXT})|*.{EXT}", style=wx.FD_SAVE|wx.FD_OVERWRITE_PROMPT)
        ret = dlg.ShowModal()
        save_file_path = dlg.GetPath()
        dlg.Destroy()
        if ret != wx.ID_OK:
            return
        try:
            shutil.copy(DATABASE_PATH, save_file_path)
        except:
            wx.MessageBox(f"파일 저장 중 오류가 발생했습니다.\n\n{format_exc()}", "안내")
        else:
            wx.MessageBox("DB 파일을 저장했습니다.", "안내")

    def __on_load_ctr(self, event):
        dlg = wx.MessageDialog(self, "기존 Cost Ctr 정보를 덮어씌웁니다.\n계속할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        dlg = wx.FileDialog(self, "Cost Ctr 파일 선택", wildcard="엑셀 파일 (*.xlsx)|*.xlsx", style=wx.FD_OPEN)
        ret = dlg.ShowModal()
        filepath = dlg.GetPath()
        dlg.Destroy()
        if ret != wx.ID_OK:
            return
        dlgp = wx.ProgressDialog("안내", "엑셀 파일을 읽는 중입니다.", parent=self)
        dlgp.Pulse()

        def success():
            self.__pn_manager.load_db_values()
            self.__on_data_updated(None)
            self.__pn_manager.redraw_data_tree()
            self.__pn_viewer.set_ctr_filter(CostCtr.get_root_ctr())
            dlgp.Destroy()
            wx.Yield()
            wx.MessageBox("Cost Ctr 정보를 엑셀 파일의 내용으로 덮어씌웠습니다.", "안내", parent=self)

        def fail(msg: str):
            dlgp.Destroy()
            wx.Yield()
            wx.MessageBox(msg, "안내", parent=self)

        def work():
            try:
                ctrs = read_ctr_excel(filepath)
                wx.CallAfter(dlgp.Pulse, "데이터베이스에 저장 중입니다.")
                with Session() as sess:
                    sess.execute(delete(CostCtr))
                    for ctr in ctrs:
                        sess.add(ctr)
                    sess.commit()
                wx.CallAfter(dlgp.Pulse, "예산을 재산정 중입니다.")
                LoadedData.cache_ctr()
                LoadedData.reload()
            except Exception as err:
                if isinstance(err, AssertionError):
                    msg = str(err)
                else:
                    msg = f"엑셀 파일을 읽던 중 예기치 않은 오류가 발생했습니다.\n\n{format_exc()}"
                wx.CallAfter(fail, msg)
            else:
                wx.CallAfter(success)

        Thread(target=work, daemon=True).start()

    def __on_load_element(self, event):
        dlg = wx.MessageDialog(self, "기존 Cost Element 및 Category 정보를 덮어씌웁니다.\n계속할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret != wx.ID_YES:
            return
        dlg = wx.FileDialog(self, "Cost Element & Category 파일 선택", wildcard="엑셀 파일 (*.xlsx)|*.xlsx", style=wx.FD_OPEN)
        ret = dlg.ShowModal()
        filepath = dlg.GetPath()
        dlg.Destroy()
        if ret != wx.ID_OK:
            return
        dlgp = wx.ProgressDialog("안내", "엑셀 파일을 읽는 중입니다.", parent=self)
        dlgp.Pulse()

        def success():
            self.__pn_manager.load_db_values()
            self.__on_data_updated(None)
            self.__pn_manager.redraw_data_tree()
            self.__pn_viewer.set_category_filter(CostCategory.get_root_category(False))
            dlgp.Destroy()
            wx.Yield()
            wx.MessageBox("Cost Element & Category 정보를 엑셀 파일의 내용으로 덮어씌웠습니다.", "안내", parent=self)

        def fail(msg: str):
            dlgp.Destroy()
            wx.Yield()
            wx.MessageBox(msg, "안내", parent=self)

        def work():
            try:
                cats, elems = read_element_excel(filepath)
                wx.CallAfter(dlgp.Pulse, "데이터베이스에 저장 중입니다.")
                with Session() as sess:
                    sess.execute(delete(CostCategory))
                    sess.execute(delete(CostElement))
                    for cat in cats:
                        sess.add(cat)
                    for elem in elems:
                        sess.add(elem)
                    sess.commit()
                wx.CallAfter(dlgp.Pulse, "예산을 재산정 중입니다.")
                LoadedData.cache_category()
                LoadedData.cache_element()
                LoadedData.reload()
            except Exception as err:
                if isinstance(err, AssertionError):
                    msg = str(err)
                else:
                    msg = f"엑셀 파일을 읽던 중 예기치 않은 오류가 발생했습니다.\n\n{format_exc()}"
                wx.CallAfter(fail, msg)
            else:
                wx.CallAfter(success)

        Thread(target=work, daemon=True).start()

    def __on_quit(self, event):
        dlg = wx.MessageDialog(self, "프로그램을 종료할까요?", "안내", style=wx.YES_NO|wx.NO_DEFAULT)
        ret = dlg.ShowModal()
        dlg.Destroy()
        if ret == wx.ID_YES:
            self.Destroy()

    def __on_data_updated(self, event):
        """데이터에 변화가 생겨서 대시보드와 뷰어를 다시 그림
        pn_manager는 필요한 경우에 별도로 호출
        """
        self.__pn_dashboard.redraw_charts()
        self.__pn_viewer.redraw_trees()
        self.__pn_viewer.update_values()

    def __on_licence(self, event):
        dlg = DialogOSSL(self)
        dlg.ShowModal()
        dlg.Destroy()
    
    def __on_info(self, event):
        dlg = DialogInfo(self)
        dlg.ShowModal()
        dlg.Destroy()
