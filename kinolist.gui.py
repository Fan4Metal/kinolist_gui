import os
import wx
import kinolist_lib as kl
import config 

api = config.KINOPOISK_API_TOKEN
APP_EXIT = 1

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(500, 500), style = wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))
        self.film_id_list = []
        self.all_searched_films = []
        
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        item = wx.MenuItem(fileMenu, APP_EXIT, "Выход\tCtrl+Q", "Выход из приложения")
        fileMenu.Append(item)
        menubar.Append(fileMenu, "&File")
        self.SetMenuBar(menubar)
        self.Bind(wx.EVT_MENU, self.onQuit, id=APP_EXIT)

        panel = wx.Panel(self)
        gr = wx.GridBagSizer(7, 3)
        
        self.l_search = wx.StaticText(panel, label='Фильм')
        gr.Add(self.l_search, pos=(0, 0), flag = wx.TOP | wx.BOTTOM | wx.LEFT, border = 10)
        
        self.t_search = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER)
        gr.Add(self.t_search, pos=(0, 1), flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT, border = 5)
        self.Bind(wx.EVT_TEXT_ENTER, self.onEnter)
        
        self.b_search = wx.Button(panel, wx.ID_ANY, size=(100, 25), label='Добавить')
        gr.Add(self.b_search, pos=(0, 2), flag = wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onSearch, id=self.b_search.GetId())
        
        self.film_list = wx.ListBox(panel, style = wx.LB_SINGLE)
        gr.Add(self.film_list, pos=(1, 1), span=(3, 1), flag = wx.EXPAND | wx.BOTTOM | wx.LEFT, border = 5)        
        
        self.b_info = wx.Button(panel, wx.ID_ANY, size=(100, 25), label='Информация')
        gr.Add(self.b_info, pos=(1, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        # self.Bind(wx.EVT_BUTTON, self.onSearch, id=b_info.GetId())
        
        self.b_change = wx.Button(panel, wx.ID_ANY, size=(100, 25), label='Изменить')
        gr.Add(self.b_change, pos=(2, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onChange, id=self.b_change.GetId())
        
        self.b_delete = wx.Button(panel, wx.ID_ANY, size=(100, 25), label='Удалить')
        gr.Add(self.b_delete, pos=(3, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onDelete, id=self.b_delete.GetId())
        
        self.b_save = wx.Button(panel, wx.ID_ANY, label='Сохранить в формате docx')
        gr.Add(self.b_save, pos=(4, 1), flag = wx.EXPAND | wx.BOTTOM | wx.LEFT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onSave, id=self.b_save.GetId())
                
        gr.AddGrowableCol(1)
        gr.AddGrowableRow(3)
        panel.SetSizer(gr)
        self.Centre()
        
        
            
    def onEnter(self, event):
        self.onSearch(self)
        
        
    def onQuit(self, event):
        self.Close()

    def onSearch(self, event):
        if self.t_search.Value:
            film = kl.find_kp_id3(self.t_search.Value, api)
            if film:
                self.film_list.Append(f'{film[0][1]} ({film[0][2]})')
                self.film_id_list.append(film[0][0])
                self.all_searched_films.append(film)
                self.t_search.Value = ""
            
            
    def onDelete(self, event):
        sel = self.film_list.GetSelection()
        if sel != -1:
            self.film_list.Delete(sel)
            del(self.film_id_list[sel])
            del(self.all_searched_films[sel])
   
            
    def onSave(self, event):
        if self.film_id_list:
            kl.make_docx(self.film_id_list, 'list.docx', 'template.docx', api)
            if os.path.exists('C:\Program Files\Microsoft Office\Office14\WINWORD.EXE'):
                os.system('start "C:\Program Files\Microsoft Office\Office14\WINWORD.EXE" list.docx')
    
    
    def onChange(self, event):
        # search = MyFrameChange(self)
        self.search_window = MyFrameChange(self)
        sel = self.film_list.GetSelection()
        if sel != -1:
            for item in self.all_searched_films[sel]:
                self.search_window.film_list.Append(f"{item[1]} ({item[2]})")
        self.search_window.Show()


class MyFrameChange(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, size=(320, 350), style = (wx.DEFAULT_FRAME_STYLE| wx.FRAME_FLOAT_ON_PARENT)
                         & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.CLOSE_BOX))
        self.Centre()
        self.panel = wx.Panel(self)
        self.box1 = wx.BoxSizer(wx.VERTICAL)
        
        
        self.film_list = wx.ListBox(self.panel, size=(310, 250), style = wx.LB_SINGLE)
        self.box1.Add(self.film_list, flag = wx.EXPAND | wx.BOTTOM | wx.LEFT | wx.RIGHT | wx.TOP, border = 10)

        self.box2 = wx.BoxSizer(wx.HORIZONTAL)
       
        self.b_ok = wx.Button(self.panel, wx.ID_ANY, size=(100, 25), label='ОК')
        self.box2.Add(self.b_ok, flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 10)
        self.Bind(wx.EVT_BUTTON, self.onOk, id=self.b_ok.GetId())
        
        self.b_cancel = wx.Button(self.panel, wx.ID_ANY, size=(100, 25), label='Отмена')
        self.box2.Add(self.b_cancel, flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 10)
        self.Bind(wx.EVT_BUTTON, self.onCancel, id=self.b_cancel.GetId())
                
        self.box1.Add(self.box2, flag = wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT, border = 10)
        
        self.panel.SetSizer(self.box1)
        
    def onOk(self, event):
        pass
    
    
    def onCancel(self,event):
        pass


app = wx.App()
top = MyFrame(None, title="Kinolist GUI")
top.Show()
# search.Show()
app.MainLoop()
