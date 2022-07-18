import os
import wx
import kinolist_lib as kl
import config 

api = config.KINOPOISK_API_TOKEN
APP_EXIT = 1

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(500, 500))
        self.film_id_list = []
        
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
        # self.Bind(wx.EVT_BUTTON, self.onSearch, id=b_change.GetId())
        
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
            film = kl.find_kp_id2(self.t_search.Value, api)
            if film:
                self.film_list.Append(f'{film[1]} ({film[2]})')
                self.film_id_list.append(film[0])
                self.t_search.Value = ""    
            
            
    def onDelete(self, event):
        sel = self.film_list.GetSelection()
        if sel != -1:
            self.film_list.Delete(sel)
            del(self.film_id_list[sel])
   
            
    def onSave(self, event):
        if self.film_id_list:
            kl.make_docx(self.film_id_list, 'list.docx', 'template.docx', api)
            if os.path.exists('C:\Program Files\Microsoft Office\Office14\WINWORD.EXE'):
                os.system('start "C:\Program Files\Microsoft Office\Office14\WINWORD.EXE" list.docx')
        

app = wx.App()
top = MyFrame(None, title="Kinolist GUI")
top.Show()
app.MainLoop()