import os
import wx
import kinolist_lib as kl
import config 

api = config.KINOPOISK_API_TOKEN
APP_EXIT = 1

def PIL2wx (image):
    width, height = image.size
    return wx.Bitmap.FromBuffer(width, height, image.tobytes())


class InfoPanel(wx.Dialog): 
    def __init__(self, parent, title): 
        super().__init__(parent, title = title, size = (650, 400))
        
        sel = parent.film_list.GetSelection()
        if sel != -1:
            filminfo = kl.get_film_info(parent.film_id_list[sel], api)
            
        poster = filminfo[9]
        poster.thumbnail((200, 300))
        
        self.panel = wx.Panel(self) 
        self.box1v = wx.BoxSizer(wx.VERTICAL)
        self.box2h = wx.BoxSizer(wx.HORIZONTAL)
        self.btn = wx.Button(self.panel, wx.ID_OK, label = "OK", size = (100, 25))
        self.box1v.Add(self.box2h, proportion = 1, flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5 )
        self.box1v.Add(self.btn, flag = wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5 )
        
        self.image = wx.StaticBitmap(self.panel, wx.ID_ANY, wx.Bitmap(PIL2wx(poster)), size=(200, 300))
        self.box3v = wx.BoxSizer(wx.VERTICAL)
        self.box2h.Add(self.image, flag = wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.box2h.Add(self.box3v, flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        
        self.l_search1 = wx.StaticText(self.panel, label = f"{filminfo[0]} ({filminfo[1]}) - Кинопоиск {filminfo[2]}")
        self.l_search2 = wx.StaticText(self.panel, label = str(filminfo[1]))
        self.l_search3 = wx.StaticText(self.panel, label = filminfo[4])
        self.box3v.Add(self.l_search1, flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.box3v.Add(self.l_search2, flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.box3v.Add(self.l_search3, proportion = 1, flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        
        self.panel.SetSizer(self.box1v)
        self.Centre()

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(500, 500), style = (wx.DEFAULT_FRAME_STYLE | 
                                                                        wx.WANTS_CHARS) & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))
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
        self.t_search.Value = "Матрица"
        
        self.b_search = wx.Button(panel, wx.ID_ANY, size=(100, 25), label='Поиск')
        gr.Add(self.b_search, pos=(0, 2), flag = wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onSearch, id=self.b_search.GetId())
        
        self.search_list = wx.ListBox(panel, wx.ID_ANY, size=(200, 70), style = wx.LB_SINGLE)
        gr.Add(self.search_list, pos=(1, 1), flag = wx.EXPAND | wx.BOTTOM | wx.LEFT, border = 5)  
        self.Bind(wx.EVT_LISTBOX_DCLICK, self.onListEnter, id=self.search_list.GetId())
        self.search_list.Bind(wx.EVT_KEY_DOWN, self.onListEnter)
        
        self.film_list = wx.ListBox(panel, style = wx.LB_SINGLE)
        gr.Add(self.film_list, pos=(2, 1), span=(3, 1), flag = wx.EXPAND | wx.BOTTOM | wx.LEFT, border = 5)        
        
        self.b_add = wx.Button(panel, wx.ID_ANY, size=(100, 25), label='Добавить')
        gr.Add(self.b_add, pos=(1, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onAdd, id=self.b_add.GetId())
        
        self.b_info = wx.Button(panel, wx.ID_ANY, size=(100, 25), label='Информация')
        gr.Add(self.b_info, pos=(2, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onInfo, id=self.b_info.GetId())
        
        self.b_delete = wx.Button(panel, wx.ID_ANY, size=(100, 25), label='Удалить')
        gr.Add(self.b_delete, pos=(3, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onDelete, id=self.b_delete.GetId())
        
        self.b_save = wx.Button(panel, wx.ID_ANY, label='Сохранить в формате docx')
        gr.Add(self.b_save, pos=(5, 1), flag = wx.EXPAND | wx.BOTTOM | wx.LEFT, border = 5)
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
        self.films = []
        self.search_list.Clear()
        if self.t_search.Value:
            self.films = kl.find_kp_id3(self.t_search.Value, api)
            if self.films:
                for film in self.films:
                    self.search_list.Append(f'{film[1]} ({film[2]})')
                self.t_search.Value = ""
                self.search_list.SetFocus()
                self.search_list.SetSelection(0)
    

    def onListEnter(self, event):
        key = event.GetKeyCode()
        if key == wx.WXK_RETURN:
            self.onAdd(self)
        event.Skip()
    
    
    def onAdd(self, event):
        sel = self.search_list.GetSelection()
        if sel != -1:
            self.film_list.Append(f'{self.films[sel][1]} ({self.films[sel][2]})')
            self.film_id_list.append(self.films[sel][0])
            self.search_list.Clear()
            self.t_search.SetFocus()
            
            
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
    
    
    def onInfo(self, event):
        sel = self.film_list.GetSelection()
        if sel != -1:
            self.info = InfoPanel(self, 'Информация о фильме').ShowModal()


def main():
    app = wx.App()
    top = MyFrame(None, title="Kinolist GUI")
    top.Show()
    app.MainLoop()

if __name__ == '__main__':
    main()

