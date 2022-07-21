import os
import winreg
import wx
import wx.adv
import kinolist_lib as kl
import config 

api = config.KINOPOISK_API_TOKEN

VER = "0.2.1"
# APP_EXIT = 1
REG_PATH = R"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"

def PIL2wx (image):
    width, height = image.size
    return wx.Bitmap.FromBuffer(width, height, image.tobytes())


def get_reg(name, reg_path):
    try:
        registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0,
                                       winreg.KEY_READ)
        value, regtype = winreg.QueryValueEx(registry_key, name)
        winreg.CloseKey(registry_key)
        return value
    except WindowsError:
        return None


class InfoPanel(wx.Dialog): 
    def __init__(self, parent, title, id): 
        super().__init__(parent, title = title, size = (650, 400))
        
        filminfo = kl.get_film_info(id, api)
            
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
        self.l_search1 = wx.StaticText(self.panel)
        if filminfo[2]:
            self.l_search1.Label = f"{filminfo[0]} ({filminfo[1]}) - Кинопоиск {filminfo[2]}"
        else:
            self.l_search1.Label = f"{filminfo[0]} ({filminfo[1]}) - нет рейтинга"
        if len(filminfo[7]) > 1:
            cast = f"{filminfo[1]}\n{', '.join(filminfo[3])}\nРежиссеры: {', '.join(filminfo[7])}\nВглавных ролях: {', '.join(filminfo[8])}"
        else:
            cast = f"{filminfo[1]}\n{', '.join(filminfo[3])}\nРежиссер: {filminfo[7][0]}\nВглавных ролях: {', '.join(filminfo[8])}"
        self.l_search2 = wx.StaticText(self.panel, label = cast)
        self.l_search3 = wx.TextCtrl(self.panel, value = filminfo[4], style = wx.TE_READONLY | wx.ALIGN_TOP | wx.TE_MULTILINE)
        self.box3v.Add(self.l_search1, flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.box3v.Add(self.l_search2, proportion = 1, flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.box3v.Add(self.l_search3, proportion = 1, flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        
        self.panel.SetSizer(self.box1v)
        self.Centre()

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(500, 500), style = (wx.DEFAULT_FRAME_STYLE | 
                                                                        wx.WANTS_CHARS) & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))
        self.film_id_list = []
        self.all_searched_films = []
                
        # ========== Меню ==========
        menubar = wx.MenuBar()
        
        fileMenu = wx.Menu()
        item_open = wx.MenuItem(fileMenu, wx.ID_OPEN, "Открыть файл\tCtrl+O")
        item_save = wx.MenuItem(fileMenu, wx.ID_SAVE, "Сохранить файл\tCtrl+S")
        item_exit = wx.MenuItem(fileMenu, wx.ID_EXIT, "Выход\tCtrl+Q")
        fileMenu.Append(item_open)
        fileMenu.Append(item_save)
        fileMenu.AppendSeparator()
        fileMenu.Append(item_exit)
        menubar.Append(fileMenu, "Файл")
        
        infoMenu = wx.Menu()
        item_about = wx.MenuItem(fileMenu, wx.ID_ANY, "О программе")
        infoMenu.Append(item_about)
        menubar.Append(infoMenu, "Справка")
        
        self.SetMenuBar(menubar)
        
        self.Bind(wx.EVT_MENU, self.onQuit, id = wx.ID_EXIT)
        self.Bind(wx.EVT_MENU, self.onOpenFile, id = wx.ID_OPEN)
        self.Bind(wx.EVT_MENU, self.onSaveFile, id = wx.ID_SAVE)
        self.Bind(wx.EVT_MENU, self.onAboutBox, id = item_about.GetId())
        

        # ========== Основные элементы ==========
        self.panel = wx.Panel(self)
        self.gr = wx.GridBagSizer(7, 3)
        
        self.l_search = wx.StaticText(self.panel, label='Фильм')
        self.gr.Add(self.l_search, pos=(0, 0), flag = wx.TOP | wx.BOTTOM | wx.LEFT, border = 10)
        
        self.t_search = wx.TextCtrl(self.panel, style=wx.TE_PROCESS_ENTER)
        self.gr.Add(self.t_search, pos=(0, 1), flag = wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT, border = 5)
        self.Bind(wx.EVT_TEXT_ENTER, self.onEnter)
        # self.t_search.Value = "Матрица"
        
        self.b_search = wx.Button(self.panel, wx.ID_ANY, size=(100, 25), label='Поиск')
        self.gr.Add(self.b_search, pos=(0, 2), flag = wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onSearch, id=self.b_search.GetId())
        
        self.search_list = wx.ListBox(self.panel, wx.ID_ANY, size=(200, 70), style = wx.LB_SINGLE)
        self.gr.Add(self.search_list, pos=(1, 1), flag = wx.EXPAND | wx.BOTTOM | wx.LEFT, border = 5)  
        self.Bind(wx.EVT_LISTBOX_DCLICK, self.onAdd, id=self.search_list.GetId())
        self.search_list.Bind(wx.EVT_KEY_DOWN, self.onListEnter)
        self.Bind(wx.EVT_LISTBOX, self.ListClick1, id=self.search_list.GetId())
        
        self.film_list = wx.ListBox(self.panel, wx.ID_ANY, style = wx.LB_SINGLE)
        self.gr.Add(self.film_list, pos=(2, 1), span=(4, 1), flag = wx.EXPAND | wx.BOTTOM | wx.LEFT, border = 5)        
        self.Bind(wx.EVT_LISTBOX, self.ListClick2, id=self.film_list.GetId())
        
        self.b_add = wx.Button(self.panel, wx.ID_ANY, size=(100, 25), label='Добавить')
        self.gr.Add(self.b_add, pos=(1, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onAdd, id=self.b_add.GetId())
        
        self.b_info = wx.Button(self.panel, wx.ID_ANY, size=(100, 25), label='Информация')
        self.gr.Add(self.b_info, pos=(2, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onInfo, id=self.b_info.GetId())
        
        self.b_delete = wx.Button(self.panel, wx.ID_ANY, size=(100, 25), label='Удалить')
        self.gr.Add(self.b_delete, pos=(3, 2), flag = wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onDelete, id=self.b_delete.GetId())
        
        self.b_up = wx.Button(self.panel, wx.ID_ANY, size=(100, 25), label='\u21e7')
        self.gr.Add(self.b_up, pos=(4, 2), flag = wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onUp, id=self.b_up.GetId())
        
        self.b_down = wx.Button(self.panel, wx.ID_ANY, size=(100, 25), label='\u21e9')
        self.gr.Add(self.b_down, pos=(5, 2), flag = wx.ALIGN_CENTER_HORIZONTAL | wx.BOTTOM | wx.LEFT | wx.RIGHT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onDown, id=self.b_down.GetId())
        
        self.b_save = wx.Button(self.panel, wx.ID_ANY, label='Сохранить в формате docx')
        self.gr.Add(self.b_save, pos=(6, 1), flag = wx.EXPAND | wx.BOTTOM | wx.LEFT, border = 5)
        self.Bind(wx.EVT_BUTTON, self.onSave, id=self.b_save.GetId())
                
        self.gr.AddGrowableCol(1)
        self.gr.AddGrowableRow(5)
        self.panel.SetSizer(self.gr)
        self.Centre()
        
        
        # ========= Статусбар ===========
        self.statusbar = self.CreateStatusBar(2, style = (wx.BORDER_NONE) & ~(wx.STB_SHOW_TIPS))
        self.statusbar.SetStatusWidths([80, -1])
        self.statusbar.SetStatusText("Фильмов: " + str(len(self.film_id_list)))
        
        # Перехват события наведения на пункт меню, чтобы отключить автоматический показ подсказок в статусбаре
        # wx.EVT_MENU_HIGHLIGHT_ALL(self, self.statusbar_status)
        wx.EvtHandler.Bind(self, wx.EVT_MENU_HIGHLIGHT_ALL, self.statusbar_status)
        
    def statusbar_status(self, event):
        pass
    
        
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
                self.film_list.SetSelection(-1)
    

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
            self.statusbar.SetStatusText("Фильмов: " + str(len(self.film_id_list)))

            
    def onDelete(self, event):
        sel = self.film_list.GetSelection()
        if sel != -1:
            self.film_list.Delete(sel)
            del(self.film_id_list[sel])
            if sel >= 1:
                self.film_list.SetSelection(sel - 1)
            if sel == 0 and self.film_list.Count > 0:
                self.film_list.SetSelection(0)
            self.statusbar.SetStatusText("Фильмов: " + str(len(self.film_id_list)))
   
            
    def onSave(self, event):
        if self.film_id_list:
            kl.make_docx(self.film_id_list, 'list.docx', 'template.docx', api)
            reg_path = R"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"
            word_path = get_reg("path", reg_path) + "winword.exe"
            if os.path.exists(word_path):
                os.system(f'start "{word_path}" list.docx')
 
    
    def onInfo(self, event):
        sel_film_list = self.film_list.GetSelection()
        sel_search_list = self.search_list.GetSelection()
        if sel_film_list != -1:
            self.info = InfoPanel(self, 'Информация о фильме', self.film_id_list[sel_film_list]).ShowModal()
        elif sel_search_list != -1:
            self.info = InfoPanel(self, 'Информация о фильме', self.films[sel_search_list][0]).ShowModal()
    
    def onUp(self, event):
        sel = self.film_list.GetSelection()
        if sel != -1 and sel != 0:
            items = self.film_list.GetItems()
            items[sel], items[sel - 1] = items[sel - 1], items[sel]
            self.film_list.SetItems(items)
            self.film_id_list[sel], self.film_id_list[sel - 1] = self.film_id_list[sel - 1], self.film_id_list[sel]
            self.film_list.SetSelection(sel - 1)
    
    def onDown(self, event):
        sel = self.film_list.GetSelection()
        if sel != -1 and sel != (self.film_list.Count - 1):
            items = self.film_list.GetItems()
            items[sel], items[sel + 1] = items[sel + 1], items[sel]
            self.film_list.SetItems(items)
            self.film_id_list[sel], self.film_id_list[sel + 1] = self.film_id_list[sel + 1], self.film_id_list[sel]
            self.film_list.SetSelection(sel + 1)
    
    def ListClick1(self, event):
        if self.search_list.Selection != -1:
            self.film_list.SetSelection(-1)
        
    def ListClick2(self, event):
        if self.film_list.Selection != -1:
            self.search_list.SetSelection(-1)
       
    
    def onAboutBox(self, event):
        description = """Программа для создания списков фильмов."""
        licence = """MIT"""
        info = wx.adv.AboutDialogInfo()
        info.SetName('Kinolist GUI')
        info.SetVersion(VER)
        info.SetDescription(description)
        info.SetCopyright('(C) 2022 Alexander Vanyunin')
        info.SetLicence(licence)
        # info.SetIcon(wx.Icon('hunter.png', wx.BITMAP_TYPE_PNG))
        # info.SetWebSite('')
        # info.AddDeveloper('Alexander Vanyunin')
        # info.AddDocWriter('')
        # info.AddArtist('')
        # info.AddTranslator('')

        wx.adv.AboutBox(info)
    
    
    def onOpenFile(self, event):
        self.film_list.Clear()
        self.film_id_list = []
        films_not_found = []
        with wx.FileDialog(self, "Открыть файл...", "", "", "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*", style=wx.FD_OPEN |
                           wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            path_name = fileDialog.GetPath()
            films_from_file = kl.file_to_list(path_name)
            self.gauge = wx.Gauge(self.statusbar, range = len(films_from_file), pos = (90, 2), size = (self.statusbar.GetSize()[0] - 90, 20), style=wx.GA_HORIZONTAL)
            self.count = 0
            for film in films_from_file:
                foundfilm = kl.find_kp_id4(film, api)
                if foundfilm:
                    self.film_list.Append(f"{foundfilm[1]} ({foundfilm[2]})")
                    self.panel.Refresh()
                    self.panel.Update()
                    self.statusbar.SetStatusText("Фильмов: " + str(len(self.film_id_list)))
                    self.film_id_list.append(foundfilm[0])
                    self.count += 1
                    self.gauge.SetValue(self.count)
                else:
                    films_not_found.append(film)
                    self.count += 1
                    self.gauge.SetValue(self.count)
            if films_not_found:
                text_not_found = '\n'.join(films_not_found)
                text = f"Следующие фильмы не найдены:\n{text_not_found}"
                wx.MessageBox(text, 'Внимание!')
            self.statusbar.SetStatusText("Фильмов: " + str(len(self.film_id_list)))
            self.gauge.Destroy()
                    
                    
    def onSaveFile(self, event):
        items_list = self.film_list.GetItems()
        if items_list:
            with wx.FileDialog(self, "Сохранить файл...", "", "", "Текстовые файлы (*.txt)|*.txt", style=wx.FD_SAVE) as fileDialog:
                if fileDialog.ShowModal() == wx.ID_CANCEL:
                    return
                path_name = fileDialog.GetPath()
                with open(path_name, "w", encoding="utf-8") as f:
                    for item in items_list:
                        f.write(item + "\n")
        
def main():
    app = wx.App()
    top = MyFrame(None, title = "Kinolist GUI")
    top.Show()
    app.MainLoop()

if __name__ == '__main__':
    main()

