import wx
from typing import Dict, Generator


class MainFrame(wx.Frame):

    def __init__(self, *args, **kw):
        super().__init__(*args, **kw)


    # dict of associate buttons with textctrls

        self.__buttons_data: Dict[wx.Button, wx.TextCtrl] = dict()

    # -------------- STYLES -------------------
        self.__GREY_SMALL_ITALIC = wx.TextAttr(wx.Colour(161, 161, 161, alpha=wx.ALPHA_OPAQUE), font=wx.Font(wx.FontInfo(pointSize=8).Italic()))
        self.__BLACK_BIG_BOLD = wx.Font(wx.FontInfo(pointSize=12).Bold())

    # -----------------------------------------

        main_sizer = wx.BoxSizer(orient=wx.VERTICAL)
        main_panel = wx.Panel(self)

    # ------ pattern contract file ----------
        pattern_file_dg_sizer = wx.BoxSizer(orient=wx.HORIZONTAL)

        txt_pattern = wx.TextCtrl(main_panel, id=wx.ID_ANY, value="Файл исходного документа", style = wx.TE_LEFT | wx.TE_READONLY | wx.TE_RICH2)
        txt_pattern.SetStyle(0, len(txt_pattern.Value), self.__GREY_SMALL_ITALIC)
        btn_open_pattern = wx.Button(main_panel, id=wx.ID_ANY, label="Загрузить")

        # associate button with TextCtrl
        self.__buttons_data[btn_open_pattern] = txt_pattern

        pattern_file_dg_sizer.AddMany([(txt_pattern, 1, wx.RIGHT, 5), (btn_open_pattern,)])

        main_sizer.Add(pattern_file_dg_sizer, flag = wx.ALL | wx.EXPAND, border=10)

    # ------ end of pattern contract file ---------
        
    # ------ edited contract file ----------
        edited_file_dg_sizer = wx.BoxSizer(orient=wx.HORIZONTAL)

        txt_edited = wx.TextCtrl(main_panel, id=wx.ID_ANY, value="Файл редактированного документа", style = wx.TE_LEFT | wx.TE_READONLY | wx.TE_RICH2)
        txt_edited.SetStyle(0, len(txt_edited.Value), self.__GREY_SMALL_ITALIC)
        btn_open_edited = wx.Button(main_panel, id=wx.ID_ANY, label="Загрузить")

        # associate button with TextCtrl
        self.__buttons_data[btn_open_edited] = txt_edited

        edited_file_dg_sizer.AddMany([(txt_edited, 1, wx.RIGHT, 5), (btn_open_edited,)])

        main_sizer.Add(edited_file_dg_sizer, flag = wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, border=10)

    # ------ end of edited contract file ---------

    # ------ label = "Результаты" ----------------
        
        lbl_results = wx.StaticText(main_panel, label = "Результаты: ")
        main_sizer.Add(lbl_results, flag = wx.TOP | wx.LEFT, border=10)
        lbl_results.SetFont(self.__BLACK_BIG_BOLD)

    # ------ textctrl = Результаты ----------------
        
        txt_results = wx.TextCtrl(main_panel, id=wx.ID_ANY, style = wx.TE_LEFT | wx.TE_READONLY | wx.TE_RICH2 | wx.TE_MULTILINE)
        main_sizer.Add(txt_results, flag = wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, border=10, proportion=1)

    # ------ button = Старт -----------------------
        
        btn_start = wx.Button(main_panel, id=wx.ID_ANY, label="Старт")
        main_sizer.Add(btn_start, flag = wx.LEFT | wx.BOTTOM | wx.RIGHT | wx.EXPAND, border=10)

    # ------ menu -----------------------------------
        
        file_menu = wx.Menu()
        file_export = file_menu.Append(id=wx.ID_ANY, item="Экспорт в docx")
        file_exit = file_menu.Append(id=wx.ID_EXIT)

        settings_menu = wx.Menu()

        menuBar = wx.MenuBar()
        menuBar.Append(file_menu, "&Файл")
        menuBar.Append(settings_menu, "&Настройки")

        self.SetMenuBar(menuBar)

    # -----------------------------------------------

        main_panel.SetSizer(main_sizer)

        self.SetMinSize(wx.Size(550, 500))
        self.SetMaxSize(wx.Size(800, 700))

    # ----- bindings ------------------------------
        self.Bind(wx.EVT_MENU, self.__export_to_docx, file_export)
        self.Bind(wx.EVT_MENU, self.__close_main_frame, file_exit)
        self.Bind(wx.EVT_BUTTON, self.__set_file_path, btn_open_pattern)
        self.Bind(wx.EVT_BUTTON, self.__set_file_path, btn_open_edited)
        settings_menu.Bind(wx.EVT_MENU_OPEN, self.__show_settings_menu)

    # ----- number generator -------------------------
        
        self.__num = self.__number_generator()


    def __show_settings_menu(self, evt):
        # open a frame with settings dialog
        print("open settings frame")

    def __export_to_docx(self, evt):

        with wx.FileDialog(self, style=wx.FD_SAVE, message="Сохранение файла docx", wildcard="DOCX files (*.docx)|*.docx", defaultFile=f"Отчет{next(self.__num)}") as fd:

            if fd.ShowModal() == wx.ID_CANCEL:
                return     # the user changed their mind
            
            print("Path = %s" % fd.GetPath())

    def __close_main_frame(self, evt):
        self.Destroy()

    def __set_file_path(self, evt: wx.CommandEvent):

        with wx.FileDialog(self, style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST, message="Открыть файл docx", wildcard="DOCX files (*.docx)|*.docx") as fd:

            if fd.ShowModal() == wx.ID_CANCEL:
                return     # the user changed their mind

            # Proceed loading the file chosen by the user
            txt: wx.TextCtrl = self.__buttons_data.get(evt.GetEventObject())
            txt.SetValue(fd.GetPath())
            txt.SetStyle(0, len(txt.Value), self.__GREY_SMALL_ITALIC)

    @staticmethod
    def __number_generator() -> Generator:
        num = 1
        while True:
            yield num
            num += 1

     
       



if __name__ == "__main__":
    app = wx.App()
    MainFrame(None, title="Проверка договора").Show()
    app.MainLoop()