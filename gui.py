import wx
from typing import Dict, Generator, List

# Data
TITLE = "Проверка договора"
TITLE_SETTINGS = "Настройки"
TXT_PATTERN_INIT_TEXT = "Файл исходного документа"
TXT_EDITED_INIT_TEXT = "Файл редактированного документа"
BTN_OPEN_LABEL = "Загрузить"
BTN_START_LABEL = "Старт"
LABEL_RESULTS = "Предпросмотр: "
SUBMENU_EXPORT_LABEL = "Экспорт в docx"
MENU_FILE_LABEL = "&Файл"
MENU_SETTINGS_LABEL = "&Настройки"


class MainFrame(wx.Frame):

    def __init__(self, controller) -> None:
        super().__init__(parent=None, title=TITLE, id=wx.ID_ANY, style=wx.DEFAULT_FRAME_STYLE)

        self.__controller = controller
        self.__current_fragments = None
        
    # dict of associate buttons with textctrls

        self.__buttons_data: Dict[wx.Button, wx.TextCtrl] = dict()

    # -------------- TEXT STYLES -------------------
        self.__GREY_SMALL_ITALIC = wx.TextAttr(wx.Colour(150, 150, 150, alpha=wx.ALPHA_OPAQUE), font=wx.Font(wx.FontInfo(pointSize=8).Italic()))
        self.__BLACK_BIG_BOLD_UNDERLINED = wx.TextAttr(wx.Colour(0, 0, 0, alpha=wx.ALPHA_OPAQUE), alignment = wx.TEXT_ALIGNMENT_CENTRE, font=wx.Font(wx.FontInfo(pointSize=12).Underlined().Bold()))
        self.__RED_MIDDLE_STRIKED = wx.TextAttr(wx.Colour(255, 0, 0, alpha=wx.ALPHA_OPAQUE), alignment = wx.TEXT_ALIGNMENT_CENTRE, font=wx.Font(wx.FontInfo(pointSize=10).Strikethrough()))
        self.__BLUE_MIDDLE_BOLD = wx.TextAttr(wx.Colour(0, 0, 255, alpha=wx.ALPHA_OPAQUE), alignment = wx.TEXT_ALIGNMENT_CENTRE, font=wx.Font(wx.FontInfo(pointSize=10).Bold()))
        self.__DEFAULT = wx.TextAttr(wx.Colour(0, 0, 0, alpha=wx.ALPHA_OPAQUE), alignment = wx.TEXT_ALIGNMENT_CENTRE, font=wx.Font(wx.FontInfo(pointSize=10)))

        self.__BIG_BOLD = wx.Font(wx.FontInfo(pointSize=12).Bold())


    # -----------------------------------------

        main_sizer = wx.BoxSizer(orient=wx.VERTICAL)
        main_panel = wx.Panel(self)

    # ------ pattern contract file ----------
        pattern_file_dg_sizer = wx.BoxSizer(orient=wx.HORIZONTAL)

        self.__txt_pattern = wx.TextCtrl(main_panel, id=wx.ID_ANY, value=TXT_PATTERN_INIT_TEXT, style = wx.TE_LEFT | wx.TE_READONLY | wx.TE_RICH2, validator = PathTxtCtrlValidator())
        self.__txt_pattern.SetStyle(0, len(self.__txt_pattern.Value), self.__GREY_SMALL_ITALIC)
        btn_open_pattern = wx.Button(main_panel, id=wx.ID_ANY, label=BTN_OPEN_LABEL)

        # associate button with TextCtrl
        self.__buttons_data[btn_open_pattern] = self.__txt_pattern

        pattern_file_dg_sizer.AddMany([(self.__txt_pattern, 1, wx.RIGHT, 5), (btn_open_pattern,)])

        main_sizer.Add(pattern_file_dg_sizer, flag = wx.ALL | wx.EXPAND, border=10)

    # ------ end of pattern contract file ---------
        
    # ------ edited contract file ----------
        edited_file_dg_sizer = wx.BoxSizer(orient=wx.HORIZONTAL)

        self.__txt_edited = wx.TextCtrl(main_panel, id=wx.ID_ANY, value=TXT_EDITED_INIT_TEXT, style = wx.TE_LEFT | wx.TE_READONLY | wx.TE_RICH2, validator = PathTxtCtrlValidator())
        self.__txt_edited.SetStyle(0, len(self.__txt_edited.Value), self.__GREY_SMALL_ITALIC)
        btn_open_edited = wx.Button(main_panel, id=wx.ID_ANY, label=BTN_OPEN_LABEL)

        # associate button with TextCtrl
        self.__buttons_data[btn_open_edited] = self.__txt_edited

        edited_file_dg_sizer.AddMany([(self.__txt_edited, 1, wx.RIGHT, 5), (btn_open_edited,)])

        main_sizer.Add(edited_file_dg_sizer, flag = wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, border=10)

    # ------ end of edited contract file ---------

    # ------ label = "Результаты" ----------------
        
        lbl_results = wx.StaticText(main_panel, label = LABEL_RESULTS)
        main_sizer.Add(lbl_results, flag = wx.TOP | wx.LEFT, border=10)
        lbl_results.SetFont(self.__BIG_BOLD)

    # ------ textctrl = Результаты ----------------
        
        self.__txt_results = wx.TextCtrl(main_panel, id=wx.ID_ANY, style = wx.TE_LEFT | wx.TE_READONLY | wx.TE_RICH2 | wx.TE_MULTILINE)
        main_sizer.Add(self.__txt_results, flag = wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, border=10, proportion=1)

        self.__txt_results.SetDefaultStyle(self.__DEFAULT)

    # ------ button = Старт -----------------------
        
        btn_start = wx.Button(main_panel, id=wx.ID_ANY, label=BTN_START_LABEL)
        main_sizer.Add(btn_start, flag = wx.LEFT | wx.BOTTOM | wx.RIGHT | wx.EXPAND, border=10)

    # ------ menu -----------------------------------
        
        file_menu = wx.Menu()
        file_export = file_menu.Append(id=wx.ID_ANY, item=SUBMENU_EXPORT_LABEL)
        file_exit = file_menu.Append(id=wx.ID_EXIT)

        settings_menu = wx.Menu()

        menuBar = wx.MenuBar()
        menuBar.Append(file_menu, MENU_FILE_LABEL)
        menuBar.Append(settings_menu, MENU_SETTINGS_LABEL)

        self.SetMenuBar(menuBar)

    # -----------------------------------------------

        main_panel.SetSizer(main_sizer)

        self.SetMinSize(wx.Size(550, 500))
        self.SetMaxSize(wx.Size(800, 700))

        self.InitDialog()

    # ----- bindings ------------------------------
        self.Bind(wx.EVT_MENU, self.__export_to_docx, file_export)
        self.Bind(wx.EVT_MENU, self.__close_main_frame, file_exit)
        self.Bind(wx.EVT_BUTTON, self.__set_file_path, btn_open_pattern)
        self.Bind(wx.EVT_BUTTON, self.__set_file_path, btn_open_edited)
        self.Bind(wx.EVT_BUTTON, self.__on_start, btn_start)
        settings_menu.Bind(wx.EVT_MENU_OPEN, self.__show_settings_menu)

    # ----- number generator -------------------------
        
        self.__num = self.__number_generator()

        self.Show()

    # ------ end constructor -------------------------

    def __on_start(self, evt) -> None:

        if self.Validate() and self.TransferDataFromWindow():
            self.__controller.get_marked_text(self.__txt_pattern.GetValue(), self.__txt_edited.GetValue())
        
    def set_result_text(self, fragments: List[tuple]) -> None:

        self.__current_fragments = fragments

        self.__txt_results.Clear()

        rus_tags = {"delete": "УДАЛЕНО", "replace": "ЗАМЕНЕНО", "insert": "ВСТАВЛЕНО"}

        for fragment in fragments:
            if fragment.tag != "equal":
            # set tag
                self.__set_text(f"{rus_tags.get(fragment.tag)}\n\n", self.__BLACK_BIG_BOLD_UNDERLINED)
            # set txt_before    
                self.__set_text(f"{fragment.txt_before}")                      
            # set old_text
                self.__set_text(f"{fragment.old_text}", self.__RED_MIDDLE_STRIKED)             
            # set new_text
                self.__set_text(fragment.new_text, self.__BLUE_MIDDLE_BOLD)
            # set text_after
                self.__set_text(f"{fragment.txt_after}\n\n")
                         
    def __set_text(self, text: str, style: wx.TextAttr = None) -> None:

        current_pos = self.__txt_results.GetLastPosition()
        self.__txt_results.AppendText(text)

        if style is not None:
            self.__txt_results.SetStyle(current_pos, self.__txt_results.GetLastPosition(), style)

    def __show_settings_menu(self, evt):
        # open a frame with settings dialog
        SettingsFrame(self)

    def __export_to_docx(self, evt) -> None:

        with wx.FileDialog(self, style=wx.FD_SAVE, message="Сохранение файла docx", wildcard="DOCX files (*.docx)|*.docx", defaultFile=f"Отчет{next(self.__num)}") as fd:

            if fd.ShowModal() == wx.ID_CANCEL:
                return     # the user changed their mind
        if self.__current_fragments is not None:    
            self.__controller.make_report(self.__current_fragments, fd.GetPath())
        else:
            wx.MessageBox(message="Сравнение документов не было выполнено! Нажмите 'Старт'. Отчет не сохранен", caption="Ошибка", style=wx.ICON_WARNING | wx.OK_DEFAULT)

    def __close_main_frame(self, evt):
        self.Destroy()

    def __set_file_path(self, evt: wx.CommandEvent):

        with wx.FileDialog(self, style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST, message="Открыть файл docx", wildcard="DOCX files (*.docx)|*.docx") as fd:

            if fd.ShowModal() == wx.ID_CANCEL:
                return     # the user changed his mind

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

class SettingsFrame(wx.Frame):

    def __init__(self, parent):
        super().__init__(parent=None, title=TITLE_SETTINGS, id=wx.ID_ANY, style=wx.DEFAULT_FRAME_STYLE)
        self.Show()

class PathTxtCtrlValidator(wx.Validator):

    def Clone(self): # type: ignore
        return PathTxtCtrlValidator()
    
    def Validate(self, parent) -> bool:
         
         textCtrl: wx.TextCtrl = self.GetWindow()
         text: str = textCtrl.GetValue()
         
         if text == TXT_PATTERN_INIT_TEXT or text == TXT_EDITED_INIT_TEXT:
             wx.MessageBox("Нужно выбрать файл docx!", "Ошибка", style=wx.ICON_WARNING | wx.OK_DEFAULT)
             textCtrl.SetBackgroundColour("pink")
             textCtrl.SetFocus()
             textCtrl.Refresh()
             return False
         else:
             textCtrl.SetBackgroundColour(wx.Colour(255, 255, 255))
             textCtrl.Refresh()
             return True
    
    def TransferToWindow(self) -> bool:
        return True
    
    def TransferFromWindow(self) -> bool:
        return True


