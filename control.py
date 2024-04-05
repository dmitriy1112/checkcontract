from typing import List
import gui, wx
import process

class Controller:

    def __init__(self) -> None:
        self.__app = wx.App()
        self.__main_frame = gui.MainFrame(controller=self)


    def run_app(self) -> None:
        self.__main_frame.Show()
        self.__app.MainLoop()

    def get_marked_text(self, evt):
        fragments: List[tuple] = process.get_marked_fragments("samples\\contract.docx", "samples\\edited.docx", 50, 50, "9. Адреса, реквизиты и подписи сторон")
        self.__main_frame.set_result_text(fragments)



if __name__ == "__main__":
    Controller().run_app()