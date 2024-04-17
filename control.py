from typing import List
import gui, wx
import process

class Controller:

    def __init__(self) -> None:
        self.__app = wx.App()
        self.__main_frame = gui.MainFrame(controller=self)

    def run_app(self) -> None:
        self.__app.MainLoop()

    def get_marked_text(self, path_pattern: str, path_edited: str):
        fragments: List[tuple] = process.get_marked_fragments(path_pattern, path_edited, 50, 50, "9. Адреса, реквизиты и подписи сторон")
        self.__main_frame.set_result_text(fragments)

    def make_report(self, fragments: List[tuple], path: str) -> None:
        process.make_report_docx(fragments, path)

if __name__ == "__main__":
    Controller().run_app()