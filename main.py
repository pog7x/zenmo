import os
import platform
import typing
from datetime import datetime
from pathlib import Path

from py_csv_xls import CSVSniffer, ExcelWorker, PyCsvXlsException

if platform.system() == "Windows":
    os.environ["KIVY_GL_BACKEND"] = "angle_sdl2"

from kivy.app import App
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle
from kivy.uix import textinput
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label

Window.size = (1200, 600)


MAIN_ZEN_FIELDS = [
    "date",
    "categoryName",
    "payee",
    "comment",
    "outcomeAccountName",
    "outcome",
    "outcomeCurrencyShortTitle",
    "incomeAccountName",
    "income",
    "incomeCurrencyShortTitle",
    "createdDate",
    "changedDate",
]


class ZenMoneyJob:
    def __init__(self, dir_path: str):
        self.__dir_path: str = dir_path

        self.__file_startswith: str = "zen_"
        self.__csv_sniffer_fields: typing.List[str] = MAIN_ZEN_FIELDS

        self.__workbook_name: str = datetime.isoformat(datetime.now())[:19].replace(
            ":", "-"
        )

        self.__csv_sniffer: CSVSniffer = CSVSniffer(
            file_startswith=self.__file_startswith,
            main_path=self.__dir_path,
            fields=self.__csv_sniffer_fields,
        )
        if self.__csv_sniffer.is_file and not self.__csv_sniffer.is_csv_file:
            self.__excel_worker: ExcelWorker = ExcelWorker(
                workbook_name=self.__dir_path,
                workbook_extension="",
                want_cleared=False,
            )
        else:
            self.__excel_worker: ExcelWorker = ExcelWorker(
                workbook_name=os.path.join(
                    str(Path(self.__dir_path).absolute()), self.__workbook_name
                ),
                workbook_extension=".xlsm",
                sheets_to_create=("Total", "Config"),
                date_cols={
                    "DD/MM/YYYY": ["A"],
                    "DD/MM/YYYY HH:MM:SS": ["K", "L"],
                },
            )

    def __prepare_data_with_dir_path(self) -> typing.List[typing.Dict]:
        return list(self.__csv_sniffer.get_dir_files_with_lines())

    def find_csv_files_and_paste_lines_to_excel(self):
        try:
            prepared_data = self.__prepare_data_with_dir_path()
            for pd in prepared_data:
                for data in pd.values():
                    for row in data:
                        try:
                            row[0] = datetime.strptime(row[0], "%Y-%m-%d")
                            row[10] = datetime.strptime(row[10], "%Y-%m-%d %H:%M:%S")
                            row[11] = datetime.strptime(row[11], "%Y-%m-%d %H:%M:%S")
                        except Exception as e:
                            print(e)
                        try:
                            row_5 = row[5]
                            row_8 = row[8]
                            row[5] = float(row_5.replace(",", ".")) if row_5 else row_5
                            row[8] = float(row_8.replace(",", ".")) if row_8 else row_8
                        except Exception as e:
                            print(e)
            self.__excel_worker.fill_workbook(all_data=prepared_data)
            return self.__excel_worker.full_workbook_name
        except PyCsvXlsException as e:
            raise ZenMoneyJobException(msg=e.msg, exc=e.exc)


class ZenMoneyJobException(Exception):
    def __init__(
        self, exc: typing.Optional[Exception] = None, msg: typing.Optional[str] = None
    ):
        self.exc = exc
        self.msg = msg


class TextInput(textinput.TextInput):
    def __init__(self, **kwargs):
        super(TextInput, self).__init__(**kwargs)
        self.padding_x = (
            [
                self.center[0]
                - self._get_text_width(
                    max(self._lines, key=len), self.tab_width, self._label_cached
                )
                / 2.0,
                0,
            ]
            if self.text
            else [self.center[0], 0]
        )
        self.padding_y = [
            self.height / 2.0 - (self.line_height / 2.0) * len(self._lines),
            0,
        ]


VIOLET = 0.20, 0.06, 0.31, 1
YELLOW = 0.988, 0.725, 0.074, 1


class ZenMoneyLayout(GridLayout):
    def __init__(self, **kwargs):
        super(ZenMoneyLayout, self).__init__(**kwargs)

        with self.canvas.before:
            Color(*YELLOW, mode="rgba")
            self.rect = Rectangle(pos=self.pos, size=self.size)
        self.bind(size=self.update_rect)

        self.cols = 1
        self.height = self.minimum_height

        self.directory_input = GridLayout(
            cols=2,
            size_hint_y=0.3,
        )
        self.directory_input.add_widget(
            Label(
                text="Путь до директории/файла:",
                color=VIOLET,
                bold=True,
                text_size=(None, None),
                font_size="20sp",
            )
        )
        self.directory = TextInput(
            multiline=True,
            hint_text="User/example/path/to/directory/...",
            is_focusable=True,
        )
        self.directory_input.add_widget(self.directory)
        self.add_widget(self.directory_input)

        self.submit = Button(
            text="Выполнить",
            background_normal="",
            background_color=VIOLET,
            size_hint_y=0.3,
            bold=True,
            text_size=(None, None),
            font_size="20sp",
            color=YELLOW,
        )

        self.submit.bind(on_press=self.press)
        self.add_widget(self.submit)

        self.error = Label(
            bold=True,
            text_size=(None, None),
            font_size="20sp",
            padding=[100, 100],
        )
        self.add_widget(self.error)

    def press(self, instance):
        try:
            zmj = ZenMoneyJob(
                dir_path=self.directory.text,
            )
            new_excel = zmj.find_csv_files_and_paste_lines_to_excel()
            self.error.color = "green"
            self.error.text = f"Успешно:\n{new_excel}"
        except ZenMoneyJobException as e:
            self.error.color = "red"
            self.error.text = f"{e.msg}\n{e.exc}"

    def update_rect(self, *args):
        self.rect.pos = self.pos
        self.rect.size = self.size


class ZenMoneyApp(App):
    icon = "zen_ico.png"
    title = "Zen Money 💰"

    def build(self):
        return ZenMoneyLayout()


if __name__ == "__main__":
    ZenMoneyApp().run()
