from tkinter import filedialog as fd

from excel_reader import excel_reader


reports = {
    "report1": None,
    "report2": None
}


def select_file_1():                        # <-- Функція для пошуку файлу Excel(Шлях до нього)
    filename = fd.askopenfilename(
        title='Open a file'
    )
    reports["report1"] = filename

def select_file_2():                        # <-- Функція для пошуку файлу Excel(Шлях до нього)
    filename = fd.askopenfilename(
        title='Open a file'
    )
    reports["report2"] = filename

# def get_data_report() -> list:
#     return excel_reader(7, 1, 9, reports["report1"])
