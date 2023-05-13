import tkinter as tk
from tkinter.ttk import Combobox

from Base_info_adder import (
    create_and_update_sales_tb,
    create_and_update_customers_tb,
    get_date_sql,
)
from excel_writer_class import Report
from funcs.GUI_funcs import select_file_1, select_file_2


def get_choosen_year(event=None):
    choosen_year = year_print.get()
    print(choosen_year)
    return choosen_year


def report_1():
    year = get_choosen_year()
    data = get_date_sql(year)
    titles = [
        "Менеджер",
        "Січень",
        "Лютий",
        "Березень",
        "Квітень",
        "Травень",
        "Червень",
        "Липень",
        "Серпень",
        "Вересень",
        "Жовтень",
        "Листопад",
        "Грудень",
        "Сума"
    ]
    res = Report(data=data, report_name="Roport for year")
    res.write_data(titles=titles)
    res.number_columns()
    res.borders()
    res.zagolovok(place = ["B2", "P2"])
    res.chart(3, 4, 11, 15, 3, 4, 11, "B13")
    res.align_centre()
    
    Report.wb.save("class.xlsx")



def run_sql():
    create_and_update_customers_tb()
    create_and_update_sales_tb()
    report_1()


root = tk.Tk()

root.title("Аналітик")
root.geometry("1200x750+360+80")
root.resizable(False, False)
root.config(bg="#505575")

# Данні для вибору місяця і року
month_value = (
    "Січень",
    "Лютий",
    "Березень",
    "Квітень",
    "Травень",
    "Червень",
    "Липень",
    "Серпень",
    "Вересень",
    "Жовтень",
    "Листопад",
    "Грудень",
)
year_value = (
    "2010",
    "2011",
    "2012",
    "2013",
    "2014",
    "2015",
    "2016",
    "2017",
    "2018",
    "2019",
    "2020",
    "2021",
    "2022",
    "2023",
    "2024",
    "2025",
)

# Поля для вибору місяця і року (Їх буде використовувати Саша при формуванні Ексель файла)
month_print = Combobox(root, values=month_value, width=12, font=("Times New Roman", 24))
year_print = Combobox(root, width=6, value=year_value, font=("Times New Roman", 24))
month_print.current(0)
year_print.current(12)


btn1 = tk.Button(
    root,
    text="  Виберіть Excel файл №1  ",
    font=60,
    bg="#979DBF",
    activebackground="#9ABADF",
    command=select_file_1,
).place(x=220, y=150)

btn2 = tk.Button(
    root,
    text="  Виберіть Excel файл №2  ",
    font=60,
    bg="#979DBF",
    activebackground="#9ABADF",
    command=select_file_2,
).place(x=220, y=190)

btn5 = tk.Button(
    root,
    text="Підтвердити",
    font=("Times New Roman", 14),
    bg="#98B3D1",
    activebackground="#9ABADF",
    command=run_sql,
).place(x=680, y=235)


month_print.place(x=458, y=22, height=41)
year_print.place(x=888, y=22, height=41)
year_print.bind("<<ComboboxSelected>>", get_choosen_year)

root.mainloop()
