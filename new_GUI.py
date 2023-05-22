import tkinter as tk
from tkinter.ttk import Combobox

from Base_info_adder import (
    create_and_update_sales_tb,
    create_and_update_customers_tb,
    get_date_sql,
    get_data_report_2,
    get_data_report_sum_2
)
from report_maker.report_1 import report_1
from report_maker.report_2 import report_2
from funcs.GUI_funcs import select_file_1, select_file_2

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

def get_choosen_year(event=None):
    choosen_year = year_print.get()
    print(choosen_year)
    return choosen_year


def get_choosen_month(event=None):
    choosen_month = month_print.get()
    print(month_value.index(choosen_month) + 1)
    return month_value.index(choosen_month) + 1

def get_choosen_month_name(event=None):
    choosen_month = month_print.get()
    print(choosen_month)
    return choosen_month


def create_reports_files():
    year = get_choosen_year()
    month = get_choosen_month()
    month_name = get_choosen_month_name()
    data1 = get_date_sql(year=year, month=month)
    data2 = get_data_report_2(year=year, month=month)
    data_sum2 = get_data_report_sum_2(month=month, year=year)
    report_1(data1, year)
    report_2(data2, month_name, data_sum2)
    

def run_sql():
    create_and_update_customers_tb()
    create_and_update_sales_tb()
    create_reports_files()


root = tk.Tk()

root.title("Аналітик")
root.geometry("1200x750+360+80")
root.resizable(False, False)
root.config(bg="#505575")

# Данні для вибору місяця і року


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
month_print.bind("<<ComboboxSelected>>", get_choosen_month)
month_print.bind("<<ComboboxSelected>>", get_choosen_month_name)


root.mainloop()
