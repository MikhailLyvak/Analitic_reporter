from excel_writer_class import Report
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.styles import Border, Side, NamedStyle, Font, Alignment


def report_1(data: list, year: int) -> None:
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
    res = Report(data=data, report_name=f"Звіт продаж менеджерів по місяцям за {year} рік")
    res.write_data(titles=titles)
    res.number_columns()
    res.borders()
    res.zagolovok(place = ["B2", "P2"])
    res.chart(3, 4, 11, 15, 3, 4, 11, "D13")
    res.align_centre()
    res.design_titles()
    res.ws.column_dimensions['C'].width=25
    res.money_format(4, 4, 11, 16)
    res.format_cols_width(11, "D", "O")
    res.ws.column_dimensions['P'].width=13
    res.ws.row_dimensions[2].height = 40
    res.name_align_left(4, 3)
    
    res.wb.save("Report0.xlsx")