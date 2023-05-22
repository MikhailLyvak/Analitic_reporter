from excel_writer_class import Report
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.styles import Border, Side, NamedStyle, Font, Alignment


def report_3(data: list, month: int, sum_data: list) -> None:
    titles = [
            "Назва міста",
            "Менеджер",
            "Сума продаж",
            "Чистий дохід"
        ]

    res2 = Report(data=data, report_name=f"Продажи по містам за {month.lower()} місяць")
    res2.write_data(titles=titles)
    res2.number_columns()
    res2.zagolovok(place=["B2", "F2"])
    res2.design_titles()
    res2.borders()
    res2.print_sum(data=sum_data, col_start=5)
    res2.ws.column_dimensions['C'].width=30
    res2.ws.column_dimensions['D'].width=30
    res2.ws.column_dimensions['E'].width=17
    res2.ws.column_dimensions['F'].width=17
    res2.align_centre()
    res2.name_align_left(4, 3, 4)
    res2.money_format(4, 5)
    res2.design_titles()
    
    
    
    
    res2.wb.save("Report3.xlsx")