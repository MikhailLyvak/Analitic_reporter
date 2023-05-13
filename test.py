from excel_writer_class import Report
from Base_info_adder import get_date_sql



def test():
    values = [
        ['John', 25, 100],
        ['Alice', 30, 50],
        ['Bob', 20, 79],
    ]

    titles = [
        "name",
        "november",
        "december"
    ]

    res = Report(data=values, report_name="Test report zagolovok)")

    res.write_data(titles=titles)
    res.number_columns()
    res.borders()
    res.zagolovok(place = ["B2", "E2"])
    res.align_centre()
    res.chart(3, 4, 6, 5)

    Report.wb.save("class.xlsx")

