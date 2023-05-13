import sqlite3 as sql
from deco import parsing_time

from excel_reader import excel_reader
from funcs.GUI_funcs import reports


@parsing_time
def create_and_update_customers_tb():
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute("DROP TABLE IF EXISTS customers")
        curs.execute("""CREATE TABLE IF NOT EXISTS 'customers' (
                        'id' INTEGER NOT NULL UNIQUE,
                        'lastName',
                        'firstName',
                        'fathersName',
                        'phone_number',
                        'e_mail',
                        'city',
                        'company',
                        'doc_numb',
                        PRIMARY KEY('id' AUTOINCREMENT)
                        );""")
        db.commit()

        curs.executemany("""INSERT INTO customers (lastName,
                                                    firstName,
                                                    fathersName,
                                                    phone_number,
                                                    e_mail,
                                                    city,
                                                    company,
                                                    doc_numb)
                                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)""", excel_reader(7, 1, 9, reports["report1"]))

@parsing_time
def create_and_update_sales_tb():
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute("DROP TABLE IF EXISTS sales")
        curs.execute("""CREATE TABLE IF NOT EXISTS 'sales' (
                        'id' INTEGER NOT NULL UNIQUE,
                        'full_name',
                        'sales',
                        'date',
                        'dogovir',
                        'manager',
                        'company',
                        PRIMARY KEY('id' AUTOINCREMENT)
                        );""")
        db.commit()

        curs.executemany("""INSERT INTO sales (full_name,
                                                    sales,
                                                    date,
                                                    dogovir,
                                                    manager,
                                                    company)
                                                    VALUES (?, ?, ?, ?, ?, ?)""", excel_reader(7, 1, 7, reports["report2"]))

def get_date_sql(year: int) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute(
            f"""SELECT manager,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '01' THEN sales ELSE 0 END), 2) AS a,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '02' THEN sales ELSE 0 END), 2) AS b,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '03' THEN sales ELSE 0 END), 2) AS c,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '04' THEN sales ELSE 0 END), 2) AS d,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '05' THEN sales ELSE 0 END), 2) AS e,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '06' THEN sales ELSE 0 END), 2) AS f,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '07' THEN sales ELSE 0 END), 2) AS g,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '08' THEN sales ELSE 0 END), 2) AS j,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '09' THEN sales ELSE 0 END), 2) AS n,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '10' THEN sales ELSE 0 END), 2) AS t,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '11' THEN sales ELSE 0 END), 2) AS o,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '12' THEN sales ELSE 0 END), 2) AS p,
                ROUND(SUM(CASE WHEN strftime('%Y', date) = '{year}' THEN sales ELSE 0 END), 2) AS suma
            FROM sales
            WHERE strftime('%Y', date) = '{year}'
            GROUP BY manager;"""
        )
    sales_report_by_month = []
    while True:
        next_row = curs.fetchone()
        if next_row:
            sales_report_by_month.append(list(next_row))
        else:
            break

    # print(sales_report_by_month)

    return sales_report_by_month

# get_date_sql(2022)
