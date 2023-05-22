import sqlite3 as sql
from deco import parsing_time

from excel_reader import excel_reader
from funcs.GUI_funcs import reports


@parsing_time
def create_and_update_customers_tb():
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute("DROP TABLE IF EXISTS customers")
        curs.execute(
            """CREATE TABLE IF NOT EXISTS 'customers' (
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
                        );"""
        )
        db.commit()

        curs.executemany(
            """INSERT INTO customers (lastName,
                                                    firstName,
                                                    fathersName,
                                                    phone_number,
                                                    e_mail,
                                                    city,
                                                    company,
                                                    doc_numb)
                                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            excel_reader(7, 1, 9, reports["report1"]),
        )


@parsing_time
def create_and_update_sales_tb():
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute("DROP TABLE IF EXISTS sales")
        curs.execute(
            """CREATE TABLE IF NOT EXISTS 'sales' (
                        'id' INTEGER NOT NULL UNIQUE,
                        'full_name',
                        'sales',
                        'date',
                        'dogovir',
                        'manager',
                        'company',
                        PRIMARY KEY('id' AUTOINCREMENT)
                        );"""
        )
        db.commit()

        curs.executemany(
            """INSERT INTO sales (full_name,
                                                    sales,
                                                    date,
                                                    dogovir,
                                                    manager,
                                                    company)
                                                    VALUES (?, ?, ?, ?, ?, ?)""",
            excel_reader(7, 1, 7, reports["report2"]),
        )


def get_date_sql(year: int = 2022, month: int = 6) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        query = f"""
            SELECT manager,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '01' AND 1 <= {month} THEN sales ELSE NULL END), 2) AS a,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '02' AND 2 <= {month} THEN sales ELSE NULL END), 2) AS b,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '03' AND 3 <= {month} THEN sales ELSE NULL END), 2) AS c,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '04' AND 4 <= {month} THEN sales ELSE NULL END), 2) AS d,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '05' AND 5 <= {month} THEN sales ELSE NULL END), 2) AS e,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '06' AND 6 <= {month} THEN sales ELSE NULL END), 2) AS f,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '07' AND 7 <= {month} THEN sales ELSE NULL END), 2) AS g,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '08' AND 8 <= {month} THEN sales ELSE NULL END), 2) AS j,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '09' AND 9 <= {month} THEN sales ELSE NULL END), 2) AS n,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '10' AND 10 <= {month} THEN sales ELSE NULL END), 2) AS t,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '11' AND 11 <= {month} THEN sales ELSE NULL END), 2) AS o,
                ROUND(SUM(CASE WHEN strftime('%m', date) = '12' AND 12 <= {month} THEN sales ELSE NULL END), 2) AS p,
                ROUND(SUM(CASE WHEN strftime('%Y-%m', date) <= '{year}-{month:02}' THEN sales ELSE NULL END), 2) AS suma
            FROM sales
            WHERE strftime('%Y', date) = '{year}'
            GROUP BY manager;
        """
        curs.execute(query)
        sales_report_by_month = []
        while True:
            next_row = curs.fetchone()
            if next_row:
                sales_report_by_month.append(list(next_row))
            else:
                break

    return sales_report_by_month



def get_data_report_2(year: int = 2022, month: int = 6) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute(
            f"""
                SELECT c.company, s.manager, SUM(s.sales) as total_sales, SUM(s.sales) * 0.18 as clear_income
                FROM customers as c
                LEFT JOIN sales as s
                ON c.company = s.company
                WHERE strftime('%m', s.date) = '{str(month).zfill(2)}' and strftime('%Y', s.date) = '{year}' and s.manager IS NOT NULL
                GROUP BY c.company
                HAVING total_sales IS NOT NULL
                ORDER BY total_sales DESC;
            """
        )
    sales_report_by_month = []
    while True:
        next_row = curs.fetchone()
        if next_row:
            sales_report_by_month.append(list(next_row))
        else:
            break

    print(sales_report_by_month)

    return sales_report_by_month


def get_data_report_sum_2(year: int = 2022, month: int = 6) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute(
            f"""
                SELECT SUM(s.sales) as total_sales, SUM(s.sales) * 0.18 as total_income
                FROM customers as c
                LEFT JOIN sales as s
                ON c.company = s.company
                WHERE strftime('%m', s.date) = '{month:02}' and strftime('%Y', s.date) = '{year}' and s.manager IS NOT NULL;
            """
        )
    sales_report_by_month = list(curs.fetchone())

    print(sales_report_by_month)

    return sales_report_by_month


def get_data_report_3(year: int = 2022, month: int = 6) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute(
            f"""
                SELECT c.city, s.manager, SUM(s.sales) as total_sales, SUM(s.sales) * 0.18 as clear_income
                FROM customers as c
                LEFT JOIN sales as s
                ON c.company = s.company
                WHERE strftime('%m', s.date) = '{str(month).zfill(2)}' and strftime('%Y', s.date) = '{year}' and s.manager IS NOT NULL
                GROUP BY c.city
                HAVING total_sales IS NOT NULL
                ORDER BY total_sales DESC;
            """
        )
    sales_report_by_month = []
    while True:
        next_row = curs.fetchone()
        if next_row:
            sales_report_by_month.append(list(next_row))
        else:
            break

    print(sales_report_by_month)

    return sales_report_by_month


def get_data_report_sum_3(year: int = 2022, month: int = 6) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute(
            f"""
                SELECT SUM(s.sales) as total_sales, SUM(s.sales) * 0.18 as total_income
                FROM customers as c
                LEFT JOIN sales as s
                ON c.company = s.company
                WHERE strftime('%m', s.date) = '{month:02}' and strftime('%Y', s.date) = '{year}' and s.manager IS NOT NULL;
            """
        )
    sales_report_by_month = list(curs.fetchone())

    print(sales_report_by_month)

    return sales_report_by_month


def get_data_report_4_1(year: int = 2022, month: int = 6) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute(
            f"""
                SELECT manager, strftime('%d-%m-%Y', date), MIN(sales) as min_sale
                FROM sales
                WHERE strftime('%m', date) = '{month:02}' and strftime('%Y', date) = '{year}'
                GROUP BY manager
                ORDER BY min_sale;
            """
        )
    sales_report_by_month = []
    while True:
        next_row = curs.fetchone()
        if next_row:
            sales_report_by_month.append(list(next_row))
        else:
            break

    print(sales_report_by_month)

    return sales_report_by_month

def get_data_report_4_2(year: int = 2022, month: int = 6) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute(
            f"""
                SELECT manager, strftime('%d-%m-%Y', date), MAX(sales) as max_sale
                FROM sales
                WHERE strftime('%m', date) = '{month:02}' and strftime('%Y', date) = '{year}'
                GROUP BY manager
                ORDER BY max_sale;
            """
        )
    sales_report_by_month = []
    while True:
        next_row = curs.fetchone()
        if next_row:
            sales_report_by_month.append(list(next_row))
        else:
            break

    print(sales_report_by_month)

    return sales_report_by_month


def get_data_report_4_3(year: int = 2022, month: int = 6) -> list:
    with sql.connect("Base_zvit1.db") as db:
        curs = db.cursor()
        curs.execute(
            f"""
                SELECT manager, SUM(sales) * 0.18 as total_income
                FROM sales
                WHERE strftime('%m', date) = '{month:02}' and strftime('%Y', date) = '{year}'
                GROUP BY manager;
            """
        )
    sales_report_by_month = []
    while True:
        next_row = curs.fetchone()
        if next_row:
            sales_report_by_month.append(list(next_row))
        else:
            break

    print(sales_report_by_month)

    return sales_report_by_month

