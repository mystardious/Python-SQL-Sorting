import csv
import pyodbc
import xlrd
import pandas as pd

server_name = 'PC-MOHAMED\SQLEXPRESS'
db_name = 'master'
input_name = 'input2.xlsx'

def run():

    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server='+server_name+';'
                          'Database='+db_name+';'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()

    # read data
    data = pd.read_excel(input_name)

    # open workbook
    book = xlrd.open_workbook(input_name)
    sheet = book.sheet_by_name("Sheet1")

    reset_table = """
    IF OBJECT_ID('temp', 'U') IS NOT NULL 
        DROP TABLE temp;
    CREATE TABLE temp (
        A varchar(255),
        B date,
        C varchar(255),
        D varchar(255),
        E varchar(255),
        F varchar(255),
        G varchar(255),
        H varchar(255),
        I varchar(255),
        J varchar(255),
        K varchar(255),
        L int,
    )"""

    cursor.execute(reset_table)

    query = """
    INSERT INTO temp (
        A,
        B,
        C,
        D,
        E,
        F,
        G,
        H,
        I,
        J,
        K,
        L
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""

    for r in range(1, sheet.nrows):
        A = sheet.cell(r, 0).value
        B = sheet.cell(r, 1).value
        C = sheet.cell(r, 2).value
        D = sheet.cell(r, 3).value
        E = sheet.cell(r, 4).value
        F = sheet.cell(r, 5).value
        G = sheet.cell(r, 6).value
        H = sheet.cell(r, 7).value
        I = sheet.cell(r, 8).value
        J = sheet.cell(r, 9).value
        K = sheet.cell(r, 10).value
        L = sheet.cell(r, 11).value

        # Assign values from each row
        values = (A, B, C, D, E, F, G, H, I, J, K, L)

        # Execute sql Query
        cursor.execute(query, values)

    conn.commit()
    conn.close()

run()

