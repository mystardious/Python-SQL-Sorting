import csv
import pyodbc

server_name = 'PC-MOHAMED\SQLEXPRESS'
db_name = 'master'
input_name = 'input.csv'

def run():

    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server='+server_name+';'
                          'Database='+db_name+';'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()

    reset_table = """
    IF OBJECT_ID('temp', 'U') IS NOT NULL 
        DROP TABLE temp;

    CREATE TABLE temp (
        A varchar(255),
        B DATE,
        C varchar(255),
        D varchar(255),
        E varchar(255),
        F varchar(255),
        G varchar(255),
        H varchar(255),
        I varchar(255),
        J varchar(255),
        K varchar(255),
        L varchar(255),
    );
    """

    cursor.execute(reset_table)

    with open(input_name, 'r') as f:
        reader = csv.reader(f)
        columns = next(reader)
        query = 'insert into temp({0}) values ({1})'
        query = query.format(','.join(columns), ','.join('?' * len(columns)))
        cursor = conn.cursor()
        for data in reader:
            cursor.execute(query, data)
        cursor.commit()


    conn.commit()
    conn.close()

run()

