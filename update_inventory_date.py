import pyodbc
import xlrd
import pandas as pd
import numpy as np

book = xlrd.open_workbook("Import/INVENTORY_SALES_093019.xlsx")
sheet = book.sheet_by_name("Sheet1")

# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()

sqlstring = """UPDATE dbo.INVTRX00 SET trx_dat=?,trx_dat_a=?
                WHERE trx_dat=?"""

trx_date = '07-31-19'
new_date = '09-30-19'
values = [new_date, new_date, trx_date]

cur.execute(sqlstring, values)


conn.commit()



# Close the database connection
conn.close()