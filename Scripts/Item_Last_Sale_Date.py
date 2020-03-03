import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
# conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
# conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

# Monthly Sales Figure
sqlstring = """SELECT item_no, item_desc_1, item_desc_2, lst_sls_dat
                FROM dbo.ITMFIL00
                WHERE DATEPART(YEAR, lst_sls_dat)=?"""

year = 2019

df = pd.read_sql(sqlstring, conn, params={year})
df['item_no'] = df['item_no'].str.strip()

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('../Data/ITEMS_LAST_SOLD.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', index=False)


writer.save()
