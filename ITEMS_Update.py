import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
# conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

# Monthly Sales Figure
sqlstring = """SELECT item_no, item_desc_1, item_desc_2, stock_unit_of_meas, item_prc_1, prc_unit_of_meas, 
                item_standard_cost, replacement_cost, weight, qty_on_hand, qty_commitd
                FROM dbo.ITMFIL00"""


df = pd.read_sql(sqlstring, conn)
df['item_no'] = df['item_no'].str.strip()
#df.applymap(lambda x: x.strip() if type(x)==str else x)


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Import/ITEMS.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', index=False)


writer.save()
