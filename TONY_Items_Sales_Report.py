import pyodbc
import xlrd
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()

sales_rep = 21
year = 2019
month = 8

sqlstring = """SELECT item_no, desc_1, desc_2, cust_no, invc_no, qty, prc, ext_prc, invc_dat FROM dbo.IHSLIN00 
            WHERE sls_rep =? AND
            DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=?
            ORDER BY item_no"""
'''
sqlstring = """SELECT item_no, desc_1, desc_2, cust_no, invc_no, qty, prc, ext_prc, invc_dat FROM dbo.IHSLIN00 
            WHERE sls_rep =? AND
            DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)<=?
            ORDER BY item_no"""

'''
df = pd.read_sql(sqlstring,conn, params=(sales_rep, year, month))
df['item_no'] = df['item_no'].str.strip()

# Create a Pandas Excel writer using XlsxWriter as the engine.
fout_name = 'Out/TONY_ITEMS_REPORT_' + str(month) + '_' + str(year) + '.xlsx'
#writer = pd.ExcelWriter('Out/ITEMS.xlsx', engine='xlsxwriter')
writer = pd.ExcelWriter(fout_name, engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', index=False)


writer.save()



print(df)

