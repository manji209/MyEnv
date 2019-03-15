import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

#sqlstring = 'SELECT invc_no, invc_dat, cust_no, item_no, invc_or_cred, desc_1, desc_2, qty, prc, ext_prc, ord_no, sls_rep FROM dbo.IHSLIN00 WHERE year(invc_dat) >= 2018 AND cust_no=?'
sqlstring = 'SELECT invc_no, invc_dat, cust_no, item_no, invc_or_cred, desc_1, desc_2, qty, prc, ext_prc, ord_no, sls_rep FROM dbo.IHSLIN00 WHERE DATEPART(YEAR, invc_dat)=? AND (DATEPART(MONTH, invc_dat)=? OR DATEPART(MONTH, invc_dat)=?)'

year = 2019
month = 1
month2 = 2


df = pd.read_sql(sqlstring, conn, params=(year, month, month2))
# df.columns = df.columns.str.strip()


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/Monthly_Invoice_History.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', index=False)


writer.save()





