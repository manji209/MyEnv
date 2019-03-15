import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np
from datetime import datetime

# Connect to SQL Server and set cursor
conn = pyodbc.connect(
    'DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT TOP (25) invc_no_alt, invc_or_cred, hdr_invc_dat, hdr_ord_dat, tot_ord_qty, tot_sls
                FROM dbo.IHSHDR00 
                WHERE hdr_cust_no =?
                ORDER BY invc_no_alt DESC"""


m = datetime.today().month
cust = 'K060'
#month = 2
year = datetime.today().year

df = pd.read_sql(sqlstring, conn, params={cust})


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/View_Invoice_History1.xlsx', engine='xlsxwriter')


#df.to_excel(writer, sheet_name='Sheet1', index=False)
df.to_excel(writer, sheet_name='Latest Invoice')

writer.save()
print('month is: ', m)
print(df)




