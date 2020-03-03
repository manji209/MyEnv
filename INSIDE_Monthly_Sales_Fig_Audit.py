import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np
from datetime import datetime

# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sales_rep = 5
year = 2019
month = 10

sqlstring = """SELECT pi_hist_hdr_invc_no, hdr_invc_dat, hdr_cust_no, cust_nam, tot_sls, freight_amt
                FROM dbo.IHSHDR00 
                WHERE hdr_sls_rep_no =? AND
                DATEPART(YEAR, hdr_invc_dat)=? AND DATEPART(MONTH, hdr_invc_dat)=?
                ORDER BY hdr_invc_dat ASC"""



df = pd.read_sql(sqlstring, conn, params=(sales_rep, year, month))
df['Before']=""

for idx, row in df.iterrows():
    df.loc[idx, 'TOTAL SALES'] = row.tot_sls + row.freight_amt


# Create a Pandas Excel writer using XlsxWriter as the engine.
fout_name = 'Out/CHI_INVOICE_REPORT_' + str(month) + '_' + str(year) + '.xlsx'

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(fout_name, engine='xlsxwriter')


#df.to_excel(writer, sheet_name='Sheet1', index=False)
df.to_excel(writer, sheet_name='Sheet1', index=False)

writer.save()
