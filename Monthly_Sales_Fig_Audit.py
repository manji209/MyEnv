import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np
from datetime import datetime

# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT pi_hist_hdr_invc_no, hdr_invc_dat, hdr_cust_no, cust_nam, tot_sls, freight_amt
                FROM dbo.IHSHDR00 
                WHERE hdr_invc_dat BETWEEN '2018-04-01' AND '2019-03-31'
                ORDER BY hdr_invc_dat ASC"""



df = pd.read_sql(sqlstring, conn)
df['Before']=""

for idx, row in df.iterrows():
    df.loc[idx, 'Before'] = row.tot_sls - row.freight_amt


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/Invoice_History.xlsx', engine='xlsxwriter')


#df.to_excel(writer, sheet_name='Sheet1', index=False)
df.to_excel(writer, sheet_name='Sheet1')

writer.save()


