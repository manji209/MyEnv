import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np
from datetime import datetime

fname = "../Data/Invc_list.xlsx"
df_invc = pd.read_excel(fname, sheet_name='Sheet1')

# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT pi_hist_hdr_invc_no, cust_nam, hdr_cust_no, invc_or_cred, appl_to_no, hdr_invc_dat, hdr_sls_rep_no, tot_sls, freight_amt
                FROM dbo.IHSHDR00 
                WHERE pi_hist_hdr_invc_no = ?"""



df_combine = pd.DataFrame(columns=['pi_hist_hdr_invc_no', 'cust_nam', 'hdr_cust_no', 'invc_or_cred', 'appl_to_no', 'hdr_invc_dat', 'hdr_sls_rep_no', 'tot_sls', 'freight_amt'])

for row in df_invc.itertuples(index=True):
    df_temp = pd.read_sql(sqlstring, conn, params={row.invc_no})

    if not df_temp.empty:
        df_combine = df_combine.append(df_temp)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('../Out/Tony_Sales_Audit_2.xlsx', engine='xlsxwriter')


#df.to_excel(writer, sheet_name='Sheet1', index=False)
df_combine.to_excel(writer, sheet_name='Sheet1')

writer.save()


