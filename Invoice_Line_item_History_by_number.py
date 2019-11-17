import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

#sqlstring = 'SELECT invc_no, invc_dat, cust_no, item_no, invc_or_cred, desc_1, desc_2, qty, prc, ext_prc, ord_no, sls_rep FROM dbo.IHSLIN00 WHERE year(invc_dat) >= 2018 AND cust_no=?'
sqlstring = 'SELECT invc_no, invc_dat, cust_no, item_no, invc_or_cred, desc_1, desc_2, qty, prc, ord_no FROM dbo.IHSLIN00 WHERE invc_no =?'
# Load prev ITEMS list
fname_invc = 'Data/100819.xlsx'
df_invc = pd.read_excel(fname_invc)
print(df_invc)

df_combine = pd.DataFrame(columns=['invc_no', 'invc_dat', 'cust_no', 'item_no', 'invc_or_cred', 'desc_1', 'desc_2', 'qty', 'prc', 'ord_no'])

for row in df_invc.itertuples(index=True):
    df_temp = pd.read_sql(sqlstring, conn, params={row.invc_no})
    #print(df_temp)
    df_combine = df_combine.append(df_temp)
    print(df_combine)




# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/Invoice_Consolidated_100819.xlsx', engine='xlsxwriter')

df_combine.to_excel(writer, sheet_name='Sheet1', index=False)


writer.save()
