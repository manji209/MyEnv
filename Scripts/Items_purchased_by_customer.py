import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

#sqlstring = 'SELECT invc_no, invc_dat, cust_no, item_no, invc_or_cred, desc_1, desc_2, qty, prc, ext_prc, ord_no, sls_rep FROM dbo.IHSLIN00 WHERE year(invc_dat) >= 2018 AND cust_no=?'
sqlstring = 'SELECT invc_dat, cust_no, item_no, invc_or_cred, desc_1, desc_2, qty FROM dbo.IHSLIN00 WHERE DATEPART(YEAR, invc_dat)=? AND cust_no=?'

year = 2019
cust_no = 'YOUNG'

df = pd.read_sql(sqlstring, conn, params=(year, cust_no))
df['item_no'] = df['item_no'].str.strip()


# df_ext_prc = df.pivot_table(values=['ext_prc', 'qty'], index=['item_no'], aggfunc={'ext_prc':np.sum, 'qty':np.sum})
# df_ext_prc.sort_values(by="ext_prc", ascending=False, inplace=True)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('../Out/Purchase_history_YOUNG_2019.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sales_Report')


writer.save()

