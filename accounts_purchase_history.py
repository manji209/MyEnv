import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

#sqlstring = 'SELECT invc_no, invc_dat, cust_no, item_no, invc_or_cred, desc_1, desc_2, qty, prc, ext_prc, ord_no, sls_rep FROM dbo.IHSLIN00 WHERE year(invc_dat) >= 2018 AND cust_no=?'
sqlstring = 'SELECT hdr_cust_no, invc_or_cred, tot_sls FROM dbo.IHSHDR00'


df = pd.read_sql(sqlstring, conn)

#df_ext_prc = df.pivot_table(values=['ext_prc', 'qty'], index=['item_no'], aggfunc={'ext_prc':np.sum, 'qty':np.sum})
#df_ext_prc.sort_values(by="ext_prc", ascending=False, inplace=True)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/accounts_pivot_out.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Raw Data', index=False)
# Write dataframe to Excel
#df_ext_prc.to_excel(writer, sheet_name='Sheet1', index=True)

'''
# Open the workbook and worksheet for editing
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,###.00'})
#center = workbook.add_format({'align': 'left'})


worksheet.set_column('B:B', 15, money)
#worksheet.set_column('C:C', 20, center)
'''
writer.save()





