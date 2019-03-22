import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT invc_no, invc_dat, cust_no, item_no, invc_or_cred, desc_1, desc_2, qty, prc, ext_prc, ord_no, sls_rep 
                FROM dbo.IHSLIN00
                WHERE cust_no=?
                AND DATEPART(YEAR, invc_dat)=?"""


cust = 'T011'
year = 2019

df = pd.read_sql(sqlstring, conn, params=(cust, year))
df.sort_values(by="invc_dat", ascending=True, inplace=True)
# df.columns = df.columns.str.strip()

'''
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/LEE_pivot_out.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Orig', index=False)

# pivot table by item_no
df_ext_prc = df.pivot_table(values=['qty'], index=['item_no', 'desc_1', 'desc_2', 'invc_dat'], aggfunc={'qty':np.sum})
df_ext_prc.sort_values(by="item_no", ascending=True, inplace=True)

# Write dataframe to Excel
df_ext_prc.to_excel(writer, sheet_name='Pivot', index=True)


writer.save()
'''

# pivot table by item_no
df_ext_prc = df.pivot_table(values=['ext_prc'], index=['item_no', 'invc_no', 'qty', 'invc_dat', 'desc_1', 'desc_2'], aggfunc={'ext_prc':np.sum})
df_ext_prc.sort_values(by=['ext_prc', 'invc_dat'], ascending=False, inplace=True)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/T011_2019_pivot_out.xlsx', engine='xlsxwriter')

# Raw
#df.to_excel(writer, sheet_name='Orig')
# Write dataframe to Excel
df_ext_prc.to_excel(writer, sheet_name='Pivot Items', index=True)

# Open the workbook and worksheet for editing
workbook = writer.book
worksheet = writer.sheets['Pivot Items']

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,###.00'})
#center = workbook.add_format({'align': 'left'})


worksheet.set_column('G:G', 15, money)
#worksheet.set_column('C:C', 20, center)

writer.save()





