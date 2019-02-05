import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT tt.invc_no, tt.invc_dat, tt.cust_no, tt.item_no, tt.invc_or_cred, tt.desc_1, tt.desc_2, tt.qty, tt.prc, tt.ext_prc, tt.ord_no, tt.sls_rep 
                FROM dbo.IHSLIN00 tt
                INNER JOIN
                    (SELECT item_no, max(invc_dat) as MaxDateTime
                    FROM dbo.IHSLIN00
                    GROUP BY item_no) groupedtt
                ON tt.item_no = groupedtt.item_no
                WHERE (tt.cust_no=? OR tt.cust_no=? OR tt.cust_no=?)
                AND tt.invc_dat = groupedtt.MaxDateTime"""


L2 = 'LEE #2'
L3 = 'LEE #3'
L4 = 'LEE #4'

df = pd.read_sql(sqlstring, conn, params=(L2, L3, L4))
# df.columns = df.columns.str.strip()

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
df_ext_prc = df.pivot_table(values=['ext_prc', 'qty'], index=['item_no'], aggfunc={'ext_prc':np.sum, 'qty':np.sum})
df_ext_prc.sort_values(by="ext_prc", ascending=False, inplace=True)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/T011_12_2018_pivot_out.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Orig')
# Write dataframe to Excel
df_ext_prc.to_excel(writer, sheet_name='Sheet1', index=True)

# Open the workbook and worksheet for editing
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,###.00'})
#center = workbook.add_format({'align': 'left'})


worksheet.set_column('B:B', 15, money)
#worksheet.set_column('C:C', 20, center)

writer.save()

'''



