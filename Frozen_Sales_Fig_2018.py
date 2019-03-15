import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

# Monthly Sales Figure
sqlstring = """SELECT
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=1 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=1 THEN ext_prc ELSE 0 END) AS jan_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=1 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=1 THEN (qty*unit_cost) ELSE 0 END) AS jan_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=2 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=2 THEN ext_prc ELSE 0 END) AS feb_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=2 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=2 THEN (qty*unit_cost) ELSE 0 END) AS feb_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=3 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=3 THEN ext_prc ELSE 0 END) AS mar_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=3 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=3 THEN (qty*unit_cost) ELSE 0 END) AS mar_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=4 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=4 THEN ext_prc ELSE 0 END) AS april_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=4 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=4 THEN (qty*unit_cost) ELSE 0 END) AS april_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=5 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=5 THEN ext_prc ELSE 0 END) AS may_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=5 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=5 THEN (qty*unit_cost) ELSE 0 END) AS may_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=6 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=6 THEN ext_prc ELSE 0 END) AS june_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=6 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=6 THEN (qty*unit_cost) ELSE 0 END) AS june_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=7 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=7 THEN ext_prc ELSE 0 END) AS july_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=7 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=7 THEN (qty*unit_cost) ELSE 0 END) AS july_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=8 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=8 THEN ext_prc ELSE 0 END) AS aug_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=8 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=8 THEN (qty*unit_cost) ELSE 0 END) AS aug_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=9 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=9 THEN ext_prc ELSE 0 END) AS sept_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=9 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=9 THEN (qty*unit_cost) ELSE 0 END) AS sept_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=10 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=10 THEN ext_prc ELSE 0 END) AS oct_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=10 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=10 THEN (qty*unit_cost) ELSE 0 END) AS oct_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=11 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=11 THEN ext_prc ELSE 0 END) AS nov_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=11 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=11 THEN (qty*unit_cost) ELSE 0 END) AS nov_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=12 THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=12 THEN ext_prc ELSE 0 END) AS dec_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=12 THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? AND DATEPART(MONTH, invc_dat)=12 THEN (qty*unit_cost) ELSE 0 END) AS dec_ext_cost,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? THEN ext_prc ELSE 0 END) AS frozen_2018_ext_prc,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, invc_dat)=? THEN (qty*unit_cost) ELSE 0 END) AS frozen_2018_ext_cost
                FROM dbo.IHSLIN00
                WHERE desc_2=?"""


# Total Yearly Figures
sqlstring_2 = """SELECT
                SUM(CASE WHEN invc_or_cred =? THEN ext_prc ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? THEN ext_prc ELSE 0 END) AS tot_2018_ext_prc,
                SUM(CASE WHEN invc_or_cred =? THEN (qty*unit_cost) ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? THEN (qty*unit_cost) ELSE 0 END) AS tot_2018_ext_cost
                FROM dbo.IHSLIN00
                WHERE DATEPART(YEAR, invc_dat)=?
                """


inv = 'I'
credit = 'C'
desc = 'FROZEN'

y4 = 2018

df_frozen = pd.read_sql(sqlstring, conn, params=(inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4,
                                          inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4,
                                          inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4, inv, y4, credit, y4,
                                          inv, y4, credit, y4, inv, y4, credit, y4, desc))

df_total = pd.read_sql(sqlstring_2, conn, params=(inv, credit, inv, credit, y4))


#df['hdr_cust_no'] = df['hdr_cust_no'].str.strip()
#df_ext_prc = df.pivot_table(values=['ext_prc', 'qty'], index=['item_no'], aggfunc={'ext_prc':np.sum, 'qty':np.sum})
#df_ext_prc.sort_values(by="ext_prc", ascending=False, inplace=True)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/frozen_figures_2018_new_7.xlsx', engine='xlsxwriter')

frames = [df_frozen, df_total]
df = pd.concat(frames, axis=1, join_axes=[df_frozen.index])

df.to_excel(writer, sheet_name='Sheet1', index=False)
# Write dataframe to Excel
#df_ext_prc.to_excel(writer, sheet_name='Sheet1', index=True)

# Open the workbook and worksheet for editing
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,###.00'})
#center = workbook.add_format({'align': 'left'})


worksheet.set_column('A2:AC2', 20, money)
#worksheet.set_column('C:C', 20, center)

writer.save()





