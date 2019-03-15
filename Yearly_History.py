import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np

# Connect to SQL Server and set cursor
conn = pyodbc.connect(
    'DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT 
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) AS yr_2013,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) AS yr_2014,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) AS yr_2015,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) AS yr_2016,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) AS yr_2017,
                SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) - SUM(CASE WHEN invc_or_cred =? AND DATEPART(YEAR, hdr_invc_dat)=? THEN tot_sls ELSE 0 END) AS yr_2018 
                FROM dbo.IHSHDR00"""


inv = 'I'
credit = 'C'
y1 = 2013
y2 = 2014
y3 = 2015
y4 = 2016
y5 = 2017
y6 = 2018

df = pd.read_sql(sqlstring, conn, params=(inv, y1, credit, y1, inv, y2, credit, y2, inv, y3, credit, y3, inv, y4, credit, y4, inv, y5, credit, y5, inv, y6, credit, y6))
# df = pd.read_sql(sqlstring, conn, params=(inv, credit, inv, y1, y4, credit, y1, y4, inv, y1, credit, y1, inv, y2, credit, y2, inv, y3, credit, y3, inv, y4, credit, y4))
# df['hdr_cust_no'] = df['hdr_cust_no'].str.strip()
# df_ext_prc = df.pivot_table(values=['ext_prc', 'qty'], index=['item_no'], aggfunc={'ext_prc':np.sum, 'qty':np.sum})
# df_ext_prc.sort_values(by="ext_prc", ascending=False, inplace=True)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/Yearly_Sales_Totals.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', index=False)
# Write dataframe to Excel
# df_ext_prc.to_excel(writer, sheet_name='Sheet1', index=True)


# Open the workbook and worksheet for editing
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,###.00'})
#center = workbook.add_format({'align': 'left'})


worksheet.set_column('A2:F2', 20, money)
#worksheet.set_column('C:C', 20, center)

writer.save()





