import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np

# Connect to SQL Server and set cursor
conn = pyodbc.connect(
    'DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT item_no, cust_no, invc_dat, qty, prc 
                FROM dbo.IHSLIN00 
                WHERE DATEPART(YEAR, invc_dat)>=?"""



item = 'FR189'
year = 2018

df = pd.read_sql(sqlstring, conn, params={year})

df['invc_dat'] = pd.to_datetime(df['invc_dat'], errors='coerce')

monthly_df = (df.groupby(['item_no', pd.Grouper(freq='M', key='invc_dat')])['qty']
                  .sum()
                  .unstack(fill_value=0))

weekly_df = (df.groupby(['item_no', pd.Grouper(freq='W', key='invc_dat')])['qty']
                  .sum()
                  .unstack(fill_value=0))

monthly_df.columns = monthly_df.columns.month_name()
weekly_df.columns = weekly_df.columns.date

# weekly_df.columns = weekly_df.columns.date
# df = pd.read_sql(sqlstring, conn, params=(inv, credit, inv, y1, y4, credit, y1, y4, inv, y1, credit, y1, inv, y2, credit, y2, inv, y3, credit, y3, inv, y4, credit, y4))
# df['hdr_cust_no'] = df['hdr_cust_no'].str.strip()
# df_ext_prc = df.pivot_table(values=['ext_prc', 'qty'], index=['item_no'], aggfunc={'ext_prc':np.sum, 'qty':np.sum})
# df_ext_prc.sort_values(by="ext_prc", ascending=False, inplace=True)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/Item_Sales_History1.xlsx', engine='xlsxwriter')


#df.to_excel(writer, sheet_name='Sheet1', index=False)
monthly_df.to_excel(writer, sheet_name='Monthly')
weekly_df.to_excel(writer, sheet_name='Weekly')
# Write dataframe to Excel
# df_ext_prc.to_excel(writer, sheet_name='Sheet1', index=True)

'''
# Open the workbook and worksheet for editing
workbook = writer.book
worksheet = writer.sheets['Sheet1']


# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,###.00'})
#center = workbook.add_format({'align': 'left'})


worksheet.set_column('A2:F2', 20, money)
#worksheet.set_column('C:C', 20, center)
'''
writer.save()





