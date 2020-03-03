import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT DISTINCT item_no FROM dbo.IHSLIN00 WHERE DATEPART(YEAR, invc_dat)=? AND disc_amt > 0
                AND (cust_no='LEE #2' OR cust_no='LEE #3' OR cust_no='LEE #4')"""

year = 2019


df = pd.read_sql(sqlstring, conn, params={year})
df['item_no'] = df['item_no'].str.strip()


# df_ext_prc = df.pivot_table(values=['ext_prc', 'qty'], index=['item_no'], aggfunc={'ext_prc':np.sum, 'qty':np.sum})
# df_ext_prc.sort_values(by="ext_prc", ascending=False, inplace=True)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('../Out/LEE_LEE_2_Percent_2019.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Discount Items')


writer.save()

