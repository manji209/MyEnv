import pyodbc
import xlrd
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = """SELECT invc_dat, item_no, desc_1, desc_2, qty, prc, ext_prc, ord_no FROM dbo.IHSLIN00 WHERE cust_no='K001' AND invc_dat BETWEEN '2018-11-15' AND '2019-02-04'"""

df = pd.read_sql(sqlstring,conn)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/history_out.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Inv History', index=False)

writer.save()


#print(df)


