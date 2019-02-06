import pyodbc
import xlrd
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

'''
sqlstring = """SELECT DISTINCT cust_no, dist_dat, journal_no, SUM(dist_amt) AS total_sum FROM dbo.ARDIST00 
                WHERE (dist_dat BETWEEN ? AND ?)
                GROUP BY cust_no, dist_dat"""
'''
sqlstring = """SELECT * FROM dbo.ARDIST00 
                WHERE (dist_dat BETWEEN ? AND ?)"""

d1 = '2018-09-28'
d2 = '2018-09-30'
df = pd.read_sql(sqlstring,conn, params=(d1,d2))

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Out/AR_out3.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='AR History', index=False)

writer.save()


#print(df)


