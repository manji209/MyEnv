import pyodbc
import xlrd
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = "SELECT cust_no, city, state, zipcode FROM dbo.CUSFIL00"

df = pd.read_sql(sqlstring,conn)


print(df)



'''
cur.execute(sqlstring)
for row in cur:
    print(row)
'''

