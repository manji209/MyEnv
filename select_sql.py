import pyodbc
import xlrd
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsmaster;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsmaster;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=DINHPC\SQLEXPRESS;DATABASE=pbsmaster;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=MULTIMEDIAPC\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = "SELECT * FROM dbo.ICPHXF00 WHERE item_no = 'VN104'"

df = pd.read_sql(sqlstring,conn)


print(df)



'''
cur.execute(sqlstring)
for row in cur:
    print(row)
'''

