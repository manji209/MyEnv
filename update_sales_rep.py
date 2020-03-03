import pyodbc
import xlrd
import pandas as pd
import numpy as np


fname = 'Data/Larry_Inactive_Customers.xlsx'
df = pd.read_excel(fname)

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


cur = conn.cursor()

'''
sql = "SELECT top 1000 * FROM dbo.CUSFIL00"
cur.execute(sql)

for row in cur:
    print('row = %r' % (row,))

'''
sqlstring = "UPDATE dbo.CUSFIL00 SET sls_rep=100 WHERE cust_no=?"
for index, row in df.iterrows():
    try:
        cur.execute(sqlstring, row.cust_no)
    except Exception as e:
        print(e)


cur.close()
conn.commit()

