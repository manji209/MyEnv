import pyodbc
import xlrd
import pandas as pd


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = """UPDATE dbo.LINITM00 SET disc_amt=2
            WHERE item_no='HM151' AND ord_no=115418"""

cur.execute(sqlstring)

conn.commit()
cur.close()
