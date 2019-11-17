import pyodbc
import xlrd
import pandas as pd
import numpy as np

# Query dbo.ICPHXF00 table for items that have not been physically counted AND Qty_on_Hand is not ZERO (Can be positive or negative Qty)


conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = """SELECT item_no FROM dbo.ICPHXF00
                WHERE qty_on_hand <> 0 AND cnt_qty_on_hand_1 = ?"""


df = pd.read_sql(sqlstring, conn, params={-99999999.99999})


print(df)


sqlstring = "UPDATE dbo.ICPHXF00 SET cnt_qty_on_hand_1=0 WHERE item_no=?"


for index, row in df.iterrows():
    values = [row.item_no]
    try:
        cur.execute(sqlstring, values)
    except Exception as e:
        print(e)


cur.close()

conn.commit()
