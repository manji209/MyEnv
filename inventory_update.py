import pyodbc
import xlrd
import pandas as pd
import numpy as np

# Update inventory from Excel sheet

file_name = "Data/Master_INV.xlsx"
df = pd.read_excel(file_name, sheet_name='Cleaned')
# Remove all NaN and replace with blank space
df = df.fillna('')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = "UPDATE dbo.ICPHXF00 SET cnt_qty_on_hand_1=? WHERE item_no=?"


for index, row in df.iterrows():
    values = [row.qty, row.item_no]
    try:
        cur.execute(sqlstring, values)
    except Exception as e:
        print(e)


cur.close()
conn.commit()
