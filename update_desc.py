import pyodbc
import xlrd
import pandas as pd
import numpy as np


file_name = "Import/Desc_Update_102119.xlsx"
df = pd.read_excel(file_name, sheet_name='Sheet1')
# Remove all NaN and replace with blank space
df = df.fillna('')


# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = """UPDATE dbo.ITMFIL00 SET item_desc_1=?, item_desc_2=?
            WHERE item_no=?"""


for index, row in df.iterrows():
    values = [row.item_desc_1, row.item_desc_2, row.item_no]
    try:
        cur.execute(sqlstring, values)
    except Exception as e:
        print(e)



cur.close()
conn.commit()
