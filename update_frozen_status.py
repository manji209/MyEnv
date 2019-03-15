import pyodbc
import xlrd
import pandas as pd
import numpy as np


file_name = "Data/FROZEN_status_code.xlsx"
df = pd.read_excel(file_name, sheet_name='Sheet1')
# Remove all NaN and replace with blank space
df = df.fillna('')

conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=LALUCKYSERVER\SQLEXPRESS;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = "UPDATE dbo.IHSLIN00 SET desc_2=? WHERE item_no=? AND DATEPART(YEAR, invc_dat)=?"

yr = 2018

for index, row in df.iterrows():
    values = [row.desc_2, row.item_no, yr]
    try:
        cur.execute(sqlstring, values)
    except Exception as e:
        print(e)



cur.close()
conn.commit()
