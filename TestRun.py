import pyodbc
import requests
import pandas as pd
import datetime
import numpy as np
from openpyxl import load_workbook

# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')

cur = conn.cursor()

# Monthly Sales Figure
sqlstring = """SELECT item_no, item_desc_1, item_desc_2, item_prc_1
                FROM dbo.ITMFIL00"""


#Get updated ITEMS list
df_new = pd.read_sql(sqlstring, conn)
df_new['item_no'] = df_new['item_no'].str.strip()
df_new['item_desc_1'] = df_new['item_desc_1'].str.strip()
df_new['item_desc_2'] = df_new['item_desc_2'].str.strip()
#df_new.applymap(lambda x: x.strip() if type(x)==str else x)

writer = pd.ExcelWriter('Data/ITEMS_prev.xlsx', engine='xlsxwriter')

df_new.to_excel(writer, sheet_name='Sheet1', index=False)

writer.save()


