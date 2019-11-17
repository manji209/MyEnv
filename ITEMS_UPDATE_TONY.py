import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np

def connect_db():
    connect = pyodbc.connect(
        #'DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
        #'DRIVER={SQL Server Native Client 11.0};SERVER=MANJI-RYZEN\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdataDEMO;UID=pbssqluser;PWD=Admin11')
        #'DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


    return connect

# Connect to SQL Server and set cursor
conn = connect_db()

try:
    cur = conn.cursor()
    sqlstring = """SELECT item_no, item_desc_1, item_desc_2, item_prc_1
                FROM dbo.ITMFIL00"""


    df = pd.read_sql(sqlstring, conn)
    df['item_no'] = df['item_no'].str.strip()
    #df['item_prc_1'] = df['item_prc_1'].map('{:,.2f}'.format)
    #df.applymap(lambda x: x.strip() if type(x)==str else x)

finally:
    conn.close()

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Data/PO_Template/ITEMS_TONY.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', index=False)


writer.save()
