import pyodbc
import xlrd
import pandas as pd
import numpy as np
from datetime import datetime

# Query dbo.ICPHXF00 table for items that have not been physically counted AND Qty_on_Hand is not ZERO (Can be positive or negative Qty)

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()

sqlstring = """SELECT item_no, item_desc_3, item_desc_4  FROM dbo.ITMFIL00"""


df = pd.read_sql(sqlstring, conn)


print(df['item_desc_3'])

sqlstring = """INSERT INTO [dbo].[ICNOTF00] (
            [sys_id]
            ,[item_no]
            ,[rest_of_sys_flds]
            ,[note_dat]
            ,[note_time]
            ,[backup_sys_id]
            ,[backup_sys_flds]
            ,[backup_dat]
            ,[backup_time]
            ,[note_lin_1]
            ) VALUES (?,?,?,?,?,?,?,?,?,?)"""


def add_note(note, item):
    # current date and time
    now = datetime.now()
    sys_id = 'IC'
    item_no = item
    rest_of_sys_flds = ''
    note_dat = now.strftime("%Y/%m/%d")
    note_time = now.toordinal()
    backup_sys_id = 'IC'
    backup_sys_flds = item
    backup_dat = now.toordinal()
    backup_time = now.toordinal()
    note_lin_1 = note

    values = (sys_id, item_no, rest_of_sys_flds, note_dat, note_time, backup_sys_id, backup_sys_flds, backup_dat, backup_time, note_lin_1)

    try:
        cur.execute(sqlstring, values)
    except Exception as e:
        print(e)


for row in df.itertuples(index=True):
    if 'None' not in str(row.item_desc_3) or 'None' not in str(row.item_desc_4):
        note_str = str(row.item_desc_3) + '\n' + str(row.item_desc_4)
        add_note(note_str, row.item_no)
        conn.commit()


cur.close()

conn.commit()
