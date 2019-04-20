import pyodbc
import xlsxwriter
import pandas as pd
import numpy as np


# Connect to SQL Server and set cursor
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=MANJI-RYZEN\SQLEXPRESS;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
cur = conn.cursor()


sqlstring = """SELECT tt.item_no, tt.desc_1, tt.desc_2
                FROM dbo.IHSLIN00 tt
                INNER JOIN
                    (SELECT item_no, MAX(invc_dat) AS MaxDateTime
                    FROM dbo.IHSLIN00
                    GROUP BY item_no) groupedtt
                ON tt.item_no = groupedtt.item_no
                WHERE (tt.cust_no=? AND DATEPART(YEAR, tt.invc_dat) >= 2018)
                AND tt.invc_dat = groupedtt.MaxDateTime"""


cust = 'K003'

df = pd.read_sql(sqlstring, conn, params={cust})

print(df)
