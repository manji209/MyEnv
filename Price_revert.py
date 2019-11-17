import pyodbc
import pandas as pd
import xlrd
from openpyxl import load_workbook, Workbook




# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


cur = conn.cursor()

order_no = 107153

sqlupdate = "UPDATE dbo.LINITM00 SET unit_prc=? WHERE seq_no=? AND ord_no=?"

fname = 'Data/107153_fixx.xlsx'
df = pd.read_excel(fname, sheet_name='Sheet1')

total_qty = 0
total_sales = 0

for idx, row in df.iterrows():
    total_qty += row.qty_ord
    total_sales += row.qty_ord * row.unit_prc
    values2 = [row.unit_prc, row.seq_no, order_no]
    try:
        cur.execute(sqlupdate, values2)
    except Exception as e:
        print(e)


conn.commit()

query_order = """UPDATE [dbo].[ORDHDR00]
                SET [tot_qty] =?
                ,[tot_sls_amt] =?
                ,[tot_gross_amt] =?
                WHERE [ord_no] =?"""

values3 = [total_qty, total_sales, total_sales, order_no]
cur.execute(query_order, values3)
conn.commit()


