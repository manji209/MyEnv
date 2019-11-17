import pyodbc
import pandas as pd
import xlrd
from openpyxl import load_workbook, Workbook




# Connect to SQL Server and set cursor
#conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=DINHPC,52052;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LALUCKYSERVER,65181;DATABASE=pbsdata00;UID=pbssqluser;PWD=Admin11')


cur = conn.cursor()

order_no = 107534

sqlstring = """SELECT ord_no, seq_no, item_no, desc_1, qty_ord, unit_prc, unit_cost FROM dbo.LINITM00 WHERE ord_no=?"""

df = pd.read_sql(sqlstring, conn, params={order_no})

df.columns = df.columns.str.strip()

df['MarkUp'] = ''
df['Total'] = ''

setprice_fname = "Data/Set_Price_06082019.xlsx"
setprice_df = pd.read_excel(setprice_fname, sheet_name='Sheet1')

found = False

for idx, row in df.iterrows():
    for idx2, row2 in setprice_df.iterrows():
        if row.item_no.strip() == row2.item_no:
            #Set sale price to fixxed price
            if row.unit_prc == 0:
                df.loc[idx, 'MarkUp'] = 0
                df.loc[idx, 'Total'] = row.unit_prc * row.qty_ord
                found = True
                break
            else:
                df.loc[idx, 'MarkUp'] = row2.price
                df.loc[idx, 'Total'] = row2.price * row.qty_ord
                found = True
                break

    if found:
        found = False
        continue
    else:
        df.loc[idx, 'MarkUp'] = row.unit_prc
        df.loc[idx, 'Total'] = row.unit_prc * row.qty_ord
        found = False



sqlupdate = "UPDATE dbo.LINITM00 SET unit_prc=? WHERE seq_no=? AND ord_no=?"

total_qty = 0
total_sales = 0


for idx3, row in df.iterrows():
    total_qty += row.qty_ord
    total_sales += row.Total
    values2 = [row.MarkUp, row.seq_no, order_no]
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


